#!/usr/bin/env python3
"""AOW Document Converter — Telegram Bot"""

import os
import re
import html
import json
import shutil
import asyncio
import logging
import tempfile
from pathlib import Path
from datetime import datetime, timedelta

from telegram import (
    Update, Document, BotCommand, MenuButtonCommands,
    InlineKeyboardButton, InlineKeyboardMarkup,
)
from telegram.constants import ChatAction, ParseMode
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, CallbackQueryHandler,
    filters, ContextTypes
)

from processing import process_docx, process_xlsx, read_xlsx_spec
from cmr import build_cmr, invoice_short
from proforma import (
    parse_facture, build_proforma, missing_fields, out_filename,
    parse_money, fmt_money, remember_refs,
)
from passport import (
    read_passport_photo, buyer_draft, ocr_available, FIELD_NAMES,
)
from buyers import load_buyers, remember_buyer, find_buyer
from translit import transliterate, has_cyrillic

logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
# httpx logs every request at INFO, and the URL contains the bot token —
# that would write the token into the systemd journal a few times a minute.
logging.getLogger('httpx').setLevel(logging.WARNING)

# ── Config ────────────────────────────────────────────────────────────────────
BOT_TOKEN      = os.environ['BOT_TOKEN']
ALLOWED_USER   = int(os.environ.get('ALLOWED_USER_ID', '0'))

# ── GitHub input-folder polling config ───────────────────────────────────────
# If GITHUB_TOKEN is unset, polling is simply disabled — the chat-upload flow
# below keeps working exactly as before.
GITHUB_TOKEN    = os.environ.get('GITHUB_TOKEN', '')
GITHUB_REPO     = os.environ.get('GITHUB_REPO', 'VIad1slav/AOW-BOT-inputs')
GITHUB_FOLDER   = os.environ.get('GITHUB_FOLDER', 'inputs')
GITHUB_POLL_SEC = int(os.environ.get('GITHUB_POLL_SEC', '45'))

BASE_DIR   = Path(__file__).parent
SEEN_FILE  = BASE_DIR / 'github_seen.json'
WORK_DIR   = BASE_DIR / 'work'          # uploaded sources, per user
STATE_FILE = BASE_DIR / 'bot_state.json'  # last used invoice / PF, for suggestions

# ── Conversation states ───────────────────────────────────────────────────────
# Two flows share one ConversationHandler. Which one you are in is decided by
# the file you send: a DOCX/XLSX starts the invoice flow, a PDF starts the Pro
# Forma flow that produces the DOCX the invoice flow needs.
COLLECT, ASK_INV, ASK_DATE, ASK_SPEC, ASK_PF, CONFIRM = range(6)
PF_COLLECT, PF_PRICE, PF_NUM, PF_DATE, PF_BUYER, PF_REFS, PF_CONFIRM = range(6, 13)
PF_BUYER_ADDR = 13          # address + date of issue, the two the MRZ lacks

# ── Commands shown in the Telegram "/" menu ──────────────────────────────────
COMMANDS = [
    BotCommand('start',    'Начать / загрузить файлы'),
    BotCommand('proforma', 'Pro Forma из платёжки (PDF)'),
    BotCommand('help',     'Как пользоваться ботом'),
    BotCommand('cancel',   'Отменить и сбросить всё'),
]

# ── Auth check ────────────────────────────────────────────────────────────────
def allowed(update: Update) -> bool:
    user = update.effective_user
    if user is None:
        return False
    return ALLOWED_USER == 0 or user.id == ALLOWED_USER

# ── Tiny persisted state (last invoice / PF) ─────────────────────────────────
def _load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding='utf-8'))
    except Exception:
        return {}

def _save_state(d: dict):
    try:
        STATE_FILE.write_text(json.dumps(d, ensure_ascii=False), encoding='utf-8')
    except Exception:
        logging.exception('Не удалось сохранить bot_state.json')

def _next_number(s):
    """'26-59' → '26-60'. Returns None if there is no trailing number."""
    m = re.match(r'^(.*?)(\d+)$', (s or '').strip())
    if not m:
        return None
    head, num = m.groups()
    return f'{head}{int(num) + 1:0{len(num)}d}'

def _year2() -> str:
    """Current year as two digits — the '26' in invoice 26-69."""
    return datetime.now().strftime('%y')

def _suggest_inv():
    """Next invoice number to offer as a one-tap button.
    When the calendar year has rolled over since the last invoice, numbering
    starts again from 1 under the new year instead of continuing the old one."""
    yy = _year2()
    last = (_load_state().get('last_inv') or '').strip()
    m = re.match(r'^(\d{2})[-/]', last)
    if m and m.group(1) != yy:
        return f'{yy}-1'
    return _next_number(last)

# ── Input normalisation ──────────────────────────────────────────────────────
def _norm_ref(raw, prefix):
    """Strip an accidentally typed 'FV'/'PF' prefix so we never build 'FVFV26-59'."""
    s = (raw or '').strip()
    s = re.sub(rf'^{prefix}[\s:№]*', '', s, flags=re.IGNORECASE)
    return s.strip()

def _norm_invoice(raw):
    """Normalise an invoice number to the 'YY-NNN' form used by AOW.

    The leading two digits are the year, so typing just the sequence number is
    enough — the current year is filled in automatically ('69' → '26-69', and
    the same input gives '27-69' once 2027 starts). A number typed out in full
    is kept as-is; a four-digit year is shortened ('2026-69' → '26-69')."""
    s = _norm_ref(raw, 'FV')
    if not s:
        return ''
    m = re.fullmatch(r'(?:19|20)(\d{2})([-/].+)', s)
    if m:
        return m.group(1) + m.group(2)
    if re.fullmatch(r'\d{1,4}', s):
        return f'{_year2()}-{s}'
    return s

def _norm_date(raw):
    """Accept 04.06.2026 / 4-6-26 / 2026-06-04 / 04062026 → '04.06.2026'.
    Returns None when the input is not a valid date."""
    s = (raw or '').strip()
    m = re.fullmatch(r'(\d{4})[-./](\d{1,2})[-./](\d{1,2})', s)
    if m:
        y, mo, d = m.groups()
    else:
        s2 = s.replace(' ', '')
        m = re.fullmatch(r'(\d{1,2})[-./](\d{1,2})[-./](\d{2}|\d{4})', s2)
        if m:
            d, mo, y = m.groups()
        else:
            m = re.fullmatch(r'(\d{2})(\d{2})(\d{4})', s2)
            if not m:
                return None
            d, mo, y = m.groups()
    if len(y) == 2:
        y = '20' + y
    try:
        return datetime(int(y), int(mo), int(d)).strftime('%d.%m.%Y')
    except ValueError:
        return None

def _iso_date(raw):
    """Same tolerant parsing as _norm_date, but in the '2026-08-14' form the
    Pro Forma uses. Returns None when the input is not a valid date."""
    s = _norm_date(raw)
    if not s:
        return None
    return datetime.strptime(s, '%d.%m.%Y').strftime('%Y-%m-%d')


def _plus_days(iso, days):
    return (datetime.strptime(iso, '%Y-%m-%d') + timedelta(days=days)).strftime('%Y-%m-%d')


def _suggest_pf():
    """Next Pro Forma number to offer as a one-tap button.

    AOW numbers them 'YY-M/NNN' — year, month, and a sequence that runs on
    across months (26-6/78 in June, 26-8/106 in August). So the month is taken
    from today and only the sequence is incremented; a year change resets it."""
    last = (_load_state().get('last_pf') or '').strip()
    m = re.fullmatch(r'(\d{2})-(\d{1,2})/(\d+)', _norm_ref(last, 'PF'))
    if not m:
        return None
    yy_last, _, seq = m.groups()
    yy, mm = _year2(), str(datetime.now().month)
    if yy_last != yy:
        return f'{yy}-{mm}/1'
    return f'{yy}-{mm}/{int(seq) + 1}'


def _spec_from_name(name: str) -> str:
    m = re.search(r'(?:Spec[_\s])?(\d+)', Path(name).stem, re.IGNORECASE)
    return m.group(1) if m else ''

# ── Pro Forma number auto-detection ──────────────────────────────────────────
_PF_NUM_RE = re.compile(r'(?:PF)?\s*(\d{1,4}[-/][\d\-/]*\d)', re.IGNORECASE)

def extract_pf_ref(docx_path):
    """Best-effort: pull the Pro Forma number out of the source document's title
    so the user can confirm it with one tap instead of retyping it.
    Returns None whenever anything at all goes wrong — it is only a suggestion."""
    try:
        from docx import Document as _Doc
        doc = _Doc(str(docx_path))
        for p in doc.paragraphs[:40]:
            t = p.text or ''
            if 'pro forma' in t.lower():
                found = _PF_NUM_RE.findall(t)
                if found:
                    return found[-1]
    except Exception:
        logging.exception('Не удалось определить номер Pro Forma из DOCX')
    return None

# ── Per-user working directory for uploaded sources ──────────────────────────
def _work_dir(uid) -> Path:
    d = WORK_DIR / str(uid)
    d.mkdir(parents=True, exist_ok=True)
    return d

def _wipe_work(uid):
    shutil.rmtree(WORK_DIR / str(uid), ignore_errors=True)

# ── GitHub polling helpers ────────────────────────────────────────────────────
def _load_seen() -> set:
    try:
        return set(json.loads(SEEN_FILE.read_text()))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()

def _save_seen(shas: set):
    SEEN_FILE.write_text(json.dumps(sorted(shas)))

def _github_headers():
    h = {'Accept': 'application/vnd.github+json', 'User-Agent': 'aow-bot'}
    if GITHUB_TOKEN:
        h['Authorization'] = f'Bearer {GITHUB_TOKEN}'
    return h

def list_github_folder():
    """Return file entries currently sitting in the GitHub inputs/ folder."""
    import requests            # only reached when GITHUB_TOKEN is configured
    url = f'https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FOLDER}'
    r = requests.get(url, headers=_github_headers(), timeout=15)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    items = r.json()
    return [i for i in items if i.get('type') == 'file']

def download_github_file(item, dest_path: Path):
    import requests
    r = requests.get(item['download_url'], headers=_github_headers(), timeout=30)
    r.raise_for_status()
    dest_path.write_bytes(r.content)

# ── XLSX → PDF conversion (headless LibreOffice) ─────────────────────────────
PDF_CONVERT_TIMEOUT = int(os.environ.get('PDF_CONVERT_TIMEOUT', '60'))

async def convert_xlsx_to_pdf(xlsx_paths, outdir: Path) -> dict:
    """Convert XLSX files to PDF via headless LibreOffice.

    Returns {xlsx Path: pdf Path} holding only the conversions that produced a
    file; a missing entry means "no PDF this time", which callers treat as a
    warning rather than a fatal error.

    All the files go through a single soffice run on purpose. Startup dominates
    the cost — a few seconds and a ~200 MB peak out of the VM's 945 MB — so
    launching it once per spec instead of once per batch was the expensive part
    of a multi-spec job."""
    xlsx_paths = list(xlsx_paths)
    if not xlsx_paths:
        return {}
    timeout = PDF_CONVERT_TIMEOUT + 20 * (len(xlsx_paths) - 1)
    try:
        proc = await asyncio.create_subprocess_exec(
            'soffice', '--headless', '--convert-to', 'pdf',
            '--outdir', str(outdir), *(str(p) for p in xlsx_paths),
            stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE,
        )
        await asyncio.wait_for(proc.communicate(), timeout=timeout)
    except FileNotFoundError:
        logging.error("LibreOffice (soffice) не установлен — PDF не сгенерирован")
        return {}
    except asyncio.TimeoutError:
        logging.error(f"Конвертация в PDF превысила {timeout} сек")
        return {}
    except Exception:
        logging.exception("Ошибка конвертации в PDF")
        return {}

    made = {}
    for src in xlsx_paths:
        pdf_path = outdir / (src.stem + '.pdf')
        if pdf_path.exists():
            made[src] = pdf_path
    return made

# ── UI helpers ────────────────────────────────────────────────────────────────
# The bot keeps exactly one "card" — the message with the current question and
# its buttons. Every new step deletes the previous card and posts a fresh one at
# the bottom of the chat, so the active buttons are always the last thing you
# see and old buttons can't be tapped by mistake.

async def _drop_card(context: ContextTypes.DEFAULT_TYPE):
    card = context.user_data.pop('card', None)
    if card is not None:
        try:
            await card.delete()
        except Exception:
            pass

async def _card(chat_id, context: ContextTypes.DEFAULT_TYPE, text, rows=None):
    await _drop_card(context)
    msg = await context.bot.send_message(
        chat_id=chat_id, text=text,
        reply_markup=InlineKeyboardMarkup(rows) if rows else None,
        parse_mode=ParseMode.HTML, disable_web_page_preview=True,
    )
    context.user_data['card'] = msg
    return msg

def _btn_back(target):
    return InlineKeyboardButton('⬅️  Назад', callback_data=f'back|{target}')

def _files_summary(ud) -> str:
    docx = ud.get('docx')
    xlsx = ud.get('xlsx', [])
    lines = []
    if docx:
        lines.append(f"📄 Pro Forma: <code>{html.escape(docx['name'])}</code>")
    else:
        lines.append("📄 Pro Forma (DOCX): <i>не загружена</i>")
    if xlsx:
        lines.append(f"📊 Спецификации ({len(xlsx)}):")
        for x in xlsx:
            lines.append(
                f"     • <code>{html.escape(x['name'])}</code>"
                f" → №<b>{html.escape(x['spec'] or '?')}</b>"
            )
    else:
        lines.append("📊 Спецификации (XLSX): <i>не загружены</i>")
    return '\n'.join(lines)

# ── Step 0: collect files ────────────────────────────────────────────────────
async def _ask_files(chat_id, context, header='<b>📦 AOW Document Converter</b>'):
    ud = context.user_data
    ud['step'] = COLLECT
    has_any = bool(ud.get('docx') or ud.get('xlsx'))

    if has_any:
        tail = "Пришлите ещё файлы или нажмите <b>Продолжить</b>."
    else:
        tail = ("Пришлите файлы одним или несколькими сообщениями:\n"
                "📄 Pro Forma <code>(DOCX)</code>\n"
                "📊 Спецификация <code>(XLSX)</code> — можно несколько\n\n"
                "<i>Нужна сама Pro Forma? Пришлите платёжку "
                "<code>(PDF)</code> — бот соберёт её. Или /proforma</i>")
        if GITHUB_TOKEN:
            tail += "\n\n<i>Или положите файлы в папку inputs/ репозитория GitHub — бот подхватит их сам.</i>"

    rows = []
    if has_any:
        rows.append([InlineKeyboardButton('▶️  Продолжить', callback_data='go')])
        rows.append([InlineKeyboardButton('🗑  Очистить список', callback_data='clear')])

    await _card(chat_id, context, f"{header}\n\n{_files_summary(ud)}\n\n{tail}", rows)

# ── Step 1: invoice number ───────────────────────────────────────────────────
async def _ask_inv(chat_id, context):
    context.user_data['step'] = ASK_INV
    yy = _year2()
    rows = []
    suggestion = _suggest_inv()
    if suggestion and len(suggestion) <= 40:
        rows.append([InlineKeyboardButton(f'➡️  FV{suggestion}  (следующий)',
                                          callback_data=f'inv|{suggestion}')])
    rows.append([_btn_back('files')])
    await _card(
        chat_id, context,
        "<b>Шаг 1 из 4 — номер инвойса</b>\n\n"
        f"Введите только номер: <code>69</code>  →  <b>FV{yy}-69</b>\n"
        f"<i>Год ({yy}) подставляется автоматически. "
        f"Нужен другой — введите полностью, например <code>25-69</code>.</i>",
        rows
    )

# ── Step 2: date ─────────────────────────────────────────────────────────────
async def _ask_date(chat_id, context):
    context.user_data['step'] = ASK_DATE
    today = datetime.now().strftime('%d.%m.%Y')
    rows = [
        [InlineKeyboardButton(f'📅  Сегодня — {today}', callback_data=f'date|{today}')],
        [_btn_back('inv')],
    ]
    await _card(
        chat_id, context,
        "<b>Шаг 2 из 4 — дата документов</b>\n\n"
        "Нажмите кнопку или введите дату: <code>04.06.2026</code>",
        rows
    )

# ── Step 3: specification numbers ────────────────────────────────────────────
async def _ask_spec(chat_id, context):
    context.user_data['step'] = ASK_SPEC
    xlsx_list = context.user_data.get('xlsx', [])
    rows = []
    if xlsx_list:
        lines = '\n'.join(
            f"     • <code>{html.escape(x['name'])}</code> → №<b>{html.escape(x['spec'] or '?')}</b>"
            for x in xlsx_list
        )
        text = ("<b>Шаг 3 из 4 — номера спецификаций</b>\n\n"
                f"Определены автоматически:\n{lines}\n\n"
                "Подтвердите кнопкой или введите номера через запятую: <code>43, 44</code>")
        rows.append([InlineKeyboardButton('✅  Номера верны', callback_data='spec_ok')])
    else:
        text = ("<b>Шаг 3 из 4 — номер спецификации</b>\n\n"
                "Введите номер, например <code>44</code>")
    rows.append([_btn_back('date')])
    await _card(chat_id, context, text, rows)

# ── Step 4: Pro Forma reference ──────────────────────────────────────────────
async def _ask_pf(chat_id, context):
    context.user_data['step'] = ASK_PF
    rows = []
    guess = context.user_data.get('pf_guess')
    if guess and len(guess) <= 40:
        rows.append([InlineKeyboardButton(f'✅  Из Pro Forma — PF{guess}',
                                          callback_data=f'pf|{guess}')])
    rows.append([_btn_back('spec')])
    await _card(
        chat_id, context,
        "<b>Шаг 4 из 4 — ссылка на Pro Forma</b>\n\n"
        "Введите номер: <code>26-05/51</code>  →  <b>PF26-05/51</b>",
        rows
    )

# ── Confirmation ─────────────────────────────────────────────────────────────
async def _ask_confirm(chat_id, context):
    ud = context.user_data
    ud['step'] = CONFIRM
    specs = ', '.join('№' + x['spec'] for x in ud.get('xlsx', []) if x['spec']) \
            or ud.get('spec', '') or '—'
    text = (
        "<b>🧾 Проверьте данные</b>\n\n"
        f"Инвойс:  <b>FV{html.escape(ud.get('inv',''))}</b>\n"
        f"Дата:  <b>{html.escape(ud.get('date',''))}</b>\n"
        f"Pro Forma:  <b>PF{html.escape(ud.get('pf',''))}</b>\n"
        f"Спецификации:  <b>{html.escape(specs)}</b>\n\n"
        f"{_files_summary(ud)}"
    )
    rows = [
        [InlineKeyboardButton('🚀  Создать документы', callback_data='run')],
        [_btn_back('pf')],
        [InlineKeyboardButton('❌  Отменить', callback_data='cancel')],
    ]
    await _card(chat_id, context, text, rows)

# ── /start ────────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    _wipe_work(update.effective_user.id)
    await _drop_card(context)
    context.user_data.clear()
    await _ask_files(update.effective_chat.id, context,
                     header='<b>👋 AOW Document Converter</b>')
    return COLLECT

# ── /help ─────────────────────────────────────────────────────────────────────
async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.effective_message.reply_text(
        "<b>ℹ️ Как пользоваться</b>\n\n"
        "Бот умеет две вещи. Что именно делать — он понимает по формату файла.\n\n"
        "<b>📕 PDF → Pro Forma</b>\n"
        "1. Пришлите платёжку <code>(PDF)</code> — <b>Ayvens</b> (Facture VO), "
        "<b>SAG</b> или <b>OPENLANE</b>. Площадку бот определит сам; можно "
        "несколько, тогда все машины попадут в одну Pro Forma.\n"
        "2. Марку, VIN, госномер и дату регистрации бот возьмёт из PDF сам. "
        "Спросит только <b>цену, номер, дату и покупателя</b>.\n"
        "3. Вернёт <b>Faktura Pro forma (DOCX)</b> — и предложит сразу сделать из неё фактуру.\n\n"
        "<b>📄 DOCX + XLSX → фактура</b>\n"
        "1. Пришлите Pro Forma <code>(DOCX)</code> и спецификации <code>(XLSX)</code> — "
        "просто перетащите их в чат, можно несколько сразу и без всяких команд.\n"
        "2. Нажмите <b>Продолжить</b> и ответьте на 4 вопроса "
        "(номер инвойса, дата, номера спецификаций, ссылка на Pro Forma).\n"
        "3. Проверьте сводку и нажмите <b>Создать документы</b>.\n"
        "Бот вернёт: <b>Faktura (DOCX)</b>, <b>Spec (XLSX)</b> и <b>Spec (PDF)</b>.\n\n"
        "<b>Команды</b>\n"
        "/start — начать заново\n"
        "/proforma — Pro Forma из платёжки\n"
        "/cancel — сбросить всё\n"
        "/help — эта справка\n\n"
        "<i>Подсказки: на каждом шаге есть кнопка «Назад». "
        "Дату можно писать как 04.06.2026, 4-6-26 или 04062026. "
        "Цену — как 5400 или 5 400,00.</i>",
        parse_mode=ParseMode.HTML,
    )
    return None   # keeps the current conversation state untouched

# ── Receive files ─────────────────────────────────────────────────────────────
async def _store_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """Download an incoming DOCX/XLSX/PDF into the user's work dir and register
    it. Returns the kind that was stored ('docx', 'xlsx' or 'pdf') so the caller
    knows which card to show, or None after reporting the problem to the chat."""
    doc: Document = update.message.document
    if not doc:
        return None

    name = doc.file_name or 'file'
    lower = name.lower()
    uid = update.effective_user.id

    if not (lower.endswith('.docx') or lower.endswith('.xlsx')
            or lower.endswith('.pdf')):
        await update.message.reply_text(
            f"⚠️ <code>{html.escape(name)}</code> — неподдерживаемый формат.\n"
            "Нужен <b>.pdf</b> (платёжка), <b>.docx</b> (Pro Forma) "
            "или <b>.xlsx</b> (спецификация).",
            parse_mode=ParseMode.HTML,
        )
        return None

    # Download straight away: the files are small, and having them on disk lets
    # us read the Pro Forma number now and skip the download at processing time.
    safe = re.sub(r'[^A-Za-z0-9._-]', '_', name)
    kind = ('pf' if lower.endswith('.docx')
            else 'vo' if lower.endswith('.pdf') else 'sp')
    dest = _work_dir(uid) / f"{kind}_{len(context.user_data.get('xlsx', []))}_{safe}"
    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(str(dest))
    except Exception as e:
        logging.exception('Не удалось скачать файл из Telegram')
        await update.message.reply_text(
            f"❌ Не удалось загрузить <code>{html.escape(name)}</code>: {html.escape(str(e))}",
            parse_mode=ParseMode.HTML)
        return None

    if lower.endswith('.pdf'):
        return await _store_facture(update, context, dest, name)

    if lower.endswith('.docx'):
        old = context.user_data.get('docx')
        if old:
            Path(old['path']).unlink(missing_ok=True)
        context.user_data['docx'] = {'path': dest, 'name': name}
        context.user_data['pf_guess'] = await asyncio.to_thread(extract_pf_ref, dest)
        return 'docx'

    xlsx_list = context.user_data.setdefault('xlsx', [])
    entry = {'path': dest, 'name': name, 'spec': _spec_from_name(name)}
    for i, x in enumerate(xlsx_list):            # re-sending a file replaces it
        if x['name'] == name:
            Path(x['path']).unlink(missing_ok=True)
            xlsx_list[i] = entry
            break
    else:
        xlsx_list.append(entry)
    return 'xlsx'

async def _store_facture(update, context, dest: Path, name: str):
    """Read an incoming platform invoice and add its cars to the basket."""
    logs = []
    try:
        head = await asyncio.to_thread(parse_facture, dest, logs.append)
    except Exception as e:
        Path(dest).unlink(missing_ok=True)
        logging.exception('Не удалось разобрать платёжку')
        await update.message.reply_text(
            f"❌ <code>{html.escape(name)}</code> — не удалось разобрать: "
            f"{html.escape(str(e))}",
            parse_mode=ParseMode.HTML)
        return None

    pdfs = context.user_data.setdefault('pdf', [])
    for i, x in enumerate(pdfs):                 # re-sending a file replaces it
        if x['name'] == name:
            Path(x['path']).unlink(missing_ok=True)
            pdfs.pop(i)
            break
    pdfs.append({'path': dest, 'name': name, 'source': head.get('source', ''),
                 'facture_no': head['facture_no'], 'vehicles': head['vehicles']})
    context.bot_data['last_log'] = logs
    return 'pdf'

async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """A document arrived while we're collecting files (or while idle).
    Whichever kind just arrived decides which of the two flows we show."""
    if not allowed(update):
        return
    kind = await _store_document(update, context)
    chat_id = update.effective_chat.id
    if kind == 'pdf' or (kind is None and context.user_data.get('step') == PF_COLLECT):
        await _ask_pdf(chat_id, context)
        return PF_COLLECT
    await _ask_files(chat_id, context)
    return COLLECT

async def receive_file_midflow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """A document arrived while we're already asking questions — take it and
    stay on the current step instead of ignoring the file."""
    if not allowed(update):
        return
    if await _store_document(update, context):
        await update.message.reply_text('✅ Файл добавлен к комплекту.')
    return await _reshow(update.effective_chat.id, context)

async def _reshow(chat_id, context):
    """Re-post the card for whichever step we're on, and return that state."""
    step = context.user_data.get('step', COLLECT)
    ask = {
        ASK_INV: _ask_inv, ASK_DATE: _ask_date, ASK_SPEC: _ask_spec,
        ASK_PF: _ask_pf, CONFIRM: _ask_confirm,
        PF_COLLECT: _ask_pdf, PF_PRICE: _ask_pf_price, PF_NUM: _ask_pf_num,
        PF_DATE: _ask_pf_date, PF_BUYER: _ask_pf_buyer,
        PF_BUYER_ADDR: _ask_buyer_addr, PF_REFS: _ask_pf_refs,
        PF_CONFIRM: _ask_pf_confirm,
    }.get(step, _ask_files)
    await ask(chat_id, context)
    return step

async def hint_collect(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Anything that isn't a document while we're collecting files."""
    if not allowed(update):
        return
    await update.effective_message.reply_text(
        "📎 Жду файлы <b>DOCX</b> / <b>XLSX</b>. "
        "Когда все загружены — нажмите <b>Продолжить</b>.",
        parse_mode=ParseMode.HTML,
    )
    return COLLECT

async def hint_text_expected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Anything that isn't text while we're asking a question."""
    if not allowed(update):
        return
    await update.effective_message.reply_text(
        "✏️ Здесь нужен текстовый ответ или кнопка выше. "
        "Начать заново — /start"
    )
    return None

# ── Periodic job: poll the GitHub inputs/ folder for new files ──────────────
async def poll_github(context: ContextTypes.DEFAULT_TYPE):
    if not GITHUB_TOKEN or ALLOWED_USER == 0:
        return
    try:
        # requests.get() is blocking — running it directly here would freeze
        # the whole bot (no replies to /start or anything else) for however
        # long the HTTP call takes. asyncio.to_thread keeps it off the event loop.
        items = await asyncio.to_thread(list_github_folder)
    except Exception:
        logging.exception("Не удалось опросить GitHub")
        return

    seen = _load_seen()
    new_items = [i for i in items if i['sha'] not in seen]
    if not new_items:
        return

    # The job is registered with user_id=ALLOWED_USER, so context.user_data is
    # the very same dict the chat handlers use.
    ud = context.user_data
    work = _work_dir(ALLOWED_USER)
    found = False

    for item in new_items:
        name = item['name']
        lower = name.lower()
        if not (lower.endswith('.docx') or lower.endswith('.xlsx')):
            seen.add(item['sha'])   # not a document — ignore silently
            continue

        dest = work / f"{item['sha'][:10]}_{re.sub(r'[^A-Za-z0-9._-]', '_', name)}"
        try:
            await asyncio.to_thread(download_github_file, item, dest)
        except Exception:
            logging.exception(f"Не удалось скачать {name} из GitHub")
            continue   # don't mark as seen — retry on next poll

        if lower.endswith('.docx'):
            ud['docx'] = {'path': dest, 'name': name}
            ud['pf_guess'] = await asyncio.to_thread(extract_pf_ref, dest)
        else:
            ud.setdefault('xlsx', []).append(
                {'path': dest, 'name': name, 'spec': _spec_from_name(name)}
            )

        seen.add(item['sha'])
        found = True

    _save_seen(seen)
    if not found:
        return

    await _card(
        ALLOWED_USER, context,
        "<b>📥 Новые файлы из GitHub</b>\n\n" + _files_summary(ud) + "\n\n"
        "Нажмите <b>Продолжить</b> или пришлите ещё файлы.",
        [[InlineKeyboardButton('▶️  Продолжить', callback_data='go')],
         [InlineKeyboardButton('🗑  Очистить список', callback_data='clear')]],
    )

# ── Callbacks: collect stage ─────────────────────────────────────────────────
async def cb_go(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    await _ask_inv(update.effective_chat.id, context)
    return ASK_INV

async def cb_clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer('Список очищен')
    _wipe_work(update.effective_user.id)
    context.user_data.pop('docx', None)
    context.user_data.pop('xlsx', None)
    context.user_data.pop('pf_guess', None)
    await _ask_files(update.effective_chat.id, context)
    return COLLECT

# ── Step handlers ────────────────────────────────────────────────────────────
async def got_inv_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    context.user_data['inv'] = update.callback_query.data.split('|', 1)[1]
    await _ask_date(update.effective_chat.id, context)
    return ASK_DATE

async def got_inv(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    value = _norm_invoice(update.message.text)
    if not value:
        await update.message.reply_text(f'⚠️ Пустой номер. Введите, например, 69 → FV{_year2()}-69')
        return ASK_INV
    context.user_data['inv'] = value
    await _ask_date(update.effective_chat.id, context)
    return ASK_DATE

async def got_date_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    context.user_data['date'] = update.callback_query.data.split('|', 1)[1]
    await _ask_spec(update.effective_chat.id, context)
    return ASK_SPEC

async def got_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    norm = _norm_date(update.message.text)
    if not norm:
        await update.message.reply_text(
            '⚠️ Не понял дату. Введите в виде 04.06.2026 или нажмите кнопку.'
        )
        return ASK_DATE
    context.user_data['date'] = norm
    await _ask_spec(update.effective_chat.id, context)
    return ASK_SPEC

async def got_spec_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    await _ask_pf(update.effective_chat.id, context)
    return ASK_PF

async def got_spec_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    raw = update.message.text.strip()
    parts = [p.strip() for p in raw.replace(',', ' ').split() if p.strip()]
    if not parts:
        await update.message.reply_text('⚠️ Введите номер(а), например: 43, 44')
        return ASK_SPEC
    xlsx_list = context.user_data.get('xlsx', [])
    if xlsx_list:
        for i, spec in enumerate(parts):
            if i < len(xlsx_list):
                xlsx_list[i]['spec'] = spec
        context.user_data['spec'] = parts[0]
    else:
        context.user_data['spec'] = raw
    await _ask_pf(update.effective_chat.id, context)
    return ASK_PF

async def got_pf_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    context.user_data['pf'] = update.callback_query.data.split('|', 1)[1]
    await _ask_confirm(update.effective_chat.id, context)
    return CONFIRM

async def got_pf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    value = _norm_ref(update.message.text, 'PF')
    if not value:
        await update.message.reply_text('⚠️ Пустая ссылка. Введите, например, 26-05/51')
        return ASK_PF
    context.user_data['pf'] = value
    await _ask_confirm(update.effective_chat.id, context)
    return CONFIRM

# ══ Pro Forma flow: platform invoice (PDF) → Faktura Pro Forma (DOCX) ═══════
# The French invoice from Ayvens already names the car; the resale price, the
# Pro Forma number, its date and the buyer do not appear on it and are asked
# for here.

def _pf_cars(ud):
    """Every car from every uploaded invoice, in upload order."""
    return [v for entry in ud.get('pdf', []) for v in entry['vehicles']]

def _car_line(v) -> str:
    bits = [b for b in (v.get('model'), v.get('plate')) if b]
    line = ' · '.join(html.escape(b) for b in bits) or '<i>без опознавательных знаков</i>'
    if v.get('price'):
        line += f"  —  <b>{fmt_money(v['price'])} EUR</b>"
    gaps = missing_fields(v)
    if gaps:
        line += f"\n       ⚠️ не найдено: <i>{html.escape(', '.join(gaps))}</i>"
    return line

def _pf_summary(ud) -> str:
    pdfs = ud.get('pdf', [])
    if not pdfs:
        return "📕 Платёжка (PDF): <i>не загружена</i>"
    lines = []
    for entry in pdfs:
        head = f"📕 <code>{html.escape(entry['name'])}</code>"
        if entry.get('source'):
            head += f" — <b>{html.escape(entry['source'])}</b>"
        if entry.get('facture_no'):
            head += f" №{html.escape(entry['facture_no'])}"
        lines.append(head)
    for i, v in enumerate(_pf_cars(ud), 1):
        lines.append(f"     {i}. {_car_line(v)}")
    return '\n'.join(lines)

# ── Step 0: collect the platform invoices ────────────────────────────────────
async def _ask_pdf(chat_id, context, header='<b>📕 Pro Forma из платёжки</b>'):
    ud = context.user_data
    ud['step'] = PF_COLLECT
    cars = _pf_cars(ud)
    rows = []
    if cars:
        tail = "Пришлите ещё платёжки или нажмите <b>Продолжить</b>."
        rows.append([InlineKeyboardButton('▶️  Продолжить', callback_data='pf_go')])
        rows.append([InlineKeyboardButton('🗑  Очистить список', callback_data='pf_clear')])
    else:
        tail = ("Пришлите платёжку в формате <code>PDF</code> — "
                "<b>Ayvens</b>, <b>SAG</b> или <b>OPENLANE</b>.\n"
                "Площадку бот определит сам; можно несколько — "
                "все машины попадут в одну Pro Forma.")
    await _card(chat_id, context, f"{header}\n\n{_pf_summary(ud)}\n\n{tail}", rows)

# ── Step 1: price ────────────────────────────────────────────────────────────
async def _ask_pf_price(chat_id, context):
    context.user_data['step'] = PF_PRICE
    cars = _pf_cars(context.user_data)
    if len(cars) == 1:
        body = ("Введите <b>конечную стоимость</b> продажи в EUR:\n"
                "<code>5400</code>  или  <code>5 400,00</code>")
    else:
        listing = '\n'.join(f"     {i}. {_car_line(v)}" for i, v in enumerate(cars, 1))
        body = (f"Машин в Pro Forma: <b>{len(cars)}</b>\n{listing}\n\n"
                f"Введите <b>{len(cars)} цены</b> в EUR — каждую с новой строки, "
                f"в том же порядке.")
    await _card(chat_id, context,
                f"<b>Шаг 1 из 4 — стоимость</b>\n\n{body}",
                [[_btn_back('pdf')]])

async def got_pf_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    cars = _pf_cars(context.user_data)
    raw = update.message.text or ''
    parts = [p for p in re.split(r'[\n;]+', raw) if p.strip()]
    if len(cars) == 1 and len(parts) == 1:
        values = [parse_money(parts[0])]
    elif len(parts) != len(cars):
        await update.message.reply_text(
            f"✏️ Нужно {len(cars)} цен, по одной на строку — получено {len(parts)}.")
        return PF_PRICE
    else:
        values = [parse_money(p) for p in parts]

    if any(v is None for v in values):
        bad = ', '.join(p.strip() for p, v in zip(parts, values) if v is None)
        await update.message.reply_text(
            f"✏️ Не похоже на сумму: {bad}. Например: <code>5400</code> "
            f"или <code>5 400,00</code>.", parse_mode=ParseMode.HTML)
        return PF_PRICE

    for car, value in zip(cars, values):
        car['price'] = value
    await _ask_pf_num(update.effective_chat.id, context)
    return PF_NUM

# ── Step 2: Pro Forma number ─────────────────────────────────────────────────
async def _ask_pf_num(chat_id, context):
    context.user_data['step'] = PF_NUM
    rows = []
    guess = _suggest_pf()
    if guess and len(guess) <= 40:
        rows.append([InlineKeyboardButton(f'➡️  PF{guess}  (следующий)',
                                          callback_data=f'pfnum|{guess}')])
    rows.append([_btn_back('price')])
    await _card(
        chat_id, context,
        "<b>Шаг 2 из 4 — номер Pro Forma</b>\n\n"
        "Введите номер: <code>26-8/107</code>  →  <b>PF26-8/107</b>",
        rows
    )

async def _set_pf_num(chat_id, context, raw):
    value = _norm_ref(raw, 'PF')
    if not value:
        return None
    context.user_data['pf_num'] = f'PF{value}'
    context.user_data['pf_num_bare'] = value
    await _ask_pf_date(chat_id, context)
    return PF_DATE

async def got_pf_num_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    return await _set_pf_num(update.effective_chat.id, context,
                             update.callback_query.data.split('|', 1)[1])

async def got_pf_num(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    state = await _set_pf_num(update.effective_chat.id, context, update.message.text)
    if state is None:
        await update.message.reply_text('✏️ Введите номер, например <code>26-8/107</code>.',
                                        parse_mode=ParseMode.HTML)
        return PF_NUM
    return state

# ── Step 3: date ─────────────────────────────────────────────────────────────
async def _ask_pf_date(chat_id, context):
    context.user_data['step'] = PF_DATE
    today = datetime.now().strftime('%d.%m.%Y')
    rows = [
        [InlineKeyboardButton(f'📅  Сегодня — {today}', callback_data=f'pfdate|{today}')],
        [_btn_back('num')],
    ]
    await _card(
        chat_id, context,
        "<b>Шаг 3 из 4 — дата выставления</b>\n\n"
        "Нажмите кнопку или введите дату: <code>14.08.2026</code>\n"
        "<i>Срок оплаты бот поставит на 7 дней позже.</i>",
        rows
    )

async def _set_pf_date(chat_id, context, raw):
    iso = _iso_date(raw)
    if not iso:
        return None
    context.user_data['pf_date'] = iso
    context.user_data['pf_termin'] = _plus_days(iso, 7)
    await _ask_pf_buyer(chat_id, context)
    return PF_BUYER

async def got_pf_date_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    return await _set_pf_date(update.effective_chat.id, context,
                              update.callback_query.data.split('|', 1)[1])

async def got_pf_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    state = await _set_pf_date(update.effective_chat.id, context, update.message.text)
    if state is None:
        await update.message.reply_text(
            '📅 Не понял дату. Например: <code>14.08.2026</code>, '
            '<code>14-8-26</code> или <code>14082026</code>.',
            parse_mode=ParseMode.HTML)
        return PF_DATE
    return state

# ── Step 4: buyer ────────────────────────────────────────────────────────────
BUYER_BUTTONS = 4          # how many saved buyers fit on the card comfortably

async def _ask_pf_buyer(chat_id, context):
    context.user_data['step'] = PF_BUYER
    saved = load_buyers()[:BUYER_BUTTONS]
    rows = [[InlineKeyboardButton(f'👤  {b["name"]}', callback_data=f'buyer|{i}')]
            for i, b in enumerate(saved)]
    rows.append([_btn_back('date_pf')])   # 'date' alone belongs to the invoice flow

    if saved:
        body = ("Выберите покупателя кнопкой — данные подставятся целиком.\n\n"
                "Новый покупатель: пришлите <b>фото паспорта</b> или введите "
                "текстом (первая строка — имя, дальше адрес и документ).")
    else:
        body = ("Пришлите <b>фото паспорта</b> — бот прочитает имя и номер, "
                "останется дописать адрес.\n"
                "Или введите текстом: <b>первая строка — имя</b>, дальше "
                "адрес и документ.\n\n"
                "<i>Например:</i>\n"
                "<code>IVANOU SIARHEI\n"
                "Belarus, Mińsk,\n"
                "ul. Szyszkina 12-7\n"
                "paszport: MP1234567 od 25.03.2022 г.</code>")

    await _card(chat_id, context, f"<b>Шаг 4 из 4 — покупатель</b>\n\n{body}", rows)

def _set_buyer(context, name, info):
    context.user_data['buyer_name'] = name
    context.user_data['buyer_info'] = info

async def got_pf_buyer_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """One of the saved buyers was tapped."""
    if not allowed(update):
        return
    await update.callback_query.answer()
    saved = load_buyers()
    try:
        chosen = saved[int(update.callback_query.data.split('|', 1)[1])]
    except (ValueError, IndexError):
        await _ask_pf_buyer(update.effective_chat.id, context)
        return PF_BUYER
    _set_buyer(context, chosen['name'], chosen.get('info', ''))
    await _ask_pf_confirm(update.effective_chat.id, context)
    return PF_CONFIRM

async def got_pf_buyer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    lines = [l.strip() for l in (update.message.text or '').splitlines()]
    lines = [l for l in lines if l]
    if not lines:
        await update.message.reply_text('✏️ Нужно хотя бы имя покупателя.')
        return PF_BUYER
    _set_buyer(context, lines[0], '\n'.join(lines[1:]))
    await _ask_pf_confirm(update.effective_chat.id, context)
    return PF_CONFIRM

# ── Step 4a: read the buyer off a passport photo ─────────────────────────────
async def got_pf_passport(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """A photo arrived on the buyer step — read its MRZ and offer a draft.

    The MRZ gives the name and the document number, and its check digits say
    whether they were read correctly. It does NOT carry the date of issue or
    an address, so the draft is handed back for the user to finish rather than
    used as-is.
    """
    if not allowed(update):
        return
    if not ocr_available():
        await update.message.reply_text(
            '📷 Распознавание фото на этом сервере недоступно. '
            'Введите покупателя текстом.')
        return PF_BUYER

    msg = update.message
    tg_obj = msg.photo[-1] if msg.photo else msg.document
    note = await msg.reply_text('⏳ Читаю паспорт…')
    dest = _work_dir(update.effective_user.id) / 'passport.jpg'
    try:
        tg_file = await tg_obj.get_file()
        await tg_file.download_to_drive(str(dest))
        mrz = await asyncio.to_thread(read_passport_photo, dest)
    except Exception as e:
        logging.exception('Не удалось прочитать паспорт')
        try:
            await note.delete()
        except Exception:
            pass
        await msg.reply_text(
            f"❌ Не получилось: {html.escape(str(e))}\n"
            "Снимите разворот так, чтобы <b>две нижние строки</b> попали в кадр "
            "целиком и без бликов — или введите покупателя текстом.",
            parse_mode=ParseMode.HTML)
        return PF_BUYER
    finally:
        Path(dest).unlink(missing_ok=True)

    try:
        await note.delete()
    except Exception:
        pass

    draft = buyer_draft(mrz)

    # The number and the dates are guarded by check digits; the name is not.
    # Saying which is which is the whole point of this message — a wrong name
    # here would go straight onto an export document.
    if mrz['ok']:
        number_note = (f"✅ <b>{html.escape(draft['passport_line'])}</b> — "
                       "сходится по контрольным цифрам")
    else:
        bad = ', '.join(FIELD_NAMES.get(f, f) for f in mrz['failed'])
        number_note = (f"⚠️ <b>{html.escape(draft['passport_line'])}</b> — "
                       f"не сошлось: <i>{html.escape(bad)}</i>, сверьте с паспортом")

    if mrz.get('name_confirmed'):
        name_note = (f"❗️ <b>{html.escape(draft['name'])}</b> — два прохода "
                     "распознавания совпали, но <b>имя в паспорте ничем не "
                     "заверено</b>: сверьте по буквам")
    else:
        name_note = (f"❗️ <b>{html.escape(draft['name'])}</b> — <b>имя прочитано "
                     "ненадёжно</b>, проходы распознавания разошлись. "
                     "Обязательно впишите его вручную с паспорта")

    context.user_data['mrz_draft'] = draft
    await msg.reply_text(
        f"📷 <b>Из паспорта</b>\n\n{number_note}\n{name_note}",
        parse_mode=ParseMode.HTML)

    # Adress and date of issue are the two things a passport MRZ cannot give.
    # If this person is already in the book, both are there — one tap and done.
    known = await asyncio.to_thread(find_buyer, draft['name'])
    rows = []
    if known:
        rows.append([InlineKeyboardButton('👤  Взять адрес из справочника',
                                          callback_data='buyer_known')])
    rows.append([InlineKeyboardButton('✏️  Ввести адрес и дату выдачи',
                                      callback_data='buyer_addr')])
    rows.append([_btn_back('date_pf')])

    tail = ("Осталось добавить <b>адрес</b> и <b>дату выдачи</b> — "
            "их в машинной зоне паспорта нет.")
    if known:
        tail += (f"\n\n<i>В справочнике для этого покупателя уже есть:</i>\n"
                 f"<code>{html.escape(known.get('info', ''))}</code>")
    await _card(update.effective_chat.id, context,
                f"<b>Покупатель — {html.escape(draft['name'])}</b>\n\n{tail}", rows)
    return PF_BUYER

async def cb_buyer_known(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Use the address already on file for the person just photographed."""
    if not allowed(update):
        return
    await update.callback_query.answer()
    draft = context.user_data.get('mrz_draft') or {}
    known = await asyncio.to_thread(find_buyer, draft.get('name', ''))
    if not known:
        await _ask_pf_buyer(update.effective_chat.id, context)
        return PF_BUYER
    _set_buyer(context, known['name'], known.get('info', ''))
    await _ask_pf_confirm(update.effective_chat.id, context)
    return PF_CONFIRM

async def cb_buyer_addr(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    await _ask_buyer_addr(update.effective_chat.id, context)
    return PF_BUYER_ADDR

async def _ask_buyer_addr(chat_id, context):
    context.user_data['step'] = PF_BUYER_ADDR
    draft = context.user_data.get('mrz_draft') or {}
    await _card(
        chat_id, context,
        f"<b>Адрес и дата выдачи</b>\n\n"
        f"Имя и номер уже есть: <b>{html.escape(draft.get('name', ''))}</b>, "
        f"{html.escape(draft.get('passport_line', ''))}\n\n"
        "Пришлите одним сообщением адрес, а последней строкой — дату выдачи:\n\n"
        "<code>Mińsk,\nul. Szyszkina 12-7\n25.03.2022</code>",
        [[_btn_back('buyer')]]
    )

async def got_buyer_addr(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Assemble the four-line buyer block from the MRZ plus what was typed."""
    if not allowed(update):
        return
    draft = context.user_data.get('mrz_draft') or {}
    lines = [l.strip() for l in (update.message.text or '').splitlines()]
    lines = [l for l in lines if l]
    if not lines:
        await update.message.reply_text('✏️ Нужен хотя бы адрес.')
        return PF_BUYER_ADDR

    issued = _norm_date(lines[-1])
    address = lines[:-1] if issued else lines
    if not address:
        await update.message.reply_text(
            '✏️ Кроме даты нужен ещё адрес — город и улица.')
        return PF_BUYER_ADDR

    # The invoice spells addresses the Polish way, so "Минск, ул. Шишкина" has
    # to become "Mińsk, ul. Szyszkina". Only the address is touched — never the
    # name, which must match the passport letter for letter.
    if any(has_cyrillic(line) for line in address):
        latin = [transliterate(line) for line in address]
        await update.message.reply_text(
            "🔤 Адрес переведён в латиницу:\n"
            f"<code>{html.escape(chr(10).join(latin))}</code>\n"
            "<i>Проверьте на карточке ниже — если не так, вернитесь кнопкой «Назад».</i>",
            parse_mode=ParseMode.HTML)
        address = latin

    country = draft.get('country', '')
    first = f"{country}, {address[0]}" if country else address[0]
    passport = draft.get('passport_line', '')
    if issued:
        passport += f' od {issued} г.'
    info = '\n'.join([first] + address[1:] + [passport])
    _set_buyer(context, draft.get('name', ''), info)
    await _ask_pf_confirm(update.effective_chat.id, context)
    return PF_CONFIRM

# ── Optional: customs code and ground clearance ──────────────────────────────
async def _ask_pf_refs(chat_id, context):
    context.user_data['step'] = PF_REFS
    cars = _pf_cars(context.user_data)
    current = cars[0] if cars else {'cn': '', 'clearance': ''}
    await _card(
        chat_id, context,
        "<b>🔧 Код CN и просвет</b>\n\n"
        "Их нет в платёжке целиком, поэтому бот берёт код по объёму двигателя, "
        "а просвет — по умолчанию. "
        "Пришлите новые одной строкой — <b>код и просвет в мм</b>:\n"
        f"<code>{html.escape(current['cn'])} {html.escape(current['clearance'])}</code>\n\n"
        "<i>Применится ко всем машинам и запомнится для этих моделей.</i>",
        [[_btn_back('confirm')]]
    )

async def got_pf_refs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    m = re.fullmatch(r'\s*(\d{6,10})\s*[,;/ ]\s*(\d{2,4})\s*(?:мм|mm)?\s*',
                     update.message.text or '')
    if not m:
        await update.message.reply_text(
            '✏️ Нужны два числа: код CN и просвет в мм, например '
            '<code>87032390 155</code>.', parse_mode=ParseMode.HTML)
        return PF_REFS
    cn, clearance = m.groups()
    for car in _pf_cars(context.user_data):
        car['cn'], car['clearance'] = cn, clearance
        await asyncio.to_thread(remember_refs, car['model'], cn, clearance)
    await _ask_pf_confirm(update.effective_chat.id, context)
    return PF_CONFIRM

# ── Confirmation ─────────────────────────────────────────────────────────────
async def _ask_pf_confirm(chat_id, context):
    ud = context.user_data
    ud['step'] = PF_CONFIRM
    cars = _pf_cars(ud)
    total = sum(c['price'] or 0 for c in cars)
    buyer = ud.get('buyer_name', '')
    if ud.get('buyer_info'):
        buyer += '\n       ' + ud['buyer_info'].replace('\n', '\n       ')
    listing = '\n'.join(f"     {i}. {_car_line(v)}" for i, v in enumerate(cars, 1))
    # Show where the customs code came from: derived from the engine capacity
    # when the platform states one, otherwise the default. It goes on the
    # customs declaration, so it should never be silently assumed.
    refs = ', '.join(sorted({
        "CN {} / {} мм{}".format(
            c['cn'], c['clearance'],
            f" (по объёму {c['capacity']} см³)" if c.get('capacity') else ' (по умолчанию)')
        for c in cars}))
    text = (
        "<b>🧾 Проверьте Pro Forma</b>\n\n"
        f"Номер:  <b>{html.escape(ud.get('pf_num', ''))}</b>\n"
        f"Дата:  <b>{html.escape(ud.get('pf_date', ''))}</b>\n"
        f"Срок оплаты:  <b>{html.escape(ud.get('pf_termin', ''))}</b>\n"
        f"Покупатель:  <b>{html.escape(buyer)}</b>\n\n"
        f"{listing}\n\n"
        f"Итого:  <b>{fmt_money(total)} EUR</b>\n"
        f"<i>{html.escape(refs)}</i>"
    )
    rows = [
        [InlineKeyboardButton('🚀  Создать Pro Forma', callback_data='pf_run')],
        [InlineKeyboardButton('🔧  Изменить CN / просвет', callback_data='pf_refs')],
        [_btn_back('buyer')],
        [InlineKeyboardButton('❌  Отменить', callback_data='cancel')],
    ]
    await _card(chat_id, context, text, rows)

async def cb_pf_refs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    await _ask_pf_refs(update.effective_chat.id, context)
    return PF_REFS

async def cb_pf_go(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    if not _pf_cars(context.user_data):
        await _ask_pdf(update.effective_chat.id, context)
        return PF_COLLECT
    await _ask_pf_price(update.effective_chat.id, context)
    return PF_PRICE

async def cb_pf_clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer('Список очищен')
    for entry in context.user_data.get('pdf', []):
        Path(entry['path']).unlink(missing_ok=True)
    context.user_data['pdf'] = []
    await _ask_pdf(update.effective_chat.id, context)
    return PF_COLLECT

async def cmd_proforma(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await _drop_card(context)
    await _ask_pdf(update.effective_chat.id, context)
    return PF_COLLECT

async def hint_pf_collect(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.effective_message.reply_text(
        "📎 Жду платёжку в формате <b>PDF</b> (Ayvens / SAG / OPENLANE). "
        "Когда все загружены — нажмите <b>Продолжить</b>.",
        parse_mode=ParseMode.HTML,
    )
    return PF_COLLECT

# ── Run: write the Pro Forma and hand it back ────────────────────────────────
async def cb_pf_run(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    # A Facture dropped in after the price step leaves its cars without a
    # price — ask again rather than failing halfway through building the file.
    if any(c.get('price') is None for c in _pf_cars(context.user_data)):
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text='➕ Появились машины без цены — укажите стоимость ещё раз.')
        await _ask_pf_price(update.effective_chat.id, context)
        return PF_PRICE
    await _drop_card(context)
    await _make_proforma(update.effective_chat.id, update.effective_user.id, context)
    return ConversationHandler.END

async def _make_proforma(chat_id, uid, context):
    ud = context.user_data
    logs = []
    prog = await context.bot.send_message(chat_id=chat_id, text='⏳ Формирую Pro Forma…')
    try:
        cars = _pf_cars(ud)
        dest = _work_dir(uid) / out_filename(ud['pf_num'])
        result = await asyncio.to_thread(build_proforma, dest, {
            'pf_num': ud['pf_num'],
            'date': ud['pf_date'],
            'termin': ud['pf_termin'],
            'buyer_name': ud.get('buyer_name', ''),
            'buyer_info': ud.get('buyer_info', ''),
            'vehicles': cars,
        }, None, logs.append)

        await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.UPLOAD_DOCUMENT)
        with open(result['path'], 'rb') as fh:
            await context.bot.send_document(chat_id=chat_id, document=fh,
                                            filename=result['path'].name)
        try:
            await prog.delete()
        except Exception:
            pass

        st = _load_state()
        st['last_pf'] = ud.get('pf_num_bare', '')
        _save_state(st)
        await asyncio.to_thread(remember_buyer, ud.get('buyer_name', ''),
                                ud.get('buyer_info', ''))

        # Keep the file so it can go straight into the invoice flow.
        for entry in ud.get('pdf', []):
            Path(entry['path']).unlink(missing_ok=True)
        ud['pdf'] = []
        ud['docx'] = {'path': result['path'], 'name': result['path'].name}
        ud['pf_guess'] = ud.get('pf_num_bare', '')
        context.bot_data['last_log'] = logs

        await context.bot.send_message(
            chat_id=chat_id,
            text=(f"✅ <b>{html.escape(ud['pf_num'])}</b> готова — "
                  f"позиций {result['count']}, итого "
                  f"<b>{fmt_money(result['total'])} EUR</b>"),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton('📄  Сделать из неё фактуру',
                                      callback_data='pf_invoice')],
                [InlineKeyboardButton('🔄  Новая Pro Forma', callback_data='pf_new')],
                [InlineKeyboardButton('🧾  Показать лог обработки', callback_data='log')],
            ]),
            parse_mode=ParseMode.HTML,
        )
    except Exception as e:
        logging.exception('Не удалось собрать Pro Forma')
        context.bot_data['last_log'] = logs
        try:
            await prog.delete()
        except Exception:
            pass
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"❌ Ошибка: <code>{html.escape(str(e))}</code>",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton('🧾  Показать лог обработки', callback_data='log')],
                [InlineKeyboardButton('🔄  Начать заново', callback_data='pf_new')],
            ]),
            parse_mode=ParseMode.HTML,
        )

async def cb_pf_invoice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Feed the Pro Forma that was just made into the invoice flow."""
    if not allowed(update):
        return
    await update.callback_query.answer()
    if not context.user_data.get('docx'):
        await _ask_files(update.effective_chat.id, context)
        return COLLECT
    await _ask_inv(update.effective_chat.id, context)
    return ASK_INV

async def cb_pf_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    _wipe_work(update.effective_user.id)
    context.user_data.clear()
    await _ask_pdf(update.effective_chat.id, context,
                   header='<b>📕 Новая Pro Forma</b>')
    return PF_COLLECT

# ── Back navigation ──────────────────────────────────────────────────────────
async def cb_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    target = update.callback_query.data.split('|', 1)[1]
    chat_id = update.effective_chat.id
    pf_steps = {
        'pdf': (_ask_pdf, PF_COLLECT), 'price': (_ask_pf_price, PF_PRICE),
        'num': (_ask_pf_num, PF_NUM), 'date_pf': (_ask_pf_date, PF_DATE),
        'buyer': (_ask_pf_buyer, PF_BUYER), 'confirm': (_ask_pf_confirm, PF_CONFIRM),
        'buyer_addr': (_ask_buyer_addr, PF_BUYER_ADDR),
    }
    if target in pf_steps:
        ask, state = pf_steps[target]
        await ask(chat_id, context)
        return state
    if target == 'files':
        await _ask_files(chat_id, context)
        return COLLECT
    if target == 'inv':
        await _ask_inv(chat_id, context)
        return ASK_INV
    if target == 'date':
        await _ask_date(chat_id, context)
        return ASK_DATE
    if target == 'spec':
        await _ask_spec(chat_id, context)
        return ASK_SPEC
    await _ask_pf(chat_id, context)
    return ASK_PF

# ── Run ──────────────────────────────────────────────────────────────────────
async def cb_run(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    await _drop_card(context)
    await _process(update.effective_chat.id, update.effective_user.id, context)
    return ConversationHandler.END

# ── Core processing ───────────────────────────────────────────────────────────
async def _process(chat_id, uid, context):
    ud = context.user_data
    inv_num  = f"FV{ud['inv']}"
    date_str = ud['date']
    pf_ref   = f"PF{ud['pf']}"
    xlsx_list = ud.get('xlsx', [])

    logs = []
    def log(m): logs.append(m)

    prog = await context.bot.send_message(chat_id=chat_id, text='⏳ Читаю спецификации…')
    async def step(text):
        try:
            await prog.edit_text(text)
        except Exception:
            pass

    sent = 0
    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)

            # Combine articles from all specs.
            # openpyxl parsing is blocking and can take seconds on big specs —
            # to_thread keeps the bot answering while it runs.
            spec_arts = set()
            items_all = []
            for entry in xlsx_list:
                arts, items = await asyncio.to_thread(
                    read_xlsx_spec, str(entry['path']), log)
                spec_arts |= arts
                items_all += items

            params_base = dict(
                invoice_num=inv_num, date=date_str,
                spec_num=', '.join(x['spec'] for x in xlsx_list if x['spec'])
                         or ud.get('spec', ''),
                pf_ref=pf_ref, _xlsx_items=items_all
            )

            output_files = []

            # Process DOCX
            docx_entry = ud.get('docx')
            if docx_entry:
                await step('⏳ Формирую фактуру (DOCX)…')
                out_docx = tmp / f'Faktura {inv_num}.docx'
                await asyncio.to_thread(
                    process_docx, str(docx_entry['path']), str(out_docx),
                    params_base, spec_arts or None, log
                )
                output_files.append(out_docx)

            # Process each XLSX
            spec_xlsx = []
            for entry in xlsx_list:
                spec_num = entry['spec']
                await step(f"⏳ Формирую спецификацию №{spec_num or '?'}…")
                params_x = dict(invoice_num=inv_num, date=date_str,
                                spec_num=spec_num, pf_ref=pf_ref)
                name = f'Spec {spec_num}.xlsx' if spec_num else f'Spec {inv_num}.xlsx'
                out_xlsx = tmp / name
                await asyncio.to_thread(process_xlsx, str(entry['path']), str(out_xlsx),
                                        params_x, log)
                spec_xlsx.append(out_xlsx)

            # CMR consignment note. Its goods block repeats what the
            # specifications' footer already says, so it is built from the
            # finished files rather than asked for again. Built before the
            # conversion so it can ride along in the same LibreOffice run.
            out_cmr = None
            if spec_xlsx:
                await step('⏳ Формирую CMR…')
                try:
                    out_cmr = tmp / f'CMR SK {invoice_short(inv_num)}.docx'
                    await asyncio.to_thread(
                        build_cmr, [str(p) for p in spec_xlsx],
                        params_base, str(out_cmr), log)
                except Exception:
                    logging.exception('Не удалось собрать CMR')
                    log('  ⚠ CMR не создан — остальные документы отправлены')
                    out_cmr = None

            # One LibreOffice run for the whole batch — see convert_xlsx_to_pdf
            if spec_xlsx:
                await step('⏳ Конвертирую в PDF…')
                pdfs = await convert_xlsx_to_pdf(
                    list(spec_xlsx) + ([out_cmr] if out_cmr else []), tmp)
                for out_xlsx in spec_xlsx:
                    output_files.append(out_xlsx)
                    pdf_path = pdfs.get(out_xlsx)
                    if pdf_path:
                        output_files.append(pdf_path)
                    else:
                        log(f'  ⚠ PDF для {out_xlsx.name} не создан '
                            '(спецификация всё равно отправлена как XLSX)')
                if out_cmr:
                    cmr_pdf = pdfs.get(out_cmr)
                    if cmr_pdf:
                        output_files.append(cmr_pdf)
                    else:
                        # Better a CMR in the wrong format than no CMR at all.
                        log('  ⚠ CMR в PDF не переведён — отправлен как DOCX')
                        output_files.append(out_cmr)

            # Send result files
            await step('📤 Отправляю готовые документы…')
            for f in output_files:
                if f.exists():
                    await context.bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
                    with open(f, 'rb') as fh:
                        await context.bot.send_document(chat_id=chat_id, document=fh,
                                                        filename=f.name)
                    sent += 1

        try:
            await prog.delete()
        except Exception:
            pass

        st = _load_state()
        st['last_inv'] = ud.get('inv', '')
        st['last_pf'] = ud.get('pf', '')
        _save_state(st)

        context.bot_data['last_log'] = logs
        await context.bot.send_message(
            chat_id=chat_id,
            text=(f"✅ <b>Готово</b> — отправлено файлов: <b>{sent}</b>\n"
                  f"Инвойс <b>{html.escape(inv_num)}</b> от {html.escape(date_str)}"),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton('🔄  Новый комплект', callback_data='new')],
                [InlineKeyboardButton('🧾  Показать лог обработки', callback_data='log')],
            ]),
            parse_mode=ParseMode.HTML,
        )

    except Exception as e:
        logging.exception("Processing error")
        context.bot_data['last_log'] = logs
        try:
            await prog.delete()
        except Exception:
            pass
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"❌ Ошибка: <code>{html.escape(str(e))}</code>",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton('🧾  Показать лог обработки', callback_data='log')],
                [InlineKeyboardButton('🔄  Начать заново', callback_data='new')],
            ]),
            parse_mode=ParseMode.HTML,
        )

    _wipe_work(uid)
    context.user_data.clear()

# ── Post-run callbacks ───────────────────────────────────────────────────────
async def cb_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    _wipe_work(update.effective_user.id)
    context.user_data.clear()
    await _ask_files(update.effective_chat.id, context,
                     header='<b>📦 Новый комплект документов</b>')
    return COLLECT

async def cb_log(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer()
    logs = context.bot_data.get('last_log') or []
    text = '\n'.join(logs) or 'Лог пуст.'
    if len(text) > 3500:
        text = text[:3500] + '\n… (лог обрезан)'
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"<b>🧾 Лог последней обработки</b>\n\n<pre>{html.escape(text)}</pre>",
        parse_mode=ParseMode.HTML,
    )
    return None

# ── Cancel / unknown ─────────────────────────────────────────────────────────
async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await _drop_card(context)
    _wipe_work(update.effective_user.id)
    context.user_data.clear()
    await update.effective_message.reply_text('❌ Отменено. Нажмите /start, чтобы начать заново.')
    return ConversationHandler.END

async def cb_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await update.callback_query.answer('Отменено')
    await _drop_card(context)
    _wipe_work(update.effective_user.id)
    context.user_data.clear()
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text='❌ Отменено. Нажмите /start, чтобы начать заново.'
    )
    return ConversationHandler.END

async def cmd_idle_hint(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fires only when no conversation is running (it is an entry point that
    deliberately returns None, so it never starts one)."""
    if not allowed(update):
        return
    await update.effective_message.reply_text(
        'Нажмите /start или просто пришлите файлы DOCX / XLSX. Справка — /help'
    )
    return None

# ── Startup: register the "/" command menu ───────────────────────────────────
async def post_init(app: Application):
    try:
        await app.bot.set_my_commands(COMMANDS)
        await app.bot.set_chat_menu_button(menu_button=MenuButtonCommands())
        logging.info('Меню команд зарегистрировано: %s',
                     ', '.join('/' + c.command for c in COMMANDS))
    except Exception:
        logging.exception('Не удалось зарегистрировать меню команд')

# ── Main ──────────────────────────────────────────────────────────────────────
def build_app():
    """Build the Application with every handler wired up.
    Kept separate from main() so the conversation routing can be exercised
    offline by tests without starting the polling loop."""
    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .connect_timeout(30)
        .read_timeout(60)
        .write_timeout(60)     # big XLSX/PDF uploads need well over the 5s default
        .pool_timeout(30)
        .build()
    )

    # Every question step accepts a late file drop, then re-asks the question;
    # anything else that isn't text gets a nudge.
    midflow_file = MessageHandler(filters.Document.ALL, receive_file_midflow)
    hint_nontext = MessageHandler(
        filters.ALL & ~filters.COMMAND & ~filters.TEXT & ~filters.Document.ALL,
        hint_text_expected)

    conv = ConversationHandler(
        entry_points=[
            CommandHandler('start', cmd_start),
            CommandHandler('proforma', cmd_proforma),
            CommandHandler('help', cmd_help),
            CommandHandler('cancel', cmd_cancel),
            # Dropping files into the chat starts the flow — no command needed.
            # A PDF routes to the Pro Forma flow, a DOCX/XLSX to the invoice one.
            MessageHandler(filters.Document.ALL, receive_file),
            CallbackQueryHandler(cb_go, pattern='^go$'),
            CallbackQueryHandler(cb_clear, pattern='^clear$'),
            CallbackQueryHandler(cb_new, pattern='^new$'),
            CallbackQueryHandler(cb_log, pattern='^log$'),
            CallbackQueryHandler(cb_pf_go, pattern='^pf_go$'),
            CallbackQueryHandler(cb_pf_clear, pattern='^pf_clear$'),
            CallbackQueryHandler(cb_pf_new, pattern='^pf_new$'),
            CallbackQueryHandler(cb_pf_invoice, pattern='^pf_invoice$'),
            # Anything else while idle → a nudge instead of silence. This must
            # stay last, and `allow_reentry` must stay off: with re-entry on,
            # PTB checks entry points BEFORE the state handlers, so this
            # catch-all would swallow every answer inside the conversation.
            MessageHandler(filters.ALL & ~filters.COMMAND, cmd_idle_hint),
        ],
        states={
            COLLECT: [
                MessageHandler(filters.Document.ALL, receive_file),
                CallbackQueryHandler(cb_go, pattern='^go$'),
                CallbackQueryHandler(cb_clear, pattern='^clear$'),
                MessageHandler(filters.ALL & ~filters.COMMAND, hint_collect),
            ],
            ASK_INV: [
                CallbackQueryHandler(got_inv_cb, pattern=r'^inv\|'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_inv),
                midflow_file, hint_nontext,
            ],
            ASK_DATE: [
                CallbackQueryHandler(got_date_cb, pattern=r'^date\|'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_date_text),
                midflow_file, hint_nontext,
            ],
            ASK_SPEC: [
                CallbackQueryHandler(got_spec_cb, pattern='^spec_ok$'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_spec_text),
                midflow_file, hint_nontext,
            ],
            ASK_PF: [
                CallbackQueryHandler(got_pf_cb, pattern=r'^pf\|'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf),
                midflow_file, hint_nontext,
            ],
            CONFIRM: [
                CallbackQueryHandler(cb_run, pattern='^run$'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                CallbackQueryHandler(cb_cancel, pattern='^cancel$'),
                midflow_file,
                MessageHandler(filters.ALL & ~filters.COMMAND, hint_text_expected),
            ],

            # ── Pro Forma flow ──
            PF_COLLECT: [
                MessageHandler(filters.Document.ALL, receive_file),
                CallbackQueryHandler(cb_pf_go, pattern='^pf_go$'),
                CallbackQueryHandler(cb_pf_clear, pattern='^pf_clear$'),
                MessageHandler(filters.ALL & ~filters.COMMAND, hint_pf_collect),
            ],
            PF_PRICE: [
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf_price),
                midflow_file, hint_nontext,
            ],
            PF_NUM: [
                CallbackQueryHandler(got_pf_num_cb, pattern=r'^pfnum\|'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf_num),
                midflow_file, hint_nontext,
            ],
            PF_DATE: [
                CallbackQueryHandler(got_pf_date_cb, pattern=r'^pfdate\|'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf_date_text),
                midflow_file, hint_nontext,
            ],
            PF_BUYER: [
                CallbackQueryHandler(got_pf_buyer_cb, pattern=r'^buyer\|'),
                CallbackQueryHandler(cb_buyer_known, pattern='^buyer_known$'),
                CallbackQueryHandler(cb_buyer_addr, pattern='^buyer_addr$'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                # A passport picture must be caught before midflow_file, which
                # would otherwise swallow a scan sent as a document.
                MessageHandler(filters.PHOTO | filters.Document.IMAGE,
                               got_pf_passport),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf_buyer),
                midflow_file, hint_nontext,
            ],
            PF_BUYER_ADDR: [
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_buyer_addr),
                midflow_file, hint_nontext,
            ],
            PF_REFS: [
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf_refs),
                midflow_file, hint_nontext,
            ],
            PF_CONFIRM: [
                CallbackQueryHandler(cb_pf_run, pattern='^pf_run$'),
                CallbackQueryHandler(cb_pf_refs, pattern='^pf_refs$'),
                CallbackQueryHandler(cb_back, pattern=r'^back\|'),
                CallbackQueryHandler(cb_cancel, pattern='^cancel$'),
                midflow_file,
                MessageHandler(filters.ALL & ~filters.COMMAND, hint_text_expected),
            ],
        },
        # /start, /cancel and /help must be reachable from inside a running
        # conversation — that is what fallbacks are for. Re-entry stays off on
        # purpose; see the note on the catch-all entry point above.
        fallbacks=[
            CommandHandler('cancel', cmd_cancel),
            CommandHandler('start', cmd_start),
            CommandHandler('proforma', cmd_proforma),
            CommandHandler('help', cmd_help),
        ],
        allow_reentry=False,
    )

    app.add_handler(conv)

    if GITHUB_TOKEN and ALLOWED_USER:
        if app.job_queue is None:
            logging.error(
                "GITHUB_TOKEN задан, но JobQueue недоступен. "
                'Установите: pip install "python-telegram-bot[job-queue]" --break-system-packages'
            )
        else:
            app.job_queue.run_repeating(
                poll_github, interval=GITHUB_POLL_SEC, first=15,
                chat_id=ALLOWED_USER, user_id=ALLOWED_USER,
            )
            logging.info(
                f"GitHub polling enabled: {GITHUB_REPO}/{GITHUB_FOLDER} every {GITHUB_POLL_SEC}s"
            )
    else:
        logging.info("GITHUB_TOKEN не задан — опрос GitHub отключён")

    return app

def main():
    logging.info("Bot started")
    build_app().run_polling(drop_pending_updates=True)

if __name__ == '__main__':
    main()
