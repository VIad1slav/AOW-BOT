#!/usr/bin/env python3
"""AOW Document Converter — Telegram Bot"""

import os
import re
import logging
import tempfile
from pathlib import Path
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, Document
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, CallbackQueryHandler,
    filters, ContextTypes
)

from processing import (
    process_docx, process_xlsx,
    get_articles_from_xlsx, get_xlsx_items_data
)

logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# ── Config ────────────────────────────────────────────────────────────────────
BOT_TOKEN      = os.environ['BOT_TOKEN']
ALLOWED_USER   = int(os.environ.get('ALLOWED_USER_ID', '0'))

# ── Conversation states ───────────────────────────────────────────────────────
COLLECT, ASK_INV, ASK_DATE, ASK_SPEC, ASK_PF = range(5)

# ── Auth check ────────────────────────────────────────────────────────────────
def allowed(update: Update) -> bool:
    uid = update.effective_user.id
    return ALLOWED_USER == 0 or uid == ALLOWED_USER

# ── /start ────────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    context.user_data.clear()
    await update.message.reply_text(
        "👋 *AOW Document Converter*\n\n"
        "Пришлите файлы:\n"
        "📄 Pro Forma `(DOCX)`\n"
        "📊 Спецификацию `(XLSX)` — можно несколько\n\n"
        "Когда все файлы загружены — нажмите *Продолжить*.",
        parse_mode='Markdown'
    )
    return COLLECT

# ── Receive files ─────────────────────────────────────────────────────────────
async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    doc: Document = update.message.document
    if not doc:
        return COLLECT

    name = doc.file_name or ''

    if name.lower().endswith('.docx'):
        context.user_data['docx'] = doc
        await update.message.reply_text(f"✅ DOCX: `{name}`", parse_mode='Markdown')

    elif name.lower().endswith('.xlsx'):
        xlsx_list = context.user_data.setdefault('xlsx', [])
        # Extract spec number from filename: "Spec_43.xlsx" → "43"
        m = re.search(r'(?:Spec[_\s])?(\d+)', Path(name).stem, re.IGNORECASE)
        spec_num = m.group(1) if m else ''
        xlsx_list.append({'doc': doc, 'name': name, 'spec': spec_num})
        await update.message.reply_text(
            f"✅ XLSX: `{name}`  (Spec №: {spec_num or '?'})\n"
            f"Всего XLSX: {len(xlsx_list)}",
            parse_mode='Markdown'
        )

    # Show Continue button if we have at least one file
    if context.user_data.get('docx') or context.user_data.get('xlsx'):
        kb = [[InlineKeyboardButton("▶️  Продолжить", callback_data='go')]]
        await update.message.reply_text(
            "Добавьте ещё файлы или нажмите *Продолжить*",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode='Markdown'
        )
    return COLLECT

# ── Continue → ask invoice number ─────────────────────────────────────────────
async def cb_go(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    await update.callback_query.answer()
    await update.callback_query.message.reply_text(
        "📋 Введите *номер инвойса*:\n_(например: `26-59` → будет FV26-59)_",
        parse_mode='Markdown'
    )
    return ASK_INV

# ── Got invoice number → ask date ─────────────────────────────────────────────
async def got_inv(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    context.user_data['inv'] = update.message.text.strip()
    today = datetime.now().strftime('%d.%m.%Y')
    kb = [[InlineKeyboardButton(f"📅  Сегодня  {today}", callback_data=f'date|{today}')]]
    await update.message.reply_text(
        "📅 Введите *дату* документов:\n_(например: `04.06.2026`)_\nили нажмите кнопку:",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode='Markdown'
    )
    return ASK_DATE

# ── Got date (text or button) → ask spec numbers ──────────────────────────────
async def got_date_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    await update.callback_query.answer()
    context.user_data['date'] = update.callback_query.data.split('|')[1]
    await _ask_spec(update.callback_query.message, context)
    return ASK_SPEC

async def got_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    context.user_data['date'] = update.message.text.strip()
    await _ask_spec(update.message, context)
    return ASK_SPEC

async def _ask_spec(msg, context):
    xlsx_list = context.user_data.get('xlsx', [])
    if not xlsx_list:
        await msg.reply_text(
            "🔢 Введите *номер спецификации*:\n_(например: `44`)_",
            parse_mode='Markdown'
        )
        return
    # Show auto-detected numbers, ask to confirm or correct
    lines = '\n'.join(
        f"  • `{x['name']}` → Spec №: *{x['spec'] or '?'}*"
        for x in xlsx_list
    )
    kb = [[InlineKeyboardButton("✅  Номера верны", callback_data='spec_ok')]]
    await msg.reply_text(
        f"🔢 Автоматически определены номера спецификаций:\n{lines}\n\n"
        "Подтвердите или введите номера через запятую:\n"
        "_(например: `43, 44, 45`)_",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode='Markdown'
    )

async def got_spec_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    await update.callback_query.answer()
    # Use auto-detected spec numbers
    await _ask_pf(update.callback_query.message, context)
    return ASK_PF

async def got_spec_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    raw = update.message.text.strip()
    parts = [p.strip() for p in raw.replace(',', ' ').split()]
    xlsx_list = context.user_data.get('xlsx', [])
    if xlsx_list:
        for i, spec in enumerate(parts):
            if i < len(xlsx_list):
                xlsx_list[i]['spec'] = spec
        context.user_data['spec'] = parts[0] if parts else raw
    else:
        context.user_data['spec'] = raw
    await _ask_pf(update.message, context)
    return ASK_PF

async def _ask_pf(msg, context):
    await msg.reply_text(
        "🔗 Введите *ссылку на Про-форму*:\n_(например: `26-05/51` → будет PF26-05/51)_",
        parse_mode='Markdown'
    )

# ── Got PF reference → process ────────────────────────────────────────────────
async def got_pf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update): return
    context.user_data['pf'] = update.message.text.strip()
    await update.message.reply_text("⏳ Обрабатываю файлы, подождите...")
    await _process(update.message, context)
    return ConversationHandler.END

# ── Core processing ───────────────────────────────────────────────────────────
async def _process(msg, context):
    ud = context.user_data
    inv_num  = f"FV{ud['inv']}"
    date_str = ud['date']
    pf_ref   = f"PF{ud['pf']}"
    xlsx_list = ud.get('xlsx', [])

    logs = []
    def log(m): logs.append(m)

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)

            # Download and process XLSX files
            xlsx_paths = []
            for i, entry in enumerate(xlsx_list):
                f = await entry['doc'].get_file()
                p = tmp / f"spec_{i}.xlsx"
                await f.download_to_drive(str(p))
                xlsx_paths.append((p, entry['spec']))

            # Combine articles from all specs
            spec_arts = set()
            items_all = []
            for path, _ in xlsx_paths:
                spec_arts |= get_articles_from_xlsx(str(path), log)
                items_all += get_xlsx_items_data(str(path), log)

            params_base = dict(
                invoice_num=inv_num, date=date_str,
                spec_num=', '.join(s for _, s in xlsx_paths if s),
                pf_ref=pf_ref, _xlsx_items=items_all
            )

            output_files = []

            # Process DOCX
            if ud.get('docx'):
                f = await ud['docx'].get_file()
                in_docx = tmp / 'input.docx'
                await f.download_to_drive(str(in_docx))
                out_docx = tmp / f'Faktura {inv_num}.docx'
                process_docx(str(in_docx), str(out_docx), params_base,
                             spec_arts or None, log)
                output_files.append(out_docx)

            # Process each XLSX
            for path, spec_num in xlsx_paths:
                params_x = dict(invoice_num=inv_num, date=date_str,
                                spec_num=spec_num, pf_ref=pf_ref)
                name = f'Spec {spec_num}.xlsx' if spec_num else f'Spec {inv_num}.xlsx'
                out_xlsx = tmp / name
                process_xlsx(str(path), str(out_xlsx), params_x, log)
                output_files.append(out_xlsx)

            # Send result files
            for f in output_files:
                if f.exists():
                    with open(f, 'rb') as fh:
                        await msg.reply_document(document=fh, filename=f.name)

            await msg.reply_text("✅ Готово!")

    except Exception as e:
        logging.exception("Processing error")
        await msg.reply_text(f"❌ Ошибка: {e}")

    context.user_data.clear()

# ── Cancel ────────────────────────────────────────────────────────────────────
async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Отменено. /start — начать заново.")
    return ConversationHandler.END

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler('start', cmd_start)],
        states={
            COLLECT: [
                MessageHandler(filters.Document.ALL, receive_file),
                CallbackQueryHandler(cb_go, pattern='^go$'),
            ],
            ASK_INV:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_inv)],
            ASK_DATE: [
                CallbackQueryHandler(got_date_cb, pattern=r'^date\|'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_date_text),
            ],
            ASK_SPEC: [
                CallbackQueryHandler(got_spec_cb, pattern='^spec_ok$'),
                MessageHandler(filters.TEXT & ~filters.COMMAND, got_spec_text),
            ],
            ASK_PF:   [MessageHandler(filters.TEXT & ~filters.COMMAND, got_pf)],
        },
        fallbacks=[CommandHandler('cancel', cmd_cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    logging.info("Bot started")
    app.run_polling(drop_pending_updates=True)

if __name__ == '__main__':
    main()
