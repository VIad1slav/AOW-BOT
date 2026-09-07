"""Прогон реального диалога через ConversationHandler без сети.

Проверяет именно маршрутизацию: что ответ на каждый вопрос попадает
в свой обработчик и состояние продвигается. Прошлая версия теста
собирала приложение, но не гоняла по нему апдейты — и пропустила
ошибку с allow_reentry.
"""
import os, sys, asyncio
from datetime import datetime, timezone

os.environ.setdefault('BOT_TOKEN', '123456:AAdummydummydummydummydummydummydumm')
UID = 421301896
os.environ.setdefault('ALLOWED_USER_ID', str(UID))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import shutil
from pathlib import Path
from telegram import Update, Message, Chat, User, CallbackQuery, Document
import bot as b

ok = True
def check(label, got, want):
    global ok
    if got != want:
        ok = False
        print(f'  FAIL {label}: got {got!r}, want {want!r}')
    else:
        print(f'  ok   {label}')

USER = User(id=UID, is_bot=False, first_name='Тест')
CHAT = Chat(id=UID, type='private')
_ids = iter(range(1000, 9000))


class StubBot:
    """Достаточно правдоподобный бот, чтобы обработчики отработали без сети."""
    def __init__(self):
        self.sent = []
        self.documents = []
        self.file_source = None   # local file the next get_file() serves
        # CommandHandler compares against bot.username, and Message.reply_text
        # reads bot.defaults — both must be real values, not the __getattr__ stub.
        self.defaults = None
        self.username = 'aow_test_bot'
        self.first_name = 'AOW'
    @property
    def id(self): return 1
    async def send_message(self, chat_id, text, **kw):
        self.sent.append(text)
        m = Message(message_id=next(_ids), date=datetime.now(timezone.utc),
                    chat=CHAT, text=text)
        m.set_bot(self)
        return m
    async def delete_message(self, *a, **kw): return True
    async def edit_message_text(self, *a, **kw): return True
    async def answer_callback_query(self, *a, **kw): return True
    async def send_chat_action(self, *a, **kw): return True
    async def send_document(self, *a, **kw):
        self.documents.append(kw.get('filename', '?'))
        return True
    async def get_file(self, file_id, *a, **kw):
        return StubFile(self.file_source)
    def __getattr__(self, name):
        async def _noop(*a, **kw): return True
        return _noop


class StubFile:
    """Stands in for telegram.File — copies a local file instead of downloading."""
    def __init__(self, source):
        self.source = source
    async def download_to_drive(self, dest):
        shutil.copyfile(self.source, dest)
        return dest


def make_doc_update(path, name=None):
    """A document upload. BOT.file_source is what the bot will actually get."""
    BOT.file_source = str(path)
    d = Document(file_id='f%d' % next(_ids), file_unique_id='u%d' % next(_ids),
                 file_name=name or Path(path).name)
    d.set_bot(BOT)
    m = Message(message_id=next(_ids), date=datetime.now(timezone.utc),
                chat=CHAT, from_user=USER, document=d)
    m.set_bot(BOT)
    return Update(update_id=next(_ids), message=m)


def make_photo_update():
    """A passport photo. The OCR itself is stubbed out in the check that uses it."""
    from telegram import PhotoSize
    ph = PhotoSize(file_id='p%d' % next(_ids), file_unique_id='q%d' % next(_ids),
                   width=1600, height=1200)
    ph.set_bot(BOT)
    m = Message(message_id=next(_ids), date=datetime.now(timezone.utc),
                chat=CHAT, from_user=USER, photo=(ph,))
    m.set_bot(BOT)
    return Update(update_id=next(_ids), message=m)


def make_text_update(text):
    m = Message(message_id=next(_ids), date=datetime.now(timezone.utc),
                chat=CHAT, from_user=USER, text=text)
    m.set_bot(BOT)
    return Update(update_id=next(_ids), message=m)


def make_cmd_update(cmd):
    from telegram import MessageEntity
    text = '/' + cmd
    m = Message(message_id=next(_ids), date=datetime.now(timezone.utc),
                chat=CHAT, from_user=USER, text=text,
                entities=[MessageEntity(type=MessageEntity.BOT_COMMAND,
                                        offset=0, length=len(text))])
    m.set_bot(BOT)
    return Update(update_id=next(_ids), message=m)


def make_cb_update(data):
    card = Message(message_id=next(_ids), date=datetime.now(timezone.utc),
                   chat=CHAT, text='card')
    card.set_bot(BOT)
    cq = CallbackQuery(id=str(next(_ids)), from_user=USER, chat_instance='ci',
                       data=data, message=card)
    cq.set_bot(BOT)
    return Update(update_id=next(_ids), callback_query=cq)


BOT = StubBot()

NAMES = {b.COLLECT: 'COLLECT', b.ASK_INV: 'ASK_INV', b.ASK_DATE: 'ASK_DATE',
         b.ASK_SPEC: 'ASK_SPEC', b.ASK_PF: 'ASK_PF', b.CONFIRM: 'CONFIRM',
         b.PF_COLLECT: 'PF_COLLECT', b.PF_PRICE: 'PF_PRICE', b.PF_NUM: 'PF_NUM',
         b.PF_DATE: 'PF_DATE', b.PF_BUYER: 'PF_BUYER', b.PF_REFS: 'PF_REFS',
         b.PF_CONFIRM: 'PF_CONFIRM', b.PF_BUYER_ADDR: 'PF_BUYER_ADDR',
         b.PP_COLLECT: 'PP_COLLECT', b.PP_NUM: 'PP_NUM', b.PP_DATE: 'PP_DATE',
         b.PP_REF: 'PP_REF', b.PP_CONFIRM: 'PP_CONFIRM',
         None: 'нет диалога'}


async def main():
    app = b.build_app()
    app.bot = BOT
    # Application.initialize() would call get_me() over the network — we only
    # need the handler machinery, so mark it initialized and skip that.
    app._initialized = True
    conv = app.handlers[0][0]

    errors = []
    async def on_error(update, context):
        errors.append(repr(context.error))
    app.add_error_handler(on_error)

    def state():
        return NAMES.get(conv._conversations.get((UID, UID)), '???')

    async def feed(update):
        await app.process_update(update)
        return state()

    ud = app.user_data[UID]

    print('— прохождение диалога —')
    check('/start → COLLECT', await feed(make_cmd_update('start')), 'COLLECT')
    check('лишний текст не ломает COLLECT',
          await feed(make_text_update('привет')), 'COLLECT')
    check('«Продолжить» → ASK_INV', await feed(make_cb_update('go')), 'ASK_INV')

    check('ответ «100» → ASK_DATE', await feed(make_text_update('100')), 'ASK_DATE')
    check('год подставлен сам', ud.get('inv'), f'{b._year2()}-100')

    check('кнопка даты → ASK_SPEC',
          await feed(make_cb_update('date|04.06.2026')), 'ASK_SPEC')
    check('дата сохранена', ud.get('date'), '04.06.2026')

    check('номера спец. → ASK_PF', await feed(make_text_update('43, 44')), 'ASK_PF')
    check('Pro Forma → CONFIRM', await feed(make_text_update('26-05/51')), 'CONFIRM')
    check('PF сохранён', ud.get('pf'), '26-05/51')

    print('— варианты ввода номера инвойса —')
    check('«Назад» → ASK_INV', await feed(make_cb_update('back|inv')), 'ASK_INV')
    check('«FV26-100» принимается', await feed(make_text_update('FV26-100')), 'ASK_DATE')
    check('  → 26-100', ud.get('inv'), '26-100')

    await feed(make_cb_update('back|inv'))
    await feed(make_text_update('2026-77'))
    check('«2026-77» → 26-77', ud.get('inv'), '26-77')

    await feed(make_cb_update('back|inv'))
    await feed(make_text_update('25-12'))
    check('«25-12» остаётся как есть', ud.get('inv'), '25-12')

    print('— команды внутри диалога —')
    check('/help не сбивает шаг', await feed(make_cmd_update('help')), 'ASK_DATE')
    check('некорректная дата не двигает шаг',
          await feed(make_text_update('чушь')), 'ASK_DATE')
    check('/cancel завершает', await feed(make_cmd_update('cancel')), 'нет диалога')
    check('после /cancel текст → подсказка',
          await feed(make_text_update('ау')), 'нет диалога')
    check('/start снова стартует', await feed(make_cmd_update('start')), 'COLLECT')

    print('— Pro Forma из Facture VO —')
    await feed(make_cmd_update('cancel'))
    SAMPLE_PDF = Path(b.BASE_DIR) / 'samples' / 'facture_sample.pdf'
    if not SAMPLE_PDF.exists():
        print('  ПРОПУЩЕНО: нет образца', SAMPLE_PDF)
    else:
        ud = app.user_data[UID]
        check('PDF → PF_COLLECT',
              await feed(make_doc_update(SAMPLE_PDF)), 'PF_COLLECT')
        check('машина распознана', len(b._pf_cars(ud)), 1)
        check('  марка', b._pf_cars(ud)[0]['model'], 'PEUGEOT 308 VP')
        check('лишний текст не ломает PF_COLLECT',
              await feed(make_text_update('привет')), 'PF_COLLECT')
        check('«Продолжить» → PF_PRICE', await feed(make_cb_update('pf_go')), 'PF_PRICE')

        check('мусор вместо цены не двигает шаг',
              await feed(make_text_update('дорого')), 'PF_PRICE')
        check('цена «5 400,00» → PF_NUM',
              await feed(make_text_update('5 400,00')), 'PF_NUM')
        check('  цена сохранена', b._pf_cars(ud)[0]['price'], 5400.0)

        check('номер → PF_DATE', await feed(make_text_update('26-8/107')), 'PF_DATE')
        check('  номер с префиксом', ud.get('pf_num'), 'PF26-8/107')
        check('кнопка даты → PF_BUYER',
              await feed(make_cb_update('pfdate|14.08.2026')), 'PF_BUYER')
        check('  дата в ISO', ud.get('pf_date'), '2026-08-14')
        check('  срок оплаты +7 дней', ud.get('pf_termin'), '2026-08-21')

        buyer = '\n'.join(['IVANOU SIARHEI', 'Belarus, Mińsk,',
                           'paszport: MP1234567'])
        check('покупатель → PF_CONFIRM',
              await feed(make_text_update(buyer)), 'PF_CONFIRM')
        check('  имя', ud.get('buyer_name'), 'IVANOU SIARHEI')

        print('— правка CN / просвета —')
        # This step writes to the CN reference file — put it back afterwards so
        # a test run never changes what the live bot will suggest.
        import proforma
        refs_backup = (proforma.REFS_FILE.read_bytes()
                       if proforma.REFS_FILE.exists() else None)
        check('кнопка → PF_REFS', await feed(make_cb_update('pf_refs')), 'PF_REFS')
        check('мусор не двигает шаг', await feed(make_text_update('ага')), 'PF_REFS')
        check('«87032390 155» → PF_CONFIRM',
              await feed(make_text_update('87032390 155')), 'PF_CONFIRM')
        check('  код применён', b._pf_cars(ud)[0]['cn'], '87032390')
        check('  просвет применён', b._pf_cars(ud)[0]['clearance'], '155')
        check('  справочник пополнился',
              proforma._load_refs().get('PEUGEOT 308 VP', {}).get('cn'), '87032390')
        if refs_backup is None:
            proforma.REFS_FILE.unlink(missing_ok=True)
        else:
            proforma.REFS_FILE.write_bytes(refs_backup)

        print('— «Назад» внутри Pro Forma —')
        check('назад к покупателю', await feed(make_cb_update('back|buyer')), 'PF_BUYER')
        check('назад к дате', await feed(make_cb_update('back|date_pf')), 'PF_DATE')
        check('назад к номеру', await feed(make_cb_update('back|num')), 'PF_NUM')
        check('назад к цене', await feed(make_cb_update('back|price')), 'PF_PRICE')
        check('назад к файлам', await feed(make_cb_update('back|pdf')), 'PF_COLLECT')

        print('— машина без цены на шаге подтверждения —')
        await feed(make_cb_update('pf_go'))
        await feed(make_text_update('5400'))
        await feed(make_text_update('26-8/107'))
        await feed(make_cb_update('pfdate|14.08.2026'))
        check('дошли до подтверждения',
              await feed(make_text_update(buyer)), 'PF_CONFIRM')
        check('дослали PDF — остаёмся на подтверждении',
              await feed(make_doc_update(SAMPLE_PDF, 'вторая.pdf')), 'PF_CONFIRM')
        check('  машин стало две', len(b._pf_cars(ud)), 2)
        check('«Создать» возвращает к цене, а не падает',
              await feed(make_cb_update('pf_run')), 'PF_PRICE')
        check('  одна цена на две машины не проходит',
              await feed(make_text_update('5400')), 'PF_PRICE')
        check('  две цены проходят',
              await feed(make_text_update('5400\n6100')), 'PF_NUM')
        check('  вторая цена сохранена', b._pf_cars(ud)[1]['price'], 6100.0)
        await feed(make_cmd_update('cancel'))

        print('— создание документа —')
        await feed(make_doc_update(SAMPLE_PDF))
        await feed(make_cb_update('pf_go'))
        await feed(make_text_update('5400'))
        await feed(make_text_update('26-8/107'))
        await feed(make_cb_update('pfdate|14.08.2026'))
        await feed(make_text_update(buyer))
        BOT.documents.clear()
        check('«Создать» завершает диалог',
              await feed(make_cb_update('pf_run')), 'нет диалога')
        check('  документ отправлен', BOT.documents,
              ['Faktura Pro forma PF26-8-107.docx'])
        check('  номер запомнен для подсказки', b._load_state().get('last_pf'), '26-8/107')
        check('  Pro Forma подставлена в поток фактуры',
              bool(app.user_data[UID].get('docx')), True)
        check('кнопка «сделать фактуру» → ASK_INV',
              await feed(make_cb_update('pf_invoice')), 'ASK_INV')
        await feed(make_cmd_update('cancel'))
        b._save_state({})

        print('— фото паспорта до готового документа —')
        import passport as psp
        import buyers as bkk
        real_read, real_ocr = psp.read_passport_photo, psp.ocr_available
        photo_backup = (bkk.BUYERS_FILE.read_bytes()
                        if bkk.BUYERS_FILE.exists() else None)
        try:
            psp.ocr_available = lambda: True
            b.ocr_available = lambda: True
            psp.read_passport_photo = lambda path: {
                'name': 'IVANOU SIARHEI', 'surname': 'IVANOU',
                'given_names': 'SIARHEI', 'number': 'MP1234567',
                'nationality': 'BLR', 'ok': True, 'failed': [],
                'name_confirmed': True,
            }
            b.read_passport_photo = psp.read_passport_photo
            bkk.BUYERS_FILE.unlink(missing_ok=True)

            await feed(make_cmd_update('start'))
            await feed(make_doc_update(SAMPLE_PDF))
            await feed(make_cb_update('pf_go'))
            await feed(make_text_update('5400'))
            await feed(make_text_update('26-8/130'))
            await feed(make_cb_update('pfdate|21.08.2026'))
            check('фото паспорта не прерывает шаг',
                  await feed(make_photo_update()), 'PF_BUYER')
            check('кнопка «ввести адрес» → свой шаг',
                  await feed(make_cb_update('buyer_addr')), 'PF_BUYER_ADDR')
            # Typed the way it is written on the registration page — the bot
            # is expected to hand back the Polish spelling the invoice uses.
            check('кириллический адрес → подтверждение',
                  await feed(make_text_update(
                      'г. Минск,\nул. Шишкина, дом 12, кв. 7\n25.03.2022')),
                  'PF_CONFIRM')
            ud = app.user_data[UID]
            check('  имя из паспорта', ud.get('buyer_name'), 'IVANOU SIARHEI')
            check('  блок собран целиком', ud.get('buyer_info'),
                  'Belarus, Mińsk,\nul. Szyszkina 12-7\n'
                  'paszport: MP1234567 od 25.03.2022 г.')
            BOT.documents.clear()
            check('документ создаётся', await feed(make_cb_update('pf_run')),
                  'нет диалога')
            check('  файл отправлен', BOT.documents,
                  ['Faktura Pro forma PF26-8-130.docx'])
            check('  покупатель сохранён полностью',
                  bkk.find_buyer('IVANOU SIARHEI')['info'].endswith('25.03.2022 г.'),
                  True)

            # Second time round: the address is already on file, one tap.
            await feed(make_cmd_update('start'))
            await feed(make_doc_update(SAMPLE_PDF))
            await feed(make_cb_update('pf_go'))
            await feed(make_text_update('5400'))
            await feed(make_text_update('26-8/131'))
            await feed(make_cb_update('pfdate|21.08.2026'))
            await feed(make_photo_update())
            check('кнопка справочника → сразу подтверждение',
                  await feed(make_cb_update('buyer_known')), 'PF_CONFIRM')
            check('  адрес подставлен из справочника',
                  app.user_data[UID].get('buyer_info').endswith('25.03.2022 г.'),
                  True)
            await feed(make_cmd_update('cancel'))
        finally:
            psp.read_passport_photo, psp.ocr_available = real_read, real_ocr
            b.read_passport_photo, b.ocr_available = real_read, real_ocr
            if photo_backup is None:
                bkk.BUYERS_FILE.unlink(missing_ok=True)
            else:
                bkk.BUYERS_FILE.write_bytes(photo_backup)

        print('— справочник покупателей —')
        import buyers as bk
        buyers_backup = (bk.BUYERS_FILE.read_bytes()
                         if bk.BUYERS_FILE.exists() else None)
        try:
            check('покупатель запомнен после создания',
                  bool(bk.find_buyer('IVANOU SIARHEI')), True)
            await feed(make_cmd_update('start'))
            await feed(make_doc_update(SAMPLE_PDF))
            await feed(make_cb_update('pf_go'))
            await feed(make_text_update('5400'))
            await feed(make_text_update('26-8/120'))
            await feed(make_cb_update('pfdate|20.08.2026'))
            check('дошли до покупателя', state(), 'PF_BUYER')
            check('кнопка сохранённого → PF_CONFIRM',
                  await feed(make_cb_update('buyer|0')), 'PF_CONFIRM')
            ud = app.user_data[UID]
            check('  имя подставлено', ud.get('buyer_name'), 'IVANOU SIARHEI')
            check('  данные подставлены целиком',
                  ud.get('buyer_info').startswith('Belarus'), True)
            check('несуществующая кнопка не роняет шаг',
                  await feed(make_cb_update('back|buyer')), 'PF_BUYER')
            check('  битый индекс возвращает на тот же шаг',
                  await feed(make_cb_update('buyer|99')), 'PF_BUYER')
            await feed(make_cmd_update('cancel'))
        finally:
            if buyers_backup is None:
                bk.BUYERS_FILE.unlink(missing_ok=True)
            else:
                bk.BUYERS_FILE.write_bytes(buyers_backup)

        print('— другие площадки —')
        SAG = Path(b.BASE_DIR) / 'samples' / 'sag_sample.pdf'
        OL = Path(b.BASE_DIR) / 'samples' / 'openlane_sample.pdf'
        if not (SAG.exists() and OL.exists()):
            print('  ПРОПУЩЕНО: нет образцов SAG / OPENLANE')
        else:
            await feed(make_cmd_update('start'))
            check('SAG PDF принят', await feed(make_doc_update(SAG)), 'PF_COLLECT')
            ud = app.user_data[UID]
            check('  площадка SAG', ud['pdf'][0]['source'], 'SAG')
            check('  модель', b._pf_cars(ud)[0]['model'], 'PEUGEOT 5008')
            check('OPENLANE PDF принят', await feed(make_doc_update(OL)), 'PF_COLLECT')
            check('  площадка OPENLANE', ud['pdf'][1]['source'], 'OPENLANE')
            check('  машин в комплекте', len(b._pf_cars(ud)), 2)
            check('  у OPENLANE нет госномера', b._pf_cars(ud)[1]['plate'], '')

            await feed(make_cb_update('pf_go'))
            check('две цены для двух площадок',
                  await feed(make_text_update('16900\n8500')), 'PF_NUM')
            await feed(make_text_update('26-8/115'))
            await feed(make_cb_update('pfdate|20.08.2026'))
            await feed(make_text_update(buyer))
            BOT.documents.clear()
            check('смешанная Pro Forma создаётся',
                  await feed(make_cb_update('pf_run')), 'нет диалога')
            check('  документ отправлен', BOT.documents,
                  ['Faktura Pro forma PF26-8-115.docx'])
            b._save_state({})
            await feed(make_cmd_update('cancel'))

        print('— чужой формат —')
        junk = Path(b.WORK_DIR) / 'junk.txt'
        junk.parent.mkdir(parents=True, exist_ok=True)
        junk.write_text('не документ', encoding='utf-8')
        await feed(make_cmd_update('proforma'))
        check('/proforma → PF_COLLECT', state(), 'PF_COLLECT')
        check('.txt отклоняется, шаг тот же',
              await feed(make_doc_update(junk, 'note.txt')), 'PF_COLLECT')
        junk.unlink(missing_ok=True)
        await feed(make_cmd_update('cancel'))

    print('— Pro Forma по спецификации (GOLFSTREAM) —')
    await feed(make_cmd_update('cancel'))
    SPEC_XLSX = Path(b.BASE_DIR) / 'samples' / 'pipes_spec_126.xlsx'
    SPEC_NAME = 'Proforma AOW - Golfstream - Specification 126 (AB01139034).xlsx'
    if not SPEC_XLSX.exists():
        print('  ПРОПУЩЕНО: нет образца', SPEC_XLSX)
    else:
        check('/proformapipes → PP_COLLECT',
              await feed(make_cmd_update('proformapipes')), 'PP_COLLECT')
        check('лишний текст не ломает PP_COLLECT',
              await feed(make_text_update('привет')), 'PP_COLLECT')
        check('XLSX принят',
              await feed(make_doc_update(SPEC_XLSX, SPEC_NAME)), 'PP_COLLECT')
        ud = app.user_data[UID]
        check('  позиций разобрано', len(ud['pipes']['items']), 38)
        check('  сумма', ud['pipes']['total'], 27827.49)
        check('  номер спецификации из имени', ud.get('pp_spec'), '126')
        check('  код без букв и нуля', ud.get('pp_code'), '1139034')
        check('  в поток фактуры файл не попал', ud.get('xlsx'), None)

        check('«Продолжить» → PP_DATE', await feed(make_cb_update('pp_go')), 'PP_DATE')
        check('чушь вместо даты не двигает шаг',
              await feed(make_text_update('позавчера')), 'PP_DATE')
        check('кнопка даты → PP_NUM',
              await feed(make_cb_update('ppdate|03.09.2026')), 'PP_NUM')
        check('  дата в ISO', ud.get('pf_date'), '2026-09-03')
        check('  срок оплаты +25 дней', ud.get('pf_termin'), '2026-09-28')

        print('— номер: год и месяц из даты выставления —')
        check('мусор вместо номера не двигает шаг',
              await feed(make_text_update('')), 'PP_NUM')
        check('один номер → PP_CONFIRM',
              await feed(make_text_update('110')), 'PP_CONFIRM')
        check('  «110» превратилось в PF26-9/110', ud.get('pf_num'), 'PF26-9/110')

        await feed(make_cb_update('back|pp_num'))
        await feed(make_text_update('26-8/109'))
        check('полный номер остаётся как есть', ud.get('pf_num'), 'PF26-8/109')
        await feed(make_cb_update('back|pp_num'))
        await feed(make_text_update('PF2026-9/111'))
        check('четырёхзначный год укорачивается', ud.get('pf_num'), 'PF26-9/111')
        await feed(make_cb_update('back|pp_num'))
        await feed(make_text_update('9/112'))
        check('«9/112» → добавляется только год', ud.get('pf_num'), 'PF26-9/112')

        # A document written in October but dated September must be numbered by
        # the date on it, not by the day it was typed.
        await feed(make_cb_update('back|pp_date'))
        await feed(make_text_update('30.09.2026'))
        await feed(make_text_update('113'))
        check('номер идёт за датой документа, а не за сегодня',
              ud.get('pf_num'), 'PF26-9/113')
        await feed(make_cb_update('back|pp_date'))
        await feed(make_cb_update('ppdate|03.09.2026'))
        await feed(make_text_update('110'))

        print('— правка нижней строки —')
        check('кнопка → PP_REF', await feed(make_cb_update('pp_ref')), 'PP_REF')
        check('одно число не проходит', await feed(make_text_update('126')), 'PP_REF')
        check('два числа → PP_CONFIRM',
              await feed(make_text_update('77 0000042')), 'PP_CONFIRM')
        check('  номер спецификации', ud.get('pp_spec'), '77')
        check('  ведущие нули срезаны', ud.get('pp_code'), '42')
        await feed(make_cb_update('pp_ref'))
        await feed(make_text_update('126 1139034'))

        print('— «Назад» внутри Pro Forma по спецификации —')
        check('назад к номеру', await feed(make_cb_update('back|pp_num')), 'PP_NUM')
        check('назад к дате', await feed(make_cb_update('back|pp_date')), 'PP_DATE')
        check('назад к файлу', await feed(make_cb_update('back|pipes')), 'PP_COLLECT')
        check('файл на месте', bool(ud.get('pipes')), True)

        print('— создание документа —')
        await feed(make_cb_update('pp_go'))
        await feed(make_cb_update('ppdate|03.09.2026'))
        await feed(make_text_update('110'))
        BOT.documents.clear()
        check('«Создать» завершает диалог',
              await feed(make_cb_update('pp_run')), 'нет диалога')
        check('  документ отправлен', BOT.documents,
              ['Faktura Pro Forma PF26-9-110.docx'])
        check('  номер запомнен для подсказки',
              b._load_state().get('last_pf'), '26-9/110')
        b._save_state({})
        await feed(make_cmd_update('cancel'))

        print('— XLSX вне этого раздела по-прежнему идёт в фактуру —')
        check('спецификация без команды → COLLECT',
              await feed(make_doc_update(SPEC_XLSX, SPEC_NAME)), 'COLLECT')
        check('  файл в списке спецификаций',
              len(app.user_data[UID].get('xlsx', [])), 1)
        await feed(make_cmd_update('cancel'))

    print('— подсказка следующего номера —')
    b._save_state({'last_inv': '26-59'})
    check('тот же год → +1', b._suggest_inv(), '26-60')
    b._save_state({'last_inv': '25-59'})
    check('другой год → сброс на 1', b._suggest_inv(), f'{b._year2()}-1')
    b._save_state({'last_pf': '26-8/106'})
    yy, mm = b._year2(), datetime.now().month
    want = f'{yy}-{mm}/107' if yy == '26' else f'{yy}-{mm}/1'
    check('следующий номер Pro Forma', b._suggest_pf(), want)
    b._save_state({'last_pf': '99-1/5'})
    check('другой год → сброс', b._suggest_pf(), f'{yy}-{mm}/1')
    b._save_state({})

    if errors:
        print('\nИСКЛЮЧЕНИЯ В ОБРАБОТЧИКАХ:')
        for e in errors:
            print('  ', e)
        return False
    return ok


res = asyncio.run(main())
print('\nВСЁ ХОРОШО' if res else '\nЕСТЬ ОШИБКИ')
sys.exit(0 if res else 1)
