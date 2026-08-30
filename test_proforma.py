#!/usr/bin/env python3
"""Offline checks for the Facture VO → Pro Forma step. No network, no Telegram.

    python3 test_proforma.py [путь-к-Facture.pdf] [путь-к-образцу-ProForma.docx]

The round-trip check is the important one: rebuilding the sample Pro Forma from
the sample Facture must reproduce the original document exactly — same text,
same bold/italic/size on every single character.
"""

import sys
import tempfile
from pathlib import Path

from docx import Document

import proforma as pf
from docx_text import iter_paragraphs

HERE = Path(__file__).parent


def _sample(arg_index, fixture, *fallbacks):
    """Prefer an explicit argument, then samples/, then wherever the file
    originally came from — so the checks run on a fresh clone too."""
    if len(sys.argv) > arg_index:
        return Path(sys.argv[arg_index])
    for candidate in [HERE / 'samples' / fixture, *fallbacks]:
        if Path(candidate).exists():
            return Path(candidate)
    return HERE / 'samples' / fixture


DOWNLOADS = Path.home() / 'Downloads' / 'Telegram Desktop'
SAMPLE_PDF = _sample(1, 'facture_sample.pdf',
                     Path.home() / 'Desktop' / 'Facture VO N° 202600028424.pdf',
                     DOWNLOADS / 'Facture VO N° 202600028424.pdf')
SAMPLE_DOCX = _sample(2, 'proforma_sample.docx',
                      Path.home() / 'Desktop' / 'Faktura Pro forma PF26-8-106.docx',
                      DOWNLOADS / 'Faktura Pro Forma PF26-8-106.docx')

BUYER_NAME = 'IVANOU SIARHEI'
BUYER_INFO = ('Belarus, Mińsk, \nul. Szyszkina 12-7\n'
              'paszport: MP1234567 od 25.03.2022 г.')

failures = []


def check(name, cond, detail=''):
    print(('  ✅ ' if cond else '  ❌ ') + name + (('  — ' + detail) if detail and not cond else ''))
    if not cond:
        failures.append(name)
    return cond


def char_map(path):
    """Per paragraph: the text, plus the formatting that applies to each character."""
    out = []
    for par in iter_paragraphs(Document(str(path))):
        text, fmt = [], []
        for run in par.runs:
            style = (run.bold, run.italic, run.font.size, run.font.name, run.underline)
            for ch in run.text:
                text.append(ch)
                fmt.append(style)
        out.append((''.join(text), fmt))
    return out


def test_money():
    print('\nСуммы')
    cases = [
        ('5400', 5400.0), ('5 400', 5400.0), ('5400,00', 5400.0),
        ('5400.50', 5400.5), ('5.400,00', 5400.0), ('5,400.00', 5400.0),
        ('5400 EUR', 5400.0), ('12 345,67', 12345.67), ('1.234.567', 1234567.0),
        ('750', 750.0), ('0', None), ('-100', None), ('abc', None),
        ('', None), ('5400 евро', None),
    ]
    for raw, want in cases:
        got = pf.parse_money(raw)
        check('parse_money(%r) → %r' % (raw, want), got == want, 'получено %r' % got)
    for value, want in [(5400, '5 400,00'), (750.5, '750,50'),
                        (1234567.89, '1 234 567,89'), (0.0, '0,00')]:
        got = pf.fmt_money(value)
        check('fmt_money(%r) → %r' % (value, want), got == want, 'получено %r' % got)


def test_parse():
    print('\nРазбор Facture VO')
    if not SAMPLE_PDF.exists():
        check('образец PDF на месте', False, str(SAMPLE_PDF))
        return None
    head = pf.parse_facture(SAMPLE_PDF)
    v = head['vehicles'][0]
    check('найдена одна машина', len(head['vehicles']) == 1, str(len(head['vehicles'])))
    check('модель', v['model'] == 'PEUGEOT 308 VP', v['model'])
    check('госномер', v['plate'] == 'FW-646-BP', v['plate'])
    check('VIN', v['vin'] == 'VF3LPHNSKLS232589', v['vin'])
    check('дата регистрации', v['reg_date'] == '23/12/2020', v['reg_date'])
    check('пробег', v['km'] == '215718', v['km'])
    check('нет пропущенных полей', pf.missing_fields(v) == [], str(pf.missing_fields(v)))
    check('номер фактуры', head['facture_no'] == '202600028424', head['facture_no'])
    check('дата фактуры', head['facture_date'] == '13/08/2026', head['facture_date'])
    return head


def test_roundtrip(head):
    print('\nСборка Pro Forma — сверка с оригиналом')
    if head is None:
        return
    if not SAMPLE_DOCX.exists():
        # The filled-in sample carries the buyer's passport, so it is kept off
        # the server. This check belongs where the template is built anyway.
        print('  ⏭  ПРОПУЩЕНО — нет образца %s' % SAMPLE_DOCX)
        return
    v = dict(head['vehicles'][0], price=5400)
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / 'rt.docx'
        res = pf.build_proforma(out, {
            'pf_num': 'PF26-8/106', 'date': '2026-08-14', 'termin': '2026-08-21',
            'buyer_name': BUYER_NAME, 'buyer_info': BUYER_INFO, 'vehicles': [v],
        })
        check('итог посчитан', res['total'] == 5400.0, str(res['total']))

        want, got = char_map(SAMPLE_DOCX), char_map(out)
        if not check('столько же абзацев', len(want) == len(got),
                     '%d против %d' % (len(want), len(got))):
            return
        bad_text = [i for i, ((a, _), (b, _)) in enumerate(zip(want, got)) if a != b]
        check('текст совпадает во всех абзацах', not bad_text,
              'расходятся абзацы %s' % bad_text[:5])
        bad_fmt = []
        for i, ((ta, fa), (tb, fb)) in enumerate(zip(want, got)):
            if ta == tb and fa != fb:
                j = next(k for k, (x, y) in enumerate(zip(fa, fb)) if x != y)
                bad_fmt.append('абзац %d, символ %d (%r)' % (i, j, ta[j]))
        check('форматирование каждого символа совпадает', not bad_fmt,
              '; '.join(bad_fmt[:3]))


FACTURE_TEXT = """AOW GROUP Sp. z o.o.
Facture VO N°
202600028424
Client N° 4288503 01-001 Warszawa
Date : 13/08/2026
Véhicules d'occasion Montant HT
Avenant n°N48183 PEUGEOT 308 VP
Immatriculé FW-646-BP
Kms Compteur 215718
Date de mise en circulation le 23/12/2020
N° chassis VF3LPHNSKLS232589
Année Millésime 2020
3888.33 - - -
Frais de gestion 415.00 - - -
TOTAL EN € 4303.33 - -
"""


def parse_text(text):
    """Run the parser over crafted invoice text, without needing a real PDF."""
    original = pf.read_pdf_text
    pf.read_pdf_text = lambda path: text
    try:
        return pf.parse_facture('не-важно.pdf')
    finally:
        pf.read_pdf_text = original


def test_parse_variants():
    print('\nРазбор — варианты фактур')

    head = parse_text(FACTURE_TEXT)
    check('эталонный текст разбирается', len(head['vehicles']) == 1)

    # Two cars on one invoice: each block must keep its own details.
    second = ("Avenant n°N48200 RENAULT CLIO V\n"
              "Immatriculé GA-123-BC\n"
              "Kms Compteur 90210\n"
              "Date de mise en circulation le 01/02/2021\n"
              "N° chassis VF1RJA00123456789\n"
              "4100.00 - - -\n")
    head = parse_text(FACTURE_TEXT.replace('Frais de gestion', second + 'Frais de gestion'))
    cars = head['vehicles']
    check('две машины найдены', len(cars) == 2, str(len(cars)))
    if len(cars) == 2:
        check('  у первой свой VIN', cars[0]['vin'] == 'VF3LPHNSKLS232589', cars[0]['vin'])
        check('  у второй своя модель', cars[1]['model'] == 'RENAULT CLIO V', cars[1]['model'])
        check('  у второй свой госномер', cars[1]['plate'] == 'GA-123-BC', cars[1]['plate'])
        check('  у второй своя дата', cars[1]['reg_date'] == '01/02/2021', cars[1]['reg_date'])
        check('  данные не перетекли между машинами',
              cars[0]['plate'] == 'FW-646-BP', cars[0]['plate'])

    # The extractor sometimes glues the amount column onto the model line.
    head = parse_text(FACTURE_TEXT.replace(
        'Avenant n°N48183 PEUGEOT 308 VP',
        'Avenant n°N48183 PEUGEOT 308 VP 3888.33 - - -'))
    check('сумма в строке модели отрезается',
          head['vehicles'][0]['model'] == 'PEUGEOT 308 VP',
          head['vehicles'][0]['model'])

    head = parse_text(FACTURE_TEXT.replace('N°', 'No'))
    check('«No» вместо «N°» тоже читается', len(head['vehicles']) == 1)

    # A car whose plate and VIN are missing must be reported, not silently blank.
    broken = FACTURE_TEXT.replace('Immatriculé FW-646-BP\n', '') \
                         .replace('N° chassis VF3LPHNSKLS232589\n', '')
    car = parse_text(broken)['vehicles'][0]
    check('пропуски перечислены', pf.missing_fields(car) == ['госномер', 'VIN'],
          str(pf.missing_fields(car)))
    check('  модель всё равно найдена', car['model'] == 'PEUGEOT 308 VP', car['model'])

    for bad, why in [('', 'пустой текст'),
                     ('просто какой-то текст', 'текст не от Ayvens'),
                     ('Facture VO N°\n123456\nDate : 01/01/2026', 'нет ни одной машины')]:
        try:
            parse_text(bad)
            check('отвергается: ' + why, False, 'исключения не было')
        except ValueError:
            check('отвергается: ' + why, True)


SAG_TEXT = """    Description VAT Base VAT Rate VAT Amount
Sale of a second-hand vehicle, the buyer being fully aware of the vehicle's condition.
415.00
14,100.00
Plate 2BTD533
Registration date 14/03/2022
Make and model PEUGEOT - 5008
Chassis VF3MRHNSUNS037337
KW 96
Capacity 1199
HP 7
Last mileage 77891
Total to pay (EUR) 14,515.00
 INVOICE SAG26/016801 Page 1/1
Date
Due date
22/06/2026
22/06/2026
"""

OPENLANE_TEXT = """Pro Forma 202607-54451
VAT PL5272934015
Data faktury 25-07-2026
Referencje 24/07/2026-11284974/5943483
Marka Citroën
Model C3
Typ Aircross SHINE 1.2
Rodzaj Sports Utility Vehicle (SUV)
Pierwsza rejestracja 29-10-2021
Podwozie VF72RHNPMM4327115
Rozmiar silnika 1199
Moc (kW) - standardowa emisja CO² 81 - EU6d
Przebieg 115377
Razem 7 214,00 0,00€
OPENLANE Europe NV
"""


def test_sag():
    print('\nSAG — английская платёжка')
    head = parse_text(SAG_TEXT)
    check('площадка определена как SAG', head['source'] == 'SAG', head.get('source'))
    check('номер документа', head['facture_no'] == 'SAG26/016801', head['facture_no'])
    check('дата документа', head['facture_date'] == '22/06/2026', head['facture_date'])
    check('итог документа', head['facture_total'] == '14,515.00', head['facture_total'])
    v = head['vehicles'][0]
    check('модель без тире', v['model'] == 'PEUGEOT 5008', v['model'])
    check('госномер', v['plate'] == '2BTD533', v['plate'])
    check('VIN', v['vin'] == 'VF3MRHNSUNS037337', v['vin'])
    check('дата регистрации', v['reg_date'] == '14/03/2022', v['reg_date'])
    check('пробег', v['km'] == '77891', v['km'])
    check('объём двигателя', v['capacity'] == '1199', v['capacity'])
    check('код CN по объёму 1199', v['cn'] == '87032290', v['cn'])
    check('пропусков нет', pf.missing_fields(v) == [], str(pf.missing_fields(v)))


def test_openlane():
    print('\nOPENLANE — польская платёжка')
    head = parse_text(OPENLANE_TEXT)
    check('площадка определена как OPENLANE',
          head['source'] == 'OPENLANE', head.get('source'))
    check('номер документа', head['facture_no'] == '202607-54451', head['facture_no'])
    check('дата приведена к ДД/ММ/ГГГГ',
          head['facture_date'] == '25/07/2026', head['facture_date'])
    check('итог документа', head['facture_total'] == '7214,00', head['facture_total'])
    v = head['vehicles'][0]
    check('марка+модель+тип склеены',
          v['model'] == 'CITROËN C3 AIRCROSS SHINE 1.2', v['model'])
    check('VIN', v['vin'] == 'VF72RHNPMM4327115', v['vin'])
    check('дата регистрации из 29-10-2021',
          v['reg_date'] == '29/10/2021', v['reg_date'])
    check('пробег', v['km'] == '115377', v['km'])
    check('госномера нет — и это не считается пропуском',
          v['plate'] == '' and pf.missing_fields(v) == [], str(pf.missing_fields(v)))


# The same platform, issued in the buyer's language. Note the doubled spaces
# inside the labels and the empty "Тип" — both come straight out of the real
# PDF and both used to break the parser.
OPENLANE_RU_TEXT = """Проформа 202505-40739
НДС PL5272934015
Дата  выставления  счета 28-05-2025
Ссылки 28/05/2025-9445188/5238604
Марка Volvo
Модель V60
Тип
Вид Break
Первая  регистрация 20-02-2015
Шасси YV1FW8481F1259266
Размер  двигателя 1560
Мощность  – норматив  загрязняющих  веществ  в  выбросах 84 - EU5
Пробег 182079
На  сумму 6 833,34€
OPENLANE Europe NV
"""


def test_openlane_languages():
    print('\nOPENLANE на другом языке')
    head = parse_text(OPENLANE_RU_TEXT)
    check('русский счёт распознан как OPENLANE',
          head['source'] == 'OPENLANE', head.get('source'))
    check('номер документа', head['facture_no'] == '202505-40739', head['facture_no'])
    check('дата из «Дата  выставления  счета»',
          head['facture_date'] == '28/05/2025', head['facture_date'])
    check('итог из «На  сумму»', head['facture_total'] == '6833,34',
          head['facture_total'])
    v = head['vehicles'][0]
    check('марка и модель', v['model'] == 'VOLVO V60', v['model'])
    check('VIN из «Шасси»', v['vin'] == 'YV1FW8481F1259266', v['vin'])
    check('дата регистрации', v['reg_date'] == '20/02/2015', v['reg_date'])
    check('пробег из «Пробег»', v['km'] == '182079', v['km'])
    check('объём из «Размер  двигателя»', v['capacity'] == '1560', v['capacity'])
    check('код CN по объёму 1560', v['cn'] == '87032390', v['cn'])

    # The label is present but its value is blank; a pattern that lets \\s+ eat
    # the newline picks up "Вид Break" from the line below instead.
    check('пустой «Тип» не утягивает следующую строку',
          'ВИД' not in v['model'] and 'BREAK' not in v['model'], v['model'])

    # And the Polish one must keep working.
    head = parse_text(OPENLANE_TEXT)
    check('польский счёт по-прежнему разбирается',
          head['vehicles'][0]['model'] == 'CITROËN C3 AIRCROSS SHINE 1.2',
          head['vehicles'][0]['model'])


def test_unreadable_message():
    print('\nСообщение об ошибке')
    # A platform the bot knows, in a language it does not: the message has to
    # say that, not "unknown platform".
    hungarian = ('OPENLANE Europe NV\nGyártmány Volvo\nTípus V60\n'
                 'Alvázszám YV1FW8481F1259266\n')
    try:
        parse_text(hungarian)
        check('нечитаемый счёт отвергается', False, 'исключения не было')
    except ValueError as e:
        check('сказано, что площадка опознана', 'OPENLANE' in str(e), str(e))
        check('  и что дело в данных машины',
              'машины' in str(e) or 'язык' in str(e), str(e))
    try:
        parse_text('какой-то посторонний документ')
        check('чужой документ отвергается', False, 'исключения не было')
    except ValueError as e:
        check('для чужого документа — список площадок',
              'Ayvens' in str(e) and 'SAG' in str(e), str(e))


def test_cn_by_capacity():
    print('\nКод CN по объёму двигателя')
    for cc, want in [('999', '87032190'), ('1199', '87032290'),
                     ('1500', '87032290'), ('1598', '87032390'),
                     ('2999', '87032390'), ('3500', '87032490')]:
        got = pf._cn_for(cc)
        check('%s см³ → %s' % (cc, want), got == want, got)
    check('без объёма — умолчание', pf._cn_for('') == pf.DEFAULT_CN, pf._cn_for(''))
    check('мусор вместо объёма — умолчание',
          pf._cn_for('abc') == pf.DEFAULT_CN, pf._cn_for('abc'))


def test_plate_line_dropped():
    print('\nСтрока госномера у машин без номера')
    from docx import Document as _Doc
    from docx_text import doc_tables as _tables
    head = parse_text(OPENLANE_TEXT)
    v = dict(head['vehicles'][0], price=8500)
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / 'noplate.docx'
        pf.build_proforma(out, {
            'pf_num': 'PF26-8/112', 'date': '2026-08-20', 'termin': '2026-08-27',
            'buyer_name': BUYER_NAME, 'buyer_info': BUYER_INFO, 'vehicles': [v]})
        cell = _tables(_Doc(str(out)))[0].rows[1].cells[1].text
        check('ячейка кончается кодом CN, без пустого хвоста',
              cell.rstrip().endswith('просвет 145 mm') and not cell.endswith('\n'),
              repr(cell[-40:]))
        check('плейсхолдер не остался', '{{' not in cell)

    # A car that does have a plate must still show it.
    ay = parse_text(FACTURE_TEXT)['vehicles'][0]
    ay['price'] = 5400
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / 'plate.docx'
        pf.build_proforma(out, {
            'pf_num': 'PF26-8/113', 'date': '2026-08-20', 'termin': '2026-08-27',
            'buyer_name': BUYER_NAME, 'buyer_info': BUYER_INFO, 'vehicles': [ay]})
        cell = _tables(_Doc(str(out)))[0].rows[1].cells[1].text
        check('госномер на месте, когда он есть', cell.rstrip().endswith('FW-646-BP'),
              repr(cell[-30:]))


def test_real_pdfs():
    print('\nНастоящие PDF всех площадок')
    for fixture, source, model in [
            ('facture_sample.pdf', 'Ayvens', 'PEUGEOT 308 VP'),
            ('facture_bmw.pdf', 'Ayvens', 'BMW X2 VP'),
            ('sag_sample.pdf', 'SAG', 'PEUGEOT 5008'),
            ('openlane_sample.pdf', 'OPENLANE', 'CITROËN C3 AIRCROSS SHINE 1.2'),
            ('openlane_ru.pdf', 'OPENLANE', 'VOLVO V60')]:
        path = HERE / 'samples' / fixture
        if not path.exists():
            print('  ⏭  ПРОПУЩЕНО — нет %s' % fixture)
            continue
        head = pf.parse_facture(path)
        check('%s: площадка' % fixture, head['source'] == source, head.get('source'))
        check('%s: модель' % fixture, head['vehicles'][0]['model'] == model,
              head['vehicles'][0]['model'])


def test_template_shape():
    print('\nШаблон')
    import re as _re
    from docx import Document as _Doc
    from docx_text import iter_paragraphs as _iter
    names = set()
    for par in _iter(_Doc(str(pf.TEMPLATE))):
        names.update(_re.findall(r'\{\{(\w+)\}\}', par.text))
    expected = {'PF_NUM', 'DATA', 'TERMIN', 'BUYER_NAME', 'BUYER_INFO', 'LP',
                'AUTO', 'REG', 'VIN', 'CN', 'CLEAR', 'PLATE', 'QTY', 'CENA',
                'WARTOSC', 'BRUTTO', 'RAZEM'}
    check('набор плейсхолдеров ровно тот, что заполняет бот',
          names == expected, 'лишние %s, недостающие %s'
          % (sorted(names - expected), sorted(expected - names)))

    body = '\n'.join(p.text for p in _iter(_Doc(str(pf.TEMPLATE))))
    for leak in ['IVANOU', 'Szyszkina', 'MP1234567', 'PEUGEOT', 'VF3LPHNSKLS232589',
                 'FW-646-BP', '5 400']:
        check('в шаблоне не осталось данных образца: ' + leak, leak not in body)


def test_output_is_valid():
    print('\nГодность готового файла')
    import zipfile
    from docx import Document as _Doc
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / pf.out_filename('PF26-9/1')
        pf.build_proforma(out, {
            'pf_num': 'PF26-9/1', 'date': '2026-09-01', 'termin': '2026-09-08',
            'buyer_name': 'ИВАНОВ ИВАН', 'buyer_info': 'Belarus\npaszport: AB123',
            'vehicles': [{'model': 'ŠKODA OCTAVIA', 'plate': 'AA-111-BB',
                          'reg_date': '05/05/2020', 'vin': 'TMBJJ7NE0J0123456',
                          'cn': '87032390', 'clearance': '155', 'price': 7250.75}],
        })
        check('файл создан', out.exists())
        check('имя без слэша', out.name == 'Faktura Pro forma PF26-9-1.docx', out.name)
        with zipfile.ZipFile(out) as z:
            check('это корректный OOXML-архив', z.testzip() is None)
            check('  внутри есть document.xml', 'word/document.xml' in z.namelist())
        body = '\n'.join(p.text for p in iter_paragraphs(_Doc(str(out))))
        for want in ['PF26-9/1', '2026-09-01', '2026-09-08', 'ИВАНОВ ИВАН',
                     'ŠKODA OCTAVIA', 'AA-111-BB', 'TMBJJ7NE0J0123456',
                     '87032390', '155 mm', '7 250,75']:
            check('в документе есть: ' + want, want in body)
        check('паспорт покупателя на месте', 'paszport: AB123' in body)
        check('плейсхолдеров не осталось', '{{' not in body)


def test_refs_memory():
    print('\nСправочник CN / просвета')
    backup = pf.REFS_FILE.read_bytes() if pf.REFS_FILE.exists() else None
    try:
        pf.REFS_FILE.unlink(missing_ok=True)
        pf.remember_refs('BMW X5 xDrive', '87033390', '210')
        car = {'model': 'bmw  x5   xdrive', 'cn': pf.DEFAULT_CN,
               'clearance': pf.DEFAULT_CLEARANCE}
        pf._apply_remembered_refs(car)
        check('запомненное применяется, регистр и пробелы не мешают',
              (car['cn'], car['clearance']) == ('87033390', '210'),
              str((car['cn'], car['clearance'])))
        other = {'model': 'FIAT PANDA', 'cn': pf.DEFAULT_CN,
                 'clearance': pf.DEFAULT_CLEARANCE}
        pf._apply_remembered_refs(other)
        check('для другой модели остаются умолчания',
              (other['cn'], other['clearance']) == (pf.DEFAULT_CN, pf.DEFAULT_CLEARANCE),
              str((other['cn'], other['clearance'])))
        pf.remember_refs('', '1', '2')
        check('пустая модель не пишется', '' not in pf._load_refs())
    finally:
        if backup is None:
            pf.REFS_FILE.unlink(missing_ok=True)
        else:
            pf.REFS_FILE.write_bytes(backup)


def test_multi(head):
    print('\nНесколько машин в одной Pro Forma')
    if head is None:
        return
    base = head['vehicles'][0]
    cars = [
        dict(base, price=5400),
        dict(base, model='RENAULT CLIO V', plate='GA-123-BC',
             vin='VF1RJA00123456789', reg_date='01/02/2021', price=4600.5),
        dict(base, model='SKODA OCTAVIA', plate='HB-777-ZZ',
             vin='TMBJJ7NE0J0123456', reg_date='15/06/2019', price=7000),
    ]
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / 'multi.docx'
        res = pf.build_proforma(out, {
            'pf_num': 'PF26-8/107', 'date': '2026-08-20', 'termin': '2026-08-27',
            'buyer_name': BUYER_NAME, 'buyer_info': BUYER_INFO, 'vehicles': cars,
        })
        check('сумма трёх машин', res['total'] == 17000.5, str(res['total']))

        doc = Document(str(out))
        from docx_text import doc_tables
        table = doc_tables(doc)[0]
        check('строк в таблице 5 (шапка + 3 + итого)', len(table.rows) == 5,
              str(len(table.rows)))
        lp = [table.rows[i].cells[0].text.strip() for i in (1, 2, 3)]
        check('нумерация позиций 1/2/3', lp == ['1', '2', '3'], str(lp))
        check('вторая строка — своя машина',
              'RENAULT CLIO V' in table.rows[2].cells[1].text
              and 'GA-123-BC' in table.rows[2].cells[1].text)
        check('своя цена во второй строке',
              table.rows[2].cells[3].text.strip() == '4 600,50',
              table.rows[2].cells[3].text)
        check('итог в строке Razem', table.rows[4].cells[4].text.strip() == '17 000,50',
              table.rows[4].cells[4].text)
        body = '\n'.join(p.text for p in iter_paragraphs(doc))
        check('итог в строке «Wartość brutto»', '17 000,50 EUR' in body)
        check('плейсхолдеров не осталось', '{{' not in body)


def test_guards():
    print('\nОшибки и защита')
    with tempfile.TemporaryDirectory() as tmp:
        junk = Path(tmp) / 'junk.pdf'
        junk.write_bytes(b'%PDF-1.4 not really a facture')
        try:
            pf.parse_facture(junk)
            check('битый PDF отвергается', False, 'исключения не было')
        except Exception as e:
            check('битый PDF отвергается', True, str(e))

        try:
            pf.build_proforma(Path(tmp) / 'x.docx', {
                'pf_num': 'PF1', 'date': 'd', 'termin': 't',
                'buyer_name': 'n', 'buyer_info': 'i', 'vehicles': []})
            check('пустой список машин отвергается', False, 'исключения не было')
        except ValueError:
            check('пустой список машин отвергается', True)

    check('шаблон на месте', pf.TEMPLATE.exists(), str(pf.TEMPLATE))
    check('имя файла', pf.out_filename('PF26-8/106') == 'Faktura Pro forma PF26-8-106.docx',
          pf.out_filename('PF26-8/106'))


def main():
    print('Проверка Facture VO → Pro Forma')
    if not SAMPLE_PDF.exists():
        # Образцы — это счета с аукционов: настоящие машины, суммы и
        # покупатели. В открытом репозитории им не место.
        print('\n⏭  Пропущено: нет samples/ — положите свои PDF, '
              'и проверки заработают.')
        return 0
    test_money()
    head = test_parse()
    test_parse_variants()
    test_sag()
    test_openlane()
    test_openlane_languages()
    test_unreadable_message()
    test_cn_by_capacity()
    test_real_pdfs()
    test_plate_line_dropped()
    test_template_shape()
    test_roundtrip(head)
    test_multi(head)
    test_output_is_valid()
    test_refs_memory()
    test_guards()
    print('\n' + ('✅ ВСЁ ПРОШЛО' if not failures
                  else '❌ ПРОВАЛЕНО %d: %s' % (len(failures), '; '.join(failures))))
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
