#!/usr/bin/env python3
"""Facture VO (the Ayvens / TEMSYS PDF) → Faktura Pro Forma (DOCX).

The French invoice already carries everything that identifies the car — model,
plate, VIN, first registration. What it does not carry is the price AOW resells
at, the Pro Forma number and its date; those three come from the user. The
buyer is asked for as well, since a Facture VO says nothing about who the car
is going to.
"""

import re
import json
import logging
from copy import deepcopy
from pathlib import Path

from docx import Document

from docx_text import doc_tables, iter_paragraphs, replace_in_paragraph

BASE_DIR = Path(__file__).parent
TEMPLATE = BASE_DIR / 'templates' / 'proforma_template.docx'
REFS_FILE = BASE_DIR / 'proforma_refs.json'

# Customs code and ground clearance are not in the Facture VO — they depend on
# engine and body. These are the values from PF26-8/106; whenever the user
# corrects them for a model, the correction is remembered in REFS_FILE.
DEFAULT_CN = '87032290'
DEFAULT_CLEARANCE = '145'


# -- Reading the Facture VO ---------------------------------------------------
def read_pdf_text(path) -> str:
    from pypdf import PdfReader
    reader = PdfReader(str(path))
    return '\n'.join(page.extract_text() or '' for page in reader.pages)


# -- Shared helpers -----------------------------------------------------------
NBSP = ' '


def _tidy(s) -> str:
    return re.sub(r'\s{2,}', ' ', (s or '').replace(NBSP, ' ')).strip(' ,;-')


def _as_date(raw) -> str:
    """Every platform writes the registration date differently — 23/12/2020,
    14/03/2022, 29-10-2021. The Pro Forma always shows DD/MM/YYYY."""
    m = re.search(r'(\d{1,2})[-./](\d{1,2})[-./](\d{4})', raw or '')
    if not m:
        return ''
    d, mo, y = m.groups()
    return '%02d/%02d/%s' % (int(d), int(mo), y)


def _digits(raw) -> str:
    return re.sub(r'\D', '', raw or '')


def _money_text(raw) -> str:
    return re.sub(r'[\s' + NBSP + r']', '', raw or '')


def _cn_for(capacity) -> str:
    """Customs code from engine capacity, for the petrol brackets of heading
    8703 — which is what every Pro Forma in the archive uses. Ayvens states no
    capacity, so there the default simply stands."""
    try:
        cc = int(capacity)
    except (TypeError, ValueError):
        return DEFAULT_CN
    if cc <= 1000:
        return '87032190'
    if cc <= 1500:
        return '87032290'
    if cc <= 3000:
        return '87032390'
    return '87032490'


def _model_name(raw) -> str:
    """One house style for every platform: 'PEUGEOT - 5008' and 'Citroën C3
    Aircross SHINE 1.2' both end up looking like Ayvens' 'PEUGEOT 308 VP'."""
    return re.sub(r'\s+-\s+', ' ', _tidy(raw)).upper()


def _car(source, model, plate='', reg_date='', vin='', km='', capacity=''):
    v = {
        'source': source,
        'model': _model_name(model),
        'plate': _tidy(plate),
        'reg_date': _as_date(reg_date),
        'vin': _tidy(vin).upper(),
        'km': _digits(km),
        'capacity': _digits(capacity),
        'cn': _cn_for(_digits(capacity)),
        'clearance': DEFAULT_CLEARANCE,
        'price': None,
    }
    _apply_remembered_refs(v)
    return v


def _blocks(text, marker):
    """One chunk per car: from each `marker` hit to the next."""
    hits = list(marker.finditer(text))
    return [(m, text[m.start():(hits[i + 1].start() if i + 1 < len(hits) else len(text))])
            for i, m in enumerate(hits)]


def _grab(rx, block, default=''):
    hit = rx.search(block)
    return hit.group(1).strip() if hit else default


# -- Ayvens / TEMSYS — "Facture VO", French ------------------------------------
_AV_MARK = re.compile(r'^\s*Avenant\s*n[°ºo]?\s*\S+\s+(.+?)\s*$',
                      re.IGNORECASE | re.MULTILINE)
_AV_PLATE = re.compile(r'Immatricul\w*\s+([A-Z0-9]{2,3}-[A-Z0-9]{2,3}-[A-Z0-9]{2,3}'
                       r'|[A-Z]{1,3}[- ]?\d{2,4}[- ]?[A-Z]{1,3})')
_AV_REG = re.compile(r'mise\s+en\s+circulation\s+le\s+(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
_AV_VIN = re.compile(r'ch[aâ]ssis\s+([A-HJ-NPR-Z0-9]{15,19})', re.IGNORECASE)
_AV_KM = re.compile(r'Kms?\s*Compteur\s+([\d ]+)', re.IGNORECASE)
_AV_NO = re.compile(r'Facture\s+VO\s+N[°ºo]?\s*\n?\s*(\d{6,})', re.IGNORECASE)
_AV_DATE = re.compile(r'\bDate\s*:\s*(\d{2}/\d{2}/\d{4})')
_AV_TOTAL = re.compile(r'TOTAL\s+EN\s*€?\s*([\d ' + NBSP + r']+[.,]\d{2})',
                       re.IGNORECASE)
# A trailing "3888.33 - - -" sometimes lands on the model line, depending on
# how the extractor orders the columns.
_AV_AMOUNT_TAIL = re.compile(r'\s+[\d ]+[.,]\d{2}(?:\s*-)*\s*$')


def _parse_ayvens(text):
    cars = [_car('Ayvens',
                 model=_AV_AMOUNT_TAIL.sub('', m.group(1)),
                 plate=_grab(_AV_PLATE, block),
                 reg_date=_grab(_AV_REG, block),
                 vin=_grab(_AV_VIN, block),
                 km=_grab(_AV_KM, block))
            for m, block in _blocks(text, _AV_MARK)]
    return {
        'facture_no': _grab(_AV_NO, text),
        'facture_date': _grab(_AV_DATE, text),
        'facture_total': _money_text(_grab(_AV_TOTAL, text)),
        'vehicles': cars,
    }


# -- SAG — "INVOICE SAG…", English ---------------------------------------------
_SG_MARK = re.compile(r'^\s*Plate\s+([A-Z0-9-]{4,12})\s*$', re.MULTILINE)
_SG_REG = re.compile(r'Registration date\s+([\d/.-]{8,10})')
_SG_MODEL = re.compile(r'Make and model\s+(.+)')
_SG_VIN = re.compile(r'Chassis\s+([A-HJ-NPR-Z0-9]{15,19})')
_SG_KM = re.compile(r'Last mileage\s+([\d ]+)')
_SG_CC = re.compile(r'Capacity\s+(\d{3,5})')
_SG_NO = re.compile(r'INVOICE\s+(SAG\S+)', re.IGNORECASE)
_SG_DATE = re.compile(r'Due date\s*\n\s*(\d{2}/\d{2}/\d{4})')
_SG_TOTAL = re.compile(r'Total to pay\s*\(EUR\)\s*([\d,. ]+\.\d{2})')


def _parse_sag(text):
    cars = []
    for m, block in _blocks(text, _SG_MARK):
        # SAG lists the vehicle details as their own lines, and the extractor
        # does not always keep them below the plate — so a field missing from
        # the block is looked up on the page.
        def field(rx):
            return _grab(rx, block) or _grab(rx, text)
        cars.append(_car('SAG',
                         model=field(_SG_MODEL),
                         plate=m.group(1),
                         reg_date=field(_SG_REG),
                         vin=field(_SG_VIN),
                         km=field(_SG_KM),
                         capacity=field(_SG_CC)))
    return {
        'facture_no': _grab(_SG_NO, text),
        'facture_date': _grab(_SG_DATE, text),
        'facture_total': _money_text(_grab(_SG_TOTAL, text)),
        'vehicles': cars,
    }


# -- OPENLANE — "Pro Forma …", in the buyer's own language ---------------------
# This one never states a plate: the cars come straight off the auction. It
# also issues each invoice in whatever language the buyer picked, so the labels
# are matched in every language AOW has seen plus the obvious neighbours — and
# the extractor pads label words with stray spaces ("Первая  регистрация"), so
# every gap inside a label has to be \s+ rather than a single space.
# Spaces but never a line break: "Тип" is left empty on some invoices, and a
# plain \s+ would happily jump the newline and swallow the next line's value.
_H = r'[^\S\r\n]'


def _label(*words):
    return r'(?:%s)' % '|'.join(w.replace(' ', _H + '+') for w in words)


def _ol_line(*words):
    return re.compile(r'^' + _H + r'*' + _label(*words) + _H + r'+(.+?)' + _H + r'*$',
                      re.MULTILINE | re.IGNORECASE)


def _ol_value(pattern, *words):
    return re.compile(_label(*words) + _H + r'*:?' + _H + r'+' + pattern,
                      re.IGNORECASE)


_OL_MARK = _ol_line('Marka', 'Марка', 'Make', 'Merk', 'Marque', 'Marke', 'Marca')
_OL_MODEL = _ol_line('Model', 'Модель', 'Modèle', 'Modell', 'Modelo', 'Modello')
_OL_TYPE = _ol_line('Typ', 'Тип', 'Type', 'Tipo')
_OL_REG = _ol_value(r'([\d/.-]{8,10})',
                    'Pierwsza rejestracja', 'Первая регистрация',
                    'First registration', 'Eerste inschrijving',
                    'Première immatriculation', 'Erstzulassung')
_OL_VIN = _ol_value(r'([A-HJ-NPR-Z0-9]{15,19})',
                    'Podwozie', 'Шасси', 'Chassis', 'Fahrgestell', 'Chasis')
_OL_KM = _ol_value(r'([\d ]+)',
                   'Przebieg', 'Пробег', 'Mileage', 'Kilometerstand',
                   'Kilométrage', 'Laufleistung')
_OL_CC = _ol_value(r'(\d{3,5})',
                   'Rozmiar silnika', 'Размер двигателя', 'Engine size',
                   'Cilinderinhoud', 'Cylindrée', 'Motorgröße', 'Hubraum')
_OL_NO = _ol_value(r'(\d{4,6}-\d+)', 'Pro Forma', 'Проформа', 'Proforma')
_OL_DATE = _ol_value(r'([\d/.-]{8,10})',
                     'Data faktury', 'Дата выставления счета', 'Invoice date',
                     'Factuurdatum', 'Date de facture', 'Rechnungsdatum')
_OL_TOTAL = _ol_value(r'([\d ' + NBSP + r']+,\d{2})',
                      'Razem', 'На сумму', 'Total', 'Totaal', 'Gesamt', 'Totale')


def _parse_openlane(text):
    cars = []
    for m, block in _blocks(text, _OL_MARK):
        def field(rx):
            return _grab(rx, block) or _grab(rx, text)
        name = ' '.join(p for p in (m.group(1), field(_OL_MODEL), field(_OL_TYPE)) if p)
        cars.append(_car('OPENLANE',
                         model=name,
                         reg_date=field(_OL_REG),
                         vin=field(_OL_VIN),
                         km=field(_OL_KM),
                         capacity=field(_OL_CC)))
    return {
        'facture_no': _grab(_OL_NO, text),
        'facture_date': _as_date(_grab(_OL_DATE, text)),
        'facture_total': _money_text(_grab(_OL_TOTAL, text)),
        'vehicles': cars,
    }


# Display name, how to recognise the platform, how to read it, and which
# identifying fields that platform actually supplies.
SOURCES = [
    ('Ayvens', re.compile(r'Facture\s+VO|Avenant\s*n[°ºo]', re.IGNORECASE),
     _parse_ayvens, ('model', 'plate', 'reg_date', 'vin')),
    ('SAG', re.compile(r'INVOICE\s+SAG|Make and model', re.IGNORECASE),
     _parse_sag, ('model', 'plate', 'reg_date', 'vin')),
    ('OPENLANE', re.compile(r'OPENLANE|Podwozie|Шасси|Rozmiar\s+silnika'
                            r'|Размер\s+двигателя', re.IGNORECASE),
     _parse_openlane, ('model', 'reg_date', 'vin')),
]

SOURCE_FIELDS = dict((name, fields) for name, _, _, fields in SOURCES)
SOURCE_NAMES = [name for name, _, _, _ in SOURCES]


def parse_facture(pdf_path, log=None):
    """Read a car purchase invoice from any platform the bot knows.

    Raises ValueError when the file is not one of them, so the bot can say so
    instead of producing a Pro Forma full of blanks.
    """
    def note(m):
        if log:
            log(m)

    text = read_pdf_text(pdf_path)
    if not text.strip():
        raise ValueError('в PDF нет текстового слоя — похоже, это скан')

    recognised = []
    for name, marker, parse, _ in SOURCES:
        if not marker.search(text):
            continue
        recognised.append(name)
        head = parse(text)
        if not head['vehicles']:
            continue
        head['source'] = name
        for v in head['vehicles']:
            note('🚗 {} · {} · VIN {}'.format(
                v['model'] or '?', v['plate'] or 'без госномера', v['vin'] or '?'))
        note('📄 {} №{} от {} — машин: {}'.format(
            name, head['facture_no'] or '?', head['facture_date'] or '?',
            len(head['vehicles'])))
        return head

    # Two different failures, and telling them apart is what makes the message
    # useful: an unknown platform is one thing, a known platform whose vehicle
    # block could not be read is quite another.
    if recognised:
        raise ValueError('это похоже на %s, но данные машины прочитать не вышло '
                         '— возможно, счёт на языке, который бот пока не знает'
                         % ' / '.join(recognised))
    raise ValueError('не удалось распознать платформу — бот понимает '
                     + ', '.join(SOURCE_NAMES))


def missing_fields(v) -> list:
    """Identifying fields the parser could not find, ignoring the ones this
    platform never supplies — OPENLANE, for one, never states a plate."""
    labels = [('model', 'модель'), ('plate', 'госномер'),
              ('reg_date', 'дата регистрации'), ('vin', 'VIN')]
    expected = SOURCE_FIELDS.get(v.get('source'), [k for k, _ in labels])
    return [label for key, label in labels
            if key in expected and not v.get(key)]


# -- Remembered CN code / clearance per model ---------------------------------
def _model_key(model) -> str:
    return re.sub(r'[^A-Z0-9]+', ' ', (model or '').upper()).strip()


def _load_refs() -> dict:
    try:
        return json.loads(REFS_FILE.read_text(encoding='utf-8'))
    except Exception:
        return {}


def _apply_remembered_refs(v):
    saved = _load_refs().get(_model_key(v['model']))
    if saved:
        v['cn'] = saved.get('cn', v['cn'])
        v['clearance'] = saved.get('clearance', v['clearance'])


def remember_refs(model, cn, clearance):
    """Keep a corrected customs code / clearance for the next time this model
    comes through, so it only has to be typed once."""
    key = _model_key(model)
    if not key:
        return
    refs = _load_refs()
    refs[key] = {'cn': cn, 'clearance': clearance}
    try:
        REFS_FILE.write_text(json.dumps(refs, ensure_ascii=False, indent=2),
                             encoding='utf-8')
    except Exception:
        logging.exception('Не удалось сохранить справочник CN/просвета')


# -- Money --------------------------------------------------------------------
def fmt_money(value) -> str:
    """5400 → '5 400,00' — the Polish grouping used throughout the Pro Forma."""
    return '{:,.2f}'.format(float(value)).replace(',', ' ').replace('.', ',')


def parse_money(raw):
    """Accept 5400 / 5 400 / 5400,50 / 5.400,00 / 5,400.00 / '5400 EUR'.
    Returns a float, or None when the text is not a positive amount."""
    s = (raw or '').replace(' ', ' ').strip().upper()
    s = s.replace('EUR', '').replace('€', '')
    s = re.sub(r'\s+', '', s)
    if not re.fullmatch(r'\d[\d.,]*', s or ''):
        return None

    if ',' in s and '.' in s:
        dec = max(s.rfind(','), s.rfind('.'))
        s = re.sub(r'[.,]', '', s[:dec]) + '.' + s[dec + 1:]
    elif ',' in s or '.' in s:
        sep = ',' if ',' in s else '.'
        tail = s.rsplit(sep, 1)[1]
        # One separator with 1-2 trailing digits is a decimal point; anything
        # else ("5.400", "1.234.567") is thousands grouping.
        if s.count(sep) == 1 and len(tail) in (1, 2):
            s = s.replace(sep, '.')
        else:
            s = s.replace(sep, '')
    try:
        value = float(s)
    except ValueError:
        return None
    return value if value > 0 else None


# -- Writing the Pro Forma ----------------------------------------------------
def _fill(paragraphs, mapping):
    for par in paragraphs:
        if '{{' not in par.text:
            continue
        for key, value in mapping.items():
            replace_in_paragraph(par, '{{%s}}' % key, value)


def _row_paragraphs(row):
    for cell in row.cells:
        for par in cell.paragraphs:
            yield par


def _drop_plate_line(row):
    """The goods cell ends with a blank line and the plate, in a paragraph of
    its own. OPENLANE never states a plate, so remove that paragraph outright —
    blanking it would leave an empty line hanging under the customs code."""
    for cell in row.cells:
        for par in list(cell.paragraphs):
            if '{{PLATE}}' not in par.text:
                continue
            alone = par.text.replace('{{PLATE}}', '').strip() == ''
            if alone and len(cell.paragraphs) > 1:
                par._p.getparent().remove(par._p)
                continue
            for pattern in ('\n\n{{PLATE}}', '\n{{PLATE}}', '{{PLATE}}'):
                if replace_in_paragraph(par, pattern, ''):
                    break


def build_proforma(out_path, ctx, template=None, log=None):
    """Write the Pro Forma. `ctx` needs: pf_num, date, termin, buyer_name,
    buyer_info and a non-empty list of vehicles that each carry a price."""
    def note(m):
        if log:
            log(m)

    vehicles = ctx['vehicles']
    if not vehicles:
        raise ValueError('нет ни одной машины для Pro Forma')

    doc = Document(str(template or TEMPLATE))
    tables = doc_tables(doc)
    if not tables:
        raise ValueError('в шаблоне нет таблицы с товаром')
    goods = tables[0]

    ITEM = 1                       # row 0 is the header, row 1 the sample item
    for _ in range(len(vehicles) - 1):
        goods.rows[ITEM]._tr.addnext(deepcopy(goods.rows[ITEM]._tr))

    total = 0.0
    for i, v in enumerate(vehicles):
        price = float(v['price'])
        total += price
        if not v.get('plate'):
            _drop_plate_line(goods.rows[ITEM + i])
        _fill(_row_paragraphs(goods.rows[ITEM + i]), {
            'LP': str(i + 1),
            'AUTO': v['model'],
            'REG': v['reg_date'],
            'VIN': v['vin'],
            'CN': v['cn'],
            'CLEAR': v['clearance'],
            'PLATE': v['plate'],
            'QTY': '1',
            'CENA': fmt_money(price),
            'WARTOSC': fmt_money(price),
            'BRUTTO': fmt_money(price),
        })

    # Everything left: the header block, the Razem row and the summary lines.
    _fill(iter_paragraphs(doc), {
        'PF_NUM': ctx['pf_num'],
        'DATA': ctx['date'],
        'TERMIN': ctx['termin'],
        'BUYER_NAME': ctx['buyer_name'],
        'BUYER_INFO': ctx['buyer_info'],
        'RAZEM': fmt_money(total),
    })

    leftover = sorted({name for par in iter_paragraphs(doc)
                       for name in re.findall(r'\{\{(\w+)\}\}', par.text)})
    if leftover:
        raise ValueError('в шаблоне остались незаполненные поля: ' + ', '.join(leftover))

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    note('✅ Pro Forma {}: позиций {}, итого {} EUR'.format(
        ctx['pf_num'], len(vehicles), fmt_money(total)))
    return {'path': out_path, 'total': total, 'count': len(vehicles)}


def out_filename(pf_num) -> str:
    safe = re.sub(r'[\\/:*?"<>|]', '-', pf_num or '').strip()
    return 'Faktura Pro forma {}.docx'.format(safe)
