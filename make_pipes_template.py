#!/usr/bin/env python3
"""Turn a filled-in pipes Pro Forma into templates/proforma_pipes_template.docx.

Run once (kept in the repo so the template can be rebuilt if the layout of the
sample ever changes):

    python make_pipes_template.py "Faktura Pro Forma PF26-9-110.docx"

This is the GOLFSTREAM line — sanitary ware, ~40 positions, a different layout
from the car Pro Forma, which is why it needs a template of its own. Everything
that varies between invoices becomes a {{PLACEHOLDER}}; the buyer block and the
Uwagi block do not vary at all on this line, so they stay as they are.

The goods table keeps its header, ONE sample item row and the Razem row; the
other 37 item rows are dropped, because the builder clones that one row as many
times as the specification needs.
"""

import sys
from copy import deepcopy
from pathlib import Path

from docx import Document

from docx_text import replace_everywhere, replace_in_cell, doc_tables

# Sample values from PF26-9/110 → placeholder. Order matters: the longer,
# more specific strings go first so they can't be eaten by a shorter one.
SIMPLE = [
    ('PF26-9/110',            '{{PF_NUM}}'),
    ('Warszawa, 2026-09-03',  'Warszawa, {{DATA}}'),
    ('2026-09-28',            '{{TERMIN}}'),
]

# The item row. Columns 6 and 7 (VAT % and VAT amount) are a literal 0 on this
# line and stay untouched. The description is not spelled out here: it carries a
# Cyrillic 'х' inside "125х2000", so it is taken from the cell itself rather
# than retyped — a mistyped homoglyph would fail silently.
ITEM = 1
ITEM_CELLS = [
    (0, '1',        '{{LP}}'),
    (1, '116060',   '{{ARTIKEL}}'),
    (3, '54',       '{{QTY}}'),
    (4, '6,05',     '{{CENA}}'),
    (5, '326,70',   '{{WARTOSC}}'),
    (8, '326,70',   '{{BRUTTO}}'),
]

# The Razem row. Quantity and money are grouped there ("11 135", "27 827,49")
# while the item rows are not — that difference is part of the layout.
TOTAL_CELLS = [
    (3, '11 135',    '{{RAZEM_QTY}}'),
    (5, '27 827,49', '{{RAZEM}}'),
    (8, '27 827,49', '{{RAZEM}}'),
]

# The last line of the document: specification number, a tab, and the buyer's
# order code. Matched together so a bare "126" elsewhere can never be hit.
FOOTER_LINE = ('126\t1139034', '{{SPEC}}\t{{CODE}}')

# The two summary lines under the table. By the time this runs the table cells
# are already placeholders, so the amount is unambiguous — and matching just the
# amount leaves the trailing " EUR", which is formatted separately, untouched.
SUMMARY_AMOUNT = ('27 827,49', '{{RAZEM}}')

# The Uwagi block, as it has to stand on every invoice of this line: the WDT
# anti-circumvention clause that the goods cannot travel without. It replaces
# what PF26-9/110 carried (a "Termin dostawy / Sale term" line), so it is
# written here rather than lifted from the sample. One paragraph, line breaks
# inside it — the same shape the sample uses; True means bold.
UWAGI = [
    [('Uwagi / Attention:', True)],
    [('Warunki dostawy / Delivery terms: ', True), ('FCA Warszaw Poland ', False)],
    [('Wewnątrzwspólnotowa dostawa towarów (WDT)', True)],
    [('(1) [Importer/Buyer] shall not sell, export or re-export - either '
      'directly or indirectly - to the Russian Federation or for use in the '
      'Russian Federation; or to Belarus or for use in Belarus goods supplied '
      'under or in connection with this Agreement. The Agreement shall apply '
      'to:', False)],
    [('- goods identified in Article 12g of Council Regulation (EU) '
      'No 833/2014; and', False)],
    [('- goods identified in Article 8g of Council Regulation (EC) '
      'No 765/2006.', False)],
    [('(2) The [Importer/Buyer] shall use its best endeavours to ensure that '
      'the purpose of paragraph (1) is not frustrated by third parties further '
      'down the trade chain, including any resellers.', False)],
    [('(3) The [Importer/Buyer] shall establish and maintain an appropriate '
      'control mechanism to detect conduct by third parties down the supply '
      'chain that would frustrate the purpose of paragraph 1.', False)],
    [('(4) Any violation of paragraphs (1), (2) or (3) shall constitute a '
      'material breach of an essential element of this Agreement and [Exporter] '
      'shall have the right to appropriate remedies, including:', False)],
    [('- termination of this Agreement; and', False)],
    [('- an appropriate penalty of 10% of the total value of this Agreement or '
      'of the price of the exported goods, whichever is greater.', False)],
    [('(5) The [Importer/Buyer] shall promptly inform the [Exporter/Seller] of '
      'any problems with the application of paragraphs (1), (2) or (3), '
      'including any relevant activities of third parties that may frustrate '
      'the purpose of paragraph (1).', False)],
    [('The [Importer/Buyer] shall provide the [Exporter/Shipper] with '
      'information relevant to the fulfilment of the obligations under '
      'paragraphs 1, 2 and 3 within two weeks of the simple request for such '
      'information.\n', False)],
]


def _rewrite_uwagi(doc, lines):
    """Replace the Uwagi paragraph with `lines`, keeping its look.

    The run properties are not rebuilt from scratch — they are copied off the
    runs already in the paragraph (one bold, one not). That carries the font
    size and the language of the original along with them, which is the whole
    reason the template is the customer's own file."""
    par = None
    for candidate in doc.paragraphs:
        if candidate.text.strip().startswith('Uwagi'):
            par = candidate
            break
    if par is None:
        raise SystemExit('в образце не найден блок Uwagi')

    def rpr_of(bold):
        for run in par.runs:
            if bool(run.bold) == bold and run._r.find(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr'
            ) is not None:
                return deepcopy(run._r.rPr)
        raise SystemExit('в блоке Uwagi нет прогона с оформлением bold=%s' % bold)

    styles = {True: rpr_of(True), False: rpr_of(False)}
    for run in list(par.runs):
        run._r.getparent().remove(run._r)
    for i, line in enumerate(lines):
        for j, (text, bold) in enumerate(line):
            run = par.add_run()
            # A leading \n becomes a real <w:br/>, which is how the sample keeps
            # the whole block inside one paragraph.
            run.text = ('\n' if (i and not j) else '') + text
            run._r.insert(0, deepcopy(styles[bold]))


def build(src: str, dest: str):
    doc = Document(src)

    for old, new in SIMPLE:
        if not replace_everywhere(doc, old, new):
            raise SystemExit(f'не найдено в образце: {old!r}')

    tables = doc_tables(doc)
    if not tables:
        raise SystemExit('в образце нет таблицы с товаром')
    goods = tables[0]
    if len(goods.rows) < 3:
        raise SystemExit('в таблице образца нет ни шапки, ни товара, ни итога')
    total_row = len(goods.rows) - 1

    # The description, straight from the cell — see the note above. Taken with
    # its trailing space, so the placeholder stands for the whole cell: the
    # builder writes the description exactly as the specification spells it.
    name = goods.rows[ITEM].cells[2].text
    if not name.strip():
        raise SystemExit('в образце пустое наименование товара')
    if not replace_in_cell(goods, ITEM, 2, name, '{{NAZWA}}'):
        raise SystemExit(f'не найдено наименование: {name!r}')

    for col, old, new in ITEM_CELLS:
        if not replace_in_cell(goods, ITEM, col, old, new):
            raise SystemExit(f'не найдено в ячейке товара R{ITEM}C{col}: {old!r}')
    for col, old, new in TOTAL_CELLS:
        if not replace_in_cell(goods, total_row, col, old, new):
            raise SystemExit(f'не найдено в строке итога R{total_row}C{col}: {old!r}')

    # Drop every item row but the first — the builder clones it.
    for row in list(goods.rows[ITEM + 1:total_row]):
        row._tr.getparent().remove(row._tr)

    old, new = FOOTER_LINE
    if replace_everywhere(doc, old, new) != 1:
        raise SystemExit(f'нижняя строка не найдена или найдена не один раз: {old!r}')

    old, new = SUMMARY_AMOUNT
    n = replace_everywhere(doc, old, new)
    if n != 2:
        raise SystemExit(f'итоговых строк найдено {n}, ожидалось 2')

    _rewrite_uwagi(doc, UWAGI)

    Path(dest).parent.mkdir(parents=True, exist_ok=True)
    doc.save(dest)
    print(f'шаблон записан: {dest}')


if __name__ == '__main__':
    sample = sys.argv[1] if len(sys.argv) > 1 else \
        str(Path.home() / 'Desktop' / 'Faktura Pro Forma PF26-9-110.docx')
    build(sample, str(Path(__file__).parent / 'templates' / 'proforma_pipes_template.docx'))
