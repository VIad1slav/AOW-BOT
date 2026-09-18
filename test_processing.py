#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Checks for reading the supplier specification and turning it into an invoice.

The invoice is the document with money on it, and the specification is what
checks it: every position the Pro Forma carries has to be in the shipment, and
the quantities and prices have to be the shipment's. All of that hangs on
finding the right columns in the supplier's sheet — and the supplier renames
its headers without telling anyone. Spec 81 said "Article", Spec 84 said
"Item", nothing else moved, and the reader went blind: no articles, no
filtering, and FV26-121 went out as a straight copy of the Pro Forma, 192 EUR
over and carrying a position that was not in the truck.

So the checks here are about exactly that seam: which column is which, and what
happens when the sheet cannot be read at all.
"""

import os
import sys
import tempfile

import openpyxl

import processing as pr

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_SPEC = os.path.join(HERE, 'samples', 'spec_item_header_84.xlsx')
SAMPLE_PF = os.path.join(HERE, 'samples', 'proforma_extra_line_111.docx')

failures = []


def check(name, got, want):
    ok = got == want
    print(('  ✅ ' if ok else '  ❌ ') + name)
    if not ok:
        print(f'       получили: {got!r}')
        print(f'       ожидали:  {want!r}')
        failures.append(name)


def raises(name, fn, needle=''):
    try:
        fn()
    except Exception as e:
        if needle and needle not in str(e):
            print('  ❌ ' + name)
            print(f'       ошибка не про то: {e}')
            failures.append(name)
        else:
            print('  ✅ ' + name)
        return
    print('  ❌ ' + name)
    print('       ошибки не было, а должна была быть')
    failures.append(name)


# ── Спецификации, собранные здесь: три раскладки, которые реально приходят ────
ROWS = [
    # артикул, описание, упаковок, штук, нетто, цена
    ('116010',    'HTEM Pipe DN/OD 125х250 mm', 6, 60, 26.22, 1.72),
    ('332000-05', 'Skolan Safe-EM Pipe DN/OD 58х150 mm', 7, 126, 36.04, 0.85),
    ('770340',    'KG2000 EM Pipe SN10 DN/OD 110х1000 mm', 5, 400, 721.42, 3.75),
]


def make_spec(path, article_header='Article', qty_header='Quantity\npcs',
              russian_column=False, totals=True):
    """A sheet shaped like the supplier's.

    `russian_column` gives the layout of the order files ("PL AOW
    Specification 124"), where a Russian description sits between the article
    and the English one and shifts everything right by one column."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['B1'] = 'Specification No 84 dated 18.09.2026 to invoice FV26-121'
    head = ['No', article_header]
    if russian_column:
        head.append('Наименование')
    head += ['Description', 'Producer', 'COO', 'HS code', 'Quantity packs',
             qty_header, 'Net weight, kg', ' Gross weight, kg', 'E-Price',
             'Total EUR']
    hdr = 3
    for c, value in enumerate(head, 1):
        ws.cell(hdr, c, value)
    # Две ловушки рядом с нужной колонкой: обе начинаются со слова Quantity.
    ws.cell(hdr, len(head) + 2, 'Pieces per pal')
    ws.cell(hdr, len(head) + 3, 'Quantity pl')

    for i, (art, desc, packs, qty, net, price) in enumerate(ROWS):
        r = hdr + 1 + i
        vals = [i + 1, art]
        if russian_column:
            vals.append('Труба полипропиленовая')
        vals += [desc, 'Gebr. Ostendorf Kunststoffe GmbH', 'DE', 39172290,
                 packs, qty, net, net + 1, price,
                 round(qty * price, 2) if totals else None]
        for c, value in enumerate(vals, 1):
            if value is not None:
                ws.cell(r, c, value)
        ws.cell(r, len(head) + 3, packs)      # Quantity pl — те самые паллеты
    wb.save(path)
    return path


def test_article_header():
    print('\nКолонка артикула: поставщик переименовывает заголовок')
    with tempfile.TemporaryDirectory() as tmp:
        for header in ('Article', 'Item', 'Artikel', 'Артикул', 'item', 'ITEM'):
            path = make_spec(os.path.join(tmp, f'{header}.xlsx'),
                             article_header=header)
            ws = openpyxl.load_workbook(path, data_only=True).active
            check(f'заголовок «{header}»', pr._find_article_col(ws), (2, 3))


def test_title_does_not_win():
    print('\nЗаголовок таблицы важнее вольного текста над ней')
    with tempfile.TemporaryDirectory() as tmp:
        path = make_spec(os.path.join(tmp, 'title.xlsx'), article_header='Item')
        wb = openpyxl.load_workbook(path)
        # Название листа с «item» внутри — раньше на нём бы всё и остановилось.
        wb.active['B1'] = 'Item list for invoice FV26-121'
        wb.save(path)
        ws = openpyxl.load_workbook(path, data_only=True).active
        check('точное совпадение выигрывает у подстроки',
              pr._find_article_col(ws), (2, 3))


def test_columns_by_header():
    print('\nКолонки ищутся по заголовку, а не по смещению от артикула')
    with tempfile.TemporaryDirectory() as tmp:
        plain = make_spec(os.path.join(tmp, 'plain.xlsx'), article_header='Item')
        arts, items = pr.read_xlsx_spec(plain, lambda m: None)
        check('артикулы прочитаны', sorted(arts), sorted(r[0] for r in ROWS))
        check('описание — английское, а не «Producer»',
              items[0]['desc'], 'HTEM Pipe DN/OD 125х250 mm')
        check('количество — штуки, а не вес нетто', items[0]['qty'], 60)
        check('количество не из «Quantity packs»', items[1]['qty'], 126)
        check('цена', items[2]['price'], 3.75)

        # Раскладка файлов-заказов: лишняя русская колонка сдвигает всё вправо.
        shifted = make_spec(os.path.join(tmp, 'shifted.xlsx'),
                            article_header='Item', russian_column=True)
        arts2, items2 = pr.read_xlsx_spec(shifted, lambda m: None)
        check('сдвинутая раскладка: описание',
              items2[0]['desc'], 'HTEM Pipe DN/OD 125х250 mm')
        check('сдвинутая раскладка: количество', items2[0]['qty'], 60)
        check('сдвинутая раскладка: цена', items2[0]['price'], 1.72)
        check('в обеих раскладках одни и те же артикулы', arts2, arts)

        # Старый заголовок количества — «Quantity\npcs».
        old = make_spec(os.path.join(tmp, 'old.xlsx'), article_header='Article')
        _, items3 = pr.read_xlsx_spec(old, lambda m: None)
        check('«Quantity pcs» тоже количество', items3[0]['qty'], 60)


def test_total_is_computed_when_missing():
    print('\nПустой «Total EUR» — формула без сохранённого результата')
    with tempfile.TemporaryDirectory() as tmp:
        path = make_spec(os.path.join(tmp, 'nototal.xlsx'),
                         article_header='Item', totals=False)
        _, items = pr.read_xlsx_spec(path, lambda m: None)
        check('сумма посчитана из количества и цены',
              items[0]['total'], round(60 * 1.72, 2))
        check('и для последней строки',
              items[2]['total'], round(400 * 3.75, 2))


def test_unreadable_sheet_stops_the_run():
    print('\nНечитаемая спецификация обязана остановить сборку')
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'junk.xlsx')
        wb = openpyxl.Workbook()
        wb.active['A1'] = 'просто какой-то текст'
        wb.active['A3'] = 'No'
        wb.save(path)
        # Раньше здесь возвращался пустой набор, фильтрация молча выключалась,
        # и фактура уходила непроверенной копией Pro Forma.
        raises('нет колонки артикула — отказ',
               lambda: pr.read_xlsx_spec(path, lambda m: None),
               'не найдена колонка артикула')
        raises('  в сообщении названы найденные заголовки',
               lambda: pr.read_xlsx_spec(path, lambda m: None), 'no')


def test_invoice_matches_specification():
    """FV26-121 целиком: Pro Forma с лишней позицией + спецификация рейса.

    Ожидаемые числа взяты из самой спецификации (80 позиций, 14 837 штук,
    37 073,06 EUR), а не из вывода бота: именно бот здесь и ошибся."""
    print('\nФактура против спецификации (PF26-9/111 + Spec 84)')
    if not (os.path.exists(SAMPLE_SPEC) and os.path.exists(SAMPLE_PF)):
        print('  ⏭  пропущено: нет образцов в samples/')
        return
    from docx import Document

    logs = []
    arts, items = pr.read_xlsx_spec(SAMPLE_SPEC, logs.append)
    check('позиций в спецификации', len(items), 80)
    check('штук в спецификации', sum(i['qty'] for i in items), 14837)
    check('сумма спецификации',
          round(sum(i['total'] for i in items), 2), 37073.06)
    check('лишней позиции в спецификации нет', '115210' in arts, False)

    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, 'FV26-121.docx')
        pr.process_docx(SAMPLE_PF, out, {
            'invoice_num': 'FV26-121', 'date': '18.09.2026',
            'pf_ref': 'PF26-9/111', 'spec_num': '84', '_xlsx_items': items,
        }, arts, logs.append)

        table = Document(out).tables[0]
        rows = [r for r in table.rows[1:]
                if 'Razem' not in ' '.join(c.text for c in r.cells)]
        check('позиций в фактуре', len(rows), 80)
        check('позиция не из этой поставки удалена',
              [r.cells[1].text.strip() for r in rows
               if r.cells[1].text.strip() == '115210'], [])
        check('итого штук', table.rows[-1].cells[3].text.strip(), '14 837')
        check('итого денег', table.rows[-1].cells[5].text.strip(), '37 073,06')
        check('брутто в итоге', table.rows[-1].cells[8].text.strip(), '37 073,06')
        check('и про удаление сказано в логе',
              any('115210' in l and 'Удалена' in l for l in logs), True)


def main():
    print('Спецификация → фактура — проверки')
    test_article_header()
    test_title_does_not_win()
    test_columns_by_header()
    test_total_is_computed_when_missing()
    test_unreadable_sheet_stops_the_run()
    test_invoice_matches_specification()
    print('\n' + ('✅ ВСЁ ПРОШЛО' if not failures
                  else '❌ ПРОВАЛЕНО %d: %s' % (len(failures), '; '.join(failures))))
    return 1 if failures else 0


if __name__ == '__main__':
    sys.path.insert(0, HERE)
    sys.exit(main())
