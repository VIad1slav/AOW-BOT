#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Checks for the CMR consignment note.

The note is not typed in — its goods block is lifted out of the finished
specification. So the parts worth checking are the seam between the two: does
the footer come across correctly, is the packing wording rewritten the way the
customer writes it, and does the form still carry its constant data afterwards.
"""

import os
import sys
import tempfile
import unicodedata

import cmr

SPEC = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    'samples', 'spec_golfstream.xlsx')

failures = []


def _norm(v):
    """Word stores accented letters decomposed: the "á" of "Geologická" is
    an "a" followed by a separate accent mark. Compare in one normal form so a
    correct document doesn't fail on an invisible difference."""
    return unicodedata.normalize('NFC', v) if isinstance(v, str) else v


def check(name, got, want):
    got, want = _norm(got), _norm(want)
    ok = got == want
    print(('  ✅ ' if ok else '  ❌ ') + name)
    if not ok:
        print(f'       получили: {got!r}')
        print(f'       ожидали:  {want!r}')
        failures.append(name)


def test_places_wording():
    print('\nЗапись мест: из спецификации в накладную')
    check('точка с запятой становится плюсом',
          cmr._fmt_places('27 crates; 8 cartons (1 pallet)', ''),
          '27 crates + 8 cartons (= 1 pallet)')
    check('название груза дописывается в конец',
          cmr._fmt_places('9 cartons (1 pallet)', 'Sealtape'),
          '9 cartons (= 1 pallet) Sealtape')
    check('скобка без числа остаётся как есть',
          cmr._fmt_places('4 cartons (part of pallets)', 'Blade'),
          '4 cartons (part of pallets) Blade')
    check('двойные пробелы схлопываются',
          cmr._fmt_places('1 crate;  2 cartons', ''), '1 crate + 2 cartons')
    check('пустая строка не ломает', cmr._fmt_places('', ''), '')


def test_weight_format():
    print('\nФормат веса, как на бланке')
    check('тысячи через пробел, дробь через запятую',
          cmr._fmt_weight(11225.9), '11 225,90')
    check('меньше тысячи', cmr._fmt_weight(109.2), '109,20')
    check('копейки не теряются', cmr._fmt_weight(0.49), '0,49')
    check('не число — пусто', cmr._fmt_weight(None), '')


def test_file_name():
    print('\nИмя файла: номер инвойса без префикса')
    check('FV26-111', cmr.invoice_short('FV26-111'), '111')
    check('FV26-99', cmr.invoice_short('FV26-99'), '99')
    check('другой год', cmr.invoice_short('FV27-5'), '5')
    check('уже без FV', cmr.invoice_short('26-111'), '111')


def test_goods_names():
    print('\nСправочник названий груза')
    names = cmr.load_goods_names()
    check('код труб', names.get('39172290'), 'PP')
    check('код фитингов', names.get('39174000'), 'Fittings')
    check('незнакомый код — пусто', names.get('00000000', ''), '')


def test_from_specification():
    print('\nГрузовой блок берётся из готовой спецификации')
    if not os.path.exists(SPEC):
        print('  ⏭  пропущено: нет образца спецификации в samples/')
        return
    groups, total, (packs, colli, note) = cmr.groups_from_specs([SPEC])
    check('групп груза', len(groups), 2)
    check('код первой группы', groups[0]['hs'], '39172290')
    check('код второй группы', groups[1]['hs'], '39174000')
    # Вес группы — её собственный итог, а не итог всей спецификации. Диапазон
    # последней группы дотягивается до общей строки внизу листа, и если брать
    # последнее совпадение, вес всей машины попадает в одну строку груза.
    check('вес первой группы', round(groups[0]['gross'], 2), 6547.29)
    check('вес второй группы', round(groups[1]['gross'], 2), 4685.92)
    check('общий вес — сумма групп', round(total, 2), 11233.21)
    check('сумма сходится с итогом спецификации',
          round(groups[0]['gross'] + groups[1]['gross'], 2), 11233.21)
    check('места переведены на английский', 'пал' not in groups[0]['places'], True)


def test_filled_form():
    print('\nЗаполненный бланк')
    if not os.path.exists(SPEC):
        print('  ⏭  пропущено: нет образца спецификации в samples/')
        return
    from docx import Document
    params = {'invoice_num': 'FV26-111', 'date': '29.08.2026'}
    with tempfile.TemporaryDirectory() as tmp:
        dst = os.path.join(tmp, 'CMR SK 111.docx')
        cmr.build_cmr([SPEC], params, dst)
        t = Document(dst).tables[0]

        check('дата приёма груза', t.rows[19].cells[3].text.strip(), '29/08/2026')
        check('дата составления', t.rows[44].cells[16].text.strip(), '29/08/2026')
        check('ссылка на инвойс', t.rows[21].cells[0].text.strip(),
              'Invoice № FV26-111 dtd 29.08.2026')
        check('первая строка груза', t.rows[24].cells[0].text.strip(),
              '27 crates + 8 cartons (= 1 pallet) + 1 transport box (= 1 pallet) PP')
        check('её код ТН ВЭД', t.rows[24].cells[24].text.strip(), '39172290')
        check('её вес', t.rows[24].cells[29].text.strip(), '6 547,29')
        check('вторая строка названа Fittings',
              t.rows[25].cells[0].text.strip().endswith('Fittings'), True)
        check('итоговый вес', t.rows[30].cells[29].text.strip(), '11 233,21')
        check('итоговая строка начинается как надо',
              t.rows[30].cells[0].text.strip().startswith('Total colli:'), True)
        check('лишние строки груза пустые', t.rows[26].cells[0].text.strip(), '')

        print('\n  Постоянные поля бланка не затронуты')
        check('отправитель', t.rows[1].cells[0].text.strip(), 'AOW GROUP SP. z o.o.')
        check('получатель', t.rows[8].cells[0].text.strip(), 'GOLFSTREAM s.r.o.')
        check('место разгрузки', t.rows[13].cells[4].text.strip(),
              'Geologická 34, 821 06 Bratislava')
        check('условия поставки', 'FCA' in t.rows[41].cells[2].text, True)


DOGRUZ = [os.path.join(os.path.dirname(os.path.abspath(__file__)), 'samples', n)
          for n in ('spec_dogruz_1.xlsx', 'spec_dogruz_2.xlsx')]


def test_dogruzy():
    """Одна машина — одна накладная, даже если счетов было несколько."""
    print('\nДогрузы: несколько спецификаций на одну машину')
    if not all(os.path.exists(p) for p in DOGRUZ):
        print('  ⏭  пропущено: нет образцов догрузов в samples/')
        return
    groups, total, (packs, colli, note) = cmr.groups_from_specs(DOGRUZ)

    codes = [g['hs'] for g in groups]
    check('строка на каждый код, без повторов', len(codes), len(set(codes)))
    check('кодов получилось три', len(groups), 3)

    by_code = {g['hs']: g for g in groups}
    # 39174000 есть в обеих спецификациях: 22,25 и 3 207,35
    check('вес общего кода сложен', round(by_code['39174000']['gross'], 2), 3229.60)
    check('код только из второй', round(by_code['73269098']['gross'], 2), 74.91)
    check('общий вес по обеим', round(total, 2), 13018.40)
    check('места сложены', colli, 81)
    check('упаковки сложены', packs, 466)
    check('описания мест объединены',
          by_code['39174000']['places'].count(';') >= 1, True)

    # Ярлыки в этих файлах написаны иначе: «TOTAL GROSS WEIGHT, KG» и
    # «TOTAL COLLI» с одной «i». Раньше из-за этого терялись целые группы.
    check('вес найден несмотря на «, KG» в подписи',
          all(isinstance(g['gross'], float) for g in groups), True)


def main():
    print('Проверка накладной CMR')
    test_dogruzy()
    test_places_wording()
    test_weight_format()
    test_file_name()
    test_goods_names()
    test_from_specification()
    test_filled_form()
    print('\n' + ('✅ ВСЁ ПРОШЛО' if not failures
                  else '❌ ПРОВАЛЕНО %d: %s' % (len(failures), '; '.join(failures))))
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
