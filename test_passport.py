#!/usr/bin/env python3
"""Checks for the passport MRZ reader.

Built on the ICAO 9303 specimen (ANNA MARIA ERIKSSON) rather than on a real
passport, so no live document data sits in the repository. The point of these
checks is the check digits: a scanner that misreads one character has to be
caught here, not on a customs form.
"""

import sys

import passport as ps

# The specimen printed in ICAO Doc 9303 Part 3.
SPECIMEN = ('P<UTOERIKSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<\n'
            'L898902C36UTO7408122F1204159ZE184226B<<<<<10')

failures = []


def check(name, got, want):
    ok = got == want
    print(('  ✅ ' if ok else '  ❌ ') + name
          + ('' if ok else '  — получено %r, ожидалось %r' % (got, want)))
    if not ok:
        failures.append(name)


def test_specimen():
    print('\nЭталон ICAO 9303')
    d = ps.parse_mrz(SPECIMEN)
    check('фамилия', d['surname'], 'ERIKSSON')
    check('имена', d['given_names'], 'ANNA MARIA')
    check('имя целиком', d['name'], 'ERIKSSON ANNA MARIA')
    check('номер документа', d['number'], 'L898902C3')
    check('гражданство', d['nationality'], 'UTO')
    check('дата рождения', d['birth_date'], '12.08.1974')
    check('пол', d['sex'], 'F')
    check('срок действия', d['expiry_date'], '15.04.2012')
    check('личный номер', d['personal_number'], 'ZE184226B')
    check('все контрольные цифры сошлись', d['ok'], True)
    check('нет провалившихся полей', d['failed'], [])


def test_check_digits():
    print('\nКонтрольные цифры ловят подмену')
    # Swap one digit of the document number: its own check and the composite
    # one must both object.
    broken = SPECIMEN.replace('L898902C36', 'L898902C46')
    d = ps.parse_mrz(broken)
    check('номер отмечен как неверный', d['checks']['number'], False)
    check('общая сумма тоже не сходится', d['checks']['composite'], False)
    check('дата рождения при этом в порядке', d['checks']['birth'], True)
    check('ok = False', d['ok'], False)
    check('провалившиеся поля названы', d['failed'], ['composite', 'number'])

    broken = SPECIMEN.replace('7408122', '7408132')
    d = ps.parse_mrz(broken)
    check('подмена даты рождения ловится', d['checks']['birth'], False)

    check('счёт по весам 7-3-1', ps.check_digit('L898902C3'), '6')
    check('  фильтр < считается нулём', ps.check_digit('<<<<<'), '0')


def test_noise():
    print('\nMRZ среди мусора от распознавания')
    noisy = ('REPUBLIC OF BELARUS\n'
             'PASSPORT No MP1234567\n'
             'some garbage line\n'
             + SPECIMEN + '\n'
             'trailing noise')
    d = ps.parse_mrz(noisy)
    check('строки найдены среди лишнего текста', d['number'], 'L898902C3')

    spaced = SPECIMEN.replace('<', ' < ')
    d = ps.parse_mrz(spaced)
    check('пробелы внутри строк не мешают', d['name'], 'ERIKSSON ANNA MARIA')

    lower = SPECIMEN.lower()
    d = ps.parse_mrz(lower)
    check('нижний регистр приводится к верхнему', d['surname'], 'ERIKSSON')

    for bad, why in [('', 'пустой текст'),
                     ('просто фотография без MRZ', 'нет MRZ'),
                     ('P<UTOERIKSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<', 'только одна строка')]:
        try:
            ps.parse_mrz(bad)
            check('отвергается: ' + why, False, True)
        except ValueError:
            check('отвергается: ' + why, True, True)


def test_ocr_confusions():
    print('\nПочинка ошибок распознавания')
    # Exactly what rapidocr returned on this VM: O read as zero in the
    # nationality, I read as one in the surname. Neither field has a check
    # digit, so without repair both would reach the invoice unnoticed.
    seen = ('P<UTOER1KSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<\n'
            'L898902C36UT07408122F1204159ZE184226B<<<<<10')
    d = ps.parse_mrz(seen)
    check('«ER1KSSON» починено в «ERIKSSON»', d['surname'], 'ERIKSSON')
    check('«UT0» починено в «UTO»', d['nationality'], 'UTO')
    check('данные при этом целы', d['number'], 'L898902C3')
    check('контрольные цифры сходятся', d['ok'], True)

    # Digits misread as letters inside the dates.
    seen = SPECIMEN.replace('7408122F', '74O8I22F')
    d = ps.parse_mrz(seen)
    check('«74O8I22» починено в дату', d['birth_date'], '12.08.1974')
    check('  и контрольная цифра сошлась', d['checks']['birth'], True)

    # A genuinely wrong digit must still be refused, not "repaired".
    d = ps.parse_mrz(SPECIMEN.replace('L898902C36', 'L898902C76'))
    check('настоящая ошибка не маскируется', d['checks']['number'], False)
    check('  и общая сумма тоже', d['ok'], False)

    check('буквы в цифровое поле', ps._as_digits('74O8I22'), '7408122')
    check('цифры в буквенное поле', ps._as_alpha('ER1KSS0N'), 'ERIKSSON')


def test_mangled_filler():
    print('\nИспорченный хвост первой строки')
    # What the light OCR path really returns: it does not know the '<' glyph
    # and renders the padding as Λ / A / 八. The non-ASCII ones are stripped,
    # which leaves the line far shorter than 44 characters.
    seen = ('P<BLRIVANOU<<SIARHEI<<<<<<<<<<<<<<<<<<<<<<<<ΛΛΛΛΛΛAΛΛΛAAΛ\n'
            'MP12345677BLR8001014M30010190010199A001PB272')
    d = ps.parse_mrz(seen)
    check('короткая первая строка всё равно находится', d['surname'], 'IVANOU')
    check('мусор из хвоста не попал в имя', d['given_names'], 'SIARHEI')
    check('вторая строка цела', d['number'], 'MP1234567')
    check('контрольные цифры сошлись', d['ok'], True)
    check('чтение признано годным', ps.looks_read(d), True)

    # How much of the padding survives varies between passes: here one '<'
    # is left, there none. Both must yield the same name, or the consensus
    # check tears itself apart on its own noise.
    one_sep = ('P<BLRPETROVA<<ALESIA<<<<<<<<<<<<<<<<<<<<<<<<ΛΛΛΛAΛAΛAAAA\n'
               'MP76543215BLR9002155F90021559002159A002PB352')
    none_sep = ('P<BLRPETROVA<<ALESIA<<<<<<<<<<<<<<<<<<<<<<<<ΛΛΛΛ\n'
                'MP76543215BLR9002155F90021559002159A002PB352')
    check('хвостовой мусор отброшен', ps.parse_mrz(one_sep)['name'],
          'PETROVA ALESIA')
    check('  и оба прохода дают одно имя',
          ps.parse_mrz(one_sep)['name'] == ps.parse_mrz(none_sep)['name'], True)
    check('составное имя не пострадало',
          ps.parse_mrz(SPECIMEN)['given_names'], 'ANNA MARIA')

    # Rows of junk between the two lines must not break the pairing.
    noisy = ('P<BLRIVANOU<<SIARHEI<<<<<<<<<<<<<<<<<<<<<<<<ΛΛΛΛ\n'
             '###\n'
             'MP12345677BLR8001014M30010190010199A001PB272')
    check('мусорная строка между MRZ не мешает',
          ps.parse_mrz(noisy)['number'], 'MP1234567')


def test_looks_read():
    print('\nКогда лёгкому пути верить нельзя')
    good = ps.parse_mrz(SPECIMEN)
    check('чистое чтение принимается', ps.looks_read(good), True)

    # Digits the letter-repair map does not cover (3, 4, 7, 9) survive in the
    # name and give the garbage away.
    mangled = ps.parse_mrz('P<UTOER9KSS4N<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<\n'
                           'L898902C36UTO7408122F1204159ZE184226B<<<<<10')
    check('  но имя с уцелевшими цифрами — нет', ps.looks_read(mangled), False)

    # Honest limit: a short mangled name that still reads like a name cannot be
    # told apart automatically — the check digits never covered the name field.
    # This is why the user always sees and edits the draft before it is used.
    plausible = ps.parse_mrz('P<UTO0U<<V1TAR<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<\n'
                             'L898902C36UTO7408122F1204159ZE184226B<<<<<10')
    check('правдоподобно испорченное имя автоматом не ловится',
          ps.looks_read(plausible), True)

    broken = ps.parse_mrz(SPECIMEN.replace('L898902C36', 'L898902C46'))
    check('несошедшаяся контрольная цифра — нет', ps.looks_read(broken), False)

    check('пустое имя — нет', ps.looks_read({'ok': True, 'surname': '',
                                             'given_names': 'X'}), False)
    check('цифры в имени — нет', ps.looks_read({'ok': True, 'surname': 'AB1',
                                                'given_names': 'CD'}), False)
    check('составное имя принимается',
          ps.looks_read({'ok': True, 'surname': 'ERIKSSON',
                         'given_names': 'ANNA MARIA'}), True)


REAL_MRZ = ('P<BLRPETROVA<<ALESIA<<<<<<<<<<<<<<<<<<<<<<<<\n'
            'MP76543215BLR9002155F90021559002159A002PB352')


def test_name_consensus():
    print('\nСверка имени двумя проходами')
    import passport

    calls = []

    def fake(readings):
        """Stand in for the OCR worker: hand back a prepared reading per call."""
        def _read(path, full=False, alt=False):
            key = 'full' if full else ('alt' if alt else 'light')
            calls.append(key)
            return readings[key]
        return _read

    original = passport.read_photo
    try:
        # Both light passes agree — no need to wake the detector at all.
        passport.read_photo = fake({'light': REAL_MRZ, 'alt': REAL_MRZ,
                                    'full': 'НЕ ДОЛЖНО ВЫЗЫВАТЬСЯ'})
        del calls[:]
        d = passport.read_passport_photo('x.jpg')
        check('согласные проходы дают имя', d['name'], 'PETROVA ALESIA')
        check('  имя помечено подтверждённым', d['name_confirmed'], True)
        check('  тяжёлый путь не запускался', 'full' in calls, False)

        # The real failure: the number survives its check digit, the name does
        # not survive at all. The passes disagree, so the detector is called.
        mangled = ('P<BLRKNRTSEZK<<NN<<<<<<<<<<<<<<<<<<<<<<<<<<<\n'
                   'MP76543215BLR9002155F90021559002159A002PB352')
        passport.read_photo = fake({'light': mangled, 'alt': REAL_MRZ,
                                    'full': REAL_MRZ})
        del calls[:]
        d = passport.read_passport_photo('x.jpg')
        check('расхождение поднимает тяжёлый путь', 'full' in calls, True)
        check('  берётся его чтение', d['name'], 'PETROVA ALESIA')
        check('  и оно подтверждено вторым проходом', d['name_confirmed'], True)

        # All three disagree: the name is handed over flagged, not silently.
        other = ('P<BLRPETROVA<<ALENA<<<<<<<<<<<<<<<<<<<<<<<<<\n'
                 'MP76543215BLR9002155F90021559002159A002PB352')
        passport.read_photo = fake({'light': mangled, 'alt': other,
                                    'full': REAL_MRZ})
        d = passport.read_passport_photo('x.jpg')
        check('несогласованное имя помечается', d['name_confirmed'], False)
        check('  номер при этом верен', d['number'], 'MP7654321')
        check('  и его контрольные цифры сошлись', d['ok'], True)

        # Light passes unusable, detector saves it.
        passport.read_photo = fake({'light': 'мусор', 'alt': 'мусор',
                                    'full': REAL_MRZ})
        d = passport.read_passport_photo('x.jpg')
        check('нечитаемые лёгкие проходы — работает тяжёлый',
              d['name'], 'PETROVA ALESIA')
        check('  без подтверждения', d['name_confirmed'], False)

        # Nothing readable anywhere.
        passport.read_photo = fake({'light': 'мусор', 'alt': 'мусор',
                                    'full': 'мусор'})
        try:
            passport.read_passport_photo('x.jpg')
            check('совсем нечитаемое фото отвергается', False, True)
        except ValueError:
            check('совсем нечитаемое фото отвергается', True, True)
    finally:
        passport.read_photo = original


def test_dates():
    print('\nДаты и век')
    import datetime
    yy = datetime.date.today().year % 100
    check('дата рождения в прошлом веке', ps._date('780907'), '07.09.1978')
    check('срок действия всегда в этом веке', ps._date('320325', future=True),
          '25.03.2032')
    check('невозможная дата отбрасывается', ps._date('780230'), '')
    check('мусор отбрасывается', ps._date('abc'), '')
    recent = '%02d0101' % max(0, yy - 1)
    check('недавний год — этот век', ps._date(recent)[-4:], str(2000 + max(0, yy - 1)))


def test_buyer_draft():
    print('\nЧерновик покупателя')
    d = ps.parse_mrz(SPECIMEN)
    b = ps.buyer_draft(d)
    check('имя как первая строка', b['name'], 'ERIKSSON ANNA MARIA')
    check('незнакомый код страны оставляется как есть', b['country'], 'UTO')
    check('строка паспорта', b['passport_line'], 'paszport: L898902C3')

    blr = dict(d, nationality='BLR', number='MP1234567')
    b = ps.buyer_draft(blr)
    check('BLR → Belarus', b['country'], 'Belarus')
    check('черновик собран', b['info'], 'Belarus,\npaszport: MP1234567')


def test_buyer_book():
    print('\nСправочник покупателей')
    import buyers as bk
    backup = bk.BUYERS_FILE.read_bytes() if bk.BUYERS_FILE.exists() else None
    try:
        bk.BUYERS_FILE.unlink(missing_ok=True)
        check('пустой справочник', bk.load_buyers(), [])

        bk.remember_buyer('IVANOU SIARHEI', 'Belarus\npaszport: MP1234567')
        bk.remember_buyer('IVANOV IVAN', 'Belarus\npaszport: AB1')
        names = [b['name'] for b in bk.load_buyers()]
        check('последний использованный — первым', names,
              ['IVANOV IVAN', 'IVANOU SIARHEI'])

        bk.remember_buyer('IVANOU SIARHEI', 'Belarus\nновый адрес')
        names = [b['name'] for b in bk.load_buyers()]
        check('повтор не плодит дубль, а поднимает наверх', names,
              ['IVANOU SIARHEI', 'IVANOV IVAN'])
        check('данные обновились',
              bk.find_buyer('IVANOU SIARHEI')['info'], 'Belarus\nновый адрес')
        check('поиск не зависит от регистра и пробелов',
              bk.find_buyer('  ivanou   siarhei ')['name'], 'IVANOU SIARHEI')

        bk.remember_buyer('', 'без имени')
        check('пустое имя не сохраняется', len(bk.load_buyers()), 2)

        bk.forget_buyer('IVANOV IVAN')
        check('удаление работает', [b['name'] for b in bk.load_buyers()],
              ['IVANOU SIARHEI'])

        bk.BUYERS_FILE.write_text('не json', encoding='utf-8')
        check('битый файл не роняет бота', bk.load_buyers(), [])
    finally:
        if backup is None:
            bk.BUYERS_FILE.unlink(missing_ok=True)
        else:
            bk.BUYERS_FILE.write_bytes(backup)


def main():
    print('Проверка чтения MRZ паспорта')
    test_buyer_book()
    test_specimen()
    test_check_digits()
    test_noise()
    test_ocr_confusions()
    test_mangled_filler()
    test_looks_read()
    test_name_consensus()
    test_dates()
    test_buyer_draft()
    print('\n' + ('✅ ВСЁ ПРОШЛО' if not failures
                  else '❌ ПРОВАЛЕНО %d: %s' % (len(failures), '; '.join(failures))))
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
