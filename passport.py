#!/usr/bin/env python3
"""Read the machine-readable zone of a passport (ICAO 9303, TD3).

The two lines across the bottom of the photo page carry the name, document
number, nationality, date of birth and expiry — each guarded by a check digit.
That last part is what makes this worth doing at all: a passport number that a
scanner misread fails its check digit, so the bot can refuse a bad read instead
of quietly putting a wrong number on an export document.

What the MRZ does NOT carry, and the Pro Forma needs anyway:
  • the date of issue ("paszport: MP1234567 od 25.03.2022") — printed only in
    the visual zone above;
  • the address — a Belarusian passport does not show one on this spread at all.
"""

import os
import re
import subprocess
import sys
from pathlib import Path

# '<' is the filler, letters count from 10. Weights repeat 7-3-1.
_WEIGHTS = (7, 3, 1)
LINE_LEN = 44


def _value(ch) -> int:
    if ch == '<':
        return 0
    if ch.isdigit():
        return int(ch)
    return ord(ch.upper()) - ord('A') + 10


def check_digit(chunk) -> str:
    total = sum(_value(ch) * _WEIGHTS[i % 3] for i, ch in enumerate(chunk))
    return str(total % 10)


# Scanners confuse the same handful of shapes every time. The MRZ says which
# fields are letters and which are digits, so most of it can simply be forced
# back — and that matters more than it looks: the check digits cover the
# document number, the dates and the personal number, but NOT the name and NOT
# the nationality, so an unrepaired "ER1KSSON" would sail through onto the
# invoice unnoticed.
_TO_ALPHA = {'0': 'O', '1': 'I', '2': 'Z', '5': 'S', '6': 'G', '8': 'B'}
_TO_DIGIT = {'O': '0', 'Q': '0', 'D': '0', 'U': '0', 'I': '1', 'L': '1',
             'Z': '2', 'A': '4', 'S': '5', 'G': '6', 'T': '7', 'B': '8'}


def _as_alpha(s) -> str:
    return ''.join(_TO_ALPHA.get(c, c) for c in s)


def _as_digits(s) -> str:
    return ''.join(_TO_DIGIT.get(c, c) for c in s)


def _repair(field, expected_check):
    """Return the reading of `field` whose check digit matches, trying the
    obvious letter/digit confusions before giving up on the raw text."""
    for candidate in (field, _as_digits(field), _as_alpha(field)):
        if check_digit(candidate) == expected_check:
            return candidate, True
    return field, False


def _date(yymmdd, future=False) -> str:
    """'780907' → '07.09.1978'. Expiry dates are always in this century."""
    m = re.fullmatch(r'(\d{2})(\d{2})(\d{2})', yymmdd or '')
    if not m:
        return ''
    yy, mm, dd = (int(x) for x in m.groups())
    if future:
        year = 2000 + yy
    else:
        import datetime
        year = 1900 + yy if yy > datetime.date.today().year % 100 else 2000 + yy
    try:
        import datetime
        datetime.date(year, mm, dd)
    except ValueError:
        return ''
    return '%02d.%02d.%d' % (dd, mm, year)


# Scanners do not know the '<' glyph and render the name padding as runs of
# Λ, A or 八. The non-ASCII ones are stripped, but a run of real letters is
# left behind and reads as an extra given name — and worse, how much of it
# survives varies between passes, which is exactly what a consensus check must
# not be tripped up by. A token of one letter repeated is never a name.
_FILLER_RUN = re.compile(r'^(.)\1+$')


def _clean_name(field) -> str:
    tokens = [t for t in _as_alpha(field or '').split('<') if t]
    while tokens and _FILLER_RUN.match(tokens[-1]):
        tokens.pop()
    return ' '.join(tokens).strip()


def find_lines(text):
    """Pick the two MRZ lines out of whatever the scanner produced.

    Only the first line is padded with '<' — the second one is dense data and
    on this passport contains no filler at all, so it must not be used as a
    signature for either line.
    """
    rows = [r for r in (re.sub(r'[^A-Z0-9<]', '', row.upper())
                        for row in (text or '').splitlines()) if r]
    # No length test on the first line: a scanner that renders the trailing
    # '<' padding as Λ or 八 leaves it far shorter than 44 once those are
    # stripped, and padding restores it. The second line is dense data, so
    # there a near-full length is exactly the right signature.
    for i, row in enumerate(rows):
        if not (row.startswith('P') and '<<' in row):
            continue
        for follower in rows[i + 1:i + 4]:
            if len(follower) >= LINE_LEN - 4:
                return (row[:LINE_LEN].ljust(LINE_LEN, '<'),
                        follower[:LINE_LEN].ljust(LINE_LEN, '<'))
    return None, None


def parse_mrz(text) -> dict:
    """Return the fields plus, per field, whether its check digit agreed.

    Raises ValueError when the two lines cannot be found at all.
    """
    top, bottom = find_lines(text)
    if not top:
        raise ValueError('не найдены две строки MRZ — нижняя часть страницы '
                         'с «P<BLR…» должна попасть в кадр целиком')

    # Name and issuing state are letters only — no check digit guards them.
    names = _as_alpha(top[5:]).split('<<')
    surname = _clean_name(names[0])
    given = _clean_name(names[1] if len(names) > 1 else '')

    number, number_ok = _repair(bottom[0:9], bottom[9])
    birth, birth_ok = _repair(_as_digits(bottom[13:19]), bottom[19])
    expiry, expiry_ok = _repair(_as_digits(bottom[21:27]), bottom[27])
    personal, personal_ok = _repair(bottom[28:42], bottom[42])

    # ICAO 9303 TD3: the composite runs over positions 1-10, 14-20 and 22-44 —
    # the sex field at position 21 is deliberately left out.
    composite = (number + bottom[9] + birth + bottom[19]
                 + expiry + bottom[27] + personal + bottom[42])
    checks = {
        'number': number_ok,
        'birth': birth_ok,
        'expiry': expiry_ok,
        'personal': personal_ok,
        'composite': check_digit(composite) == bottom[43],
    }

    return {
        'surname': surname,
        'given_names': given,
        'name': (surname + ' ' + given).strip(),
        'country': _as_alpha(top[2:5]).replace('<', ''),
        'number': number.replace('<', ''),
        'nationality': _as_alpha(bottom[10:13]).replace('<', ''),
        'birth_date': _date(birth),
        'sex': bottom[20].replace('<', ''),
        'expiry_date': _date(expiry, future=True),
        'personal_number': personal.replace('<', ''),
        'checks': checks,
        'ok': all(checks.values()),
        'failed': sorted(k for k, v in checks.items() if not v),
    }


# -- Reading a photo ----------------------------------------------------------
OCR_TIMEOUT = int(os.environ.get('OCR_TIMEOUT', '120'))


def ocr_available() -> bool:
    try:
        import importlib.util
        return importlib.util.find_spec('rapidocr_onnxruntime') is not None
    except Exception:
        return False


def read_photo(path, full=False, alt=False) -> str:
    """Text off a passport photo, via the separate worker process.

    Keeping the engine out of this process is the whole point — see the note
    at the top of ocr_worker.py.
    """
    worker = Path(__file__).with_name('ocr_worker.py')
    flags = ['--full'] if full else (['--alt'] if alt else [])
    cmd = [sys.executable, str(worker), str(path)] + flags
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=OCR_TIMEOUT)
    except subprocess.TimeoutExpired:
        raise ValueError('распознавание не уложилось в %d с' % OCR_TIMEOUT)
    if proc.returncode != 0:
        detail = (proc.stderr or b'').decode('utf-8', 'replace').strip()
        raise ValueError(detail or 'распознаватель не смог прочитать фото')
    return (proc.stdout or b'').decode('utf-8', 'replace')


_NAME_PART = re.compile(r'^[A-Z]{2,}(?: [A-Z]+)*$')


def looks_read(mrz) -> bool:
    """Is this reading good enough to skip the heavy second pass?

    The check digits vouch for the number and the dates, but nothing in a TD3
    MRZ guards the name — so it gets its own sanity test before the cheap path
    is trusted.
    """
    return (mrz.get('ok')
            and bool(_NAME_PART.match(mrz.get('surname') or ''))
            and bool(_NAME_PART.match(mrz.get('given_names') or '')))


def read_passport_photo(path) -> dict:
    """Photo → parsed MRZ, with an honest note on how far to trust the name.

    The number and the dates carry check digits and are either right or
    rejected. The name carries nothing: a TD3 MRZ has no check digit over it,
    so a scanner that drops letters — PETROVA read as KNRTSEZK — produces
    a wrong name that every other test passes. The only handle available is
    agreement, so the photo is read twice with the slices falling in different
    places, and the name is called confirmed only when both readings match.

    `name_confirmed` says which of the two happened. It is never a promise that
    the name is right, only that two independent readings agreed on it.
    """
    readings, light_error = [], None
    for alt in (False, True):
        try:
            mrz = parse_mrz(read_photo(path, alt=alt))
            if looks_read(mrz):
                readings.append(mrz)
        except ValueError as e:
            light_error = e

    if len(readings) == 2 and readings[0]['name'] == readings[1]['name']:
        readings[0]['name_confirmed'] = True
        return readings[0]

    # No agreement — bring in the detector and see what it makes of it.
    try:
        heavy = parse_mrz(read_photo(path, full=True))
    except ValueError:
        if readings:
            readings[0]['name_confirmed'] = False
            return readings[0]
        raise light_error or ValueError('не удалось прочитать MRZ')

    heavy['name_confirmed'] = any(r['name'] == heavy['name'] for r in readings)
    return heavy


# ISO country code → what the Pro Forma writes on the buyer's address line.
COUNTRIES = {
    'BLR': 'Belarus', 'RUS': 'Rosja', 'UKR': 'Ukraina', 'KAZ': 'Kazachstan',
    'POL': 'Polska', 'LTU': 'Litwa', 'LVA': 'Łotwa', 'EST': 'Estonia',
    'GEO': 'Gruzja', 'ARM': 'Armenia', 'AZE': 'Azerbejdżan', 'MDA': 'Mołdawia',
    'UZB': 'Uzbekistan', 'KGZ': 'Kirgistan', 'TJK': 'Tadżykistan',
}

FIELD_NAMES = {
    'number': 'номер паспорта', 'birth': 'дата рождения',
    'expiry': 'срок действия', 'personal': 'личный номер',
    'composite': 'общая контрольная сумма',
}


def buyer_draft(mrz) -> dict:
    """Turn a parsed MRZ into the buyer block the Pro Forma asks for.

    `info` is only a draft: the street address is not in the passport, and
    neither is the date of issue, so both are left for the user to add.
    """
    country = COUNTRIES.get(mrz.get('nationality'), mrz.get('nationality', ''))
    return {
        'name': mrz.get('name', ''),
        'country': country,
        'passport_line': 'paszport: %s' % mrz.get('number', ''),
        'info': '\n'.join(x for x in (country + ',' if country else '',
                                      'paszport: %s' % mrz.get('number', '')) if x),
    }
