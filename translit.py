#!/usr/bin/env python3
"""Cyrillic address → the Polish spelling the Pro Forma uses.

Reading the registration page of a passport would not actually help: it says
"Шишкина", and the invoice says "Szyszkina". The gap between the two is not
optical, it is orthographic — and that is the part worth automating, because
it is the part a person has to stop and think about.

Nothing here is authoritative. The result is always shown on the confirmation
card before the document is written, and can be corrected there.
"""

import re

VOWELS = 'аеёиоуыэюяіў'
# After a vowel or at the start of a word the first form; after a consonant the
# second. е is the odd one out: after a consonant Polish just writes e.
IOTATED = {'я': ('ja', 'ia'), 'ю': ('ju', 'iu'),
           'ё': ('jo', 'io'), 'е': ('je', 'e')}

# Belarusian г is /h/, which is why Гомель is Homel and not Gomel in Polish.
LETTERS = [
    ('щ', 'szcz'), ('ш', 'sz'), ('ч', 'cz'), ('ж', 'ż'), ('х', 'ch'),
    ('э', 'e'), ('ы', 'y'),
    ('й', 'j'), ('ц', 'c'), ('ў', 'u'), ('і', 'i'), ('ъ', ''), ('ь', ''),
    ('а', 'a'), ('б', 'b'), ('в', 'w'), ('г', 'h'), ('д', 'd'),
    ('з', 'z'), ('и', 'i'), ('к', 'k'), ('л', 'l'), ('м', 'm'), ('н', 'n'),
    ('о', 'o'), ('п', 'p'), ('р', 'r'), ('с', 's'), ('т', 't'), ('у', 'u'),
    ('ф', 'f'),
]

# Place names and abbreviations whose Polish forms are conventions, not
# transliterations — no letter rule turns Минск into Mińsk.
EXCEPTIONS = {
    'минск': 'Mińsk', 'мінск': 'Mińsk',
    'гомель': 'Homel', 'гомель': 'Homel',
    'брест': 'Brześć', 'брэст': 'Brześć',
    'витебск': 'Witebsk', 'віцебск': 'Witebsk',
    'гродно': 'Grodno', 'гродна': 'Grodno',
    'могилёв': 'Mohylew', 'могилев': 'Mohylew', 'магілёў': 'Mohylew',
    'бобруйск': 'Bobrujsk', 'борисов': 'Borysów', 'барысаў': 'Borysów',
    'беларусь': 'Belarus', 'белоруссия': 'Belarus',
    'россия': 'Rosja', 'украина': 'Ukraina', 'казахстан': 'Kazachstan',
    'ул': 'ul', 'улица': 'ul', 'вул': 'ul',
    'пр': 'al', 'проспект': 'al', 'просп': 'al',
    'пер': 'zauł', 'переулок': 'zauł',
    'корп': 'korp',
}

# Polish does not write szi/czi/żi — the vowel turns into y, which is the whole
# difference between Sziszkina and Szyszkina. Case-insensitive, because the
# first letter of a street name is capital.
_HARDENED = re.compile(r'(sz|cz|ż|rz|dz)i', re.IGNORECASE)

# "дом 12, кв. 7" is written 30-3 on the invoice.
_HOUSE_FLAT = re.compile(r'(?:дом|буд|д)\.?\s*(\d+[а-яА-Я]?)\s*,?\s*'
                         r'(?:корп\.?\s*[-\d]+\s*,?\s*)?'
                         r'(?:кв|кват|апт)\.?\s*(\d+)', re.IGNORECASE)
_HOUSE_ONLY = re.compile(r'(?:дом|буд)\.?\s*(\d+[а-яА-Я]?)', re.IGNORECASE)
# A leading "г." just means "city of".
_CITY_PREFIX = re.compile(r'\b(?:г|гор|м)\.\s*(?=[А-ЯЁІ])', re.IGNORECASE)


def _word(word: str) -> str:
    lowered = word.lower()
    if lowered in EXCEPTIONS:
        out = EXCEPTIONS[lowered]
        return out if word[:1].islower() else out[:1].upper() + out[1:]

    result = []
    for i, ch in enumerate(word):
        upper = ch.isupper()
        low = ch.lower()
        # Polish writes ja/ju/jo at the start of a word and after a vowel, but
        # ia/iu/io after a consonant: Цнянская is Cnianskaja, not Cnjanskaja.
        if low in IOTATED:
            after_vowel = i == 0 or word[i - 1].lower() in VOWELS + 'ьъ'
            dst = IOTATED[low][0 if after_vowel else 1]
        else:
            dst = next((d for s, d in LETTERS if s == low), ch)
        result.append(dst.capitalize() if upper and dst else dst)
    out = ''.join(result)
    # Applied twice: "шишкина" produces two candidates in a row.
    for _ in range(2):
        out = _HARDENED.sub(lambda m: m.group(1) + 'y', out)
    return out


def transliterate(text: str) -> str:
    """Rewrite any Cyrillic in `text`; Latin parts and digits are left alone."""
    text = _CITY_PREFIX.sub('', text or '')
    text = _HOUSE_FLAT.sub(lambda m: '%s-%s' % (m.group(1), m.group(2)), text)
    text = _HOUSE_ONLY.sub(lambda m: m.group(1), text)
    text = re.sub(r'[А-Яа-яЁёІіЎўЭэ]+', lambda m: _word(m.group(0)), text)
    text = re.sub(r'\s+,', ',', re.sub(r'\s{2,}', ' ', text))
    # "ul. Szyszkina, 30-3" reads as "ul. Szyszkina 12-7" on the invoice.
    return re.sub(r',\s*(\d+[a-zA-Zа-яА-Я]?-\d+)', r' \1', text).strip()


def has_cyrillic(text: str) -> bool:
    return bool(re.search(r'[А-Яа-яЁёІіЎў]', text or ''))
