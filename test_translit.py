#!/usr/bin/env python3
"""Checks for the Cyrillic → Polish address rewrite.

The invoice says "ul. Szyszkina 12-7" where the passport's registration page
says "ул. Шишкина, дом 12, кв. 7". Getting from one to the other is the part a
person has to stop and think about, so it is the part worth checking.
"""

import sys

from translit import has_cyrillic, transliterate as t

failures = []


def check(name, got, want):
    ok = got == want
    print(('  ✅ ' if ok else '  ❌ ') + name
          + ('' if ok else '  — получено %r, ожидалось %r' % (got, want)))
    if not ok:
        failures.append(name)


def test_real_address():
    print('\nАдрес из паспорта покупателя')
    check('как записано в фактуре',
          t('Минск, ул. Шишкина 12-7'), 'Mińsk, ul. Szyszkina 12-7')
    check('как написано в прописке',
          t('г. Минск, ул. Шишкина, дом 12, кв. 7'), 'Mińsk, ul. Szyszkina 12-7')
    check('с прочерком в корпусе',
          t('ул. Шишкина, дом 12, корп. -, кв. 7'), 'ul. Szyszkina 12-7')


def test_orthography():
    print('\nПольская орфография')
    # The whole point: a letter-by-letter map gives Sziszkina, which is wrong.
    check('ш → sz, и после него → y', t('Шишкина'), 'Szyszkina')
    check('ч → cz с тем же правилом', t('Чижовка'), 'Czyżowka')
    check('щ → szcz', t('Щорса'), 'Szczorsa')
    check('ж → ż', t('Жодино'), 'Żodino')
    check('х → ch', t('Хатынь'), 'Chatyn')
    check('ц → c', t('Цнянская'), 'Cnianskaja')
    check('ю и я', t('Юбилейная'), 'Jubilejnaja')
    check('и после к остаётся i', t('Кировская'), 'Kirowskaja')


def test_place_names():
    print('\nГорода — это соглашения, а не правила')
    for src, want in [('Минск', 'Mińsk'), ('Мінск', 'Mińsk'),
                      ('Гомель', 'Homel'), ('Брест', 'Brześć'),
                      ('Витебск', 'Witebsk'), ('Гродно', 'Grodno'),
                      ('Могилёв', 'Mohylew'), ('Беларусь', 'Belarus')]:
        check('%s → %s' % (src, want), t(src), want)


def test_left_alone():
    print('\nЧего трогать нельзя')
    check('уже латиница не меняется',
          t('Mińsk, ul. Szyszkina 12-7'), 'Mińsk, ul. Szyszkina 12-7')
    check('цифры и дефисы целы', t('30-3'), '30-3')
    check('пустая строка', t(''), '')
    check('смешанный текст', t('ul. Шишкина 12'), 'ul. Szyszkina 12')

    check('кириллица найдена', has_cyrillic('Минск'), True)
    check('латиница — не кириллица', has_cyrillic('Minsk'), False)
    check('пусто — не кириллица', has_cyrillic(''), False)


def main():
    print('Проверка транслитерации адреса')
    test_real_address()
    test_orthography()
    test_place_names()
    test_left_alone()
    print('\n' + ('✅ ВСЁ ПРОШЛО' if not failures
                  else '❌ ПРОВАЛЕНО %d: %s' % (len(failures), '; '.join(failures))))
    return 1 if failures else 0


if __name__ == '__main__':
    sys.exit(main())
