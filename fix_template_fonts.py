#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Убрать из бланка CMR несуществующие имена шрифтов.

Образец клиента размечен шрифтом «Times New Roman Bold». Такого семейства нет:
жирность — это начертание внутри «Times New Roman», а не отдельный шрифт. Word
на Windows догадывается и рисует нужное начертание, LibreOffice на сервере — нет:
имя не находится, и прогон уезжает в подстановку (DejaVu Sans). В шаблоне так
размечена часть прогонов, поэтому в PDF соседние ячейки одной строки выходили
разными шрифтами: «2 941,20» шире остальных весов и переносилось на вторую
строку. Здесь имя заменяется настоящим, а жирность выносится в <w:b/>.
Заодно «Arial CYR» — такой же устаревший псевдоним — становится «Arial».
"""

import re
import shutil
import sys
import zipfile

FAKE_BOLD = 'Times New Roman Bold'
REAL = 'Times New Roman'
ALIASES = {'Arial CYR': 'Arial'}

_RPR = re.compile(r'<w:rPr>.*?</w:rPr>', re.S)
_RFONTS = re.compile(r'<w:rFonts\b[^>]*/>')


def _fix_rpr(block):
    if FAKE_BOLD not in block:
        return block
    block = block.replace(FAKE_BOLD, REAL)
    # Жирность держалась на имени шрифта — без <w:b/> текст стал бы обычным.
    if not re.search(r'<w:b(?:\s[^>]*)?/>', block):
        block = _RFONTS.sub(lambda m: m.group(0) + '<w:b/>', block, count=1)
    return block


def fix(path):
    src = zipfile.ZipFile(path)
    parts = {n: src.read(n) for n in src.namelist()}
    infos = src.infolist()
    src.close()

    changed = 0
    for name, data in parts.items():
        if not name.endswith('.xml') or 'word/' not in name:
            continue
        xml = data.decode('utf-8')
        new = _RPR.sub(lambda m: _fix_rpr(m.group(0)), xml)
        for old, real in ALIASES.items():
            new = new.replace(f'"{old}"', f'"{real}"')
        if new != xml:
            parts[name] = new.encode('utf-8')
            changed += 1

    shutil.copy2(path, path + '.bak-fonts')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as out:
        for info in infos:
            out.writestr(info, parts[info.filename])
    print(f'{path}: изменено частей {changed}')


if __name__ == '__main__':
    for p in sys.argv[1:]:
        fix(p)
