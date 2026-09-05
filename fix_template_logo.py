#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Поставить «CMR» ровно в середину овала на бланке.

В образце клиента надпись держалась на пустых местах: текстовая рамка вдвое
шире овала, абзац прижат влево, а «CMR» отодвинуто вправо четырьмя пробелами.
В Word ширина пробела совпадала с задуманной, и буквы попадали в овал; при
конвертации на сервере шрифт подставляется другой, ширина пробела уезжает —
и надпись вылезает за левый край овала.

Пробелы здесь заменены настоящим выравниванием: рамка ставится точно по овалу
(та же геометрия в координатах группы), абзац центрируется по горизонтали, а
`v-text-anchor:middle` центрует его по вертикали. Теперь положение надписи не
зависит от того, каким шрифтом её нарисуют.
"""

import re
import shutil
import sys
import zipfile

TEXT_RECT = 'Rectangles 3'
OVAL = 'Oval 4'

# Абзац без ведущих пробелов: одно слово, центрированное в рамке.
PARAGRAPH = (
    '<w:txbxContent><w:p w14:paraId="3ACCB341" w14:textId="77777777"'
    ' w:rsidR="00B80D66" w:rsidRDefault="003E2A11"><w:pPr>'
    '<w:pStyle w:val="6"/><w:spacing w:before="0" w:after="0"/>'
    '<w:jc w:val="center"/></w:pPr>'
    '<w:r><w:t>CMR</w:t></w:r></w:p></w:txbxContent>'
)


def _geometry(xml, shape_id):
    m = re.search(r'<v:\w+ id="%s"[^>]*style="([^"]*)"' % re.escape(shape_id), xml)
    if not m:
        raise SystemExit(f'не найдена фигура {shape_id}')
    return dict(p.split(':', 1) for p in m.group(1).split(';') if ':' in p)


def fix(path):
    src = zipfile.ZipFile(path)
    parts = {n: src.read(n) for n in src.namelist()}
    infos = src.infolist()
    src.close()

    xml = parts['word/document.xml'].decode('utf-8')
    oval = _geometry(xml, OVAL)
    style = ('position:absolute;left:%s;top:%s;width:%s;height:%s;'
             'v-text-anchor:middle' % (oval['left'], oval['top'],
                                       oval['width'], oval['height']))

    start = xml.index(f'<v:rect id="{TEXT_RECT}"')
    end = xml.index('</v:rect>', start) + len('</v:rect>')
    block = xml[start:end]
    block = re.sub(r'style="[^"]*"', 'style="%s"' % style, block, count=1)
    block = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', PARAGRAPH, block,
                   count=1, flags=re.S)
    parts['word/document.xml'] = (xml[:start] + block + xml[end:]).encode('utf-8')

    shutil.copy2(path, path + '.bak-logo')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as out:
        for info in infos:
            out.writestr(info, parts[info.filename])
    print(f'{path}: надпись CMR выровнена по овалу ({style})')


if __name__ == '__main__':
    for p in sys.argv[1:]:
        fix(p)
