#!/usr/bin/env python3
"""Read text off a passport photo. Run as a separate process, on purpose.

    python3 ocr_worker.py photo.jpg          # light path
    python3 ocr_worker.py photo.jpg --alt    # light path, sliced differently
    python3 ocr_worker.py photo.jpg --full   # detector path, as a fallback

Memory is the whole design here. This VM has 945 MB, and importing the OCR
engine inside the bot would leave the bot sitting on hundreds of megabytes for
the rest of its life — so the engine lives in this short-lived process instead.

Two paths, because the detector is what costs the memory. Its config resizes
an image so the *smaller* side reaches 736 px, which blows a wide 1600x400 MRZ
strip up to roughly 2900x736 and costs ~530 MB. The light path skips the
detector: it finds the two MRZ lines by counting dark pixels per row, slices
each line where there is no ink, and sends the pieces straight to the
recogniser — about 200 MB. Slicing matters because the recogniser squeezes a
line into 320 px of width, and 44 characters across 320 px is unreadable mush.

The caller checks the MRZ check digits and asks for --full only when the light
path came back wrong.
"""

import sys

MAX_SIDE = 1600          # phone photos are far larger than OCR needs
MRZ_BAND = 0.40          # bottom of the page, where the two MRZ lines sit
CHUNK_ASPECT = 7.0       # width-to-height a text slice may reach
ALT_ASPECT = 4.5         # a second, differently-sliced reading, for --alt
DARK_MARGIN = 25         # how much darker than the page counts as ink


def _load(path):
    from PIL import Image, ImageOps
    img = Image.open(path)
    # JPEG can be decoded straight at a reduced scale, so the full-size bitmap
    # never has to exist in memory.
    img.draft('RGB', (MAX_SIDE, MAX_SIDE))
    img = ImageOps.exif_transpose(img).convert('RGB')
    if max(img.size) > MAX_SIDE:
        scale = float(MAX_SIDE) / max(img.size)
        img = img.resize((int(img.width * scale), int(img.height * scale)),
                         Image.BILINEAR)
    return img


def _band(img):
    return img.crop((0, int(img.height * (1 - MRZ_BAND)), img.width, img.height))


def _ink(img):
    import numpy as np
    g = np.asarray(img.convert('L'), dtype=np.int16)
    return g < (int(g.mean()) - DARK_MARGIN)


def _runs(flags, min_len):
    out, start = [], None
    for i, on in enumerate(flags):
        if on and start is None:
            start = i
        elif not on and start is not None:
            out.append((start, i))
            start = None
    if start is not None:
        out.append((start, len(flags)))
    return [r for r in out if r[1] - r[0] >= min_len]


def _text_rows(band):
    """The last two dense rows of the strip — the MRZ lines."""
    rows = _ink(band).sum(axis=1)
    return _runs(rows > max(6, band.width * 0.06), 8)[-2:]


def _cuts(line, aspect):
    """Where to slice a line so no character is cut in half: at the columns
    carrying the least ink, near evenly spaced targets."""
    columns = _ink(line).sum(axis=0)
    pieces = max(1, int(round(line.width / float(line.height) / aspect)))
    width = line.width
    points = []
    for k in range(1, pieces):
        want = int(width * k / float(pieces))
        lo, hi = max(1, want - 45), min(width - 1, want + 45)
        if lo < hi:
            points.append(min(range(lo, hi), key=lambda x: (columns[x], abs(x - want))))
    return [0] + sorted(set(points)) + [width]


def _rec_line(engine, line, aspect=CHUNK_ASPECT):
    import numpy as np
    out = []
    bounds = _cuts(line, aspect)
    for i in range(len(bounds) - 1):
        piece = line.crop((bounds[i], 0, bounds[i + 1], line.height))
        if piece.width < 10:
            continue
        result, _ = engine(np.asarray(piece))
        if result:
            first = result[0]
            out.append(first[0] if isinstance(first, (list, tuple)) else str(first))
    return ''.join(out)


def read_light(path, aspect=CHUNK_ASPECT) -> str:
    """No detector: locate the lines here, recognise them piece by piece.

    `aspect` changes where the slices fall. Reading twice with different values
    and comparing is the only handle there is on the name, which no check digit
    in a TD3 MRZ covers.
    """
    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR(use_det=False, use_cls=False)
    band = _band(_load(path))
    lines = []
    for top, bottom in _text_rows(band):
        pad = 6
        line = band.crop((0, max(0, top - pad), band.width,
                          min(band.height, bottom + pad)))
        lines.append(_rec_line(engine, line, aspect))
    return '\n'.join(lines)


def read_full(path) -> str:
    """The detector path: slower and much heavier, but it copes with photos
    the light path cannot segment."""
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR(use_cls=False)
    img = _load(path)
    band = _band(img)
    result, _ = engine(np.asarray(band))
    text = '\n'.join(line[1] for line in (result or []))
    if 'P<' in text.replace(' ', '').upper():
        return text
    result, _ = engine(np.asarray(img))
    return text + '\n' + '\n'.join(line[1] for line in (result or []))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if not args:
        print('нужен путь к изображению', file=sys.stderr)
        return 2
    try:
        if '--full' in sys.argv:
            text = read_full(args[0])
        else:
            text = read_light(args[0], ALT_ASPECT if '--alt' in sys.argv
                              else CHUNK_ASPECT)
        sys.stdout.write(text)
    except Exception as e:                    # noqa: BLE001 - reported upstream
        print('ошибка распознавания: %s' % e, file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
