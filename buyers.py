#!/usr/bin/env python3
"""The buyer book.

Every Pro Forma in the archive so far went to the same person, so retyping the
four-line buyer block each time is pure waste. Whoever a Pro Forma was made out
to is remembered here and offered as a button next time, most recent first.
"""

import json
import logging
from datetime import datetime
from pathlib import Path

BUYERS_FILE = Path(__file__).with_name('buyers.json')
KEEP = 12                      # more than this and the card stops being useful


def _key(name) -> str:
    return ' '.join((name or '').upper().split())


def load_buyers() -> list:
    try:
        data = json.loads(BUYERS_FILE.read_text(encoding='utf-8'))
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    return [b for b in data if isinstance(b, dict) and b.get('name')]


def save_buyers(buyers) -> None:
    try:
        BUYERS_FILE.write_text(
            json.dumps(buyers[:KEEP], ensure_ascii=False, indent=2),
            encoding='utf-8')
    except Exception:
        logging.exception('Не удалось сохранить справочник покупателей')


def remember_buyer(name, info) -> list:
    """Put this buyer at the front of the book, replacing any earlier entry
    under the same name."""
    name = (name or '').strip()
    if not name:
        return load_buyers()
    entry = {'name': name, 'info': (info or '').strip(),
             'used': datetime.now().strftime('%Y-%m-%d')}
    rest = [b for b in load_buyers() if _key(b['name']) != _key(name)]
    buyers = [entry] + rest
    save_buyers(buyers)
    return buyers[:KEEP]


def forget_buyer(name) -> list:
    buyers = [b for b in load_buyers() if _key(b['name']) != _key(name)]
    save_buyers(buyers)
    return buyers


def find_buyer(name):
    for b in load_buyers():
        if _key(b['name']) == _key(name):
            return b
    return None
