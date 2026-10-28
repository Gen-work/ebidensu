"""GIFTデータvsGFIXデータ: does every DF picture compare the files the receive
logs stored? (K-53)

A DF window's title bar and its [1] / [2] header lines carry both file paths:
    DF - [...\\DATA\\GIFT\\LJODWAF1\\F202609280005.csv - ...\\GFIX\\...\\F202610260009.cs]
OCR the top of each DF picture, read the F<12 digits> name next to GIFT and
next to GFIX, and check each against the "file stored : [...]" lines of
GIFTReceive.log (GIFT受信結果) / GFIXReceive.log (GFIX受信結果).

10/27 LJODWAF1: the DF of STEP MQFX0030 compared GIFT F202609280003.csv --
KJODWAF1's empty file -- while the GIFT log stored ...0004 / ...0005. The
hand-pasted right DF (...0005) had been deleted and the tool's old one kept:
both empty files, both "同一内容", only the title tells them apart.

Needs Tesseract, like gift_list.py.
"""
import re

from drawings import sheet_items
import gift_list

STORED = re.compile(r'file stored : \[[^\]]*?(F\d{12})\.\w+\]')
SIDE = re.compile(r'(GIFT|GFIX)\W{0,3}\w*\W{0,3}(F\s*\d[\d ]{11,14})')
MIN_WIDTH = 300     # the DF's "..." wave between two pictures is 31 px wide
# (top px, bottom px, scale): one read misses a digit now and then (10/23
# C/F/MJODWAF1: 20261023 read as 20261028 at one scale only), so a side is
# right when ANY read names a stored file, and wrong only when none does
READS = ((0, 110, 2), (0, 110, 4), (0, 30, 2))


def stored_files(ws):
    """'file stored' names (without extension) of a receive-log sheet, in order."""
    out = []
    for row in ws.iter_rows(min_col=2, max_col=2):
        v = row[0].value
        if isinstance(v, str):
            for m in STORED.finditer(v):
                if m.group(1) not in out:
                    out.append(m.group(1))
    return out


def read_pictures(bk, part):
    """-> [(picture name, anchor row, {'GIFT': {names}, 'GFIX': {names}})] for
    every DF picture on the sheet (pictures underneath another included)."""
    out = []
    for it in sheet_items(bk, part):
        if it['kind'] != 'pic' or not it.get('media'):
            continue
        try:
            im = gift_list._picture(bk, it)
        except Exception:
            continue
        if im.width < MIN_WIDTH:
            continue
        names = {'GIFT': set(), 'GFIX': set()}
        for top, bottom, scale in READS:
            text = gift_list._ocr(im.crop((0, top, im.width, min(im.height, bottom))), scale=scale).replace('¥', '\\')
            for side, raw in SIDE.findall(text):
                d = re.sub(r'\s', '', raw)
                if len(d) == 13:
                    names[side].add(d)
        if names['GIFT'] or names['GFIX']:
            a = it.get('anchor')
            out.append((it['name'], a[1] + 1 if a else 0, names))
    return out


def judge(pictures, stored):
    """-> [(level, picture name, row, message)]: a DF side whose file the
    receive log did not store is NG (exact name: ...0003 and ...0005 are one
    digit apart, so no OCR slack)."""
    out = []
    for name, row, names in pictures:
        for side in ('GIFT', 'GFIX'):
            got, log = names[side], stored.get(side) or []
            if not got or not log:
                continue
            if not got & set(log):
                out.append(('NG', name, row,
                            'DF %s side is %s.csv, but %sReceive.log stored %s: another job\'s file? '
                            '(empty files all compare 同一内容; a hand-pasted DF deleted instead of the old one?)'
                            % (side, '/'.join(sorted(got)), side, ', '.join(f + '.csv' for f in log))))
    return out
