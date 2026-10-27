"""GIFT受信結果: is the red box on the GIFT transfer list on the row(s) the
Detail pages show? (K-50)

The list ("Transfer status inquiry results") gives two lines per transfer:
the times, then the Msgid, e.g.
    A2009999A000000000000001402026092902153027600001
and a Detail page gives SENDDATETIMECHAR, e.g. 202609290215302760 -- the
same digits, inside that Msgid. So: OCR the part of the list picture under
each red box, read the Msgid(s) there, OCR every Detail picture, and every
Detail's SENDDATETIMECHAR must be in a boxed Msgid.

10/27 JJODWAB1: box on row 9 (11:00:41), Detail and log on row 10 (11:15:30).
10/27 JJODWU33: box across the lower half of row 8 and the upper half of
row 9; the three files are rows 7-9.

Needs Tesseract (apt install tesseract-ocr; pip install pytesseract). When
it is missing the check reports one WARN and does nothing else.
"""
import io
import re

from drawings import sheet_items

MSGID = re.compile(r'A\s*2009999\s*A[0-9 ]{20,}')
SDT = re.compile(r'SENDDATETIMECHAR\W*([0-9][0-9 ]{11,})')
RECVFILE = re.compile(r'RECVFILENAME\W*\S*?([0-9]{14,})')


def ocr_available():
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _ocr(im, scale=2):
    import pytesseract
    return pytesseract.image_to_string(im.resize((im.width * scale, im.height * scale)), config='--psm 6')


def _picture(bk, it):
    from PIL import Image
    p = Image.open(io.BytesIO(bk.z.read(it['media']))).convert('RGB')
    c = it.get('crop') or {}
    if c:
        w, h = p.size
        p = p.crop((int(w * c.get('l', 0) / 1e5), int(h * c.get('t', 0) / 1e5),
                    int(w * (1 - c.get('r', 0) / 1e5)), int(h * (1 - c.get('b', 0) / 1e5))))
    return p


def _under(im, pic_box, box):
    """The part of the picture a box (sheet EMU) covers, in picture pixels."""
    w, h = im.size
    x0, y0, x1, y1 = pic_box
    sx, sy = w / max(1, x1 - x0), h / max(1, y1 - y0)
    return im.crop((max(0, int((box[0] - x0) * sx)), max(0, int((box[1] - y0) * sy)),
                    min(w, int((box[2] - x0) * sx)), min(h, int((box[3] - y0) * sy))))


def near_in(needle, hay, slack=1):
    """needle is in hay with at most `slack` digits read differently."""
    n = len(needle)
    for i in range(0, len(hay) - n + 1):
        if sum(1 for a, b in zip(needle, hay[i:i + n]) if a != b) <= slack:
            return True
    return False


def read_sheet(bk, part):
    """-> dict: details [(name, sdt digits or '', recv digits or '')],
    lists [(name, [boxed msgid digits per box])], text of every picture."""
    items = sheet_items(bk, part)
    pics = [i for i in items if i['kind'] == 'pic' and i.get('media')]
    boxes = [i for i in items if i['kind'] == 'sp' and (i.get('line') or '').upper() == 'FF0000']
    details, lists = [], []
    for pc in pics:
        try:
            im = _picture(bk, pc)
        except Exception:
            continue
        text = _ocr(im)
        flat = text.replace(' ', '')
        if 'SENDDATETIMECHAR' in flat:
            m, r = SDT.search(text), RECVFILE.search(text)
            details.append((pc['name'], re.sub(r'\D', '', m.group(1)) if m else '',
                            r.group(1) if r else ''))
            continue
        if len(MSGID.findall(text)) < 2:
            continue
        boxed = []
        for b in boxes:
            bb, pb = b['box'], pc['box']
            if not (bb[0] < pb[2] and bb[2] > pb[0] and bb[1] < pb[3] and bb[3] > pb[1]):
                continue
            t = _ocr(_under(im, pb, bb), scale=3)
            boxed.append((b['name'], [re.sub(r'\D', '', x) for x in MSGID.findall(t)]))
        lists.append((pc['name'], boxed))
    return {'details': details, 'lists': lists}


def judge(info):
    """-> [(level, message)] for the list box vs the Detail pages."""
    out = []
    details = [d for d in info['details'] if d[1]]
    if not info['lists'] or not details:
        if info['details'] and not details:
            out.append(('WARN', 'Detail page(s) found, but SENDDATETIMECHAR could not be read -- check the list box by eye'))
        return out
    boxed = [(lname, bname, ids) for lname, bxs in info['lists'] for bname, ids in bxs]
    if not boxed:
        out.append(('NG', 'no red box on the GIFT transfer list picture'))
        return out
    all_ids = [i for _, _, ids in boxed for i in ids]
    for name, sdt, _ in details:
        if not any(near_in(sdt, i) for i in all_ids):
            seen = ', '.join(i[-24:] for i in all_ids) or 'no Msgid read under the box'
            out.append(('NG', 'Detail %r SENDDATETIMECHAR %s is not on a boxed list row (boxed Msgid ...%s): '
                              'the box is on another row, or does not cover the whole row' % (name, sdt, seen)))
    for lname, bname, ids in boxed:
        for i in ids:
            if not any(near_in(sdt, i) for _, sdt, _ in details):
                out.append(('WARN', 'boxed list row Msgid ...%s has no Detail page' % i[-24:]))
    return out
