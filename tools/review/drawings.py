"""Read the pictures and shapes of an evidence workbook (.xlsx) and render them.

Positions come from the cell anchors (<xdr:from>/<xdr:to>, or <xdr:from> +
<xdr:ext> for a oneCellAnchor), NOT from the shape's own <a:xfrm>: Excel
places a drawing by its anchor, and the xfrm is often stale after rows were
inserted or resized. On 10/23 two red boxes looked misplaced (KJODWWBV,
CJODWAF1) only because the xfrm was used; see KNOWLEDGE K-42.

Usage:
  python3 drawings.py list   <book.xlsx>
  python3 drawings.py render <book.xlsx> <out_dir>

render writes one PNG per sheet that has drawings: the pictures pasted at
their anchored place and size, red/other shape outlines drawn on top. Cell
text is NOT drawn -- read it with check_book.py.
"""
import os
import sys
import zipfile
import posixpath
import xml.etree.ElementTree as ET

NS = {
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RID = '{%s}id' % NS['r']
EMU_PX = 9525           # EMU per pixel at 96 dpi
EMU_PT = 12700          # EMU per point (row heights are in points)


def _tag(e):
    return e.tag.split('}')[1]


class Book:
    """Thin reader over the xlsx zip: sheet list, rels, row geometry, drawings."""

    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)

    def xml(self, name):
        return ET.fromstring(self.z.read(name))

    def exists(self, name):
        return name in self.z.namelist()

    def rels(self, part):
        d, b = posixpath.split(part)
        rp = posixpath.join(d, '_rels', b + '.rels')
        out = {}
        if self.exists(rp):
            for r in self.xml(rp):
                out[r.get('Id')] = posixpath.normpath(posixpath.join(d, r.get('Target')))
        return out

    def sheets(self):
        """[(name, part, state)] in tab order."""
        wb = self.xml('xl/workbook.xml')
        rel = self.rels('xl/workbook.xml')
        return [(s.get('name'), rel[s.get(RID)], s.get('state', 'visible'))
                for s in wb.find(M + 'sheets')]

    def row_y(self, sheet_part):
        """Function (row0, rowOff_emu) -> y in EMU, from the sheet's row heights."""
        s = self.xml(sheet_part)
        fmt = s.find(M + 'sheetFormatPr')
        dh = float(fmt.get('defaultRowHeight', '15')) if fmt is not None else 15.0
        h = {}
        for r in s.iter(M + 'row'):
            if r.get('hidden') == '1':
                h[int(r.get('r'))] = 0.0
            elif r.get('ht'):
                h[int(r.get('r'))] = float(r.get('ht'))
        cache = {}

        def y(row0, off):
            if row0 not in cache:
                cache[row0] = sum(h.get(r, dh) for r in range(1, row0 + 1)) * EMU_PT
            return cache[row0] + int(off)
        return y

    def drawing_parts(self, sheet_part):
        return [t for t in self.rels(sheet_part).values()
                if '/drawings/' in t and t.endswith('.xml')]


def _collect(node, tf, out, media):
    """Walk pic/sp/grpSp under node. tf maps child xfrm coords to parent coords:
    (ox, oy, sx, sy, chx, chy)."""
    for ch in node:
        t = _tag(ch)
        if t not in ('pic', 'sp', 'grpSp', 'cxnSp'):
            continue
        pr = ch.find('xdr:grpSpPr', NS) if t == 'grpSp' else ch.find('xdr:spPr', NS)
        x = pr.find('a:xfrm', NS) if pr is not None else None
        if x is None or x.find('a:off', NS) is None:
            continue
        o, e = x.find('a:off', NS), x.find('a:ext', NS)
        X, Y, W, H = int(o.get('x')), int(o.get('y')), int(e.get('cx')), int(e.get('cy'))

        def m(px, py):
            return (tf[0] + (px - tf[4]) * tf[2], tf[1] + (py - tf[5]) * tf[3])
        x0, y0 = m(X, Y)
        x1, y1 = m(X + W, Y + H)
        nv = ch.find('.//xdr:cNvPr', NS)
        item = {'kind': t, 'name': nv.get('name', ''), 'descr': nv.get('descr', ''),
                'box': [x0, y0, x1, y1]}
        if t == 'grpSp':
            co, ce = x.find('a:chOff', NS), x.find('a:chExt', NS)
            sx = (x1 - x0) / max(1, int(ce.get('cx')))
            sy = (y1 - y0) / max(1, int(ce.get('cy')))
            _collect(ch, (x0, y0, sx, sy, int(co.get('x')), int(co.get('y'))), out, media)
            continue
        if t == 'pic':
            b = ch.find('.//a:blip', NS)
            item['media'] = media.get(b.get('{%s}embed' % NS['r'])) if b is not None else None
            src = ch.find('.//a:srcRect', NS)
            item['crop'] = {k: int(src.get(k, 0)) for k in 'ltrb'} if src is not None else {}
        else:
            ln = ch.find('xdr:spPr/a:ln', NS)
            col = None
            if ln is not None:
                c = ln.find('.//a:srgbClr', NS)
                col = c.get('val') if c is not None else ('none' if ln.find('a:noFill', NS) is not None else None)
            item['line'] = col
            item['text'] = ''.join(tt.text or '' for tt in ch.findall('.//a:t', NS))
        out.append(item)


def sheet_items(book, sheet_part):
    """Every picture / shape on the sheet with its box in EMU, anchor-corrected.
    Each item also carries 'anchor' = (col0, row0) of its top-left cell."""
    Y = book.row_y(sheet_part)
    out = []
    for dp in book.drawing_parts(sheet_part):
        media = book.rels(dp)
        for z, anc in enumerate(book.xml(dp)):
            n0 = len(out)
            _collect(anc, (0, 0, 1, 1, 0, 0), out, media)
            if len(out) == n0:
                continue
            f, t = anc.find('xdr:from', NS), anc.find('xdr:to', NS)
            for it in out[n0:]:
                it['z'] = z
            if f is None:           # absoluteAnchor: xfrm is the truth
                continue
            col0, row0 = int(f.find('xdr:col', NS).text), int(f.find('xdr:row', NS).text)
            ay0 = Y(row0, f.find('xdr:rowOff', NS).text)
            if t is not None:
                ay1 = Y(int(t.find('xdr:row', NS).text), t.find('xdr:rowOff', NS).text)
            else:
                ay1 = ay0 + int(anc.find('xdr:ext', NS).get('cy'))
            top = min(i['box'][1] for i in out[n0:])
            bot = max(i['box'][3] for i in out[n0:])
            sc = (ay1 - ay0) / max(1, bot - top)
            for it in out[n0:]:
                b = it['box']
                it['box'] = [b[0], ay0 + (b[1] - top) * sc, b[2], ay0 + (b[3] - top) * sc]
                it['anchor'] = (col0, row0)
    return out


def col_name(c0):
    s, c = '', c0 + 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def px(v):
    return int(round(v / EMU_PX))


def list_book(path):
    bk = Book(path)
    for name, part, state in bk.sheets():
        items = sheet_items(bk, part)
        if not items:
            continue
        print('### %s%s' % (name, '' if state == 'visible' else ' [%s]' % state))
        for it in items:
            a = it.get('anchor')
            where = '%s%d' % (col_name(a[0]), a[1] + 1) if a else 'abs'
            box = [px(v) for v in it['box']]
            extra = ''
            if it['kind'] == 'pic':
                extra = 'media=%s' % posixpath.basename(it.get('media') or '')
            elif it['kind'] in ('sp', 'cxnSp'):
                extra = 'line=%s text=%r' % (it.get('line'), it.get('text', '')[:60])
            print('  z%-2s %-6s %-4s %-28r px%s %s' % (it.get('z', ''), where, it['kind'],
                                                    it['name'][:28], box, extra))


def render_book(path, out_dir, scale=1.0):
    from PIL import Image, ImageDraw
    import io
    bk = Book(path)
    os.makedirs(out_dir, exist_ok=True)
    written = []
    for i, (name, part, state) in enumerate(bk.sheets()):
        items = sheet_items(bk, part)
        if not items:
            continue
        W = int(max(it['box'][2] for it in items) / EMU_PX * scale) + 20
        H = int(max(it['box'][3] for it in items) / EMU_PX * scale) + 20
        im = Image.new('RGB', (W, H), 'white')
        d = ImageDraw.Draw(im)

        def sb(b):
            return [int(v / EMU_PX * scale) for v in b]
        for it in items:
            if it['kind'] != 'pic':
                continue
            x0, y0, x1, y1 = sb(it['box'])
            try:
                p = Image.open(io.BytesIO(bk.z.read(it['media']))).convert('RGBA')
                bg = Image.new('RGBA', p.size, 'white')
                p = Image.alpha_composite(bg, p).convert('RGB')
                c = it.get('crop') or {}
                if c:
                    pw, ph = p.size
                    p = p.crop((int(pw * c.get('l', 0) / 1e5), int(ph * c.get('t', 0) / 1e5),
                                int(pw * (1 - c.get('r', 0) / 1e5)), int(ph * (1 - c.get('b', 0) / 1e5))))
                im.paste(p.resize((max(1, x1 - x0), max(1, y1 - y0))), (x0, y0))
            except Exception as ex:   # unreadable media (wmf etc.): mark the spot
                d.rectangle((x0, y0, x1, y1), outline='blue')
                d.text((x0 + 3, y0 + 3), 'IMG? %s' % str(ex)[:40], fill='blue')
        for it in items:
            if it['kind'] not in ('sp', 'cxnSp'):
                continue
            x0, y0, x1, y1 = sb(it['box'])
            col = '#' + it['line'] if it.get('line') and it['line'] != 'none' else '#00AA00'
            d.rectangle((x0, y0, max(x0 + 1, x1), max(y0 + 1, y1)), outline=col, width=max(1, int(2 * scale)))
            if it.get('text'):
                d.text((x0 + 2, y0 + 2), it['text'][:60], fill='purple')
        fn = os.path.join(out_dir, '%d_%s.png' % (i + 1, name))
        im.save(fn)
        written.append(fn)
    return written


if __name__ == '__main__':
    if len(sys.argv) >= 3 and sys.argv[1] == 'list':
        list_book(sys.argv[2])
    elif len(sys.argv) >= 4 and sys.argv[1] == 'render':
        for f in render_book(sys.argv[2], sys.argv[3]):
            print(f)
    else:
        print(__doc__)
        sys.exit(2)
