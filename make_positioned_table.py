"""Write the .doc files of the positioned-table reproducer, byte by byte, from [MS-DOC] and [MS-CFB].

No application writes these files: this script does. Text is invented. Licence: MIT OR Apache-2.0.

    python make_positioned_table.py      # writes the three .doc files next to this script

positioned-table-row.doc          the row carries sprmTPc, sprmTDxaAbs, sprmTDyaAbs,
                                  sprmTDxaFromText and sprmTDxaFromTextRight
positioned-table-cell-frames.doc  only the cell paragraphs carry the frame sprms
                                  (sprmPPc, sprmPDxaAbs, sprmPDyaAbs, sprmPWr, sprmPDxaFromText)
positioned-table-inline.doc       neither: an ordinary table

Each has three paragraphs, a three-row table centred between the margins, five empty paragraphs
and a paragraph with text whose first word is wider than the space beside the table.

Sources: [MS-DOC] (the FIB, the Clx and its PlcPcd, ChpxFkp and PapxFkp with their PlcBte,
the STSH, PlcfSed and SEPX, PlcfHdd, the Dop, the sprm tables, TDefTableOperand and
PositionCodeOperand) and [MS-CFB] (a version 3 compound file).
"""
import struct

ISTD_NIL = 0x0FFF
TEXT_FC = 0x400
PAGE = 512


def sprm(op, operand):
    return struct.pack('<H', op) + operand


def u8(v):
    return struct.pack('<B', v)


def i16(v):
    return struct.pack('<h', v)


def u16(v):
    return struct.pack('<H', v)


def u32(v):
    return struct.pack('<I', v)


# ---- common sprms ([MS-DOC] §2.6.1-2.6.4)
def p_in_table():
    return sprm(0x2416, u8(1)) + sprm(0x6649, u32(1))


def p_ttp():
    return sprm(0x2417, u8(1))


def p_after(tw):
    return sprm(0xA414, u16(tw))


def p_before(tw):
    return sprm(0xA413, u16(tw))


def p_line(dya, multiple=True):
    return sprm(0x6412, i16(dya) + i16(1 if multiple else 0))


def c_size(halfpts):
    return sprm(0x4A43, u16(halfpts))


def c_font(ftc):
    return sprm(0x4A4F, u16(ftc)) + sprm(0x4A51, u16(ftc))


def t_def_table(widths, left=0):
    """sprmTDefTable: itcMac, rgdxaCenter, rgTc80 with single 1/2 pt borders."""
    n = len(widths)
    xs = [left]
    for w in widths:
        xs.append(xs[-1] + w)
    body = u8(n) + b''.join(i16(x) for x in xs)
    brc = bytes([4, 1, 0, 0])  # Brc80: dptLineWidth 4 (1/2 pt), brcType 1 (single), ico 0
    for w in widths:
        body += u16(0) + u16(w) + brc * 4
    return sprm(0xD608, u16(len(body) + 1) + body)


def t_row_height(tw):
    """sprmTDyaRowHeight: positive = at least, negative = exact."""
    return sprm(0x9407, i16(tw))


def t_gap_half(tw):
    return sprm(0x9602, i16(tw))


class Para:
    def __init__(self, text='', istd=0, papx=b'', chpx=b'', mark_chpx=None):
        self.text = text
        self.istd = istd
        self.papx = papx
        self.chpx = chpx
        self.mark_chpx = chpx if mark_chpx is None else mark_chpx


class Table:
    """rows: list of (cells, tapx); cells: list of lists of Para."""

    def __init__(self, rows, cell_papx=b''):
        self.rows = rows
        self.cell_papx = cell_papx


class Doc:
    def __init__(self):
        self.body = []
        self.header = None  # list of blocks (odd/default header)
        self.footer = None
        self.sepx = (sprm(0xB01F, u16(11906)) + sprm(0xB020, u16(16838)) + sprm(0xB021, u16(1134))
                     + sprm(0xB022, u16(1134)) + sprm(0x9023, i16(1417)) + sprm(0x9024, i16(1134))
                     + sprm(0xB017, u16(708)) + sprm(0xB018, u16(708)))
        self.fonts = ['Times New Roman', 'Arial', 'Tahoma']
        # istd -> (name, sti, base, papx grpprl, chpx grpprl)
        self.styles = {
            0: ('Normal', 0, ISTD_NIL, p_after(200) + p_line(276), c_size(20) + c_font(1)),
            13: ('header', 31, 0, b'', b''),
            14: ('footer', 32, 0, b'', b''),
        }


class _Text:
    def __init__(self):
        self.units = []
        self.paras = []  # (end_cp, istd, papx)
        self.runs = []  # (start_cp, end_cp, chpx)

    def cp(self):
        return len(self.units)

    def run(self, s, chpx):
        a = self.cp()
        self.units.extend(ord(c) for c in s)
        if self.cp() > a:
            self.runs.append((a, self.cp(), chpx))

    def mark(self, ch, istd, papx, chpx):
        self.run(ch, chpx)
        self.paras.append((self.cp(), istd, papx))

    def blocks(self, blocks):
        for b in blocks:
            if isinstance(b, Para):
                self.run(b.text, b.chpx)
                self.mark('\r', b.istd, b.papx, b.mark_chpx)
            else:
                for cells, tapx in b.rows:
                    for cell in cells:
                        for k, p in enumerate(cell):
                            self.run(p.text, p.chpx)
                            last = k == len(cell) - 1
                            self.mark('\x07' if last else '\r', p.istd, p_in_table() + b.cell_papx + p.papx, p.mark_chpx)
                    self.mark('\x07', 0, p_in_table() + p_ttp() + tapx, b'')


def _fkp_pages(entries, entry_size, encode, max_runs):
    """entries: (fc_start, fc_end, data-in-page-form). Returns (pages, fcs)."""
    pages, fcs = [], []
    cur, low, stored = [], PAGE - 1, {}
    buf = bytearray(PAGE)

    def header(n):
        return (n + 1) * 4 + n * entry_size

    def flush(end_fc):
        nonlocal buf, cur, low, stored
        crun = len(cur)
        for i, (fc, _) in enumerate(cur):
            buf[i * 4:i * 4 + 4] = u32(fc)
        buf[crun * 4:crun * 4 + 4] = u32(end_fc)
        for i, (_, off) in enumerate(cur):
            buf[(crun + 1) * 4 + i * entry_size] = off
        buf[PAGE - 1] = crun
        pages.append(bytes(buf))
        fcs.append(cur[0][0])
        buf, cur, low, stored = bytearray(PAGE), [], PAGE - 1, {}

    for fc_a, fc_b, data in entries:
        data = encode(data)
        for attempt in range(2):
            ok = len(cur) < max_runs
            off = 0
            if ok and data:
                if data in stored:
                    off = stored[data]
                    ok = header(len(cur) + 1) <= low
                else:
                    at = (low - len(data)) & ~1
                    ok = at >= header(len(cur) + 1)
                    if ok:
                        buf[at:at + len(data)] = data
                        low = at
                        off = at // 2
                        stored[data] = off
            elif ok:
                ok = header(len(cur) + 1) <= low
            if ok:
                cur.append((fc_a, off))
                break
            flush(fc_a)
    flush(entries[-1][1])
    fcs.append(entries[-1][1])
    return pages, fcs


def _chpx_enc(g):
    return (u8(len(g)) + g) if g else b''


def _papx_enc(istd_g):
    istd, g = istd_g
    body = u16(istd) + g
    if len(body) % 2 == 1:
        return u8((len(body) + 1) // 2) + body
    return u8(0) + u8(len(body) // 2) + body


def _std(istd, name, sti, stk, base, papx, chpx):
    if stk == 1:
        upx = [u16(istd) + papx, chpx]
    elif stk == 2:
        upx = [chpx]
    elif stk == 3:
        upx = [b'', u16(istd), b'']
    else:
        upx = [u16(istd)]
    nxt = istd if stk == 1 else ISTD_NIL
    v = u16(sti & 0x0FFF) + u16((stk & 0xF) | ((base & 0x0FFF) << 4)) + u16((len(upx) & 0xF) | (nxt << 4))
    v += u16(0) + u16(0)  # bchUpe (patched), grfstd
    units = name.encode('utf-16-le')
    v += u16(len(name)) + units + u16(0)
    for u in upx:
        v += u16(len(u)) + u
        if len(u) % 2:
            v += b'\0'
    v = v[:6] + u16(len(v)) + v[8:]
    return v


def _stsh(doc):
    slots = {i: None for i in range(15)}
    slots[10] = ('Default Paragraph Font', 65, 2)
    slots[11] = ('Normal Table', 105, 3)
    slots[12] = ('No List', 107, 4)
    sti_max = 0x5B
    cb_stshi = 18 + 2 + 2 + 4 * sti_max
    cstd = max(max(slots), max(doc.styles)) + 1
    v = u16(cb_stshi) + u16(cstd) + u16(0x0A) + u16(1) + u16(sti_max) + u16(15) + u16(0)
    v += u16(1) + u16(1) + u16(1) + u16(1)  # rgftcStandardChpStsh, ftcBi
    v += u16(4) + b'\0' * (4 * sti_max)
    for istd in range(cstd):
        if istd in doc.styles:
            name, sti, base, papx, chpx = doc.styles[istd]
            s = _std(istd, name, sti, 1, base, papx, chpx)
        elif slots.get(istd):
            name, sti, stk = slots[istd]
            s = _std(istd, name, sti, stk, ISTD_NIL, b'', b'')
        else:
            s = b''
        v += u16(len(s)) + s
    return v


def _fonts(names):
    v = u16(len(names)) + u16(0)
    for n in names:
        sym = n.lower() in ('symbol', 'wingdings')
        fam = 5 if sym else (2 if n in ('Arial', 'Tahoma', 'Calibri') else 1)
        f = u8(2 | 4 | (fam << 4)) + u16(400) + u8(2 if sym else 0) + u8(0) + b'\0' * 34
        f += n.encode('utf-16-le') + u16(0)
        v += u8(len(f)) + f
    return v


def _sttb_ext(n):
    return u16(0xFFFF) + u16(n) + u16(0) + u16(0) * n


def _dop():
    v = bytearray(500)
    v[0] = 1 << 5
    v[2:4] = u16(1 << 2)
    v[10:12] = u16(720)
    v[14:16] = u16(360)
    v[32:34] = u16(1)
    v[52:54] = u16(1 << 2)
    v[54:56] = u16(3)
    v[82:84] = u16((100 << 3) | (1 << 12))
    return bytes(v)


def build(doc, extra_streams=None):
    t = _Text()
    t.blocks(doc.body)
    if not t.paras or t.paras[-1][0] != t.cp() or not isinstance(doc.body[-1], Para):
        t.mark('\r', 0, b'', b'')
    ccp_text = t.cp()
    ccp_hdd = 0
    hdd_cps = []
    if doc.header or doc.footer:
        start = t.cp()
        hdd_cps = [0] * 6
        for story in (None, doc.header, None, doc.footer, None, None):
            hdd_cps.append(t.cp() - start)
            if story is None:
                continue
            t.blocks(story)
            t.mark('\r', 0, b'', b'')  # guard mark ending the story
        t.mark('\r', 0, b'', b'')
        ccp_hdd = t.cp() - start
        t.mark('\r', 0, b'', b'')  # one more mark after the last subdocument
    ccp_all = t.cp()

    word = bytearray(TEXT_FC)
    for u in t.units:
        word += u16(u)

    def fc(cp):
        return TEXT_FC + cp * 2

    while len(word) % PAGE:
        word.append(0)
    chpx_pages, chpx_fcs = _fkp_pages([(fc(a), fc(b), g) for a, b, g in t.runs], 1, _chpx_enc, 0x65)
    chpx_pn = len(word) // PAGE
    for p in chpx_pages:
        word += p
    papx_entries = []
    start = 0
    for end, istd, g in t.paras:
        papx_entries.append((fc(start), fc(end), (istd, g)))
        start = end
    papx_pages, papx_fcs = _fkp_pages(papx_entries, 13, _papx_enc, 0x1D)
    papx_pn = len(word) // PAGE
    for p in papx_pages:
        word += p
    while len(word) % 2:
        word.append(0)
    sepx_at = len(word)
    word += u16(len(doc.sepx)) + doc.sepx
    cb_mac = len(word)

    tbl = bytearray()
    pairs = {}

    def put(i, b):
        while len(tbl) % 2:
            tbl.append(0)
        pairs[i] = (len(tbl), len(b))
        tbl.extend(b)

    stsh = _stsh(doc)
    put(1, stsh)
    pairs[0] = pairs[1]
    sed = u16(0) + u32(sepx_at) + u16(0) + u32(0xFFFFFFFF)
    put(6, u32(0) + u32(ccp_text) + sed)
    if ccp_hdd:
        cps = hdd_cps + [ccp_hdd - 1, ccp_hdd]
        put(11, b''.join(u32(c) for c in cps))

    def bte(fcs, pn):
        return b''.join(u32(f) for f in fcs) + b''.join(u32(pn + i) for i in range(len(fcs) - 1))

    put(12, bte(chpx_fcs, chpx_pn))
    put(13, bte(papx_fcs, papx_pn))
    put(31, _dop())
    put(32, _sttb_ext(18))
    plc = u32(0) + u32(ccp_all) + u16(0) + u32(TEXT_FC) + u16(0)
    put(33, b'\x02' + u32(len(plc)) + plc)
    put(15, _fonts(doc.fonts))

    fib = bytearray(0x9A + 0x5D * 8 + 2)
    struct.pack_into('<H', fib, 0, 0xA5EC)
    struct.pack_into('<H', fib, 2, 0x00C1)
    struct.pack_into('<H', fib, 6, 0x0409)
    struct.pack_into('<H', fib, 0x0A, 0x0200 | 0x1000)
    struct.pack_into('<H', fib, 0x0C, 0x00BF)
    struct.pack_into('<H', fib, 0x20, 0x000E)
    struct.pack_into('<H', fib, 0x22 + 26, 0x0409)
    struct.pack_into('<H', fib, 0x3E, 0x0016)
    struct.pack_into('<I', fib, 0x40, cb_mac)
    struct.pack_into('<I', fib, 0x40 + 12, ccp_text)
    struct.pack_into('<I', fib, 0x40 + 20, ccp_hdd)
    struct.pack_into('<H', fib, 0x98, 0x005D)
    for i, (a, n) in pairs.items():
        struct.pack_into('<II', fib, 0x9A + i * 8, a, n)
    word[:len(fib)] = fib
    streams = [('WordDocument', bytes(word)), ('1Table', bytes(tbl))]
    if extra_streams:
        streams += extra_streams
    return cfb(streams)


def cfb(streams):
    """[MS-CFB] version 3; every stream padded to the 4096-byte cutoff so no mini stream is needed."""
    sec = 512
    data_secs = []  # (name, start, nsec, size)
    body = bytearray()
    nsec = 0
    padded = []
    for name, d in streams:
        d = d + b'\0' * max(0, 4096 - len(d))
        padded.append((name, d))
    for name, d in padded:
        n = (len(d) + sec - 1) // sec
        data_secs.append((name, nsec, n, len(d)))
        body += d + b'\0' * (n * sec - len(d))
        nsec += n
    dir_start = nsec
    n_entries = 1 + len(streams)
    dir_secs = (n_entries * 128 + sec - 1) // sec
    nsec += dir_secs
    fat_secs = 1
    while (nsec + fat_secs) > fat_secs * (sec // 4):
        fat_secs += 1
    fat_start = nsec
    total = nsec + fat_secs
    fat = [0xFFFFFFFF] * (fat_secs * (sec // 4))
    for _, s, n, _ in data_secs:
        for k in range(n):
            fat[s + k] = s + k + 1 if k < n - 1 else 0xFFFFFFFE
    for k in range(dir_secs):
        fat[dir_start + k] = dir_start + k + 1 if k < dir_secs - 1 else 0xFFFFFFFE
    for k in range(fat_secs):
        fat[fat_start + k] = 0xFFFFFFFD

    # directory: root, then streams; a red-black tree with the middle entry as the root's child
    def key(n):
        return (len(n), n.upper())

    order = sorted(range(len(streams)), key=lambda i: key(streams[i][0]))
    left = {}
    right = {}
    color = {}
    mid = len(order) // 2

    def place(lo, hi, depth):
        if lo > hi:
            return 0xFFFFFFFF
        m = (lo + hi) // 2
        i = order[m]
        left[i] = place(lo, m - 1, depth + 1)
        right[i] = place(m + 1, hi, depth + 1)
        color[i] = 1
        return i + 1

    assert len(streams) <= 3, 'tree colouring below covers depth 2 only'
    root_child = place(0, len(order) - 1, 0)
    # the root's child is black, its children red: equal black heights on every path
    for i in range(len(streams)):
        color[i] = 1 if i + 1 == root_child else 0

    def entry(name, typ, col, l, r, child, start, size):
        nm = name.encode('utf-16-le') + b'\0\0'
        e = nm.ljust(64, b'\0') + u16(len(nm)) + u8(typ) + u8(col)
        e += u32(l) + u32(r) + u32(child) + b'\0' * 16 + u32(0) + b'\0' * 16
        e += u32(start) + u32(size) + u32(0)
        return e

    d = entry('Root Entry', 5, 1, 0xFFFFFFFF, 0xFFFFFFFF, root_child, 0xFFFFFFFE, 0)
    for i, (name, s, n, size) in enumerate(data_secs):
        d += entry(name, 2, color[i], left[i], right[i], 0xFFFFFFFF, s, size)
    d = d.ljust(dir_secs * sec, b'\0')
    # unused directory slots must be empty entries with sibling ids NOSTREAM
    for k in range(n_entries, dir_secs * sec // 128):
        off = k * 128
        d = d[:off + 68] + u32(0xFFFFFFFF) * 3 + d[off + 80:]
    hdr = bytes.fromhex('D0CF11E0A1B11AE1') + b'\0' * 16 + u16(0x3E) + u16(3) + u16(0xFFFE) + u16(9) + u16(6)
    hdr += b'\0' * 6 + u32(0) + u32(fat_secs) + u32(dir_start) + u32(0) + u32(4096)
    hdr += u32(0xFFFFFFFE) + u32(0) + u32(0xFFFFFFFE) + u32(0)
    difat = [fat_start + k for k in range(fat_secs)] + [0xFFFFFFFF] * (109 - fat_secs)
    hdr += b''.join(u32(x) for x in difat)
    out = hdr + bytes(body) + d + b''.join(u32(x) for x in fat)
    assert len(out) == sec * (1 + total)
    return out


# ---- the reproducer
TAPX_POS = (sprm(0x360D, u8(0x60))      # sprmTPc: vertical from the paragraph, horizontal from the margins
            + sprm(0x940E, i16(-4))     # sprmTDxaAbs: centred
            + sprm(0x940F, i16(120))    # sprmTDyaAbs: twips below the anchor paragraph
            + sprm(0x9410, i16(144))    # sprmTDxaFromText
            + sprm(0x941E, i16(144)))   # sprmTDxaFromTextRight
PAPX_FRAME = (sprm(0x261B, u8(0x60)) + sprm(0x8418, i16(-4)) + sprm(0x8419, i16(120))
              + sprm(0x2423, u8(2)) + sprm(0x842F, i16(144)))


def make(variant):
    d = Doc()
    d.body = [Para('Invented paragraph %d before the table.' % i) for i in range(3)]
    tapx = TAPX_POS if variant == 'row' else b''
    cell_papx = PAPX_FRAME if variant == 'cell-frames' else b''
    rows = []
    for k, h in enumerate((400, 700, 1000)):  # row heights, "at least"
        cells = [[Para('Row %d cell %d' % (k, c), papx=p_after(0))] for c in range(3)]
        rows.append((cells, t_def_table([2200, 3800, 2300]) + t_gap_half(108) + t_row_height(h) + tapx))
    d.body.append(Table(rows, cell_papx=cell_papx))
    d.body += [Para('', papx=p_after(0) + p_line(240)) for _ in range(5)]
    d.body.append(Para('Afterwards, a paragraph with text.'))
    return build(d)


if __name__ == '__main__':
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    for v in ('row', 'cell-frames', 'inline'):
        with open(os.path.join(here, 'positioned-table-%s.doc' % v), 'wb') as f:
            f.write(make(v))