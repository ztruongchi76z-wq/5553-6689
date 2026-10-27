# -*- coding: utf-8 -*-
"""Công cụ sửa trực tiếp tệp Word thuyết minh (dùng cho tiep_thu_gop_y.py).

Quy ước chuỗi định dạng (markup) trong mọi hàm: ^{...} chỉ số trên, _{...} chỉ số dưới, **...** chữ đậm,
//...// chữ nghiêng, \\n xuống dòng trong đoạn.
"""
import copy
import re

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

NHAT_KY = []


def ghi(msg):
    NHAT_KY.append(msg)
    print('  •', msg)


# --------------------------------------------------------------------------- đọc
def van_ban(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def kieu(p):
    s = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
    return s.get(qn('w:val')) if s is not None else None


def cac_doan(body):
    return list(body.iterchildren(qn('w:p')))


def tim_doan(body, dau, style=None, chua=None, so_luong=1):
    """Đoạn (con trực tiếp của body, sau MỤC LỤC/danh mục) bắt đầu bằng `dau` (và chứa `chua` nếu có)."""
    kq = []
    for p in body.iterchildren(qn('w:p')):
        k = kieu(p) or ''
        if k.startswith('TOC') or k == 'TableofFigures':
            continue
        t = van_ban(p)
        if t.startswith(dau) and (chua is None or chua in t) and (style is None or k == style):
            kq.append(p)
    if so_luong is not None and len(kq) != so_luong:
        raise ValueError('Tìm thấy %d đoạn bắt đầu bằng %r' % (len(kq), dau))
    return kq[0] if so_luong == 1 else kq


def tim_bang(body, dau_chu_thich):
    """(đoạn chú thích, bảng) với chú thích kiểu Chuthichbang bắt đầu bằng `dau_chu_thich`."""
    cap = tim_doan(body, dau_chu_thich, style='Chuthichbang')
    tbl = cap.getnext()
    assert tbl.tag == qn('w:tbl'), dau_chu_thich
    return cap, tbl


def hang(tbl):
    return tbl.findall(qn('w:tr'))


def o(tbl, r, c):
    return hang(tbl)[r].findall(qn('w:tc'))[c]


# --------------------------------------------------------------------------- tạo run từ markup
_RX_KHOI = re.compile(r'\*\*(.+?)\*\*|//(.+?)//', re.S)
_RX = re.compile(r'\^\{([^}]*)\}|_\{([^}]*)\}|([^\^_]+|[\^_])')


def _dat_rpr(rpr, dam=None, nghieng=None, vert=None):
    """Thêm/bỏ b, i, vertAlign trong rPr theo đúng thứ tự lược đồ (b, i ở đầu; vertAlign sau sz)."""
    for tag, bat in (('w:b', dam), ('w:i', nghieng)):
        if bat is None:
            continue
        cu = rpr.find(qn(tag))
        if bat and cu is None:
            e = OxmlElement(tag)
            # sau rStyle/rFonts/b, trước các phần tử khác
            truoc = [x for x in rpr if x.tag in (qn('w:rStyle'), qn('w:rFonts'), qn('w:b'), qn('w:bCs'))]
            if truoc:
                truoc[-1].addnext(e)
            else:
                rpr.insert(0, e)
        elif not bat and cu is not None:
            rpr.remove(cu)
    for e in rpr.findall(qn('w:vertAlign')):
        rpr.remove(e)
    if vert:
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), vert)
        sau = [x for x in rpr if x.tag in (qn('w:lang'), qn('w:eastAsianLayout'), qn('w:specVanish'),
                                           qn('w:oMath'))]
        if sau:
            sau[0].addprevious(va)
        else:
            rpr.append(va)


def _runs_doan(chu_doan, rpr_mau, dam, nghieng):
    out = []
    for m in _RX.finditer(chu_doan):
        vert = None
        if m.group(1) is not None:
            chu, vert = m.group(1), 'superscript'
        elif m.group(2) is not None:
            chu, vert = m.group(2), 'subscript'
        else:
            chu = m.group(3)
        r = OxmlElement('w:r')
        rpr = copy.deepcopy(rpr_mau) if rpr_mau is not None else OxmlElement('w:rPr')
        _dat_rpr(rpr, dam=dam, nghieng=nghieng, vert=vert)
        if len(rpr):
            r.append(rpr)
        for i, phan in enumerate(chu.split('\n')):
            if i:
                r.append(OxmlElement('w:br'))
            if phan:
                t = OxmlElement('w:t')
                t.text = phan
                t.set(XML_SPACE, 'preserve')
                r.append(t)
        out.append(r)
    return out


def tao_runs(markup, rpr_mau=None, dam_mac_dinh=False, nghieng_mac_dinh=False):
    """Run từ markup: **đậm**, //nghiêng// (bên trong vẫn dùng được ^{..}, _{..}), ^{trên}, _{dưới}."""
    out = []
    vt = 0
    for m in _RX_KHOI.finditer(markup):
        if m.start() > vt:
            out += _runs_doan(markup[vt:m.start()], rpr_mau, dam_mac_dinh, nghieng_mac_dinh)
        if m.group(1) is not None:
            out += _runs_doan(m.group(1), rpr_mau, True, nghieng_mac_dinh)
        else:
            out += _runs_doan(m.group(2), rpr_mau, dam_mac_dinh, True)
        vt = m.end()
    if vt < len(markup):
        out += _runs_doan(markup[vt:], rpr_mau, dam_mac_dinh, nghieng_mac_dinh)
    return out


def rpr_cua(p):
    """rPr của run chữ đầu tiên trong đoạn (đã bỏ b, i, vertAlign) – làm mẫu cỡ chữ."""
    for r in p.iter(qn('w:r')):
        if r.find(qn('w:t')) is not None:
            rp = r.find(qn('w:rPr'))
            if rp is None:
                return None
            rp = copy.deepcopy(rp)
            for tag in ('w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:vertAlign'):
                for e in rp.findall(qn(tag)):
                    rp.remove(e)
            return rp
    return None


# --------------------------------------------------------------------------- sửa đoạn
def _runs_chu(p):
    return [r for r in p.iterchildren(qn('w:r')) if r.find(qn('w:t')) is not None]


def _dat_chu(r, s):
    ts = r.findall(qn('w:t'))
    ts[0].text = s
    ts[0].set(XML_SPACE, 'preserve')
    for t in ts[1:]:
        r.remove(t)


def thay(p, cu, moi, lan=1):
    """Thay chuỗi `cu` (có thể trải qua nhiều run, phải xuất hiện đúng `lan` lần) bằng markup `moi`.
    Run mới thừa hưởng định dạng (rPr) của run chứa đầu chuỗi cũ."""
    runs = _runs_chu(p)
    chu = [''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in runs]
    toan = ''.join(chu)
    if toan.count(cu) != lan:
        raise ValueError('Chuỗi %r xuất hiện %d lần (cần %d) trong đoạn: %r' % (cu, toan.count(cu), lan, toan[:90]))
    a = toan.index(cu)
    b = a + len(cu)
    vt = 0
    i0 = i1 = None
    for i, c in enumerate(chu):
        if i0 is None and a < vt + len(c):
            i0, o0 = i, a - vt
        if b <= vt + len(c):
            i1, o1 = i, b - vt
            break
        vt += len(c)
    r0, r1 = runs[i0], runs[i1]
    dau, cuoi = chu[i0][:o0], chu[i1][o1:]
    rp = r0.find(qn('w:rPr'))
    rp_mau = copy.deepcopy(rp) if rp is not None else None
    dam0 = rp_mau is not None and rp_mau.find(qn('w:b')) is not None
    ngh0 = rp_mau is not None and rp_mau.find(qn('w:i')) is not None
    if rp_mau is not None:
        for e in rp_mau.findall(qn('w:vertAlign')):
            rp_mau.remove(e)
    if dau:
        rd = copy.deepcopy(r0)
        _dat_chu(rd, dau)
        r0.addprevious(rd)
    for r in tao_runs(moi, rp_mau, dam_mac_dinh=dam0, nghieng_mac_dinh=ngh0):
        r0.addprevious(r)
    if cuoi:
        rc = copy.deepcopy(r1)
        _dat_chu(rc, cuoi)
        r0.addprevious(rc)
    for r in runs[i0:i1 + 1]:
        p.remove(r)


def thay_tu(p, cu_dau, moi):
    """Thay từ vị trí chuỗi `cu_dau` đến hết đoạn bằng `moi`."""
    toan = van_ban(p)
    a = toan.index(cu_dau)
    thay(p, toan[a:], moi)


def them_cuoi(p, markup):
    runs = _runs_chu(p)
    neo = runs[-1]
    rp = rpr_cua(p)
    for r in tao_runs(markup, rp):
        neo.addnext(r)
        neo = r


def dat_noi_dung(p, markup):
    """Xóa mọi run/hyperlink/công thức của đoạn (giữ pPr, bookmark) rồi ghi markup mới."""
    rp = rpr_cua(p)
    for e in list(p):
        if e.tag in (qn('w:r'), qn('w:hyperlink'), '{%s}oMath' % M, '{%s}oMathPara' % M):
            p.remove(e)
    for r in tao_runs(markup, rp):
        p.append(r)


def doan_moi(markup, mau, kieu_doan=None):
    """Đoạn mới (pPr sao từ `mau`; đổi kiểu nếu có) – chưa chèn."""
    p = OxmlElement('w:p')
    ppr = mau.find(qn('w:pPr'))
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for e in ppr.findall(qn('w:numPr')):
            if kieu_doan is not None:
                ppr.remove(e)
        p.append(ppr)
    if kieu_doan is not None:
        ppr = p.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr')
            p.append(ppr)
        ps = ppr.find(qn('w:pStyle'))
        if ps is None:
            ps = OxmlElement('w:pStyle')
            ppr.insert(0, ps)
        ps.set(qn('w:val'), kieu_doan)
    for r in tao_runs(markup, rpr_cua(mau)):
        p.append(r)
    return p


def chen_sau(neo, *phan_tu):
    for e in phan_tu:
        neo.addnext(e)
        neo = e
    return neo


def chen_truoc(neo, *phan_tu):
    for e in phan_tu:
        neo.addprevious(e)
    return phan_tu[-1]


# --------------------------------------------------------------------------- bảng
def dat_o(tc, markup, dam=None, can=None):
    """Ghi nội dung ô (giữ pPr, rPr cỡ chữ của ô; nếu ô rỗng lấy rPr của ô khác cùng hàng)."""
    p = tc.find(qn('w:p'))
    rp = rpr_cua(p)
    if rp is None:
        for tc2 in tc.getparent().findall(qn('w:tc')):
            rp = rpr_cua(tc2.find(qn('w:p')))
            if rp is not None:
                break
    dam0 = False
    for r in p.iter(qn('w:r')):
        if r.find(qn('w:t')) is not None:
            rr = r.find(qn('w:rPr'))
            dam0 = rr is not None and rr.find(qn('w:b')) is not None
            break
    if dam is not None:
        dam0 = dam
    for e in list(p):
        if e.tag in (qn('w:r'), qn('w:hyperlink')):
            p.remove(e)
    for ph in tc.findall(qn('w:p'))[1:]:
        tc.remove(ph)
    for r in tao_runs(markup, rp, dam_mac_dinh=dam0):
        p.append(r)
    if can:
        jc = p.find(qn('w:pPr') + '/' + qn('w:jc'))
        if jc is not None:
            jc.set(qn('w:val'), can)


def them_hang(tbl, sau_hang, gia_tri, mau_hang=None, dam=None):
    """Chèn hàng mới (sao từ mau_hang, mặc định là hàng sau_hang) ngay sau hàng chỉ số `sau_hang`."""
    rows = hang(tbl)
    mau = rows[sau_hang if mau_hang is None else mau_hang]
    moi = copy.deepcopy(mau)
    for tc, gt in zip(moi.findall(qn('w:tc')), gia_tri):
        dat_o(tc, gt, dam=dam)
    rows[sau_hang].addnext(moi)
    return moi


def bang_moi(tieu_de, cot, dong, rong, mau_cap, can=None, co_chu=20, dam_cot_dau=False):
    """(chú thích, bảng) mới theo đúng định dạng các bảng của thuyết minh (TableGrid, đầu bảng nền F2F2F2,
    cỡ chữ co_chu/2 pt, dòng đơn, cách 1 pt). `cot`: tiêu đề cột; `dong`: danh sách hàng (markup);
    `rong`: bề rộng cột (dxa); `can`: căn lề từng cột ('left'/'center'/'right')."""
    cap = copy.deepcopy(mau_cap)
    for e in list(cap):
        if e.tag != qn('w:pPr'):
            cap.remove(e)
    so, ten = tieu_de.split(' ', 2)[:2], tieu_de.split(' ', 2)[2]
    r1 = tao_runs(' '.join(so) + ' ')[0]
    cap.append(r1)
    for r in tao_runs(ten):
        cap.append(r)
    tbl = OxmlElement('w:tbl')
    tpr = etree.SubElement(tbl, qn('w:tblPr'))
    for tag, att in (('w:tblStyle', {'w:val': 'TableGrid'}), ('w:tblW', {'w:type': 'auto', 'w:w': '0'}),
                     ('w:jc', {'w:val': 'center'}), ('w:tblLayout', {'w:type': 'fixed'}),
                     ('w:tblLook', {'w:firstColumn': '1', 'w:firstRow': '1', 'w:lastColumn': '0', 'w:lastRow': '0',
                                    'w:noHBand': '0', 'w:noVBand': '1', 'w:val': '04A0'})):
        e = etree.SubElement(tpr, qn(tag))
        for k, v in att.items():
            e.set(qn(k), v)
    grid = etree.SubElement(tbl, qn('w:tblGrid'))
    for w in rong:
        etree.SubElement(grid, qn('w:gridCol')).set(qn('w:w'), str(w))
    can = can or (['left'] + ['center'] * (len(cot) - 1))

    def _hang(gia_tri, dau):
        tr = etree.SubElement(tbl, qn('w:tr'))
        trpr = etree.SubElement(tr, qn('w:trPr'))
        if dau:
            etree.SubElement(trpr, qn('w:tblHeader')).set(qn('w:val'), 'true')
        etree.SubElement(trpr, qn('w:cantSplit')).set(qn('w:val'), 'true')
        for j, (gt, w) in enumerate(zip(gia_tri, rong)):
            tc = etree.SubElement(tr, qn('w:tc'))
            tcpr = etree.SubElement(tc, qn('w:tcPr'))
            tw = etree.SubElement(tcpr, qn('w:tcW'))
            tw.set(qn('w:type'), 'dxa')
            tw.set(qn('w:w'), str(w))
            if dau:
                sh = etree.SubElement(tcpr, qn('w:shd'))
                sh.set(qn('w:val'), 'clear')
                sh.set(qn('w:color'), 'auto')
                sh.set(qn('w:fill'), 'F2F2F2')
            p = etree.SubElement(tc, qn('w:p'))
            ppr = etree.SubElement(p, qn('w:pPr'))
            sp = etree.SubElement(ppr, qn('w:spacing'))
            for k, v in (('w:before', '20'), ('w:after', '20'), ('w:line', '240'), ('w:lineRule', 'auto')):
                sp.set(qn(k), v)
            etree.SubElement(ppr, qn('w:jc')).set(qn('w:val'), 'center' if dau else can[j])
            rpr = OxmlElement('w:rPr')
            etree.SubElement(rpr, qn('w:sz')).set(qn('w:val'), str(co_chu))
            for r in tao_runs(gt, rpr, dam_mac_dinh=dau or (dam_cot_dau and j == 0)):
                p.append(r)
    _hang(cot, True)
    for d in dong:
        _hang(d, False)
    kc = OxmlElement('w:p')
    kpr = OxmlElement('w:pPr')
    ksp = OxmlElement('w:spacing')
    ksp.set(qn('w:after'), '40')
    kpr.append(ksp)
    kc.append(kpr)
    return cap, tbl, kc


def doi_so_chu_thich(body, cu, moi):
    """Đổi số trong chú thích bảng/hình 'cu' -> 'moi' (vd 'Bảng 4.2' -> 'Bảng 4.3')."""
    n = 0
    for p in body.iterchildren(qn('w:p')):
        if kieu(p) in ('Chuthichbang', 'Chuthichhinh') and van_ban(p).startswith(cu + ' '):
            thay(p, cu + ' ', moi + ' ')
            n += 1
    return n


def doi_tham_chieu(body, cu, moi):
    """Đổi mọi tham chiếu 'cu' trong lời văn và bảng (trừ chú thích, danh mục) – theo từng đoạn."""
    n = 0
    rx = re.compile(re.escape(cu) + r'(?![0-9])')
    for p in body.iter(qn('w:p')):
        k = kieu(p) or ''
        if k in ('Chuthichbang', 'Chuthichhinh') or k.startswith('TOC') or k == 'TableofFigures':
            continue
        t = van_ban(p)
        for _ in range(len(rx.findall(t))):
            # thay lần lượt từng chỗ (chuỗi có thể nằm trên nhiều run)
            runs = _runs_chu(p)
            chu = ''.join(''.join(x.text or '' for x in r.findall(qn('w:t'))) for r in runs)
            m = rx.search(chu)
            if not m:
                break
            _thay_vi_tri(p, m.start(), m.end(), moi)
            n += 1
    return n


def _thay_vi_tri(p, a, b, moi_text):
    runs = _runs_chu(p)
    chu = [''.join(t.text or '' for t in r.findall(qn('w:t'))) for r in runs]
    vt = 0
    i0 = i1 = None
    for i, c in enumerate(chu):
        if i0 is None and a < vt + len(c):
            i0, o0 = i, a - vt
        if b <= vt + len(c):
            i1, o1 = i, b - vt
            break
        vt += len(c)
    if i0 == i1:
        _dat_chu(runs[i0], chu[i0][:o0] + moi_text + chu[i0][o1:])
    else:
        _dat_chu(runs[i0], chu[i0][:o0] + moi_text)
        _dat_chu(runs[i1], chu[i1][o1:])
        for r in runs[i0 + 1:i1]:
            _dat_chu(r, '')


# --------------------------------------------------------------------------- hình
def thay_anh(d, ten_goc, duong_dan_jpg):
    """Thay ảnh có tên gốc (pic:cNvPr name bắt đầu bằng ten_goc + '_') bằng tệp JPG mới; giữ bề rộng,
    cập nhật chiều cao theo tỷ lệ ảnh mới. Trả về (cx, cy)."""
    from PIL import Image
    body = d.element.body
    kq = None
    for pic in body.iter('{%s}pic' % PIC):
        ten = pic.find('{%s}nvPicPr/{%s}cNvPr' % (PIC, PIC)).get('name')
        if re.fullmatch(re.escape(ten_goc) + r'_[0-9a-f]{10}\.jpg', ten):
            rid = pic.find('{%s}blipFill/{%s}blip' % (PIC, A)).get('{%s}embed' % R)
            part = d.part.related_parts[rid]
            with open(duong_dan_jpg, 'rb') as f:
                part._blob = f.read()
            w, h = Image.open(duong_dan_jpg).size
            inline = pic
            while inline.tag != '{%s}inline' % WP:
                inline = inline.getparent()
            ext = inline.find('{%s}extent' % WP)
            cx = int(ext.get('cx'))
            cy = int(round(cx * h / w))
            ext.set('cy', str(cy))
            pic.find('{%s}spPr/{%s}xfrm/{%s}ext' % (PIC, A, A)).set('cy', str(cy))
            kq = (cx, cy)
    if kq is None:
        raise ValueError('Không thấy ảnh ' + ten_goc)
    return kq


def doan_anh(d, mau_doan_anh, duong_dan, rong_emu, ten):
    """Đoạn ảnh mới (sao định dạng đoạn ảnh mẫu) chứa ảnh `duong_dan` rộng rong_emu."""
    from PIL import Image
    rid, _ = d.part.get_or_add_image(duong_dan)
    p = copy.deepcopy(mau_doan_anh)
    w, h = Image.open(duong_dan).size
    cy = int(round(rong_emu * h / w))
    inline = p.find('.//{%s}inline' % WP)
    inline.find('{%s}extent' % WP).set('cx', str(rong_emu))
    inline.find('{%s}extent' % WP).set('cy', str(cy))
    dp = inline.find('{%s}docPr' % WP)
    ids = [int(x.get('id')) for x in d.element.body.iter('{%s}docPr' % WP)]
    dp.set('id', str(max(ids) + 1))
    dp.set('name', 'Picture %d' % (max(ids) + 1))
    pic = inline.find('.//{%s}pic' % PIC)
    pic.find('{%s}nvPicPr/{%s}cNvPr' % (PIC, PIC)).set('name', ten)
    pic.find('{%s}blipFill/{%s}blip' % (PIC, A)).set('{%s}embed' % R, rid)
    ext = pic.find('{%s}spPr/{%s}xfrm/{%s}ext' % (PIC, A, A))
    ext.set('cx', str(rong_emu))
    ext.set('cy', str(cy))
    return p


# --------------------------------------------------------------------------- công thức (OMML)
def _mr(chu, thuong=False):
    r = etree.Element('{%s}r' % M)
    rp = etree.SubElement(r, '{%s}rPr' % M)
    if thuong:
        etree.SubElement(rp, '{%s}sty' % M).set('{%s}val' % M, 'p')
    wr = etree.SubElement(r, qn('w:rPr'))
    f = etree.SubElement(wr, qn('w:rFonts'))
    f.set(qn('w:ascii'), 'Cambria Math')
    f.set(qn('w:hAnsi'), 'Cambria Math')
    t = etree.SubElement(r, '{%s}t' % M)
    t.text = chu
    t.set(XML_SPACE, 'preserve')
    return r


def _m(tag, *con):
    e = etree.Element('{%s}%s' % (M, tag))
    for c in con:
        e.append(c)
    return e


def _nhom(x):
    """Danh sách phần tử OMML từ chuỗi (dùng _{..}, ^{..}) hoặc phần tử có sẵn."""
    if isinstance(x, (list, tuple)):
        out = []
        for y in x:
            out += _nhom(y)
        return out
    if not isinstance(x, str):
        return [x]
    out = []
    # 'X_{ab}' -> sSub ; 'X^{ab}' -> sSup ; chữ thường -> m:r
    for m in re.finditer(r'([^\s_^{}·=+−×()\[\]≥≤/,;]+|.)(?:_\{([^}]*)\})?(?:\^\{([^}]*)\})?', x):
        goc, duoi, tren = m.group(1), m.group(2), m.group(3)
        if duoi is None and tren is None:
            if out and out[-1].tag == '{%s}r' % M:
                out[-1].find('{%s}t' % M).text += goc
            else:
                out.append(_mr(goc))
        elif tren is None:
            out.append(_m('sSub', _m('e', _mr(goc)), _m('sub', _mr(duoi))))
        elif duoi is None:
            out.append(_m('sSup', _m('e', _mr(goc)), _m('sup', _mr(tren))))
        else:
            out.append(_m('sSubSup', _m('e', _mr(goc)), _m('sub', _mr(duoi)), _m('sup', _mr(tren))))
    return out


def phan_so(tu, mau):
    return _m('f', _m('num', *_nhom(tu)), _m('den', *_nhom(mau)))


def doan_cong_thuc(mau_doan_ct, noi_dung, so):
    """Đoạn công thức đánh số như các công thức của thuyết minh: tab giữa – công thức – tab phải – (so)."""
    p = OxmlElement('w:p')
    p.append(copy.deepcopy(mau_doan_ct.find(qn('w:pPr'))))
    r = OxmlElement('w:r')
    r.append(OxmlElement('w:tab'))
    p.append(r)
    om = _m('oMath', *_nhom(noi_dung))
    p.append(om)
    r = OxmlElement('w:r')
    r.append(OxmlElement('w:tab'))
    t = OxmlElement('w:t')
    t.text = '(%s)' % so
    r.append(t)
    p.append(r)
    return p


# --------------------------------------------------------------------------- mục lục
def go_muc_luc(body):
    """Đưa 3 trường TOC đã điền về dạng rỗng (để chen_muc_luc.py điền lại) và xóa bookmark _Toc cũ."""
    n_dong = 0
    for ma in ('TOC \\o "1-3" \\h \\z \\u', 'TOC \\h \\z \\t "Chu thich hinh,1"',
               'TOC \\h \\z \\t "Chu thich bang,1"'):
        dau = None
        for p in body.iterchildren(qn('w:p')):
            if any((it.text or '').strip() == ma for it in p.iter(qn('w:instrText'))):
                dau = p
                break
        if dau is None:
            raise ValueError('Không thấy trường ' + ma)
        st = kieu(dau)

        def cung_nhom(e):
            if e is None or e.tag != qn('w:p'):
                return False
            k = kieu(e) or ''
            return k == st or (st.startswith('TOC') and k.startswith('TOC'))
        ds = [dau]
        nx = dau.getnext()
        while cung_nhom(nx):
            ds.append(nx)
            nx = nx.getnext()
        p = OxmlElement('w:p')
        ppr = OxmlElement('w:pPr')
        ps = OxmlElement('w:pStyle')
        ps.set(qn('w:val'), st)
        ppr.append(ps)
        p.append(ppr)
        for loai, instr, chu in (('begin', None, None), (None, ' %s ' % ma, None), ('separate', None, None),
                                 (None, None, 'Nhấn F9 để cập nhật'), ('end', None, None)):
            r = OxmlElement('w:r')
            if loai:
                f = OxmlElement('w:fldChar')
                f.set(qn('w:fldCharType'), loai)
                r.append(f)
            if instr:
                it = OxmlElement('w:instrText')
                it.set(XML_SPACE, 'preserve')
                it.text = instr
                r.append(it)
            if chu:
                t = OxmlElement('w:t')
                t.text = chu
                r.append(t)
            p.append(r)
        dau.addprevious(p)
        for x in ds:
            body.remove(x)
        n_dong += len(ds)
    ids = set()
    for bs in list(body.iter(qn('w:bookmarkStart'))):
        if (bs.get(qn('w:name')) or '').startswith('_Toc'):
            ids.add(bs.get(qn('w:id')))
            bs.getparent().remove(bs)
    for be in list(body.iter(qn('w:bookmarkEnd'))):
        if be.get(qn('w:id')) in ids:
            be.getparent().remove(be)
    return n_dong, len(ids)
