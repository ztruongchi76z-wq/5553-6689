# -*- coding: utf-8 -*-
"""
ĐIỀN SẴN MỤC LỤC, DANH MỤC HÌNH, DANH MỤC BẢNG (CÓ SỐ TRANG) VÀO TỆP WORD

Bản dựng bằng python-docx chỉ có 3 trường TOC rỗng ("Nhấn F9 để cập nhật"), nên khi mở bằng
Google Docs, WPS, điện thoại hoặc khi chọn "No" ở hộp thoại cập nhật trường của Word thì không thấy
mục lục và số trang. Chương trình này:
  1. gắn bookmark _Toc… cho mọi đề mục (Heading 1–3, Heading 1N) và chú thích hình, bảng;
  2. ghi kết quả trường TOC giống hệt cách Word ghi: mỗi dòng là liên kết tới bookmark, tab có dấu
     chấm và trường PAGEREF chứa số trang – Word vẫn cập nhật lại được bằng F9;
  3. lấy số trang thật bằng cách xuất PDF qua LibreOffice (bookmark -> named destination),
     đọc số trang in ở chân trang, điền vào rồi xuất lại và kiểm tra từng dòng.

  python3 chen_muc_luc.py <vào.docx> <ra.docx> [<ra.pdf>]
Cần LibreOffice (soffice, có gói libreoffice-math để hiện công thức) và PyMuPDF.
"""
import copy
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

import docx
import pymupdf
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TAB_PHAI = 9062  # vị trí tab phải (twip) = bề rộng vùng chữ 16 cm


# ---------------------------------------------------------------------------
def van_ban(p):
    return ''.join(t.text or '' for t in p.iter(qn('w:t')))


def kieu(p):
    s = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
    return s.get(qn('w:val')) if s is not None else None


def dinh_dang_so(numbering_el, num_id='28'):
    """lvlText của 3 cấp đầu theo numbering.xml (vd 'CHƯƠNG %1', '%1.%2', '%1.%2.%3')."""
    num = [n for n in numbering_el.findall(qn('w:num')) if n.get(qn('w:numId')) == num_id][0]
    aid = num.find(qn('w:abstractNumId')).get(qn('w:val'))
    absn = [a for a in numbering_el.findall(qn('w:abstractNum')) if a.get(qn('w:abstractNumId')) == aid][0]
    out = {}
    for lvl in absn.findall(qn('w:lvl')):
        i = int(lvl.get(qn('w:ilvl')))
        if i < 3:
            out[i] = lvl.find(qn('w:lvlText')).get(qn('w:val'))
    return out


def co_danh_so(p):
    """Đề mục có đánh số tự động? (numId = 0 trên đoạn hoặc kiểu Heading1N/BKHeading1 -> không)."""
    if kieu(p) in ('Heading1N', 'BKHeading1'):
        return False
    np_ = p.find(qn('w:pPr') + '/' + qn('w:numPr'))
    if np_ is not None:
        nid = np_.find(qn('w:numId'))
        if nid is not None and nid.get(qn('w:val')) == '0':
            return False
    return True


CAP_DE_MUC = {'Heading1': 1, 'Heading1N': 1, 'BKHeading1': 1, 'Heading2': 2, 'Heading3': 3}


def thu_thap(body, numbering_el):
    """Danh sách đề mục và chú thích theo thứ tự xuất hiện (bỏ qua phần trước MỤC LỤC)."""
    fmt = dinh_dang_so(numbering_el)
    dem = [0, 0, 0]
    de_muc, hinh, bang = [], [], []
    bat_dau = False
    for p in body.iterchildren(qn('w:p')):
        t = van_ban(p).strip()
        if not bat_dau:
            if t == 'MỤC LỤC':
                bat_dau = True
            continue
        k = kieu(p)
        if k in CAP_DE_MUC and t:
            cap = CAP_DE_MUC[k]
            so = ''
            if co_danh_so(p):
                dem[cap - 1] += 1
                for j in range(cap, 3):
                    dem[j] = 0
                so = fmt[cap - 1]
                for j in range(cap):
                    so = so.replace('%%%d' % (j + 1), str(dem[j]))
            de_muc.append(dict(p=p, cap=cap, so=so, chu=t))
        elif k == 'Chuthichhinh' and t:
            hinh.append(dict(p=p, cap=1, so='', chu=t))
        elif k == 'Chuthichbang' and t:
            bang.append(dict(p=p, cap=1, so='', chu=t))
    return de_muc, hinh, bang


def gan_bookmark(muc, ten, bid):
    p = muc['p']
    runs = [r for r in p.iterchildren(qn('w:r'))]
    bs = OxmlElement('w:bookmarkStart')
    bs.set(qn('w:id'), str(bid))
    bs.set(qn('w:name'), ten)
    be = OxmlElement('w:bookmarkEnd')
    be.set(qn('w:id'), str(bid))
    runs[0].addprevious(bs)
    runs[-1].addnext(be)
    muc['bm'] = ten


def _run(text=None, fld=None, instr=None, tab=False):
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rpr.append(OxmlElement('w:noProof'))
    r.append(rpr)
    if fld:
        f = OxmlElement('w:fldChar')
        f.set(qn('w:fldCharType'), fld)
        r.append(f)
    if instr:
        it = OxmlElement('w:instrText')
        it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        it.text = instr
        r.append(it)
    if tab:
        r.append(OxmlElement('w:tab'))
    if text is not None:
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        r.append(t)
    return r


def dong_muc_luc(muc, style, so_trang):
    """Một dòng mục lục kiểu Word: <hyperlink anchor> chữ … tab PAGEREF </hyperlink>."""
    p = OxmlElement('w:p')
    ppr = OxmlElement('w:pPr')
    ps = OxmlElement('w:pStyle')
    ps.set(qn('w:val'), style)
    ppr.append(ps)
    tabs = OxmlElement('w:tabs')
    tb = OxmlElement('w:tab')
    tb.set(qn('w:val'), 'right')
    tb.set(qn('w:leader'), 'dot')
    tb.set(qn('w:pos'), str(TAB_PHAI))
    tabs.append(tb)
    ppr.append(tabs)
    p.append(ppr)
    hl = OxmlElement('w:hyperlink')
    hl.set(qn('w:anchor'), muc['bm'])
    hl.set(qn('w:history'), '1')
    if muc['so']:
        hl.append(_run(muc['so'] + ' '))
    # chép các run chữ của đề mục/chú thích (giữ chỉ số trên/dưới), bỏ định dạng đậm/nghiêng
    for r in muc['p'].iterchildren(qn('w:r')):
        if r.find(qn('w:t')) is None:
            continue
        rc = copy.deepcopy(r)
        rpr = rc.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            rc.insert(0, rpr)
        for tag in ('w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps'):
            for e in rpr.findall(qn(tag)):
                rpr.remove(e)
        # w:noProof phải đứng trước vertAlign, sz… theo thứ tự lược đồ CT_RPr (Word báo lỗi nếu sai)
        truoc = [e for e in rpr if e.tag in (qn('w:rStyle'), qn('w:rFonts'))]
        vt = rpr.index(truoc[-1]) + 1 if truoc else 0
        rpr.insert(vt, OxmlElement('w:noProof'))
        hl.append(rc)
    hl.append(_run(tab=True))
    hl.append(_run(fld='begin'))
    hl.append(_run(instr=' PAGEREF %s \\h ' % muc['bm']))
    hl.append(_run(fld='separate'))
    hl.append(_run(so_trang))
    hl.append(_run(fld='end'))
    p.append(hl)
    return p


def thay_ket_qua_truong(body, ma_truong, dong):
    """Thay đoạn chứa trường TOC (ma_truong) bằng các dòng `dong`; giữ nguyên mã trường."""
    goc = None
    for p in body.iterchildren(qn('w:p')):
        for it in p.iter(qn('w:instrText')):
            if (it.text or '').strip() == ma_truong:
                goc = p
    if goc is None:
        raise ValueError('Không thấy trường ' + ma_truong)
    dau = dong[0]
    hl = dau.find(qn('w:hyperlink'))
    for r in (_run(fld='begin'), _run(instr=' %s ' % ma_truong), _run(fld='separate')):
        hl.addprevious(r)
    dong[-1].append(_run(fld='end'))
    for d in dong:
        goc.addprevious(d)
    goc.getparent().remove(goc)


def dung(vao, ra, trang):
    """Dựng tệp ra từ tệp vào; trang: dict bookmark -> số trang (chuỗi) hoặc None (chưa biết)."""
    d = docx.Document(vao)
    body = d.element.body
    # Danh mục hình, bảng: cách đoạn sau 6 pt theo hướng dẫn trình bày của Trường (mục "Các danh mục")
    tof = d.styles['table of figures'].element
    ppr = tof.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        tof.append(ppr)
    sp = ppr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        ppr.append(sp)
    sp.set(qn('w:after'), '120')
    de_muc, hinh, bang = thu_thap(body, d.part.numbering_part.element)
    bid = 900
    ds = []
    for nhom, ten in ((de_muc, 'tm'), (hinh, 'hinh'), (bang, 'bang')):
        for i, m in enumerate(nhom):
            gan_bookmark(m, '_Toc%s%03d' % ({'tm': '1', 'hinh': '2', 'bang': '3'}[ten], i), bid)
            bid += 1
            ds.append(m)
    so = lambda m: (trang or {}).get(m['bm']) or '0'
    dong_tm = [dong_muc_luc(m, 'TOC%d' % m['cap'], so(m)) for m in de_muc]
    dong_h = [dong_muc_luc(m, 'TableofFigures', so(m)) for m in hinh]
    dong_b = [dong_muc_luc(m, 'TableofFigures', so(m)) for m in bang]
    thay_ket_qua_truong(body, 'TOC \\o "1-3" \\h \\z \\u', dong_tm)
    thay_ket_qua_truong(body, 'TOC \\h \\z \\t "Chu thich hinh,1"', dong_h)
    thay_ket_qua_truong(body, 'TOC \\h \\z \\t "Chu thich bang,1"', dong_b)
    d.save(ra)
    return [(m['bm'], (m['so'] + ' ' + m['chu']).strip()) for m in ds]


# ---------------------------------------------------------------------------
def xuat_pdf(docx_path, out_dir):
    loc = 'pdf:writer_pdf_Export:' + json.dumps({'ExportBookmarksToPDFDestination': {'type': 'boolean', 'value': 'true'}})
    subprocess.run(['soffice', '--headless', '--convert-to', loc, '--outdir', out_dir, docx_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')


def so_trang_in(page):
    """Số trang in ở chân trang (chữ số hoặc số La Mã); None nếu không có."""
    h = page.rect.height
    for b in sorted(page.get_text('blocks'), key=lambda b: -b[1]):
        if b[1] > h - 75:
            s = b[4].strip()
            if re.fullmatch(r'\d+|[ivxlc]+', s):
                return s
    return None


def doc_trang(pdf_path):
    doc = pymupdf.open(pdf_path)
    names = doc.resolve_names()
    out = {}
    for ten, dest in names.items():
        # LibreOffice mã hóa '_' đầu tên thành '#5F'; PyMuPDF trả về '5F…'
        ten = re.sub(r'^(#?5F)', '_', ten)
        if ten.startswith('_Toc'):
            out[ten] = so_trang_in(doc[dest['page']])
    return out, doc.page_count


def kiem_tra(pdf_path, danh_sach, trang):
    """So số trang ghi trong mục lục (trang i…) với vị trí thật của bookmark."""
    doc = pymupdf.open(pdf_path)
    thuc, _ = doc_trang(pdf_path)
    sai = [(bm, trang.get(bm), thuc.get(bm)) for bm, _ in danh_sach if trang.get(bm) != thuc.get(bm)]
    return sai, doc.page_count


def main(vao, ra, ra_pdf=None):
    tmp = tempfile.mkdtemp(prefix='muc_luc_')
    trang = None
    for lan in range(1, 4):
        buoc = os.path.join(tmp, 'buoc%d.docx' % lan)
        danh_sach = dung(vao, buoc, trang)
        pdf = xuat_pdf(buoc, tmp)
        moi, n = doc_trang(pdf)
        thieu = [bm for bm, _ in danh_sach if not moi.get(bm)]
        if thieu:
            raise RuntimeError('Không xác định được trang của %d mục: %s' % (len(thieu), thieu[:5]))
        if trang == moi:
            print('Lần %d: số trang ổn định (%d trang PDF).' % (lan, n))
            break
        print('Lần %d: %d mục, %d trang PDF – cập nhật số trang.' % (lan, len(danh_sach), n))
        trang = moi
    danh_sach = dung(vao, ra, trang)
    pdf = xuat_pdf(ra, tmp)
    sai, n = kiem_tra(pdf, danh_sach, trang)
    print('Kiểm tra cuối: %d dòng mục lục/danh mục, %d dòng sai số trang, %d trang.' % (len(danh_sach), len(sai), n))
    for s in sai:
        print('  SAI', s)
    if ra_pdf:
        shutil.copy(pdf, ra_pdf)
        print('PDF:', ra_pdf)
    with open(os.path.splitext(ra)[0] + '_muc_luc.txt', 'w', encoding='utf-8') as f:
        for bm, chu in danh_sach:
            f.write('%-6s %s\n' % (trang[bm], chu))
    shutil.rmtree(tmp, ignore_errors=True)
    return 0 if not sai else 1


if __name__ == '__main__':
    sys.exit(main(*sys.argv[1:4]))
