# -*- coding: utf-8 -*-
"""Công thức Word (OMML) và bảng chữ nhỏ dùng chung cho các file Word thuyết minh (cùng cách viết với mục 2.3)."""
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

from xuat_bang_tham_docx import chu, doan, hang_lap

M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def _m(tag, *con, **thuoc):
    e = etree.Element('{%s}%s' % (M, tag))
    for k, v in thuoc.items():
        e.set('{%s}%s' % (M, k), v)
    for c in con:
        e.append(c)
    return e


def t(chu_, dung=False):
    """m:r – dung=True: chữ đứng (tên mực nước, chữ viết tắt)."""
    r = _m('r')
    if dung:
        r.append(_m('rPr', _m('sty', val='p')))
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), 'Cambria Math')
    rf.set(qn('w:hAnsi'), 'Cambria Math')
    rpr.append(rf)
    r.append(rpr)
    tt = _m('t')
    tt.text = chu_
    tt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(tt)
    return r


def _ds(x):
    return x if isinstance(x, list) else [x]


def sub(goc, cs):
    return _m('sSub', _m('e', *_ds(goc)), _m('sub', *_ds(cs)))


def sup(goc, mu):
    return _m('sSup', _m('e', *_ds(goc)), _m('sup', *_ds(mu)))


def sub_sup(goc, cs, mu):
    return _m('sSubSup', _m('e', *_ds(goc)), _m('sub', *_ds(cs)), _m('sup', *_ds(mu)))


def phan_so(tu, mau):
    return _m('f', _m('num', *_ds(tu)), _m('den', *_ds(mau)))


def can(x, bac=None):
    if bac is None:
        return _m('rad', _m('radPr', _m('degHide', val='1')), _m('deg'), _m('e', *_ds(x)))
    return _m('rad', _m('deg', *_ds(bac)), _m('e', *_ds(x)))


def ngoac(x, mo='(', dong=')'):
    return _m('d', _m('dPr', _m('begChr', val=mo), _m('endChr', val=dong)), _m('e', *_ds(x)))


def gach(x):
    return _m('bar', _m('barPr', _m('pos', val='top')), _m('e', *_ds(x)))


def phuong_trinh(d, phan, so_ct=None):
    """Công thức giữa dòng, số hiệu canh phải."""
    p = d.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(4), Pt(4), 1.2
    pf.keep_with_next = False
    ppr = p._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    for vt, kieu in ((4535, 'center'), (9071, 'right')):
        tb = OxmlElement('w:tab')
        tb.set(qn('w:val'), kieu)
        tb.set(qn('w:pos'), str(vt))
        tabs.append(tb)
    ppr.append(tabs)
    p.add_run('\t')
    p._p.append(_m('oMath', *phan))
    if so_ct:
        chu(p, '\t(%s)' % so_ct)
    return p


def o_bang(cell, markup, rong, co=10, dam=False, can_=WD_ALIGN_PARAGRAPH.CENTER, nen=None):
    tcpr = cell._tc.get_or_add_tcPr()
    tcw = tcpr.find(qn('w:tcW'))
    if tcw is None:
        tcw = OxmlElement('w:tcW')
        tcpr.append(tcw)
    tcw.set(qn('w:w'), str(rong))
    tcw.set(qn('w:type'), 'dxa')
    if nen:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), nen)
        tcpr.append(shd)
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), 'center')
    tcpr.append(va)
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.alignment, pf.space_before, pf.space_after, pf.line_spacing = can_, Pt(1), Pt(1), 1.0
    chu(p, markup, co, dam)


def bang(d, chu_thich, tieu_de, du_lieu, cot, co=10, cot_trai=(0,), dau=None, le=None):
    """Bảng có chú thích ở trên (nghiêng, giữa); cot: bề rộng cột (dxa); dau: các dòng chữ trước bảng (ghi chú số liệu);
    le: lề trái, phải trong ô (dxa) – bảng nhiều cột dùng lề nhỏ để số không xuống dòng."""
    p = doan(d, chu_thich, co=12, can=WD_ALIGN_PARAGRAPH.CENTER, nghieng=True, truoc=6, sau=4)
    p.paragraph_format.keep_with_next = True
    for dong in dau or []:
        q = doan(d, dong, co=11, can=WD_ALIGN_PARAGRAPH.LEFT, truoc=0, sau=0)
        q.paragraph_format.keep_with_next = True
    tb = d.add_table(rows=1 + len(du_lieu), cols=len(cot))
    tb.style = 'Table Grid'
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    if le is not None:
        mar = OxmlElement('w:tblCellMar')
        for canh in ('left', 'right'):
            e = OxmlElement('w:%s' % canh)
            e.set(qn('w:w'), str(le))
            e.set(qn('w:type'), 'dxa')
            mar.append(e)
        tb._tbl.tblPr.append(mar)
    for gc, w in zip(tb._tbl.tblGrid.findall(qn('w:gridCol')), cot):
        gc.set(qn('w:w'), str(w))
    for j, (c, w) in enumerate(zip(tb.rows[0].cells, cot)):
        o_bang(c, tieu_de[j], w, co, dam=True, nen='F2F2F2')
    hang_lap(tb.rows[0], tieu_de=True)
    for i, dong in enumerate(du_lieu):
        for j, (c, w) in enumerate(zip(tb.rows[i + 1].cells, cot)):
            o_bang(c, dong[j], w, co, can_=WD_ALIGN_PARAGRAPH.LEFT if j in cot_trai else WD_ALIGN_PARAGRAPH.CENTER)
        hang_lap(tb.rows[i + 1])
    if len(du_lieu) <= 12:                      # bảng ngắn: giữ trên một trang
        for row in tb.rows[:-1]:
            for c in row.cells:
                c.paragraphs[0].paragraph_format.keep_with_next = True
    return tb
