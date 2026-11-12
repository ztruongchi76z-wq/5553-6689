# -*- coding: utf-8 -*-
"""
THƯ VIỆN DỰNG BẢN VẼ A0 CHO ĐATN HỒ CÀ GIÂY (ezdxf)

Quy ước:
  - Mỗi tờ là một tệp DXF (AutoCAD 2018), vẽ trong Model ở đơn vị mm GIẤY (khung A0 1189 × 841 mm,
    tỷ lệ in 1:1). Mỗi hình chiếu có tỷ lệ riêng ghi ở tên hình; kích thước tự ghi theo cm thật
    (DIMLFAC = mẫu số tỷ lệ / 10), cao trình ghi theo m.
  - Các hàm vẽ từng hình chiếu được viết theo bố cục A1 cũ; khi dựng tờ A0, mỗi hình chiếu/bảng là một
    "khối" được đặt bằng To.dat(): đo khung bao, tịnh tiến và phóng đồng dạng hệ số f (chữ, nét, ký hiệu
    phóng theo). Tỷ lệ ghi trong chữ "TỶ LỆ ..." của khối được đổi tự động (1/2500 ở f = 1,25 -> 1/2000);
    f chỉ chọn sao cho tỷ lệ mới là tỷ lệ chuẩn.
  - Khung bản vẽ theo TCVN 7285:2003 (ISO 5457): lề trái 20 mm, các lề khác 10 mm; khung tên góc phải dưới.
  - Phông Arial (arial.ttf) – có sẵn trên Windows; bản PDF xem trước dùng Liberation Sans (cùng kích thước chữ).
"""
import math
import os
import re
from contextlib import contextmanager

import ezdxf
from ezdxf import bbox as _bbox
from ezdxf import units
from ezdxf.enums import TextEntityAlignment
from ezdxf.math import Vec2

A1 = (841.0, 594.0)
A0 = (1189.0, 841.0)
KHO = A0                      # khổ giấy của bộ bản vẽ
LE_TRAI, LE = 20.0, 10.0
DE_TAI = 'TÍNH TOÁN, THIẾT KẾ SỬA CHỮA NÂNG CẤP CỤM CÔNG TRÌNH ĐẦU MỐI HỒ CHỨA NƯỚC CÀ GIÂY – TỈNH LÂM ĐỒNG'
TONG_SO_TO = 7

# (tên lớp, màu ACI, độ dày nét 1/100 mm, kiểu đường)
LOP = [
    ('KHUNG', 7, 50, 'CONTINUOUS'), ('NET_CHINH', 7, 50, 'CONTINUOUS'), ('NET_MANH', 8, 18, 'CONTINUOUS'),
    ('NET_KHUAT', 7, 25, 'DASHED'), ('TRUC', 1, 18, 'CENTER'), ('KICH_THUOC', 3, 18, 'CONTINUOUS'),
    ('CHU', 7, 25, 'CONTINUOUS'), ('CHU_DAM', 7, 35, 'CONTINUOUS'), ('HATCH', 8, 13, 'CONTINUOUS'),
    ('NUOC', 5, 25, 'CONTINUOUS'), ('DIA_CHAT', 30, 25, 'DASHED'), ('THAM', 1, 35, 'CONTINUOUS'),
    ('CUNG_TRUOT', 6, 35, 'CONTINUOUS'), ('SUA_CHUA', 1, 50, 'CONTINUOUS'), ('NEN', 9, 13, 'CONTINUOUS'),
    ('BANG', 7, 25, 'CONTINUOUS'),
]


def so(v, nd=2, dau=False):
    """Định dạng số kiểu Việt Nam: dấu phẩy thập phân, có thể thêm dấu +."""
    s = ('%.' + str(nd) + 'f') % v
    s = s.replace('.', ',')
    if dau and v >= 0:
        s = '+' + s
    return s


class To:
    """Một tờ bản vẽ (mặc định A0)."""

    def __init__(self, so_to, ten_ban_ve, ty_le_chung, ma, kho=None, khung=True):
        self.W, self.H = kho or KHO
        self._f, self._src, self._dst = 1.0, (0.0, 0.0), (0.0, 0.0)
        self.k_chu = 1.0          # hệ số phóng chiều cao chữ (không phóng hình) – dùng khi tờ A0 còn chỗ
        self.doc = ezdxf.new('R2018', setup=True)
        self.doc.units = units.MM
        self.doc.header['$LWDISPLAY'] = 1
        self.msp = self.doc.modelspace()
        self.so_to, self.ten, self.ty_le_chung, self.ma = so_to, ten_ban_ve, ty_le_chung, ma
        for ten, mau, lw, lt in LOP:
            self.doc.layers.add(ten, color=mau, lineweight=lw, linetype=lt)
        st = self.doc.styles
        for ten, f in (('VN', 'arial.ttf'), ('VN_DAM', 'arialbd.ttf'), ('VN_NGHIENG', 'ariali.ttf')):
            if ten not in st:
                st.add(ten, font=f)
        ds = self.doc.dimstyles.new('KT')
        ds.dxf.update(dict(dimtxt=2.5, dimasz=1.8, dimexo=1.0, dimexe=1.5, dimgap=0.8, dimtad=1, dimdec=0,
                           dimtxsty='VN', dimclrd=3, dimclre=3, dimclrt=7, dimtih=0, dimtoh=0, dimtix=0,
                           dimdsep=ord(','), dimzin=8))
        ds.set_arrows(blk='ARCHTICK')
        self.doc.header['$LTSCALE'] = 1.0
        if khung:
            self.ve_khung()

    # ------------------------------------------------------------------ khối (đặt + phóng đồng dạng)
    def _p(self, p):
        f, (sx, sy), (dx, dy) = self._f, self._src, self._dst
        return (dx + (p[0] - sx) * f, dy + (p[1] - sy) * f)

    def _kw(self, kw):
        kw = dict(kw)
        for k in ('const_width', 'ltscale'):
            if k in kw:
                kw[k] = kw[k] * self._f
        return kw

    def _doi_ty_le(self, s):
        """Trong khối phóng f, đổi 1/N trong chữ 'TỶ LỆ ...' thành 1/(N/f); số lẻ thì đánh dấu '?' để soát."""
        if self._f == 1.0 or 'TỶ LỆ' not in s.upper():
            return s

        def doi(m):
            v = float(m.group(1).replace(',', '.')) / self._f
            return '1/%d' % round(v) if abs(v - round(v)) < 1e-6 else m.group(0) + '?'
        return re.sub(r'1/(\d+(?:[.,]\d+)?)', doi, s)

    @contextmanager
    def khoi(self, f, src, dst):
        cu = (self._f, self._src, self._dst)
        self._f, self._src, self._dst = float(f), tuple(src), tuple(dst)
        try:
            yield self
        finally:
            self._f, self._src, self._dst = cu

    @staticmethod
    def khung_bao(ham, *args, k_chu=1.0, **kw):
        """Khung bao (xmin, ymin, xmax, ymax) của khối do ham(to, ...) vẽ, đo trên một tờ nháp không khung."""
        nhap = To(0, '', '', '', khung=False)
        nhap.k_chu = k_chu
        ham(nhap, *args, **kw)
        b = _bbox.extents(nhap.msp, fast=True)
        return (b.extmin.x, b.extmin.y, b.extmax.x, b.extmax.y)

    def dat(self, ham, *args, f=1.0, neo='TL', xy=(0.0, 0.0), khung=None, k_chu=None, **kw):
        """Vẽ khối ham(to, ...) phóng f, điểm neo của khung bao ('TL','TR','BL','BR','TC','BC','C')
        đặt tại xy. k_chu: hệ số chiều cao chữ riêng cho khối (vd 0,625 ở khối f = 2 để chữ chỉ phóng 1,25).
        Trả về khung bao mới (xmin, ymin, xmax, ymax) trên tờ."""
        kc = self.k_chu if k_chu is None else k_chu
        x0, y0, x1, y1 = khung or To.khung_bao(ham, *args, k_chu=kc, **kw)
        ax = {'L': x0, 'R': x1, 'C': (x0 + x1) / 2}[neo[1] if len(neo) > 1 else 'C']
        ay = {'T': y1, 'B': y0, 'C': (y0 + y1) / 2}[neo[0]]
        cu = self.k_chu
        self.k_chu = kc
        try:
            with self.khoi(f, (ax, ay), xy):
                ham(self, *args, **kw)
        finally:
            self.k_chu = cu
        X0, Y0 = xy[0] + (x0 - ax) * f, xy[1] + (y0 - ay) * f
        return (X0, Y0, X0 + (x1 - x0) * f, Y0 + (y1 - y0) * f)

    # ------------------------------------------------------------------ cơ bản
    def line(self, p1, p2, lop='NET_CHINH', **kw):
        return self.msp.add_line(self._p(p1), self._p(p2), dxfattribs=dict(layer=lop, **self._kw(kw)))

    def pline(self, pts, lop='NET_CHINH', dong=False, **kw):
        return self.msp.add_lwpolyline([self._p(q) for q in pts], close=dong,
                                       dxfattribs=dict(layer=lop, **self._kw(kw)))

    def rect(self, x, y, w, h, lop='NET_CHINH', **kw):
        return self.pline([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], lop, dong=True, **kw)

    def circle(self, c, r, lop='NET_CHINH', **kw):
        return self.msp.add_circle(self._p(c), r * self._f, dxfattribs=dict(layer=lop, **self._kw(kw)))

    def arc(self, c, r, a0, a1, lop='NET_CHINH', **kw):
        return self.msp.add_arc(self._p(c), r * self._f, a0, a1, dxfattribs=dict(layer=lop, **self._kw(kw)))

    def text(self, s, p, h=2.5, can='LEFT', lop='CHU', style='VN', rot=0.0, **kw):
        al = {'LEFT': TextEntityAlignment.BOTTOM_LEFT, 'CENTER': TextEntityAlignment.BOTTOM_CENTER,
              'RIGHT': TextEntityAlignment.BOTTOM_RIGHT, 'MC': TextEntityAlignment.MIDDLE_CENTER,
              'ML': TextEntityAlignment.MIDDLE_LEFT, 'MR': TextEntityAlignment.MIDDLE_RIGHT,
              'TL': TextEntityAlignment.TOP_LEFT, 'TC': TextEntityAlignment.TOP_CENTER}[can]
        t = self.msp.add_text(self._doi_ty_le(s), height=h * self._f * self.k_chu, rotation=rot,
                              dxfattribs=dict(layer=lop, style=style, **kw))
        t.set_placement(self._p(p), align=al)
        return t

    def mtext(self, s, p, h=2.5, rong=100.0, lop='CHU', style='VN', cach_dong=1.25):
        m = self.msp.add_mtext(s, dxfattribs=dict(layer=lop, style=style, char_height=h * self._f * self.k_chu,
                                                  width=rong * self._f, line_spacing_factor=cach_dong))
        m.set_location(self._p(p), attachment_point=1)
        return m

    def hatch(self, pts, mau='ANSI31', ty_le=1.0, lop='HATCH', goc=0.0, mau_to=None):
        h = self.msp.add_hatch(color=mau_to if mau_to is not None else 256, dxfattribs=dict(layer=lop))
        if mau == 'SOLID':
            h.set_solid_fill(color=mau_to if mau_to is not None else 8)
        else:
            h.set_pattern_fill(mau, color=mau_to if mau_to is not None else 7, scale=ty_le * self._f, angle=goc)
        h.paths.add_polyline_path([self._p(q) for q in pts], is_closed=True)
        return h

    def dim(self, p1, p2, d, mau_so_ty_le, goc=0.0, text=None):
        """Kích thước thẳng giữa p1, p2 (mm giấy); đường kích thước cách p1 khoảng d theo pháp tuyến.
        Ghi số cm thật (DIMLFAC = mẫu số / 10)."""
        p1, p2 = Vec2(self._p(p1)), Vec2(self._p(p2))
        u = Vec2.from_deg_angle(goc)
        n = Vec2(-u.y, u.x)
        base = p1 + n * d * self._f
        ov = dict(dimlfac=mau_so_ty_le / 10.0 / self._f, dimscale=self._f, dimtxt=2.5 * self.k_chu)
        dim = self.msp.add_linear_dim(base=base, p1=p1, p2=p2, angle=goc, dimstyle='KT', text=text or '<>',
                                      override=ov, dxfattribs=dict(layer='KICH_THUOC'))
        dim.render()
        return dim

    # ------------------------------------------------------------------ ký hiệu
    def cao_trinh(self, x, y, gia_tri, ben='R', h=2.5, lop='CHU', nuoc=False, chu=None):
        """Ký hiệu cao trình: tam giác đỉnh chạm (x, y), dòng chữ cao trình phía trên."""
        s = chu if chu is not None else so(gia_tri, 2, True)
        a = 1.6
        tri = [(x, y), (x - a, y + a * 1.2), (x + a, y + a * 1.2)]
        self.pline(tri, 'NUOC' if nuoc else 'NET_MANH', dong=True)
        if not nuoc:
            hh = self.msp.add_hatch(color=7, dxfattribs=dict(layer='NET_MANH'))
            hh.set_solid_fill(color=7)
            hh.paths.add_polyline_path([self._p(q) for q in tri], is_closed=True)
        dai = max(len(s) * h * self.k_chu * 0.62, 10)
        if ben == 'R':
            self.line((x - a, y + a * 1.2), (x + dai, y + a * 1.2), 'NUOC' if nuoc else 'NET_MANH')
            self.text(s, (x + 0.6, y + a * 1.2 + 0.6), h, 'LEFT', lop)
        else:
            self.line((x + a, y + a * 1.2), (x - dai, y + a * 1.2), 'NUOC' if nuoc else 'NET_MANH')
            self.text(s, (x - 0.6, y + a * 1.2 + 0.6), h, 'RIGHT', lop)

    def muc_nuoc(self, x0, x1, y, ten, gia_tri, nhan_x=None):
        """Đường mực nước + ký hiệu ▽ (lớp NUOC)."""
        self.line((x0, y), (x1, y), 'NUOC')
        xm = nhan_x if nhan_x is not None else x0 + 0.15 * (x1 - x0)
        for k in (1, 2):
            self.line((xm - 3 + k, y - 1.0 * k), (xm + 3 - k, y - 1.0 * k), 'NUOC')
        self.cao_trinh(xm, y, gia_tri, nuoc=True, chu='%s %s' % (ten, so(gia_tri, 2, True)))

    def ten_hinh(self, s, cx, y, ty_le=None, h=5.0, can='CENTER'):
        """Tên hình chiếu gạch chân + tỷ lệ (can = 'CENTER' hoặc 'LEFT' – khi đó cx là mép trái)."""
        self.text(s, (cx, y), h, can, 'CHU_DAM', 'VN_DAM')
        w = len(s) * h * self.k_chu * 0.62
        x1 = cx - w / 2 if can == 'CENTER' else cx
        self.line((x1, y - 1.2), (x1 + w, y - 1.2), 'CHU_DAM')
        if ty_le:
            self.text('TỶ LỆ: ' + ty_le, (cx, y - 6.5), 3.0, can)

    def chu_thich_mui_ten(self, p_dau, p_chu, s, h=2.5, rong=60.0):
        """Đường dẫn chú thích từ điểm p_dau tới dòng chữ."""
        self.line(p_dau, p_chu, 'NET_MANH')
        self.circle(p_dau, 0.5, 'NET_MANH')
        x2 = p_chu[0] + (rong if p_chu[0] >= p_dau[0] else -rong)
        self.line(p_chu, (x2, p_chu[1]), 'NET_MANH')
        if p_chu[0] >= p_dau[0]:
            self.text(s, (p_chu[0] + 0.8, p_chu[1] + 0.7), h, 'LEFT')
        else:
            self.text(s, (p_chu[0] - 0.8, p_chu[1] + 0.7), h, 'RIGHT')

    def bang(self, x, y_tren, cot, dong, h=2.5, cao_dong=6.0, tieu_de=None, can=None):
        """Bảng: cot = danh sách độ rộng (mm), dong = danh sách các hàng (list chuỗi). Trả về y đáy."""
        y = y_tren
        if tieu_de:
            self.text(tieu_de, (x + sum(cot) / 2, y + 2.0), 3.5, 'CENTER', 'CHU_DAM', 'VN_DAM')
        rong = sum(cot)
        for i, hang in enumerate(dong):
            nd = max(1, max(str(c).count('\n') + 1 for c in hang))
            ch = cao_dong * nd if nd > 1 else cao_dong
            self.line((x, y), (x + rong, y), 'BANG')
            xx = x
            for j, (w, c) in enumerate(zip(cot, hang)):
                cc = (can[j] if can else ('L' if j == 0 else 'C'))
                for k, sdong in enumerate(str(c).split('\n')):
                    yy = y - ch / 2 - (k - (nd - 1) / 2) * cao_dong * 0.95
                    if cc == 'L':
                        self.text(sdong, (xx + 1.2, yy), h, 'ML', 'CHU_DAM' if i == 0 else 'CHU',
                                  'VN_DAM' if i == 0 else 'VN')
                    else:
                        self.text(sdong, (xx + w / 2, yy), h, 'MC', 'CHU_DAM' if i == 0 else 'CHU',
                                  'VN_DAM' if i == 0 else 'VN')
                xx += w
            y -= ch
        self.line((x, y), (x + rong, y), 'BANG')
        xx = x
        for w in [0] + cot:
            xx += w
            self.line((xx, y_tren), (xx, y), 'BANG')
        return y

    # ------------------------------------------------------------------ khung
    def ve_khung(self):
        W, H = self.W, self.H
        self.rect(0, 0, W, H, 'NET_MANH')
        self.rect(LE_TRAI, LE, W - LE_TRAI - LE, H - 2 * LE, 'KHUNG', const_width=0.0)
        # khung tên 185 × 58 mm ở góc phải dưới
        x0, y0, w = W - LE - 185.0, LE, 185.0
        hs = [8.0, 7.0, 12.0, 13.0, 6.0, 6.0, 6.0]
        ys = [y0]
        for hh in reversed(hs):
            ys.append(ys[-1] + hh)
        ytop = ys[-1]
        self.rect(x0, y0, w, ytop - y0, 'KHUNG')
        for yy in ys[1:-1]:
            self.line((x0, yy), (x0 + w, yy), 'BANG')
        # 3 hàng trên cùng (tên trường, khoa), đề tài, tên bản vẽ
        y_truong, y_khoa, y_detai, y_ten = ys[7], ys[6], ys[5], ys[4]
        self.text('TRƯỜNG ĐẠI HỌC THỦY LỢI – PHÂN HIỆU TẠI TP. HỒ CHÍ MINH', (x0 + w / 2, y_truong - 4.0), 3.2,
                  'MC', 'CHU_DAM', 'VN_DAM')
        self.text('KHOA CÔNG TRÌNH  –  ĐỒ ÁN TỐT NGHIỆP KỸ SƯ', (x0 + w / 2, y_khoa - 4.0), 3.0, 'MC')
        self.text('ĐỀ TÀI: TÍNH TOÁN, THIẾT KẾ SỬA CHỮA NÂNG CẤP CỤM CÔNG TRÌNH ĐẦU MỐI', (x0 + w / 2, y_detai - 3.8),
                  2.6, 'MC')
        self.text('HỒ CHỨA NƯỚC CÀ GIÂY – TỈNH LÂM ĐỒNG', (x0 + w / 2, y_detai - 8.6), 2.6, 'MC')
        dong_ten = self.ten.split('\n')
        for k, s in enumerate(dong_ten):
            self.text(s, (x0 + w / 2, y_ten - 6.5 - (k - (len(dong_ten) - 1) / 2) * 4.6), 3.2 if len(dong_ten) > 1 else 3.8,
                      'MC', 'CHU_DAM', 'VN_DAM')
        # 3 hàng dưới: chức danh | họ tên | chữ ký | ô phải
        cx = [x0, x0 + 30, x0 + 85, x0 + 115, x0 + 145, x0 + w]
        for xx in cx[1:-1]:
            self.line((xx, y0), (xx, ys[3]), 'BANG')
        rows = [('GV HƯỚNG DẪN', 'TS. Phạm Ngọc Thịnh', 'TỶ LỆ', self.ty_le_chung),
                ('SV THỰC HIỆN', 'Trần Quốc Thanh', 'BẢN VẼ SỐ', '%02d / %02d' % (self.so_to, TONG_SO_TO)),
                ('NGÀNH', 'KTXD công trình thủy', 'KÝ HIỆU', self.ma)]
        for i, (a, b, c, d) in enumerate(rows):
            yy = ys[3 - i] - 3.0
            self.text(a, (cx[0] + 1.5, yy), 2.2, 'ML')
            self.text(b, (cx[1] + 1.5, yy), 2.4, 'ML')
            self.text(c, (cx[3] + 1.5, yy), 2.2, 'ML')
            self.text(d, (cx[4] + 20, yy), 2.4, 'MC', 'CHU_DAM', 'VN_DAM')
        self.text('Tháng 10/2026', (cx[3] - 15, y0 + 3.0), 2.0, 'MC')
        self.khung_ten_tren = ytop
        self.khung_ten_trai = x0

    # ------------------------------------------------------------------ xuất
    def luu(self, duong_dan_dxf, duong_dan_pdf=None, duong_dan_png=None):
        os.makedirs(os.path.dirname(duong_dan_dxf) or '.', exist_ok=True)
        self.doc.saveas(duong_dan_dxf)
        if duong_dan_pdf or duong_dan_png:
            xuat_hinh(self.doc, duong_dan_pdf, duong_dan_png, kho=(self.W, self.H))


def xuat_hinh(doc, pdf=None, png=None, dpi_png=90, kho=None):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    from ezdxf.addons.drawing.config import (Configuration, BackgroundPolicy, ColorPolicy, LineweightPolicy,
                                             HatchPolicy)
    from ezdxf.fonts import fonts as _fonts
    # máy không có Arial: dựng PDF xem trước bằng Liberation Sans (cùng kích thước chữ) thay cho phông dự phòng
    # mặc định của ezdxf (DejaVu Sans, rộng hơn); máy có arial.ttf thì giữ Arial.
    _fonts.font_manager.add_synonyms({'LiberationSans-Regular.ttf': 'arial.ttf', 'LiberationSans-Bold.ttf': 'arialbd.ttf',
                                      'LiberationSans-Italic.ttf': 'ariali.ttf'}, reverse=False)
    W, H = kho or KHO
    fig = plt.figure(figsize=(W / 25.4, H / 25.4))
    ax = fig.add_axes([0, 0, 1, 1])
    cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.COLOR,
                        lineweight_policy=LineweightPolicy.ABSOLUTE, lineweight_scaling=1.0,
                        min_lineweight=0.12)
    Frontend(RenderContext(doc), MatplotlibBackend(ax, adjust_figure=False), config=cfg).draw_layout(
        doc.modelspace(), finalize=False)
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.set_aspect('equal')
    ax.axis('off')
    if pdf:
        fig.savefig(pdf)
    if png:
        fig.savefig(png, dpi=dpi_png)
    plt.close(fig)


class KhungNhin:
    """Phép biến đổi tọa độ thật (m) -> mm giấy cho một hình chiếu: X = x0 + (x - gx)·1000/tl."""

    def __init__(self, x0, y0, gx, gy, tl_ngang, tl_dung=None):
        self.x0, self.y0, self.gx, self.gy = x0, y0, gx, gy
        self.kx = 1000.0 / tl_ngang
        self.ky = 1000.0 / (tl_dung or tl_ngang)
        self.tl = tl_ngang

    def __call__(self, x, z):
        return (self.x0 + (x - self.gx) * self.kx, self.y0 + (z - self.gy) * self.ky)

    def pts(self, xs, zs):
        return [self(x, z) for x, z in zip(xs, zs)]
