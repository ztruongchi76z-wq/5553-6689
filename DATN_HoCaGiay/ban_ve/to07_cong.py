# -*- coding: utf-8 -*-
"""BV-07: CỐNG LẤY NƯỚC – THAY MÁY ĐÓNG MỞ, KIỂM TRA THỦY LỰC VÀ LỰC ĐÓNG MỞ (Chương 6).

Gồm: cắt dọc cống với đường mặt nước Q = 6,54 và 10 m³/s (ngang 1/250, đứng 1/50, sơ đồ) và bảng số liệu; mặt cắt
ngang cống tại tim đập với sơ đồ tải trọng (1/50); sơ đồ cửa van – trục vít – máy đóng mở VĐ20 với giá đỡ trung gian
(không tỷ lệ); các bảng 6.1 ÷ 6.3, tiêu năng, khối lượng và ghi chú. Kích thước kết cấu không có trong hồ sơ ghi
"giả thiết".
"""
import math

import du_lieu_ban_ve as D
from bv_lib import To, KhungNhin, so

C = D.CONG


def z_day(x):
    return C['Z_nguong'] - C['i'] * x


def mui_ten(to, p0, p1, lop='NET_CHINH', dai=2.0, **kw):
    to.line(p0, p1, lop, **kw)
    a = math.atan2(p1[1] - p0[1], p1[0] - p0[0])
    for da in (2.7, -2.7):
        to.line(p1, (p1[0] + dai * math.cos(a + da), p1[1] + dai * math.sin(a + da)), lop, **kw)


# ====================================================================== cắt dọc
CX = lambda x: 40.0 + (x + 8.0) * 4.0                # ngang 1/250
CY = lambda z: 425.0 + (z - 62.0) * 20.0             # đứng 1/50


def cat_doc(to):
    to.ten_hinh('CẮT DỌC CỐNG LẤY NƯỚC – ĐƯỜNG MẶT NƯỚC TRONG THÂN CỐNG', 40, 576, None, h=4.5, can='LEFT')
    to.text('TỶ LỆ: NGANG 1/250, ĐỨNG 1/50 (sơ đồ)', (40, 569.5), 3.0)
    xr = 33.0
    to.line((xr, CY(62.0)), (xr, CY(68.5)), 'NET_CHINH')
    for k in range(13):
        z = 62.0 + 0.5 * k
        to.line((xr, CY(z)), (xr + (1.5 if k % 2 == 0 else 0.8), CY(z)), 'NET_CHINH')
        if k % 2 == 0:
            to.text(so(z, 1), (xr - 0.8, CY(z)), 1.7, 'MR')
    L, H, t = C['L'], C['H'], C['t_nap']
    # thân cống: bản đáy, bản nắp (dày giả thiết 0,4 m)
    for z0, z1 in ((-t, 0.0), (H, H + t)):
        p = [(CX(0), CY(C['Z_nguong'] + z0)), (CX(L), CY(z_day(L) + z0)), (CX(L), CY(z_day(L) + z1)),
             (CX(0), CY(C['Z_nguong'] + z1))]
        to.hatch(p, 'AR-CONC', 0.03)
        to.pline(p, 'NET_CHINH', dong=True)
    # tháp van (sơ đồ) + khe van
    tw = [(-4.5, 63.6), (0.0, 63.6), (0.0, C['Z_nguong'] - t), (-4.5, C['Z_nguong'] - t)]
    p = [(CX(x), CY(z)) for x, z in tw]
    to.hatch(p, 'AR-CONC', 0.03)
    to.pline(p, 'NET_CHINH', dong=True)
    for xa, xb in ((-1.6, -1.1), (-0.5, 0.0)):
        p = [(CX(xa), CY(C['Z_nguong'] + H)), (CX(xb), CY(C['Z_nguong'] + H)), (CX(xb), CY(68.4)),
             (CX(xa), CY(68.4))]
        to.hatch(p, 'AR-CONC', 0.03)
        to.pline(p, 'NET_CHINH', dong=True)
    to.line((CX(-1.6), CY(68.4)), (CX(0.0), CY(68.4)), 'NET_MANH', linetype='DASHED', ltscale=0.3)
    to.text('THÁP VAN', (CX(-0.8), CY(68.4) + 2.0), 2.0, 'CENTER')
    to.pline([(CX(-1.0), CY(C['Z_nguong'])), (CX(-0.6), CY(C['Z_nguong'])), (CX(-0.6), CY(C['Z_nguong'] + H)),
              (CX(-1.0), CY(C['Z_nguong'] + H))], 'NET_MANH', dong=True, linetype='DASHED', ltscale=0.2)
    to.chu_thich_mui_ten((CX(-0.8), CY(66.2)), (CX(-5.0), CY(63.0)), 'Cửa van 1,8 × 2,0 m', h=1.8, rong=20)
    # mực nước hồ
    to.muc_nuoc(CX(-8.0), CX(-1.6), CY(D.MN['MNC']), 'MNC', D.MN['MNC'], nhan_x=CX(-7.4))
    to.text('MNDBT +74,70, MNLTK +77,63: ngoài phạm vi hình vẽ', (CX(-8.0), CY(68.4) + 7.0), 1.8, 'LEFT', 'NUOC')
    # đường mặt nước
    for pts, mau, kieu, ten in ((D.CONG_DMN_6_54, 5, 'CONTINUOUS', 'Q = 6,54 m³/s'),
                                (D.CONG_DMN_10_00, 1, 'DASHED', 'Q = 10,0 m³/s')):
        to.pline([(CX(x), CY(z_day(x) + h)) for x, h in pts], 'NUOC', color=mau, linetype=kieu, ltscale=0.3)
    to.chu_thich_mui_ten((CX(12.0), CY(z_day(12.0) + 1.079)), (CX(16.0), CY(68.0)),
                         'Q = 10,0 m³/s (QTVH): h = 1,45 → 0,86 m', h=1.8, rong=62)
    to.chu_thich_mui_ten((CX(21.0), CY(z_day(21.0) + 0.719)), (CX(25.0), CY(63.3)),
                         'Q = 6,54 m³/s: h = 1,09 → 0,62 m', h=1.8, rong=52)
    # tim đập
    xt = C['x_tim_dap']
    to.line((CX(xt), CY(62.3)), (CX(xt), CY(68.5)), 'TRUC')
    to.text('TIM ĐẬP', (CX(xt) - 1.0, CY(68.5) - 1.0), 2.0, 'TL', rot=90)
    to.text('Đỉnh đập sau sửa chữa +78,85; đất đắp trên nắp H = 11,75 m', (CX(xt) + 1.5, CY(68.5) - 1.6), 1.8, 'LEFT')
    to.text('(đỉnh nắp ≈ +67,10, nắp dày 0,4 m – giả thiết)', (CX(xt) + 1.5, CY(68.5) - 4.4), 1.8, 'LEFT')
    # cửa ra, bể tiêu năng (sơ đồ), kênh hạ lưu
    zr = z_day(L)
    to.line((CX(L), CY(zr)), (CX(L + 1.0), CY(zr)), 'NET_CHINH')
    d_be = 1.2
    be = [(L + 1.0, zr), (L + 2.5, zr - d_be), (L + 10.0, zr - d_be), (L + 11.5, zr), (L + 18.0, zr)]
    to.pline([(CX(x), CY(z)) for x, z in be], 'NET_CHINH', linetype='DASHED', ltscale=0.3)
    to.text('Bể tiêu năng hiện trạng (sơ đồ) – đối chiếu hồ sơ hoàn công:', (CX(L + 0.5), CY(zr - d_be) - 3.6), 1.8, 'LEFT')
    to.text('yêu cầu d ≈ 0,7 ÷ 1,2 m; Lb ≈ 5,7 ÷ 8,3 m', (CX(L + 0.5), CY(zr - d_be) - 6.4), 1.8, 'LEFT')
    to.line((CX(L + 10.0), CY(zr + C['hh'])), (CX(L + 18.0), CY(zr + C['hh'])), 'NUOC', linetype='DASHED', ltscale=0.3)
    to.text('MN kênh hạ lưu, hh = 1,2 m (giả thiết)', (CX(L + 18.0), CY(zr + C['hh']) + 1.0), 1.8, 'RIGHT', 'NUOC')
    # cao trình, kích thước
    to.cao_trinh(CX(1.2), CY(C['Z_nguong']), 0, h=1.8, chu='+65,60 ngưỡng')
    to.cao_trinh(CX(L - 1.2), CY(zr), 0, ben='L', h=1.8, chu='+63,59')
    to.text('i = 0,03 (QTVH; một số tài liệu ghi 3‰ – cần xác minh)', (CX(44.0), CY(z_day(44.0) - t) - 3.2), 1.8, 'CENTER',
            rot=-math.degrees(math.atan(C['i'] * 20 / 4)))
    yk = CY(62.0) + 1.0
    to.dim((CX(0.0), yk), (CX(L), yk), 0, 250)
    to.dim((CX(-2.2), CY(C['Z_nguong'])), (CX(-2.2), CY(C['Z_nguong'] + H)), 0, 50, goc=90)
    # bảng số liệu
    y0, hd = CY(62.0) - 4.0, 6.0
    xs = [0, 6, 12, 18, 24, 30, 36, 42, 48, 54, 60, 67]
    h1 = dict(D.CONG_DMN_6_54)
    h2 = dict(D.CONG_DMN_10_00)

    def nh(dic, x):
        ks = sorted(dic)
        for a, b in zip(ks[:-1], ks[1:]):
            if a <= x <= b:
                return dic[a] + (dic[b] - dic[a]) * (x - a) / (b - a)
        return dic[ks[-1]]
    hang = [('x (m)', lambda x: '%d' % x), ('Cao độ đáy (m)', lambda x: so(z_day(x), 2)),
            ('h Q = 6,54 (m)', lambda x: so(nh(h1, x), 3)), ('h Q = 10,0 (m)', lambda x: so(nh(h2, x), 3)),
            ('Lưu không Q = 10 (m)', lambda x: so(H - nh(h2, x), 2))]
    x_b0 = 25.0
    for k, (ten, f) in enumerate(hang):
        yt, yb = y0 - k * hd, y0 - (k + 1) * hd
        to.line((x_b0, yt), (CX(70.0), yt), 'BANG')
        to.text(ten, (x_b0 + 1.0, 0.5 * (yt + yb)), 1.9, 'ML', 'CHU_DAM' if k == 0 else 'CHU',
                'VN_DAM' if k == 0 else 'VN')
        for x in xs:
            to.text(f(x), (CX(x), 0.5 * (yt + yb)), 1.7, 'MC')
    yb = y0 - len(hang) * hd
    to.line((x_b0, yb), (CX(70.0), yb), 'BANG')
    for x in (x_b0, CX(-3.5), CX(70.0)):
        to.line((x, y0), (x, yb), 'BANG')


# ====================================================================== mặt cắt ngang + tải trọng
def mat_cat_ngang(to, x0, y0):
    to.ten_hinh('MẶT CẮT NGANG CỐNG TẠI TIM ĐẬP – SƠ ĐỒ TẢI TRỌNG', x0 - 45, y0 + 152, '1/25', h=3.8, can='LEFT')
    B, H, t = C['B'], C['H'], C['t_nap']
    kn = KhungNhin(x0, y0, -B / 2 - t, 0.0, 25)          # gốc: mép ngoài trái, đáy ngoài
    ngoai = [(-B / 2 - t, 0), (B / 2 + t, 0), (B / 2 + t, H + 2 * t), (-B / 2 - t, H + 2 * t)]
    trong = [(-B / 2, t), (B / 2, t), (B / 2, t + H), (-B / 2, t + H)]
    for a, b in ((ngoai[0], (-B / 2, H + 2 * t)), ((B / 2, 0), ngoai[2])):
        p = [kn(a[0], a[1]), kn(b[0], a[1]), kn(b[0], b[1]), kn(a[0], b[1])]
        to.hatch(p, 'AR-CONC', 0.02)
    for a, b in (((-B / 2, 0), (B / 2, t)), ((-B / 2, t + H), (B / 2, H + 2 * t))):
        p = [kn(a[0], a[1]), kn(b[0], a[1]), kn(b[0], b[1]), kn(a[0], b[1])]
        to.hatch(p, 'AR-CONC', 0.02)
    to.pline([kn(*q) for q in ngoai], 'NET_CHINH', dong=True)
    to.pline([kn(*q) for q in trong], 'NET_CHINH', dong=True)
    # nước trong cống (Q = 10 m³/s, h ≤ 1,45 m)
    to.line(kn(-B / 2, t + 1.45), kn(B / 2, t + 1.45), 'NUOC', color=1, linetype='DASHED', ltscale=0.3)
    to.text('h ≤ 1,45 m', (kn(0, 0)[0], kn(0, t + 1.45)[1] - 3.0), 1.8, 'CENTER', 'NUOC')
    # áp lực đất đứng q = 242 kPa
    yq = kn(0, H + 2 * t)[1]
    xa, xb = kn(-B / 2 - t, 0)[0], kn(B / 2 + t, 0)[0]
    to.rect(xa, yq + 8.0, xb - xa, 0.01, 'CUNG_TRUOT')
    for k in range(7):
        x = xa + (xb - xa) * k / 6
        mui_ten(to, (x, yq + 8.0), (x, yq + 0.3), 'CUNG_TRUOT', 1.4)
    to.text('q = 242 kPa', (0.5 * (xa + xb), yq + 9.5), 2.4, 'CENTER', 'CUNG_TRUOT')
    # áp lực ngang e1 = 163, e2 = 202 kPa (hai bên)
    for sgn, xx in ((-1, xa), (1, xb)):
        ytop, ybot = yq, kn(0, 0)[1]
        e1, e2 = 163 / 202 * 18.0, 18.0
        to.pline([(xx + sgn * e1, ytop), (xx + sgn * e2, ybot)], 'CUNG_TRUOT')
        for k in range(5):
            y = ytop + (ybot - ytop) * k / 4
            e = e1 + (e2 - e1) * k / 4
            mui_ten(to, (xx + sgn * e, y), (xx + sgn * 0.3, y), 'CUNG_TRUOT', 1.3)
    to.text('e1 = 163 kPa', (xa - 16.0, yq + 2.0), 2.0, 'RIGHT', 'CUNG_TRUOT')
    to.text('e2 = 202 kPa', (xa - 19.0, kn(0, 0)[1] - 3.5), 2.0, 'RIGHT', 'CUNG_TRUOT')
    # phản lực nền
    yb = kn(0, 0)[1]
    for k in range(7):
        x = xa + (xb - xa) * k / 6
        mui_ten(to, (x, yb - 6.0), (x, yb - 0.3), 'CUNG_TRUOT', 1.2)
    to.text('Phản lực nền (đá lớp II)', (0.5 * (xa + xb), yb - 9.5), 2.0, 'CENTER', 'CUNG_TRUOT')
    # kích thước
    to.dim(kn(-B / 2, t), kn(B / 2, t), 8.0, 25)
    to.dim(kn(B / 2, t), kn(B / 2, t + H), -8.0, 25, goc=90)
    to.dim(kn(-B / 2 - t, 0), kn(-B / 2, 0), -6.0, 25, text='40')
    to.text('Nắp, đáy, tường dày 40 (giả thiết)', (0.5 * (xa + xb), yb - 15.0), 2.0, 'CENTER')
    to.text('Nâng đỉnh +78,35 → +78,85: q tăng 10,3 kPa (4,4 %)', (0.5 * (xa + xb), yb - 19.0), 2.0, 'CENTER')
    to.text('Đường bão hòa tại tim đập +65,7 ÷ +73,4 m: kể áp lực nước ngoài, đẩy nổi', (0.5 * (xa + xb), yb - 23.0),
            2.0, 'CENTER')


# ====================================================================== sơ đồ cửa van – trục vít – máy đóng mở
def so_do_van(to, x0, y0):
    to.ten_hinh('SƠ ĐỒ CỬA VAN – TRỤC VÍT – MÁY ĐÓNG MỞ VĐ20', x0, y0 + 160, 'không tỷ lệ', h=3.8, can='LEFT')
    xc = x0 + 45
    # khe van (2 bên), cửa van
    for sgn in (-1, 1):
        to.rect(xc + sgn * 20 - (6 if sgn > 0 else 0), y0, 6, 120, 'NET_CHINH')
    to.hatch([(xc - 20, y0), (xc - 14, y0), (xc - 14, y0 + 120), (xc - 20, y0 + 120)], 'AR-CONC', 0.03)
    to.hatch([(xc + 14, y0), (xc + 20, y0), (xc + 20, y0 + 120), (xc + 14, y0 + 120)], 'AR-CONC', 0.03)
    gv = [(xc - 14, y0 + 2), (xc + 14, y0 + 2), (xc + 14, y0 + 30), (xc - 14, y0 + 30)]
    to.pline(gv, 'NET_CHINH', dong=True)
    for k in range(1, 4):
        to.line((xc - 14, y0 + 2 + 7 * k), (xc + 14, y0 + 2 + 7 * k), 'NET_MANH')
    to.text('CỬA VAN', (xc, y0 + 16), 2.2, 'MC', 'CHU_DAM', 'VN_DAM')
    to.text('1,8 × 2,0 m', (xc, y0 + 12.5), 1.9, 'MC')
    for yy in (y0 + 6, y0 + 25):
        for sgn in (-1, 1):
            to.rect(xc + sgn * 14 - (1.6 if sgn > 0 else 0), yy - 1.5, 1.6, 3.0, 'SUA_CHUA')
    to.chu_thich_mui_ten((xc + 13.2, y0 + 25), (xc + 30, y0 + 34),
                         'Thay tấm trượt: đồng – graphit hoặc UHMW, f ≤ 0,3', h=1.8, rong=74)
    to.chu_thich_mui_ten((xc + 14, y0 + 4), (xc + 30, y0 + 10), 'Thay gioăng cao su kín nước', h=1.8, rong=44)
    # trục vít D85
    yv0, yv1 = y0 + 30, y0 + 122
    to.line((xc - 1.2, yv0), (xc - 1.2, yv1), 'NET_CHINH')
    to.line((xc + 1.2, yv0), (xc + 1.2, yv1), 'NET_CHINH')
    for k in range(int((yv1 - yv0) / 2.0)):
        y = yv0 + 2.0 * k
        to.line((xc - 1.2, y), (xc + 1.2, y + 1.2), 'NET_MANH')
    # giá đỡ trung gian (đề xuất)
    ym = 0.5 * (yv0 + yv1)
    p = [(xc - 20, ym - 1.5), (xc + 20, ym - 1.5), (xc + 20, ym + 1.5), (xc - 20, ym + 1.5)]
    to.hatch(p, 'ANSI31', 0.3, lop='SUA_CHUA', mau_to=1)
    to.pline(p, 'SUA_CHUA', dong=True)
    to.circle((xc, ym), 2.4, 'SUA_CHUA')
    to.chu_thich_mui_ten((xc + 20, ym), (xc + 30, ym + 10),
                         'Bổ sung giá đỡ dẫn hướng trung gian: L = 6 m → 2 × 3 m', h=1.8, rong=82)
    to.text('(Kođ 0,67 ÷ 1,62 → 2,69 ÷ 6,48 ≥ 2)', (xc + 30.8, ym + 7.2), 1.8, 'LEFT')
    to.dim((xc - 26, yv0), (xc - 26, ym), 0, 1, goc=90, text='L = 3,0 m')
    to.dim((xc - 26, ym), (xc - 26, yv1), 0, 1, goc=90, text='L = 3,0 m')
    to.text('Trục vít D85 (d chân ren ≈ 72 mm)', (xc + 4.0, yv0 + 16.0), 1.8, 'LEFT')
    # sàn đặt máy, máy VĐ20, tủ điện
    to.rect(xc - 30, yv1, 60, 3, 'NET_CHINH')
    to.hatch([(xc - 30, yv1), (xc + 30, yv1), (xc + 30, yv1 + 3), (xc - 30, yv1 + 3)], 'AR-CONC', 0.03)
    may = [(xc - 9, yv1 + 3), (xc + 9, yv1 + 3), (xc + 9, yv1 + 17), (xc - 9, yv1 + 17)]
    to.hatch(may, 'ANSI31', 0.5, lop='SUA_CHUA', mau_to=1)
    to.pline(may, 'SUA_CHUA', dong=True)
    to.text('VĐ20', (xc, yv1 + 10), 2.6, 'MC', 'CHU_DAM', 'VN_DAM')
    to.rect(xc + 16, yv1 + 3, 10, 15, 'SUA_CHUA')
    to.text('Tủ điện', (xc + 21, yv1 + 20.0), 1.8, 'CENTER')
    to.chu_thich_mui_ten((xc - 9, yv1 + 14), (xc - 24, yv1 + 25),
                         'Máy vít VĐ20 (20 T, 4,5 kW, 380 V/3 pha) – thay mới', h=1.8, rong=74)
    to.text('+ công tắc hành trình, rơ le bảo vệ quá tải', (xc - 24.8, yv1 + 22.5), 1.8, 'RIGHT')


# ====================================================================== biểu đồ
def khung_bieu_do(to, x0, y0, w, h, xs, ys, nhan_x, nhan_y, tieu_de):
    """Khung trục: xs = (xmin, xmax, [ticks]); ys = (ymin, ymax, [ticks]). Trả về hàm đổi tọa độ."""
    fx = lambda v: x0 + w * (v - xs[0]) / (xs[1] - xs[0])
    fy = lambda v: y0 + h * (v - ys[0]) / (ys[1] - ys[0])
    to.rect(x0, y0, w, h, 'NET_CHINH')
    to.text(tieu_de, (x0 + w / 2, y0 + h + 4.0), 2.4, 'CENTER', 'CHU_DAM', 'VN_DAM')
    for v in xs[2]:
        to.line((fx(v), y0), (fx(v), y0 + h), 'NET_MANH')
        to.text(so(v, 1) if isinstance(v, float) and v % 1 else str(v).replace('.', ','), (fx(v), y0 - 2.8), 1.8, 'CENTER')
    for v in ys[2]:
        to.line((x0, fy(v)), (x0 + w, fy(v)), 'NET_MANH')
        to.text(str(v).replace('.', ','), (x0 - 1.2, fy(v)), 1.8, 'MR')
    to.text(nhan_x, (x0 + w / 2, y0 - 7.0), 2.0, 'CENTER')
    to.text(nhan_y, (x0 - 9.0, y0 + h / 2), 2.0, 'MC', rot=90)
    return fx, fy


def bieu_do_Q(to, x0, y0):
    """Khả năng tháo khi mở hết cửa: Q = m·b·√(2g)·H^1,5 (6-1), m = 0,35; b = 1,8 m."""
    fx, fy = khung_bieu_do(to, x0, y0, 150.0, 90.0, (66.0, 69.0, [66.0, 66.5, 67.0, 67.5, 68.0, 68.5, 69.0]),
                           (0, 16, [0, 4, 8, 12, 16]), 'Mực nước hồ (m)', 'Q (m³/s)',
                           'KHẢ NĂNG THÁO CỦA CỐNG KHI MỞ HẾT CỬA (công thức 6-1)')
    Q = lambda z: 0.35 * C['B'] * math.sqrt(2 * 9.81) * max(z - C['Z_nguong'], 0) ** 1.5
    zs = [66.0 + 0.05 * k for k in range(61)]
    to.pline([(fx(z), fy(Q(z))) for z in zs], 'NUOC', color=5)
    for q, ten in ((6.54, 'Q = 6,54'), (10.0, 'Q = 10,0')):
        to.line((fx(66.0), fy(q)), (fx(69.0), fy(q)), 'SUA_CHUA', linetype='DASHED', ltscale=0.3)
        to.text(ten, (fx(66.05), fy(q) + 0.8), 1.8, 'LEFT', 'SUA_CHUA')
    for z, ten in ((D.MN['MNC'], 'MNC +67,50: Q = 7,31'), (67.94, '+67,94: Q = 10,0')):
        to.line((fx(z), fy(0)), (fx(z), fy(Q(z))), 'NET_CHINH', linetype='DASHED', ltscale=0.3)
        to.circle((fx(z), fy(Q(z))), 0.9, 'NET_CHINH')
        to.text(ten, (fx(z) + 1.2, fy(Q(z)) - 3.2), 1.8, 'LEFT')


def bieu_do_luc(to, x0, y0):
    """Lực mở cửa van (T) theo hệ số ma sát f (Bảng 6.3) và hệ số ổn định trục vít."""
    fx, fy = khung_bieu_do(to, x0, y0, 150.0, 100.0, (0, 3, []), (0, 30, [0, 5, 10, 15, 20, 25, 30]),
                           'Hệ số ma sát gối trượt f', 'P mở (T)', 'LỰC MỞ CỬA VAN THEO HỆ SỐ MA SÁT (Bảng 6.3)')
    du_lieu = {'MNDBT': [14.34, 17.84, 21.34], 'MNLTK': [18.53, 23.29, 28.06]}
    MAU = {5: 151, 1: 11, 6: 211, 3: 91}
    for k, f in enumerate(('0,30', '0,40', '0,50')):
        xc = fx(k + 0.5)
        to.text('f = ' + f, (xc, y0 - 2.8), 1.8, 'CENTER')
        for j, (ten, mau) in enumerate((('MNDBT', 5), ('MNLTK', 1))):
            v = du_lieu[ten][k]
            xa = xc - 13.0 + j * 13.0
            p = [(xa, fy(0)), (xa + 11.0, fy(0)), (xa + 11.0, fy(v)), (xa, fy(v))]
            to.hatch(p, 'SOLID', lop='NET_MANH', mau_to=MAU[mau if v <= 20 else 6])
            to.pline(p, 'NET_CHINH', dong=True)
            to.text(so(v, 1), (xa + 5.5, fy(v) + 1.0), 1.7, 'CENTER')
    to.line((fx(0), fy(20)), (fx(3), fy(20)), 'SUA_CHUA')
    to.text('Sức nâng máy VĐ20 = 20 T', (fx(0.05), fy(20) + 1.0), 1.9, 'LEFT', 'SUA_CHUA')
    for j, (ten, mau) in enumerate((('MNDBT +74,70', 5), ('MNLTK +77,63', 1), ('Vượt sức nâng 20 T', 6))):
        xl = x0 + 2.0 + j * 42.0
        p = [(xl, y0 - 14.5), (xl + 5, y0 - 14.5), (xl + 5, y0 - 11.5), (xl, y0 - 11.5)]
        to.hatch(p, 'SOLID', lop='NET_MANH', mau_to=MAU[mau])
        to.pline(p, 'NET_CHINH', dong=True)
        to.text(ten, (xl + 6.5, y0 - 13.0), 1.8, 'ML')
    # hệ số ổn định trục vít
    x1 = x0 + 190.0
    fx2, fy2 = khung_bieu_do(to, x1, y0, 160.0, 100.0, (0, 6, []), (0, 7, [0, 1, 2, 3, 4, 5, 6, 7]),
                             'Trường hợp (MN hồ – f)', 'Kođ', 'ỔN ĐỊNH TRỤC VÍT KHI ĐÓNG CƯỠNG BỨC')
    nhan = ['DBT 0,3', 'DBT 0,4', 'DBT 0,5', 'LTK 0,3', 'LTK 0,4', 'LTK 0,5']
    K6 = [1.62, 1.19, 0.94, 1.13, 0.84, 0.67]
    K3 = [6.48, 4.77, 3.77, 4.53, 3.38, 2.69]
    for k in range(6):
        xc = fx2(k + 0.5)
        to.text(nhan[k], (xc, y0 - 2.8), 1.6, 'CENTER')
        for j, (v, mau) in enumerate(((K6[k], 1), (K3[k], 3))):
            xa = xc - 10.0 + j * 10.0
            p = [(xa, fy2(0)), (xa + 9.0, fy2(0)), (xa + 9.0, fy2(v)), (xa, fy2(v))]
            to.hatch(p, 'SOLID', lop='NET_MANH', mau_to=MAU[mau])
            to.pline(p, 'NET_CHINH', dong=True)
            to.text(so(v, 2), (xa + 4.5, fy2(v) + 1.0), 1.6, 'CENTER')
    to.line((fx2(0), fy2(2)), (fx2(6), fy2(2)), 'SUA_CHUA')
    to.text('[Kođ] = 2', (fx2(6) + 1.0, fy2(2)), 1.9, 'ML', 'SUA_CHUA')
    for j, (ten, mau) in enumerate((('Trục tự do L = 6 m (hiện trạng)', 1), ('L = 3 m (có giá đỡ trung gian)', 3))):
        xl = x1 + 2.0 + j * 72.0
        p = [(xl, y0 - 14.5), (xl + 5, y0 - 14.5), (xl + 5, y0 - 11.5), (xl, y0 - 11.5)]
        to.hatch(p, 'SOLID', lop='NET_MANH', mau_to=MAU[mau])
        to.pline(p, 'NET_CHINH', dong=True)
        to.text(ten, (xl + 6.5, y0 - 13.0), 1.8, 'ML')



# ====================================================================== bảng, ghi chú
def cac_bang(to):
    xL = 430.0
    cot1 = [86, 16, 18, 52]
    hang1 = [['Thông số', 'Ký hiệu', 'Đơn vị', 'Giá trị'],
             ['Lưu lượng thiết kế', 'Qtk', 'm³/s', '6,54 (BCNCKT); 10,0 (QTVH)'],
             ['Mực nước thượng lưu nhỏ nhất / lớn nhất khi lấy nước', 'ZTL', 'm', '+67,50 / +74,70'],
             ['Cao trình ngưỡng cửa vào', 'Zng', 'm', '+65,60'],
             ['Kích thước mặt cắt b × h', '–', 'm', '1,8 × 2,0'],
             ['Chiều dài thân cống; độ dốc đáy', 'L ; i', 'm ; –', '67 ; 0,03 (cần xác minh)'],
             ['Hệ số nhám bê tông', 'n', '–', '0,014'],
             ['Chiều sâu nước kênh hạ lưu (giả thiết)', 'hh', 'm', '1,20'],
             ['Máy đóng mở', '–', '–', 'VĐ20 (20 T), trục vít D85, L = 6 m']]
    y = to.bang(xL, 576, cot1, hang1, h=1.9, cao_dong=4.6, tieu_de='THÔNG SỐ TÍNH TOÁN CỐNG (Bảng 6.1)')
    to.text('Khả năng tháo khi mở hết cửa (đập tràn đỉnh rộng, m = 0,35): ở MNC H0 = 1,90 m → Q = 7,31 m³/s > 6,54; '
            'Q = 10 m³/s cần MN hồ ≥ +67,94 m.', (xL, y - 3.5), 1.8)
    cot2 = [20, 24, 24, 26, 32, 18]
    hang2 = [['Q (m³/s)', 'h đầu cống (m)', 'h cuối cống (m)', 'v cuối cống (m/s)', 'Lưu không nhỏ nhất (m)', 'h/H']] + \
            [list(r) for r in D.CONG_BANG62]
    to.bang(xL + 186, 576, cot2, hang2, h=1.9, cao_dong=4.6, tieu_de='ĐƯỜNG MẶT NƯỚC TRONG CỐNG (Bảng 6.2)')
    to.text('i = 0,03: h0 = 0,59 ÷ 0,80 m < hk = 1,10 ÷ 1,47 m → dòng xiết, không áp;',
            (xL + 186, 576 - 3 * 4.6 - 3.5), 1.8)
    to.text('nếu i = 3‰: h0 = 1,39 ÷ 1,96 m > hk – dòng êm, gần đầy cống (phải tính lại).',
            (xL + 186, 576 - 3 * 4.6 - 6.3), 1.8)
    cot3 = [20, 20, 16, 20, 18, 22]
    hang3 = [['Q (m³/s)', 'h1 (m)', 'Fr1', 'h″ (m)', 'σ = hh/h″', 'Ln = 4,5h″ (m)']] + [list(r) for r in D.CONG_TIEU_NANG]
    y3 = to.bang(xL + 186, 538, cot3, hang3, h=1.9, cao_dong=4.6, tieu_de='TIÊU NĂNG SAU CỐNG')
    to.text('σ < 1,05: nước nhảy phóng xa nếu không có bể; bể cần d ≈ 1,05h″ − hh ≈ 0,7 ÷ 1,2 m.', (xL + 186, y3 - 3.5), 1.8)
    cot4 = [27, 11, 15, 16, 16, 14, 16, 16, 18, 18]
    hang4 = [['Mực nước hồ', 'f', 'W (kN)', 'T1 + T2\n(kN)', 'P mở\n(kN)', 'P mở\n(T)', 'So với\n20 T',
              'P đóng\n(kN)', 'Kođ trục\nL = 6 m', 'Kođ trục\nL = 3 m']] + [list(r) for r in D.CONG_BANG63]
    y4 = to.bang(xL, 500, cot4, hang4, h=1.9, cao_dong=4.6,
                 tieu_de='LỰC ĐÓNG MỞ CỬA VAN VÀ ỔN ĐỊNH TRỤC VÍT (Bảng 6.3)')
    to.text('P mở = nG·G + nms·(T1 + T2) + Wh; G = 20 kN (ước tính); Kođ = Pth/P đóng ≥ [Kođ] = 2 '
            '(Euler, μ = 0,7, d = 72 mm).', (xL, y4 - 3.5), 1.8)
    cot5 = [118, 16, 22]
    hang5 = [['Hạng mục (Bảng 7.1 – III)', 'Đơn vị', 'Khối lượng'],
             ['Máy đóng mở VĐ20 (kèm tủ điện, công tắc hành trình)', 'bộ', '2'],
             ['Thay gioăng, tấm trượt, giá đỡ trục trung gian; sơn cửa van', 'bộ', '2']]
    to.bang(xL, 440, cot5, hang5, h=1.9, cao_dong=4.6, tieu_de='KHỐI LƯỢNG SỬA CHỮA CỐNG')
    to.text('Số bộ theo TKCS – đối chiếu số cửa van trong tháp van (mô tả hiện trạng chỉ nêu một cửa van).',
            (xL, 440 - 3 * 4.6 - 3.5), 1.8)


def ghi_chu(to, x, y):
    gc = ('GHI CHÚ: \\P'
          '1. Cao trình ghi bằng m, kích thước ghi bằng cm. Cống hộp BTCT 1,8 × 2,0 m, dài 67 m, chảy không áp, đặt trên '
          'đá lớp II; tháp van thượng lưu, cửa van phẳng; bể tiêu năng và cụm chia nước hạ lưu còn tốt. Tháp van, bể tiêu '
          'năng, chiều dày kết cấu vẽ sơ đồ (giả thiết). \\P'
          '2. Kết cấu cống giữ nguyên. Thay máy đóng mở hư hỏng bằng máy vít VĐ20 (trục vít D85 dài 6,0 m, 4,5 kW, '
          '380 V/3 pha/50 Hz, kèm tủ điện, công tắc hành trình, rơ le bảo vệ quá tải). \\P'
          '3. Thay tấm trượt bằng vật liệu ma sát thấp (đồng – graphit hoặc polime UHMW, f ≤ 0,3) và gioăng cao su; bổ '
          'sung giá đỡ dẫn hướng trung gian chia trục vít thành 2 nhịp ≤ 3 m; sơn chống gỉ cửa van. \\P'
          '4. Vận hành: chỉ vận hành cống khi mực nước hồ ≤ MNDBT, trừ trường hợp khẩn cấp; với nhiệm vụ Q = 10 m³/s, khi '
          'MN hồ +67,50 ÷ +67,94 m cống chỉ tháo được 7,3 ÷ 10 m³/s (ghi vào QTVH). \\P'
          '5. Kiểm tra bên trong thân cống bằng camera/robot khi đóng cửa van (thời gian ngừng cấp nước giữa hai vụ); nếu '
          'phát hiện nứt, rò nước ở khớp nối thì tính toán kết cấu và xử lý bằng vữa, keo epoxy hoặc ống lót. \\P'
          '6. Phải xác minh độ dốc đáy cống (3 % hay 3‰) bằng cao trình đáy cửa vào – cửa ra trên hồ sơ hoàn công trước '
          'khi chốt kết quả thủy lực; đối chiếu kích thước bể tiêu năng với yêu cầu ở bảng tiêu năng.')
    to.mtext(gc, (x, y), 2.0, rong=196)


def ve(thu_muc):
    to = To(7, 'CỐNG LẤY NƯỚC – THAY MÁY ĐÓNG MỞ\nKIỂM TRA THỦY LỰC VÀ LỰC ĐÓNG MỞ', 'Như hình vẽ', 'HCG-BV07')
    cat_doc(to)
    mat_cat_ngang(to, 95.0, 200.0)
    so_do_van(to, 250.0, 40.0)
    bieu_do_Q(to, 48.0, 40.0)
    cac_bang(to)
    ghi_chu(to, 432.0, 405.0)
    bieu_do_luc(to, 450.0, 130.0)
    to.luu(thu_muc + '/BV07_CongLayNuoc.dxf', thu_muc + '/BV07_CongLayNuoc.pdf', thu_muc + '/BV07_CongLayNuoc.png')
    return to


if __name__ == '__main__':
    import sys
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
