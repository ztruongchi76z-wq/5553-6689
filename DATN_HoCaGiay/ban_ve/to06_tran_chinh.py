# -*- coding: utf-8 -*-
"""BV-06: TRÀN XẢ LŨ CHÍNH – SỬA CHỮA DỐC NƯỚC, KIỂM TRA THỦY LỰC VÀ ỔN ĐỊNH NGƯỠNG (Chương 4).

Gồm: cắt dọc tràn chính với đường mặt nước P = 1 %, 0,2 % (ngang 1/250, đứng 1/100) và bảng số liệu; mặt bằng sơ đồ
(1/250); mặt cắt ngang dốc nước 1-1 tại x = 20 m (1/100); mặt cắt ngưỡng và sơ đồ lực kiểm tra ổn định (1/50);
chi tiết xử lý mặt đáy dốc và khe co giãn; các bảng 4.1, 4.2, 4.4, mũi phun, tràn sự cố và khối lượng.
Hình học ngưỡng, dốc nước theo hồ sơ (tinh_toan/so_lieu.py); các kích thước không có trong hồ sơ ghi "giả thiết".
"""
import math

import du_lieu_ban_ve as D
from bv_lib import To, KhungNhin, so

T = D.TC
G = 9.81


def b_doc(x):
    return T['B_dau'] + (T['B_cuoi'] - T['B_dau']) * min(max(x, 0.0), T['L_chuyen']) / T['L_chuyen']


def z_day(x):
    return T['Z_dau_doc'] - T['i'] * x


def h_tuong(x):
    p = D.TC_TUONG
    for (a, ha), (b, hb) in zip(p[:-1], p[1:]):
        if a <= x <= b:
            return ha + (hb - ha) * (x - a) / (b - a)
    return p[-1][1]


def mui_ten(to, p0, p1, lop='NET_CHINH', dai=2.5, **kw):
    to.line(p0, p1, lop, **kw)
    a = math.atan2(p1[1] - p0[1], p1[0] - p0[0])
    for da in (2.7, -2.7):
        to.line(p1, (p1[0] + dai * math.cos(a + da), p1[1] + dai * math.sin(a + da)), lop, **kw)


# ====================================================================== cắt dọc
CX = lambda x: 40.0 + (x + 14.0) * 4.0              # ngang 1/250 (phóng 1,25 → 1/200)
CY = lambda z: 412.0 + (z - 66.0) * 8.0             # đứng 1/125 (đặt lên A0 phóng 1,25 → 1/100)
R_MP, TH_MP = 5.0, math.radians(25.0)               # mũi phun (sơ đồ): bán kính, góc hắt giả thiết


def mui_phun_hh():
    """Tọa độ mép mũi phun (sơ đồ)."""
    x0, z0 = T['L_doc'], z_day(T['L_doc'])
    return x0 + R_MP * math.sin(TH_MP), z0 + R_MP * (1 - math.cos(TH_MP))


def cat_doc(to):
    to.ten_hinh('CẮT DỌC TRÀN CHÍNH – ĐƯỜNG MẶT NƯỚC VÀ NỘI DUNG SỬA CHỮA', 40, CY(79.5) + 26.0, None, h=4.5,
                can='LEFT')
    to.text('TỶ LỆ: NGANG 1/250, ĐỨNG 1/125 (sơ đồ)', (40, CY(79.5) + 19.5), 3.0)
    # thước cao độ
    xr = 33.0
    to.line((xr, CY(66.0)), (xr, CY(79.5)), 'NET_CHINH')
    for z in range(66, 80):
        to.line((xr, CY(z)), (xr + 1.5, CY(z)), 'NET_CHINH')
        to.text('%d,0' % z, (xr - 0.8, CY(z)), 1.7, 'MR')
    # sân thượng lưu, lớp bồi lắng, ngưỡng
    xs0 = -14.0
    to.line((CX(xs0), CY(T['Z_san'])), (CX(-5.05), CY(T['Z_san'])), 'NET_CHINH')
    boi = [(xs0, 74.0), (xs0, 74.45), (-11.0, 74.38), (-9.0, 74.30), (-7.0, 74.18), (-5.05, 74.10), (-5.05, 74.0)]
    to.hatch([(CX(x), CY(z)) for x, z in boi], 'AR-SAND', 0.08)
    to.pline([(CX(x), CY(z)) for x, z in boi[1:-1]], 'NET_MANH', linetype='DASHED', ltscale=0.3)
    to.chu_thich_mui_ten((CX(-10.0), CY(74.33)), (CX(-8.0), CY(71.0)), 'Nạo vét đến +74,00', h=1.8, rong=20)
    p = [(CX(x), CY(z)) for x, z in D.TC_NGUONG]
    to.hatch(p, 'AR-CONC', 0.03)
    to.pline(p, 'NET_CHINH', dong=True)
    # nền đá dưới sân và dốc (vạch nền)
    to.line((CX(xs0), CY(73.5)), (CX(-5.05), CY(73.5)), 'NET_MANH', linetype='DASHED', ltscale=0.3)
    # bản đáy dốc nước (sơ đồ, dày 0,5 m) và mũi phun
    xm, zm = mui_phun_hh()
    tren = [(0.0, z_day(0.0)), (T['L_doc'], z_day(T['L_doc']))]
    cung = [(T['L_doc'] + R_MP * math.sin(a), z_day(T['L_doc']) + R_MP * (1 - math.cos(a)))
            for a in [TH_MP * k / 6 for k in range(1, 7)]]
    tren += cung
    duoi = [(xm + 0.6, zm), (xm + 0.6, z_day(T['L_doc']) - 1.5), (T['L_doc'] - 1.0, z_day(T['L_doc']) - 1.5),
            (T['L_doc'] - 1.0, z_day(T['L_doc'] - 1.0) - 0.5), (0.0, z_day(0.0) - 1.0)]
    vb = [(CX(x), CY(z)) for x, z in tren + duoi]
    to.hatch(vb, 'AR-CONC', 0.03)
    to.pline(vb, 'NET_CHINH', dong=True)
    # lớp vữa sửa chữa mặt đáy dốc (nét đỏ)
    to.pline([(CX(0.0), CY(z_day(0.0)) + 0.6), (CX(T['L_doc']), CY(z_day(T['L_doc'])) + 0.6)], 'SUA_CHUA')
    # đỉnh tường bên
    tw = [(x, z_day(x) + h_tuong(x)) for x in (0.0, 5.0, 10.0, 60.0)]
    to.pline([(CX(x), CY(z)) for x, z in tw], 'NET_CHINH', linetype='DASHDOT', ltscale=0.4)
    to.text('Đỉnh tường bên hiện trạng', (CX(38.0), CY(z_day(38.0) + 2.5) + 1.2), 1.9, 'LEFT')
    # mực nước thượng lưu và đường mặt nước
    for ten, Z, H, hc, cot_h, cot_hk, kieu, mau in (
            ('MNLTK', D.MN['MNLTK'], 2.93, 0.777, 2, 4, 'CONTINUOUS', 5),
            ('MNLKT', D.MN['MNLKT'], 3.84, 1.023, 5, 7, 'DASHED', 1)):
        dau = [(xs0, Z), (-9.0, Z - 0.05), (-6.0, Z - 0.25), (-4.0, Z - 0.55),
               (-2.7, T['Z_nguong'] + 0.62 * H), (-1.2, T['Z_nguong'] + 0.30 * H - 0.6), (0.0, z_day(0.0) + hc)]
        doc = [(r[0], z_day(r[0]) + r[cot_h]) for r in D.TC_BANG43]
        to.pline([(CX(x), CY(z)) for x, z in dau + doc[1:]], 'NUOC', linetype=kieu, ltscale=0.3, color=mau)
        hk = [(r[0], z_day(r[0]) + r[cot_hk]) for r in D.TC_BANG43]
        if ten == 'MNLKT':
            to.pline([(CX(x), CY(z)) for x, z in hk], 'NUOC', linetype='DOT', ltscale=0.3, color=1)
        to.muc_nuoc(CX(xs0), CX(-6.5), CY(Z), ten, Z, nhan_x=CX(-12.6) if ten == 'MNLTK' else CX(-9.0))
    to.cao_trinh(CX(-2.7), CY(T['Z_nguong']), 0, ben='L', h=1.8, chu='Ngưỡng +74,70')
    to.cao_trinh(CX(0.4), CY(z_day(0.0)), 0, h=1.8, chu='+73,30')
    to.cao_trinh(CX(T['L_doc'] - 0.4), CY(z_day(T['L_doc'])), 0, ben='L', h=1.8, chu='+68,50')
    to.cao_trinh(CX(-4.0), CY(T['Z_day_khoi']), 0, ben='L', h=1.7, chu='+72,30 (giả thiết)')
    # tia mũi phun, mực nước hạ lưu giả thiết
    z_hl = zm - 2.0
    to.line((CX(xm - 1.0), CY(z_hl)), (CX(80.0), CY(z_hl)), 'NUOC', linetype='DASHED', ltscale=0.3)
    to.text('MN hạ lưu giả thiết (a = 2,0 m dưới mũi phun)', (CX(80.0), CY(z_hl) - 3.6), 1.8, 'RIGHT', 'NUOC')
    for (lab, v, h, q, L, t), mau in zip(D.TC_MUI_PHUN, (5, 1)):
        y0 = 2.0 + 0.5 * h * math.cos(TH_MP)
        pts = []
        tt = 0.0
        while True:
            xj = 0.9 * v * math.cos(TH_MP) * tt
            yj = y0 + v * math.sin(TH_MP) * tt - 0.5 * G * tt * tt
            pts.append((xm + xj, z_hl + yj))
            if yj <= 0:
                break
            tt += 0.02
        to.pline([(CX(x), CY(z)) for x, z in pts], 'NUOC', color=mau)
        to.text('L = %s m (%s)' % (so(L, 1), lab), (CX(xm + L) + 1.0, CY(z_hl) + 1.0 + (3.0 if mau == 1 else 0.0)),
                1.8, 'LEFT', 'NUOC')
    to.text('Mũi phun (sơ đồ, θ = 25° giả thiết)', (CX(xm) - 1.0, CY(zm) + 9.0), 1.8, 'RIGHT')
    to.line((CX(xm), CY(zm) + 1.0), (CX(xm) - 1.0, CY(zm) + 8.6), 'NET_MANH')
    # chú thích sửa chữa
    to.chu_thich_mui_ten((CX(25.0), CY(z_day(25.0)) + 0.6), (CX(30.0), CY(71.0)),
                         'Đục 2 ÷ 3 cm BT mài mòn, trát vữa XM – polime ≥ 40 MPa, W8 dày 3 cm (chi tiết B)', h=1.8, rong=110)
    to.chu_thich_mui_ten((CX(12.0), CY(z_day(12.0)) - 1.0), (CX(15.0), CY(68.6)),
                         'Khe co giãn: thay vật liệu chèn bằng ma tít đàn hồi; thông lỗ thoát nước dưới bản đáy (chi tiết C)',
                         h=1.8, rong=124)
    to.chu_thich_mui_ten((CX(-1.5), CY(72.6)), (CX(4.0), CY(70.4)), 'Neo D25 (xem MC ngưỡng)', h=1.8, rong=34)
    to.text('KÊNH DẪN', (CX(-10.0), CY(75.3)), 2.0, 'CENTER')
    to.text('DỐC NƯỚC L = 60 m, i = 0,08', (CX(30.0), CY(z_day(30.0)) - 5.5), 2.2, 'CENTER', 'CHU_DAM', 'VN_DAM',
            rot=-math.degrees(math.atan(T['i'] * 8 / 4)))
    # bảng số liệu
    y0, hd = CY(66.0) - 2.0, 6.0
    hang = [('x (m)', lambda r: so(r[0], 0)), ('B (m)', lambda r: so(r[1], 2)),
            ('Cao độ đáy (m)', lambda r: so(z_day(r[0]), 2)),
            ('h / hhk P = 1% (m)', lambda r: '%s/%s' % (so(r[2], 2), so(r[4], 2))),
            ('h / hhk P = 0,2% (m)', lambda r: '%s/%s' % (so(r[5], 2), so(r[7], 2))),
            ('v P = 0,2% (m/s)', lambda r: so(r[6], 2)),
            ('Tường bên / vượt cao (m)', lambda r: '%s/%s' % (so(r[8], 2), so(r[9], 2)))]
    x_bang0 = 25.0
    for k, (ten, f) in enumerate(hang):
        yt, yb = y0 - k * hd, y0 - (k + 1) * hd
        to.line((x_bang0, yt), (CX(64.0), yt), 'BANG')
        to.text(ten, (x_bang0 + 1.0, 0.5 * (yt + yb)), 1.9, 'ML', 'CHU_DAM' if k == 0 else 'CHU',
                'VN_DAM' if k == 0 else 'VN')
        for r in D.TC_BANG43:
            to.text(f(r), (CX(r[0]), 0.5 * (yt + yb)), 1.55, 'MC')
    yb = y0 - len(hang) * hd
    to.line((x_bang0, yb), (CX(64.0), yb), 'BANG')
    for x in (x_bang0, CX(-3.0), CX(64.0)):
        to.line((x, y0), (x, yb), 'BANG')
    to.text('Hàm khí hhk = h·(1 + v/100); tường bên sau x = 40 m giả thiết 2,5 m; độ vượt cao tính với lũ kiểm tra.',
            (x_bang0, yb - 3.2), 1.8)


# ====================================================================== mặt bằng
PY0 = 280.0
PYY = lambda y: PY0 + y * 4.0                         # 1/250, y = tọa độ ngang dòng chảy


def mat_bang(to):
    to.ten_hinh('MẶT BẰNG TRÀN CHÍNH (sơ đồ)', 40, 355, '1/250', h=4.5, can='LEFT')
    B0 = T['B_dau'] / 2
    xs0 = -14.0
    xm, zm = mui_phun_hh()
    # kênh dẫn + vùng nạo vét
    p = [(CX(xs0), PYY(-B0)), (CX(-5.05), PYY(-B0)), (CX(-5.05), PYY(B0)), (CX(xs0), PYY(B0))]
    to.hatch(p, 'AR-SAND', 0.2)
    to.text('NẠO VÉT KÊNH DẪN', (CX(-9.5), PYY(2.0)), 2.0, 'MC')
    to.text('đến +74,00 (≈ 500 m³)', (CX(-9.5), PYY(-1.0)), 1.9, 'MC')
    # tường bên (2 nét, dày 0,5 m sơ đồ)
    xs = [-5.05, 0.0, T['L_chuyen'], T['L_doc'], xm]
    for sgn in (1, -1):
        trong = [(CX(x), PYY(sgn * b_doc(x) / 2)) for x in xs]
        ngoai = [(CX(x), PYY(sgn * (b_doc(x) / 2 + 0.5))) for x in xs]
        to.pline(trong, 'NET_CHINH')
        to.pline(ngoai, 'NET_CHINH')
        to.line((CX(xs0), PYY(sgn * B0)), (CX(-5.05), PYY(sgn * B0)), 'NET_MANH', linetype='DASHED', ltscale=0.3)
    # ngưỡng: 9 khoang, 8 trụ pin (sơ đồ)
    nb, bk = T['n_khoang'], T['b_khoang']
    tp = (T['B_dau'] - nb * bk) / (nb - 1)
    y = -B0
    for k in range(nb):
        y_end = y + bk
        if k < nb - 1:
            a, b = y_end, y_end + tp
            pp = [(CX(-5.05), PYY(a + 0.25 * tp)), (CX(-5.6), PYY(a + 0.5 * tp)), (CX(-5.05), PYY(b - 0.25 * tp)),
                  (CX(-5.05), PYY(b)), (CX(0.0), PYY(b)), (CX(0.0), PYY(a)), (CX(-5.05), PYY(a))]
            to.hatch(pp, 'AR-CONC', 0.03)
            to.pline(pp, 'NET_CHINH', dong=True)
        y = y_end + tp
    to.line((CX(-2.95), PYY(-B0)), (CX(-2.95), PYY(B0)), 'NET_MANH')
    to.line((CX(-2.45), PYY(-B0)), (CX(-2.45), PYY(B0)), 'NET_MANH')
    to.text('NGƯỠNG +74,70 – 9 khoang × 2,40 m', (CX(-5.05), PYY(-B0 - 2.2)), 2.0, 'LEFT')
    # vùng xử lý mặt đáy dốc
    vung = [(CX(x), PYY(b_doc(x) / 2)) for x in (0.0, T['L_chuyen'], T['L_doc'])] + \
           [(CX(x), PYY(-b_doc(x) / 2)) for x in (T['L_doc'], T['L_chuyen'], 0.0)]
    to.hatch(vung, 'ANSI31', 1.6, lop='SUA_CHUA', mau_to=1)
    to.text('XỬ LÝ MẶT ĐÁY DỐC: 1.210 m²', (CX(42.0), PYY(2.2)), 2.4, 'MC', 'CHU_DAM', 'VN_DAM')
    to.text('(vữa XM – polime dày 3 cm, chi tiết B)', (CX(42.0), PYY(-1.6)), 2.0, 'MC')
    # mũi phun
    to.line((CX(T['L_doc']), PYY(-b_doc(60) / 2)), (CX(T['L_doc']), PYY(b_doc(60) / 2)), 'NET_MANH')
    to.line((CX(xm), PYY(-b_doc(60) / 2 - 0.5)), (CX(xm), PYY(b_doc(60) / 2 + 0.5)), 'NET_CHINH')
    to.text('MŨI PHUN', (CX(xm) + 2.0, PYY(0.0)), 2.0, 'ML', rot=90)
    # kích thước bề rộng
    to.dim((CX(1.0), PYY(-B0)), (CX(1.0), PYY(B0)), 0, 250, goc=90)
    to.dim((CX(T['L_doc'] - 1.5), PYY(-T['B_cuoi'] / 2)), (CX(T['L_doc'] - 1.5), PYY(T['B_cuoi'] / 2)), 0, 250, goc=90)
    yk = PYY(-B0 - 4.0)
    to.dim((CX(0.0), yk), (CX(T['L_chuyen']), yk), 0, 250)
    to.dim((CX(T['L_chuyen']), yk), (CX(T['L_doc']), yk), 0, 250)
    # mặt cắt 1-1, hướng dòng chảy
    for sgn in (1, -1):
        yy = PYY(sgn * (b_doc(20) / 2 + 3.0))
        to.pline([(CX(20.0), yy), (CX(20.0), yy - sgn * 4.0)], 'NET_CHINH', const_width=0.4)
        to.text('1', (CX(20.0) + 1.5, yy + (0.5 if sgn > 0 else -3.0)), 3.0, 'LEFT', 'CHU_DAM', 'VN_DAM')
    mui_ten(to, (CX(4.0), PYY(-9.5)), (CX(14.0), PYY(-9.5)), 'NUOC', 2.5)
    to.text('Dòng chảy', (CX(9.0), PYY(-9.5) - 3.5), 1.9, 'CENTER', 'NUOC')


# ====================================================================== mặt cắt 1-1
def mat_cat_11(to):
    to.ten_hinh('MẶT CẮT 1-1 (x = 20 m)', 470, 577, '1/100', h=4.0, can='LEFT')
    B = T['B_cuoi']
    kn = KhungNhin(560.0, 535.0, 0.0, 0.0, 100)          # gốc: tim dốc, đáy dốc
    ht = 2.5
    # bản đáy + tường (sơ đồ: đáy dày 0,5 m, tường dày 0,5 m)
    p = [(-B / 2 - 0.5, ht), (-B / 2, ht), (-B / 2, 0.0), (B / 2, 0.0), (B / 2, ht), (B / 2 + 0.5, ht),
         (B / 2 + 0.5, -0.5), (-B / 2 - 0.5, -0.5)]
    pp = [kn(*q) for q in p]
    to.hatch(pp, 'AR-CONC', 0.03)
    to.pline(pp, 'NET_CHINH', dong=True)
    to.pline([kn(-B / 2, 0.03), kn(B / 2, 0.03)], 'SUA_CHUA', const_width=0.35)
    for h, ten, mau, kieu in ((1.305, 'P = 1%: hhk = 1,30', 5, 'CONTINUOUS'), (1.775, 'P = 0,2%: hhk = 1,77', 1, 'DASHED')):
        to.line(kn(-B / 2, h), kn(B / 2, h), 'NUOC', color=mau, linetype=kieu, ltscale=0.3)
        to.text(ten, (kn(-B / 2 + 0.4, 0)[0], kn(0, h)[1] + 0.8), 1.8, 'LEFT', 'NUOC')
    to.dim(kn(B / 2, 0.0), kn(B / 2, ht), 6.0, 100, goc=90)
    to.dim(kn(B / 2, 1.775), kn(B / 2, ht), -6.0, 100, goc=90, text='73')
    to.dim(kn(-B / 2, -0.5), kn(B / 2, -0.5), -6.0, 100)
    to.text('Độ vượt cao 0,73 m ≥ 0,5 m (lũ kiểm tra)', (kn(-B / 2 + 0.4, 0)[0], kn(0, 2.2)[1]), 1.8, 'LEFT')
    to.text('Vữa XM – polime 3 cm', (kn(-2.0, 0)[0], kn(0, -0.5)[1] - 7.0), 1.8, 'CENTER', 'SUA_CHUA')


# ====================================================================== mặt cắt ngưỡng
def nguong(to, x0, y0):
    kn = KhungNhin(x0, y0, -6.5, 72.0, 50)
    to.ten_hinh('MẶT CẮT NGƯỠNG TRÀN CHÍNH – SƠ ĐỒ KIỂM TRA ỔN ĐỊNH', x0 - 5, y0 + 115, '1/50', h=3.8, can='LEFT')
    p = [kn(x, z) for x, z in D.TC_NGUONG]
    to.hatch(p, 'AR-CONC', 0.02)
    to.pline(p, 'NET_CHINH', dong=True)
    # sân thượng lưu, đáy dốc (đoạn), nền đá
    to.line(kn(-6.5, 74.0), kn(-5.05, 74.0), 'NET_CHINH')
    to.line(kn(0.0, 73.3), kn(1.6, 73.3 - 0.08 * 1.6), 'NET_CHINH')
    to.line(kn(-6.5, 72.3), kn(1.6, 72.3), 'NET_MANH', linetype='DASHED', ltscale=0.3)
    nen = [kn(-6.5, 72.3), kn(1.6, 72.3), kn(1.6, 70.0), kn(-6.5, 70.0)]
    to.hatch(nen, 'ANSI33', 1.6)
    to.text('Nền đá lớp II', (kn(-6.2, 0)[0], kn(0, 70.4)[1]), 1.9, 'LEFT')
    # neo D25 (bổ sung khi hồ sơ hoàn công không có thép neo)
    for xn in (-4.3, -2.8, -1.3):
        to.line(kn(xn, 72.9), kn(xn, 70.3), 'SUA_CHUA', linetype='DASHED', ltscale=0.25)
    to.dim(kn(-4.3, 70.3), kn(-2.8, 70.3), -5.0, 50)
    to.dim(kn(-1.3, 72.3), kn(-1.3, 70.3), -12.0, 50, goc=90)
    to.chu_thich_mui_ten(kn(-1.3, 71.0), kn(1.8, 70.6), 'Neo D25, 2 hàng a 150, ngàm 200 vào đá II (hạng mục thiết kế); đầu thanh thấp hơn mặt '
                         'ngưỡng 6 cm, trám vữa XM – polime ≥ W8', h=1.8, rong=88)
    # mực nước (cột nước có ngắt)
    for Z, ten, mau in ((D.MN['MNLTK'], 'MNLTK +77,63', 5), (D.MN['MNLKT'], 'MNLKT +78,54', 1)):
        yv = kn(0, 75.55 if mau == 5 else 75.95)[1]
        to.line((kn(-6.5, 0)[0], yv), (kn(-3.6, 0)[0], yv), 'NUOC', color=mau)
        to.text(ten + ' (vẽ ngắt)', (kn(-6.4, 0)[0], yv + 0.7), 1.8, 'LEFT', 'NUOC')
    # sơ đồ lực
    def luc(p0, p1, ten, p_chu):
        mui_ten(to, kn(*p0), kn(*p1), 'CUNG_TRUOT', 1.8)
        to.text(ten, kn(*p_chu), 2.6, 'MC', 'CUNG_TRUOT', 'VN_DAM')
    luc((-1.6, 74.2), (-1.6, 73.2), 'G', (-1.2, 73.6))
    luc((-4.0, 75.0), (-4.0, 74.45), 'Wn', (-4.4, 74.85))
    luc((-6.3, 73.0), (-5.15, 73.0), 'H', (-5.75, 73.25))
    for xu, hu in ((-4.8, 1.25), (-3.0, 1.0), (-1.0, 0.75), (0.0, 0.62)):
        mui_ten(to, kn(xu, 72.3 - hu), kn(xu, 72.3), 'CUNG_TRUOT', 1.4)
    to.pline([kn(-5.05, 72.3 - 1.35), kn(0.0, 72.3 - 0.6)], 'CUNG_TRUOT')
    to.text('U (đẩy ngược)', (kn(0.3, 0)[0], kn(0, 71.85)[1]), 2.0, 'LEFT', 'CUNG_TRUOT')
    # kích thước
    to.dim(kn(-5.05, 72.3), kn(0.0, 72.3), -16.0, 50)
    to.dim(kn(-2.95, 74.7), kn(-2.45, 74.7), 5.0, 50)
    to.dim(kn(-5.05, 74.0), kn(-5.05, 74.7), 10.0, 50, goc=90)
    to.dim(kn(0.0, 73.3), kn(0.0, 74.7), -8.0, 50, goc=90)
    to.text('1:3', (kn(-4.2, 0)[0], kn(0, 74.55)[1]), 2.0, 'CENTER')
    to.text('1:1,75', (kn(-0.7, 0)[0], kn(0, 74.2)[1]), 2.0, 'LEFT')
    to.cao_trinh(*kn(-2.7, 74.7), 0, h=1.8, chu='+74,70')
    to.cao_trinh(*kn(-6.2, 74.0), 0, ben='L', h=1.8, chu='+74,00')
    to.cao_trinh(*kn(0.6, 73.25), 0, h=1.8, chu='+73,30')
    to.cao_trinh(*kn(1.0, 72.3), 0, h=1.8, chu='+72,30')


# ====================================================================== chi tiết B, C
def chi_tiet(to, x0, y0):
    """Chi tiết B: xử lý mặt đáy dốc (1/10); chi tiết C: khe co giãn và lỗ thoát nước (không tỷ lệ)."""
    to.ten_hinh('CHI TIẾT B – XỬ LÝ MẶT ĐÁY DỐC', x0, y0 + 60, '1/10', h=3.6, can='LEFT')
    k = 1000.0 / 10 / 100.0                              # mm giấy / cm thật
    L = 80.0
    # bê tông bản đáy cũ (dày thể hiện 25 cm), lớp đục bỏ 3 cm, vữa 3 cm
    yb = y0
    pts = [(x0, yb), (x0 + L, yb), (x0 + L, yb + 25 * k), (x0, yb + 25 * k)]
    to.hatch(pts, 'AR-CONC', 0.05)
    to.pline(pts, 'NET_CHINH', dong=True)
    to.line((x0, yb + 28 * k), (x0 + L, yb + 28 * k), 'NET_MANH', linetype='DASHED', ltscale=0.3)
    pv = [(x0, yb + 25 * k), (x0 + L, yb + 25 * k), (x0 + L, yb + 28 * k), (x0, yb + 28 * k)]
    to.hatch(pv, 'DOTS', 0.25, lop='SUA_CHUA', mau_to=1)
    to.pline(pv, 'SUA_CHUA', dong=True)
    to.dim((x0 + L, yb + 25 * k), (x0 + L, yb + 28 * k), 6.0, 10, goc=90)
    to.chu_thich_mui_ten((x0 + 20, yb + 28 * k), (x0 + 30, yb + 28 * k + 12),
                         'Mặt BT cũ trước khi đục (mài mòn, gồ ghề)', h=1.8, rong=62)
    to.chu_thich_mui_ten((x0 + 45, yb + 26.5 * k), (x0 + 52, yb + 28 * k + 6),
                         'Vữa XM – polime R ≥ 40 MPa, W8, dày 3 cm, gồ ghề ≤ 3 mm', h=1.8, rong=66)
    to.chu_thich_mui_ten((x0 + 60, yb + 25 * k), (x0 + 66, yb - 6),
                         'Mặt BT chắc sau khi đục 2 ÷ 3 cm, làm nhám, phun nước áp lực, quét lớp kết dính', h=1.8,
                         rong=100)
    to.text('Bản đáy BT hiện trạng', (x0 + 2.0, yb + 12 * k), 1.9, 'ML')
    # chi tiết C
    yc = y0 - 70
    to.ten_hinh('CHI TIẾT C – KHE CO GIÃN, LỖ THOÁT NƯỚC', x0, yc + 50, 'không tỷ lệ', h=3.6, can='LEFT')
    w, hd, g = 50.0, 22.0, 2.0
    for xa, xb in ((x0, x0 + w), (x0 + w + g, x0 + 2 * w + g)):
        p = [(xa, yc), (xb, yc), (xb, yc + hd), (xa, yc + hd)]
        to.hatch(p, 'AR-CONC', 0.05)
        to.pline(p, 'NET_CHINH', dong=True)
    xk = x0 + w
    to.rect(xk, yc + hd - 3.0, g, 3.0, 'SUA_CHUA')
    to.hatch([(xk, yc + hd - 3.0), (xk + g, yc + hd - 3.0), (xk + g, yc + hd), (xk, yc + hd)], 'SOLID', lop='SUA_CHUA',
             mau_to=1)
    to.circle((xk + g / 2, yc + hd - 4.2), 0.9, 'NET_CHINH')
    to.line((xk + g / 2, yc + 3.0), (xk + g / 2, yc + hd - 6.0), 'NET_CHINH', linetype='DASHED', ltscale=0.2)
    to.chu_thich_mui_ten((xk + g, yc + hd - 1.5), (xk + 14, yc + hd + 8), 'Ma tít đàn hồi (thay mới)', h=1.8, rong=44)
    to.chu_thich_mui_ten((xk + g / 2 + 0.9, yc + hd - 4.2), (xk + 30, yc + hd + 2.5), 'Thanh đệm xốp', h=1.8, rong=26)
    to.chu_thich_mui_ten((xk + g / 2, yc + 10.0), (xk + 30, yc + 8.0), 'Khớp nối hiện trạng (giữ nguyên)', h=1.8,
                         rong=50)
    # lỗ thoát nước dưới bản đáy
    xd = x0 + 18.0
    to.rect(xd - 1.2, yc, 2.4, hd, 'NUOC')
    to.hatch([(xd - 6, yc), (xd + 6, yc), (xd + 4, yc - 6), (xd - 4, yc - 6)], 'GRAVEL', 0.06)
    to.pline([(xd - 6, yc), (xd + 6, yc), (xd + 4, yc - 6), (xd - 4, yc - 6)], 'NET_MANH', dong=True)
    to.chu_thich_mui_ten((xd + 1.2, yc + 12), (xd + 6, yc + 30), 'Lỗ thoát nước: thông tắc, súc rửa', h=1.8, rong=48)
    to.text('Tầng lọc', (xd - 7.0, yc - 3.0), 1.8, 'MR')


# ====================================================================== bảng, ghi chú
def cac_bang(to):
    x0 = 700.0
    # mũi phun
    cot = [20, 16, 14, 18, 16, 22]
    hang = [['Tần suất', 'v (m/s)', 'h (m)', 'q (m³/s/m)', 'L (m)', 't Veronese (m)']] + \
           [[a, so(v, 2), so(h, 2), so(q, 2), so(L, 1), so(t, 1)] for a, v, h, q, L, t in D.TC_MUI_PHUN]
    y = to.bang(x0 - 10, 575, cot, hang, h=1.9, cao_dong=4.6, tieu_de='TIÊU NĂNG MŨI PHUN')
    to.text('t chỉ tham khảo (công thức cho nền rời);', (x0 - 10, y - 3.0), 1.7)
    to.text('quan trắc hố xói sau mỗi trận lũ lớn.', (x0 - 10, y - 5.6), 1.7)
    xL = 430.0
    cot1 = [40, 18, 16, 17, 19, 17, 19, 18]
    hang1 = [['Trường hợp', 'Qmax đến\n(m³/s)', 'Zmax\n(m)', 'H tràn\nchính (m)', 'q tràn\nchính (m³/s)',
              'H tràn\nsự cố (m)', 'q tràn\nsự cố (m³/s)', 'q tổng\n(m³/s)']] + [list(r) for r in D.TC_BANG41]
    y = to.bang(xL, 503, cot1, hang1, h=1.9, cao_dong=4.6, tieu_de='TRỊ SỐ THIẾT KẾ CÔNG TRÌNH THÁO LŨ (Bảng 4.1)')
    cot2 = [16, 16, 20, 22]
    hang2 = [['H0 (m)', 'ε', 'Q (m³/s)', 'q (m³/s/m)']] + [[a, b, c, d] for a, b, c, d in D.TC_BANG42]
    to.bang(xL + 175, 503, cot2, hang2, h=1.9, cao_dong=4.6, tieu_de='KHẢ NĂNG THÁO (Bảng 4.3)')
    to.text('m = 0,45; n = 9; b = 2,4 m', (xL + 175, 503 - 8 * 4.6 - 3.0), 1.7)
    cot4 = [27, 10, 15, 15, 15, 15, 15, 14, 14, 13, 22]
    hang4 = [['Mực nước', 'c\n(kPa)', 'G\n(kN)', 'W nước\n(kN)', 'U\n(kN)', 'ΣV\n(kN)', 'ΣH\n(kN)', 'K trượt',
              '[K]', 'K lật', 'σ gót / mũi\n(kPa)']] + [list(r) for r in D.TC_BANG44]
    y = to.bang(xL, 445, cot4, hang4, h=1.9, cao_dong=4.6,
                tieu_de='ỔN ĐỊNH NGƯỠNG TRÀN CHÍNH – DẢI RỘNG 1 m (Bảng 4.5)')
    to.text('K = (f·ΣV + c·B)/ΣH; f = 0,60; B = 5,05 m; [K] = γn·γlc/γc = 1,15·1,0/0,95 = 1,21 (cơ bản), 1,15·0,9/0,95 = 1,09 '
            '(đặc biệt) – QCVN 04-05:2022 CT (1), Bảng B.2. c = 0: K < [K] khi tháo lũ → cần c ≥ 7,6 kPa (MNLTK), 10,4 kPa (MNLKT).', (xL, y - 3.2), 1.8)
    to.text('Neo D25 thụ động (2 hàng, a = 1,5 m, ngàm 2,0 m vào đá lớp II), c = 0: lực kéo cần 48 / 65 kN/thanh ≤ 38% '
            'Ntt = 171,8 kN; K = 2,23 / 1,81 khi mọi neo đạt Ntt (Bảng 4.6).', (xL, y - 6.4), 1.8)
    cot5 = [18, 18, 30, 22, 18, 30]
    hang5 = [['Tần suất', 'Q (m³/s)', 'h đầu dốc (m)', 'h cuối dốc (m)', 'v cuối (m/s)', 'Tường yêu cầu (m)']] + \
            [list(r) for r in D.TSC_TL]
    y = to.bang(xL, 385, cot5, hang5, h=1.9, cao_dong=4.6, tieu_de='THỦY LỰC TRÀN SỰ CỐ (b = 5 m, dốc 32 m, i = 0,08)')
    to.text('Đầu dốc h = hk (P = 0, mặt cắt khống chế); tường yêu cầu = hhk + 0,5 m – đối chiếu chiều cao tường thực tế.',
            (xL, y - 3.2), 1.7)
    cot6 = [100, 16, 22]
    hang6 = [['Hạng mục (Bảng 7.1 – II)', 'Đơn vị', 'Khối lượng'],
             ['Xử lý bề mặt đáy dốc tràn chính bằng vữa XM – polime dày 3 cm', 'm²', '1.210'],
             ['Nạo vét kênh dẫn thượng lưu tràn chính (ước tính)', 'm³', '≈ 500'],
             ['Bóc bỏ, lát lại tấm BTCT mái tràn sự cố', 'm²', '1.474'],
             ['Thay vật liệu chèn khe co giãn (ma tít đàn hồi)', 'khe', 'theo HT']]
    to.bang(xL + 165, 385, cot6, hang6, h=1.9, cao_dong=4.6, tieu_de='KHỐI LƯỢNG SỬA CHỮA TRÀN')


def ghi_chu(to, x, y):
    gc = ('GHI CHÚ: \\P'
          '1. Cao trình ghi bằng m, kích thước ghi bằng cm. Hình học tràn theo hồ sơ hiện trạng: 9 khoang × 2,4 m, ngưỡng '
          'thực dụng +74,70 chảy tự do, dốc nước dài 60 m, i = 0,08, thu hẹp 31 → 18 m trên 20 m đầu, tiêu năng mũi phun. '
          'Bản đáy, tường, mũi phun vẽ sơ đồ; kích thước ghi "giả thiết" phải đối chiếu hồ sơ hoàn công. \\P'
          '2. Kết cấu ngưỡng, trụ pin, tường bên, cầu giao thông còn tốt – giữ nguyên. Độ vượt cao tường bên nhỏ nhất '
          '0,73 m (x = 20 m, lũ kiểm tra) ≥ 0,5 m: không nâng tường. \\P'
          '3. Đáy dốc: đục bỏ lớp BT bề mặt mài mòn 2 ÷ 3 cm đến lớp BT chắc, làm sạch bằng phun nước áp lực, quét lớp '
          'kết dính, trát vữa XM – polime R ≥ 40 MPa, W8 (TCVN 4116:2023 mục 5.1.9) dày 3 cm, hoàn thiện phẳng (gồ ghề cục bộ ≤ 3 mm) để hạn chế xâm '
          'thực ở lưu tốc 9 ÷ 12 m/s. Thi công mùa khô, theo từng dải, bảo dưỡng ẩm theo chỉ dẫn nhà sản xuất. \\P'
          '4. Khe co giãn: cạo bỏ vật liệu chèn cũ, chèn lại ma tít đàn hồi; kiểm tra, thông các lỗ thoát nước dưới bản '
          'đáy. \\P'
          '5. Nạo vét kênh dẫn thượng lưu đến cao trình sân +74,00 (thấp hơn ngưỡng 0,7 m) để bảo đảm điều kiện chảy '
          'tới ngưỡng như tính toán. \\P'
          '6. Ổn định ngưỡng: khi bỏ qua lực dính, K trượt khi tháo lũ < [K]; ngưỡng đạt khi lực dính mặt tiếp xúc BT – đá '
          'c ≥ 10,4 kPa – chưa kiểm chứng, nên khoan neo là hạng mục thiết kế: thép CB400-V D25, 2 hàng, a = 150 cm, khoan '
          'D76 qua khối ngưỡng, ngàm 200 cm vào đá lớp II, bơm vữa xi măng không co ngót (≈ 36 thanh, 144 m khoan, 0,6 t); '
          'neo thụ động (ma sát – cắt): với c = 0 lực kéo cần 48 ÷ 65 kN/thanh ≤ 38% Ntt = min(thép 171,8; bám dính vữa – '
          'đá 191; thép – vữa 314) = 171,8 kN; K = 2,23 (MNLTK) / 1,81 (MNLKT) khi mọi neo đạt Ntt (Bảng 4.6). Thí '
          'nghiệm kéo nhổ ≥ 3 neo đến 100 kN (≈ 1,5 lần lực cần lớn nhất) trước thi công đại trà; kiểm tra ổn định chỉ dùng '
          'lực neo không vượt lực đã thử. Chống ăn mòn: thanh có định tâm, vữa bao ≥ 25 mm; đầu thanh thấp hơn mặt ngưỡng '
          '60 mm, trám vữa XM – polime ≥ W8 (TCVN 4116:2023 mục 6.3.3, 7.4). Số lượng điều chỉnh theo hồ sơ hoàn công '
          '(nếu đã có thép neo, chân khay ngàm đá). \\P'
          '7. Mũi phun và hố xói: không sửa chữa; quan trắc địa hình hố xói sau mỗi trận lũ lớn.')
    to.mtext(gc, (x, y), 2.0, rong=196)


def ve(thu_muc):
    to = To(6, 'TRÀN XẢ LŨ CHÍNH – SỬA CHỮA DỐC NƯỚC\nKIỂM TRA THỦY LỰC VÀ ỔN ĐỊNH NGƯỠNG', 'Như hình vẽ', 'HCG-BV06')
    a = to.dat(cat_doc, f=1.25, neo='TL', xy=(30.0, 818.0))                       # 1/200, 1/100
    b = to.dat(mat_bang, f=1.25, neo='TL', xy=(30.0, a[1] - 22.0))               # 1/200
    c = to.dat(nguong, 50.0, 70.0, f=2.0, k_chu=0.625, neo='TL', xy=(30.0, b[1] - 22.0))     # 1/50 -> 1/25
    xr = max(a[2], b[2], c[2]) + 30.0
    m = to.dat(mat_cat_11, f=2.0, k_chu=0.625, neo='TL', xy=(xr, 818.0))         # 1/100 -> 1/50
    t = to.dat(cac_bang, f=1.25, neo='TL', xy=(xr, m[1] - 22.0))
    g = to.dat(ghi_chu, 432.0, 340.0, f=1.25, neo='TL', xy=(xr, t[1] - 20.0))
    to.dat(chi_tiet, 250.0, 128.0, f=1.0, k_chu=1.25, neo='TL', xy=(xr, g[1] - 22.0))     # 1/10
    to.luu(thu_muc + '/BV06_TranChinh.dxf', thu_muc + '/BV06_TranChinh.pdf', thu_muc + '/BV06_TranChinh.png')
    return to


if __name__ == '__main__':
    import sys
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
