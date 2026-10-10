# -*- coding: utf-8 -*-
"""BV-04: CHI TIẾT SỬA CHỮA MẶT ĐẬP VÀ HỆ THỐNG RÃNH THOÁT NƯỚC MÁI HẠ LƯU."""
import math

import du_lieu_ban_ve as D
from bv_lib import To, KhungNhin, so

BT, DAT, CP, DA = ('AR-CONC', 0.03), ('AR-SAND', 0.12), ('GRAVEL', 0.15), ('AR-RROOF', 0.08)


def vung(to, kn, pts, mau=None, lop='NET_CHINH'):
    p = [kn(*q) for q in pts]
    if mau:
        to.hatch(p, mau[0], mau[1])
    to.pline(p, lop, dong=True)
    return p


def chi_tiet_A(to, x0, y0):
    """Chi tiết A – mặt đập sau sửa chữa, TL 1:25 (gốc x = 0 tại mặt trong tường chắn sóng)."""
    TL = 25
    kn = KhungNhin(x0, y0, -1.2, 77.0, TL)
    zc0, zboc, zdap = D.DINH_HT, 78.20, 78.60
    i = 0.02
    xs = 4.30                       # mép hạ lưu tấm BT
    z_top = lambda x: D.DINH_SC - i * x
    # nền đập cũ (đất đắp) – vẽ tới +77,0
    nen = [(-1.2, 77.0), (6.6, 77.0), (6.6, 78.35 - (6.6 - 4.7) / 3.0), (4.7, zboc), (0.0, zboc), (0.0, 77.6),
           (-0.3, 77.6), (-0.3, 77.0)]
    to.hatch([kn(*p) for p in nen], DAT[0], DAT[1])
    to.pline([kn(-1.2, 77.0), kn(6.6, 77.0)], 'NET_MANH', linetype='DASHED')
    # tường chắn sóng hiện trạng
    vung(to, kn, [(-0.3, 77.6), (0.0, 77.6), (0.0, D.TUONG_CS), (-0.3, D.TUONG_CS)], ('ANSI31', 0.5))
    # mặt đập hiện trạng (nét đứt) và đường bóc
    to.pline([kn(0, zc0), kn(4.7, zc0)], 'NET_KHUAT', linetype='DASHED', ltscale=0.4)
    # đất đắp bù K >= 0,97
    dap = [(0.0, zboc), (4.7, zboc), (4.7, zdap - i * 4.7), (0.0, zdap)]
    vung(to, kn, dap, ('ANSI31', 0.9))
    # cấp phối đá dăm đệm 5 cm + BT M250 20 cm
    vung(to, kn, [(0.0, zdap), (xs, zdap - i * xs), (xs, zdap + 0.05 - i * xs), (0.0, zdap + 0.05)], CP, 'NET_MANH')
    vung(to, kn, [(0.0, zdap + 0.05), (xs, zdap + 0.05 - i * xs), (xs, z_top(xs)), (0.0, z_top(0))], BT, 'SUA_CHUA')
    to.line(kn(0.0, zdap + 0.05), kn(xs, zdap + 0.05 - i * xs), 'NET_MANH', linetype='DASHED', ltscale=0.2)
    # gờ chắn bánh 20 × 25 cm (đặt trên tấm BT kéo dài)
    zg = zdap - i * 4.5
    vung(to, kn, [(xs, zg), (4.50, zg), (4.50, z_top(xs) + 0.25), (xs, z_top(xs) + 0.25)], BT, 'SUA_CHUA')
    # tấm lát BT M200 dày 10 cm trên mái m = 3 tới rãnh 20 × 20
    xa0, za0 = 4.50, zg
    xa1 = xa0 + 0.75
    za1 = za0 - 0.75 / 3.0
    vung(to, kn, [(xa0, za0), (xa1, za1), (xa1, za1 - 0.10), (xa0, za0 - 0.10)], BT, 'SUA_CHUA')
    # rãnh 20 × 20, thành 10 cm, đáy 10 cm
    xr0 = xa1
    zr = za1 - 0.10
    ranh = [(xr0, zr), (xr0, za1 - 0.10 - 0.30), (xr0 + 0.40, za1 - 0.40), (xr0 + 0.40, za1 - 0.05),
            (xr0 + 0.30, za1 - 0.05), (xr0 + 0.30, za1 - 0.30), (xr0 + 0.10, za1 - 0.30), (xr0 + 0.10, zr)]
    vung(to, kn, ranh, BT, 'SUA_CHUA')
    # mái đất + cỏ phía sau rãnh
    to.pline([kn(xr0 + 0.40, za1 - 0.05), kn(6.6, za1 - 0.05 - (6.6 - xr0 - 0.4) / 3.0)], 'NET_CHINH')
    for k in range(8):
        xx = xr0 + 0.5 + k * 0.18
        zz = za1 - 0.05 - (xx - xr0 - 0.4) / 3.0
        to.line(kn(xx, zz), kn(xx + 0.04, zz + 0.09), 'NET_MANH')
    # cao trình
    to.cao_trinh(*kn(-0.15, D.TUONG_CS), D.TUONG_CS, ben='L')
    to.cao_trinh(*kn(0.05, z_top(0.05)), D.DINH_SC, ben='R', chu='%s (mép thượng lưu)' % so(D.DINH_SC, 2, True))
    to.cao_trinh(*kn(xs - 0.05, z_top(xs)), z_top(xs), ben='L', chu='%s (mép tấm)' % so(z_top(xs), 2, True))
    to.cao_trinh(*kn(2.2, zc0), zc0, ben='R', chu='+78,35 (mặt đập hiện trạng)')
    to.cao_trinh(*kn(1.0, zboc), zboc, ben='R', chu='+78,20 (sau khi bóc)')
    to.text('i = 2%', kn(1.9, z_top(1.9) + 0.09), 2.8, 'CENTER', rot=-math.degrees(math.atan(i * 1.0)))
    # kích thước (cm)
    to.dim(kn(-0.3, D.TUONG_CS), kn(4.7, D.TUONG_CS), 12, TL)
    to.dim(kn(0.0, 77.0), kn(xs, 77.0), -8, TL)
    to.dim(kn(xs, 77.0), kn(4.5, 77.0), -8, TL)
    to.dim(kn(4.5, 77.0), kn(xa1, 77.0), -8, TL)
    to.dim(kn(xr0, 77.0), kn(xr0 + 0.40, 77.0), -8, TL)
    to.dim(kn(4.5, z_top(xs)), kn(4.5, z_top(xs) + 0.25), -7, TL, goc=90)
    # chú thích lớp
    cts = [((2.8, z_top(2.8) - 0.1), (3.0, 80.9), 'Bê tông M250 đá 1×2 dày 20 cm, khe co giãn 5 m/khe chèn ma tít'),
           ((2.2, zdap + 0.03 - i * 2.2), (2.4, 80.5), 'Nilon tái sinh + cấp phối đá dăm đệm dày 5 cm'),
           ((1.6, 78.40), (1.8, 80.1), 'Đất đắp bù (bãi VL lớp 1) đầm K ≥ 0,97 đến +78,60'),
           ((-0.15, 79.1), (0.05, 81.3), 'Tường chắn sóng hiện trạng – giữ nguyên'),
           ((4.4, z_top(xs) + 0.2), (5.9, 80.3), 'Gờ chắn bánh BT M250 20×25 cm, đoạn 1,0 m cách 0,30 m'),
           ((xr0 + 0.2, za1 - 0.2), (7.0, 79.2), 'Rãnh ngang đỉnh mái 20×20 cm, BT M200'),
           ((5.0, za0 - 0.18), (7.0, 78.6), 'Tấm lát BT M200 dày 10 cm, lót nilon'),
           ((3.9, 78.27), (7.0, 77.75), 'Bóc lớp đá dăm láng nhựa hư hỏng (~15 cm)')]
    for p, q, s in cts:
        to.chu_thich_mui_ten(kn(*p), kn(*q), s, 2.5, rong=len(s) * 1.5)
    to.ten_hinh('CHI TIẾT A – MẶT ĐẬP SAU SỬA CHỮA', kn(2.6, 0)[0], kn(0, 76.25)[1], '1/25')


def mat_bang_go(to, x0, y0):
    """Mặt bằng gờ chắn bánh, TL 1:50."""
    TL = 50
    kn = KhungNhin(x0, y0, 0, 0, TL)
    to.pline([kn(0, 0), kn(6.6, 0)], 'NET_MANH')
    to.pline([kn(0, 0.2), kn(6.6, 0.2)], 'NET_MANH')
    x = 0.15
    while x + 1.0 <= 6.6:
        vung(to, kn, [(x, 0), (x + 1.0, 0), (x + 1.0, 0.2), (x, 0.2)], BT, 'SUA_CHUA')
        x += 1.3
    to.dim(kn(0.15, 0.2), kn(1.15, 0.2), 6, TL)
    to.dim(kn(1.15, 0.2), kn(1.45, 0.2), 6, TL)
    to.dim(kn(1.45, 0.2), kn(2.45, 0.2), 6, TL)
    to.dim(kn(6.6, 0.0), kn(6.6, 0.2), 6, TL, goc=90)
    to.text('Khe hở 30 cm để nước mặt đập chảy vào rãnh ngang đỉnh mái', kn(0.15, -0.35), 2.5, 'LEFT')
    to.ten_hinh('MẶT BẰNG GỜ CHẮN BÁNH', kn(3.3, 0)[0], kn(0, 1.35)[1], '1/50', h=4.0)


def mat_cat_ranh(to, x0, y0, ten, b, h, t, tl=20, da_hoc=False, mai=0.0):
    """Mặt cắt rãnh chữ nhật (hoặc hình thang đá hộc xây). b, h: lòng rãnh (m); t: chiều dày."""
    kn = KhungNhin(x0, y0, 0, 0, tl)
    if not da_hoc:
        ngoai = [(-t, -t), (b + t, -t), (b + t, h), (b, h), (b, 0), (0, 0), (0, h), (-t, h)]
        vung(to, kn, ngoai, BT, 'SUA_CHUA')
        to.pline([kn(-t - 0.25, h), kn(-t, h)], 'NET_CHINH')
        to.pline([kn(b + t, h), kn(b + t + 0.25, h)], 'NET_CHINH')
        to.line(kn(-t - 0.05, -t - 0.02), kn(b + t + 0.05, -t - 0.02), 'NET_MANH', linetype='DASHED', ltscale=0.15)
        to.dim(kn(0, h), kn(b, h), 5, tl)
        to.dim(kn(-t, -t), kn(b + t, -t), -6, tl)
        to.dim(kn(b + t, 0), kn(b + t, h), 6, tl, goc=90)
        to.dim(kn(b + t, -t), kn(b + t, 0), 6, tl, goc=90)
        ghi = 'BT M200 đá 1×2, lót nilon'
    else:
        m = mai
        tr = [(-m * h, h), (0, 0), (b, 0), (b + m * h, h)]
        ngoai = [(-m * h - t * 1.41, h), (-t * 0.41, -t), (b + t * 0.41, -t), (b + m * h + t * 1.41, h)]
        vung(to, kn, ngoai + list(reversed(tr)), DA, 'NET_CHINH')
        to.dim(kn(0, 0), kn(b, 0), -8, tl)
        to.dim(kn(-m * h, h), kn(b + m * h, h), 5, tl)
        to.dim(kn(b + m * h + t * 1.41, 0), kn(b + m * h + t * 1.41, h), 6, tl, goc=90)
        ghi = 'Đá hộc xây vữa M100 dày 25 cm'
    to.text(ghi, kn(b / 2, -t - 0.35), 2.5, 'TC')
    to.ten_hinh(ten, kn(b / 2, 0)[0], kn(0, h)[1] + 20, '1/%d' % tl, h=3.6)


def mat_bang_khai_trien(to, x0, y0):
    """Mặt bằng khai triển mái hạ lưu – khoang điển hình 40 m (chiều dài thật theo mái), TL 1:500."""
    TL = 500
    kn = KhungNhin(x0, y0, 0, 0, TL)
    h = D.mat_cat(D.DINH_SC)
    L1 = (h['Zc'] - h['Zcd']) * math.sqrt(1 + 3.0 ** 2)          # mái trên m = 3
    L2 = (h['Zcd'] - h['Z_ltr']) * math.sqrt(1 + 3.5 ** 2)       # mái dưới m = 3,5
    s = [0.0, L1, L1 + 14.0, L1 + 14.0 + L2, L1 + 14.0 + L2 + 3.0]  # mốc dọc mái (từ mép đỉnh)
    B = 40.0
    # biên khoang và các đới
    for a, b, ten in ((s[0], s[1], 'MÁI m = 3,0'), (s[1], s[2], 'CƠ +68,00 (rộng 14 m)'),
                      (s[2], s[3], 'MÁI m = 3,5'), (s[3], s[4], 'ĐỈNH LĂNG TRỤ +59,00')):
        to.rect(*kn(0, -b), B * 1000 / TL, (b - a) * 1000 / TL, 'NET_MANH')
        to.text(ten, kn(B + 10.5, -(a + b) / 2), 2.4, 'ML')
    # rãnh xiên 20×20 dạng mắt cáo ô 4×4 m (xoay 45°) trên hai đới mái
    d = 4.0 * math.sqrt(2)
    for a, b in ((s[0] + 0.6, s[1] - 0.6), (s[2] + 0.6, s[3] - 0.6)):
        for sign in (1, -1):
            k = -40
            while k < 40:
                c = k * d
                pts = []
                # đường y = sign·x + c cắt hình chữ nhật [0,B]×[-b,-a]
                for xx in (0.4, B - 0.4):
                    yy = -(a) - (sign * xx + c)
                    pts.append((xx, yy))
                (x1, y1), (x2, y2) = pts
                # cắt theo đới
                seg = []
                for t in [i / 200 for i in range(201)]:
                    xx, yy = x1 + (x2 - x1) * t, y1 + (y2 - y1) * t
                    if -b <= yy <= -a:
                        seg.append((xx, yy))
                if len(seg) > 1:
                    to.line(kn(*seg[0]), kn(*seg[-1]), 'SUA_CHUA')
                k += 1
    # rãnh dọc 40×40 ở hai biên khoang, rãnh ngang đỉnh, cơ, đỉnh lăng trụ, chân đập
    for xx in (0.0, B):
        to.pline([kn(xx - 0.2, 0), kn(xx - 0.2, -s[4]), kn(xx + 0.2, -s[4]), kn(xx + 0.2, 0)], 'SUA_CHUA', dong=True)
    for yy, w in ((0.0, 0.2), (s[1] + 1.0, 0.4), (s[3] + 1.0, 0.4)):
        to.rect(*kn(0, -yy - w), B * 1000 / TL, w * 1000 / TL, 'SUA_CHUA')
    to.rect(*kn(0, -s[4] - 1.6), B * 1000 / TL, 0.5 * 1000 / TL, 'NET_CHINH')
    to.text('Rãnh chân đập 50×50 đá hộc xây (dẫn nước ra hai vai)', kn(B / 2, -s[4] - 3.4), 2.4, 'TC')
    # nhãn
    to.chu_thich_mui_ten(kn(0, -L1 * 0.55), kn(-6, -L1 * 0.45), 'Rãnh dọc 40×40 @ 40 m', 2.4, rong=34)
    to.chu_thich_mui_ten(kn(8.5, -10.3), kn(-6, -6), 'Rãnh xiên 20×20, ô 4×4 m', 2.4, rong=34)
    to.chu_thich_mui_ten(kn(20, -0.1), kn(14, 3.5), 'Rãnh ngang đỉnh mái 20×20', 2.4, rong=40)
    to.chu_thich_mui_ten(kn(12, -s[1] - 1.2), kn(-6, -s[1] - 4.5), 'Rãnh ngang trên cơ 40×40', 2.4, rong=34)
    to.text('Trong ô: đất màu 10 cm + trồng cỏ', kn(20, -L1 * 0.3), 2.4, 'MC')
    to.dim(kn(0, 0), kn(B, 0), 9, TL)
    for a, b in zip(s[:-1], s[1:]):
        to.dim(kn(B, -a), kn(B, -b), 14, TL, goc=90)
    to.text('Chiều dài ghi theo mặt mái (khai triển)', kn(B + 10.5, 3.0), 2.4, 'LEFT')
    to.ten_hinh('MẶT BẰNG KHAI TRIỂN HỆ RÃNH MÁI HẠ LƯU – KHOANG 40 m', kn(B / 2, 0)[0], kn(0, 12.0)[1], '1/500',
                h=4.0)


def xu_ly_ho_xoi(to, x0, y0):
    """Chi tiết xử lý hố xói mái hạ lưu, TL 1:50."""
    TL = 50
    kn = KhungNhin(x0, y0, 0, 0, TL)
    m = 3.0
    mai = lambda x: -x / m
    # mái thiết kế
    to.pline([kn(-1.0, mai(-1.0)), kn(7.0, mai(7.0))], 'NET_CHINH')
    # hố xói hiện trạng (nét đứt)
    hx = [(1.0, mai(1.0)), (1.6, mai(1.6) - 0.9), (2.8, mai(2.8) - 1.25), (4.0, mai(4.0) - 0.8), (4.7, mai(4.7))]
    to.pline([kn(*p) for p in hx], 'NET_KHUAT', linetype='DASHED', ltscale=0.3)
    # phạm vi đào mở rộng: mái m = 1, sâu thêm 0,3 m
    zday = mai(2.8) - 1.25 - 0.30
    dao = [(0.2, mai(0.2)), (0.2 + (mai(0.2) - zday), zday), (5.3 - (mai(5.3) - zday), zday), (5.3, mai(5.3))]
    p = [kn(*q) for q in dao]
    to.hatch(p + [kn(5.3, mai(5.3)), kn(0.2, mai(0.2))], 'ANSI31', 0.9)
    to.pline(p, 'SUA_CHUA')
    # các lớp đắp 25 cm
    zz = zday + 0.25
    while zz < mai(0.2) - 0.05:
        xa = max(0.2 + (mai(0.2) - zz), -m * zz if zz < 0 else 0)
        xb = min(5.3 - (mai(5.3) - zz), -m * zz)
        if xb > xa:
            to.line(kn(xa, zz), kn(xb, zz), 'NET_MANH', linetype='DASHED', ltscale=0.15)
        zz += 0.25
    to.chu_thich_mui_ten(kn(2.8, mai(2.8) - 1.0), kn(5.5, mai(2.8) - 1.9),
                         'Đắp đất cùng loại từng lớp 25 cm, đầm cóc K ≥ 0,95', 2.4, rong=70)
    to.chu_thich_mui_ten(kn(2.9, mai(2.9) - 1.25), kn(-0.5, mai(2.9) - 2.4), 'Đáy hố xói hiện trạng', 2.4, rong=34)
    to.text('m = 1,0', kn(0.2 + 0.4 * (mai(0.2) - zday), mai(0.2) - 0.55), 2.4, 'RIGHT')
    to.dim(kn(2.8, mai(2.8) - 1.25), kn(2.8, zday), -9, TL, goc=90)
    to.ten_hinh('CHI TIẾT XỬ LÝ HỐ XÓI MÁI HẠ LƯU', kn(3.0, 0)[0], kn(0, 1.3)[1], '1/50', h=4.0)


def cac_ranh(to):
    """Bốn mặt cắt rãnh vẽ ở 1/25, đặt lên A0 phóng 1,25 → 1/20."""
    mat_cat_ranh(to, 495.0, 420.0, 'RÃNH DỌC 40×40', 0.40, 0.40, 0.10, tl=25)
    mat_cat_ranh(to, 560.0, 420.0, 'RÃNH XIÊN 20×20', 0.20, 0.20, 0.08, tl=25)
    mat_cat_ranh(to, 620.0, 420.0, 'RÃNH CƠ, ĐỈNH LĂNG TRỤ 40×40', 0.40, 0.40, 0.10, tl=25)
    mat_cat_ranh(to, 715.0, 420.0, 'RÃNH CHÂN ĐẬP 50×50', 0.50, 0.50, 0.25, tl=25, da_hoc=True, mai=1.0)


def bang_khoi_luong(to):
    cot = [10, 96, 14, 22]
    hang = [['TT', 'Hạng mục', 'Đơn vị', 'Khối lượng'],
            ['1', 'Cào bóc mặt đập đá dăm láng nhựa (dày 15 cm)', 'm³', '740'],
            ['2', 'Đất đắp bù mặt đập K ≥ 0,97', 'm³', '1.972'],
            ['3', 'Nilon tái sinh lót móng', 'm²', '4.930'],
            ['4', 'Bê tông M250 mặt đập dày 20 cm', 'm³', '986'],
            ['5', 'Bê tông M250 gờ chắn bánh', 'm³', '61'],
            ['6', 'Cột đèn chiếu sáng cao 9 m, 150 W', 'cột', '30'],
            ['7', 'Phát quang, bóc lớp mặt mái HL dày 10 cm', 'm³', '3.952'],
            ['8', 'Rãnh dọc BT M200 40×40', 'm', '1.212'],
            ['9', 'Rãnh xiên BT M200 20×20 (ô 4×4 m)', 'm', '19.759'],
            ['10', 'Rãnh ngang trên cơ, trên lăng trụ 40×40', 'm', '1.972'],
            ['11', 'Rãnh ngang đỉnh mái 20×20', 'm', '986'],
            ['12', 'Rãnh chân đập đá hộc xây VM100 50×50', 'm', '1.106'],
            ['13', 'Trồng cỏ mái hạ lưu (kể đất màu 10 cm)', 'm²', '36.357']]
    to.bang(480, 345, cot, hang, h=2.4, cao_dong=5.5, tieu_de='KHỐI LƯỢNG CHỦ YẾU – MẶT ĐẬP, MÁI HẠ LƯU (Bảng 7.1)')


def ghi_chu(to):
    gc = ('GHI CHÚ:\\P'
          '1. Cao trình ghi bằng m, kích thước ghi bằng cm.\\P'
          '2. Phương án Đ3 (Chương 3): đắp bù đến +78,60, mặt BT đỉnh +78,85 (mép thượng lưu), dốc ngang 2% về hạ lưu;\\P'
          '    cấu tạo lớp theo TKCS (đá dăm đệm 5 cm + nilon + BT M250 20 cm).\\P'
          '3. Đắp đất từng lớp ≤ 25 cm, độ ẩm sai lệch ±2% so với độ ẩm tốt nhất; kiểm tra độ chặt từng lớp.\\P'
          '4. Rãnh dọc chạy theo mái có lưu tốc lớn (≈ 9 m/s khi đầy): BT liền khối, có gờ tiêu năng tại chỗ nối với cơ\\P'
          '    và lăng trụ. Khả năng thoát của các rãnh lớn hơn nhiều lần lưu lượng thiết kế (Bảng 5.6).\\P'
          '5. Ô 4×4 m giảm ứng suất tiếp dòng chảy tràn trên mái từ ≈ 6,6 Pa xuống ≈ 1,8 Pa (Bảng 5.5).')
    to.mtext(gc, (480, 255), 2.4, rong=170)


def ve(thu_muc):
    to = To(4, 'CHI TIẾT SỬA CHỮA MẶT ĐẬP\nVÀ HỆ THỐNG RÃNH THOÁT NƯỚC MÁI HẠ LƯU', 'Như hình vẽ', 'HCG-BV04')
    a = to.dat(chi_tiet_A, 40.0, 395.0, f=1.25, neo='TL', xy=(30.0, 816.0))          # 1/25 -> 1/20
    to.dat(mat_bang_khai_trien, 100.0, 262.0, f=2.0, k_chu=0.625, neo='TL', xy=(30.0, a[1] - 30.0))  # 1/500 -> 1/250
    xr = a[2] + 40.0
    g = to.dat(mat_bang_go, 520.0, 525.0, f=2.0, k_chu=0.625, neo='TL', xy=(xr, 816.0))              # 1/50 -> 1/25
    r = to.dat(cac_ranh, f=1.25, neo='TL', xy=(xr, g[1] - 22.0))
    h = to.dat(xu_ly_ho_xoi, 300.0, 200.0, f=2.0, k_chu=0.625, neo='TL', xy=(xr, r[1] - 28.0))       # 1/50 -> 1/25
    b = to.dat(bang_khoi_luong, f=1.4, neo='TL', xy=(xr, h[1] - 25.0))
    to.dat(ghi_chu, f=1.4, neo='TL', xy=(b[2] + 22.0, h[1] - 25.0))
    to.luu(thu_muc + '/BV04_ChiTietMatDap_Ranh.dxf', thu_muc + '/BV04_ChiTietMatDap_Ranh.pdf',
           thu_muc + '/BV04_ChiTietMatDap_Ranh.png')
    return to


if __name__ == '__main__':
    import sys
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
