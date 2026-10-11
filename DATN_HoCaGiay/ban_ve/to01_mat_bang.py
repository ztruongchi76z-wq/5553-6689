# -*- coding: utf-8 -*-
"""BV-01: MẶT BẰNG TỔNG THỂ CỤM ĐẦU MỐI – BỐ TRÍ CÁC HẠNG MỤC SỬA CHỮA, NÂNG CẤP.

Nền: bình đồ TKCS (DWG khoan phụt mang tràn sự cố, hệ VN-2000), trích các lớp mái, tường, chân mái, đường, suối,
tim tuyến vào du_lieu/binh_do_tkcs.json. Vẽ ở tỷ lệ 1/2500 rồi đặt lên tờ A0 phóng 1,25 → 1/2000; hình trích
vẽ ở 1/1250 phóng 1,25 → 1/1000. Tim đập: DP (487477,39; 1255393,26) → C3 → C2 → DT, dài 1100 m;
đập đất từ C23 (lý trình 63,98) đến C2 (1050,73).
"""
import json
import math
import os

from bv_lib import To, so

TL = 2500.0
K = 1000.0 / TL                                     # mm giấy / m
GOC = (486700.0, 1255180.0)
CUA_SO = (20.0, 20.0, 1980.0, 590.0)                # cửa sổ (m, tương đối gốc)
X0, Y0 = 24.0, 340.0
TL_TRICH, MO_RONG = 1000.0, (1.0, 1.0)             # tỷ lệ vẽ hình trích, hệ số mở rộng cửa sổ (ngang, đứng)
DP = (487477.39, 1255393.26)
C3 = (488520.49, 1255279.67)
C2 = (488521.93, 1255279.51)
DT = (488570.92, 1255274.18)


def P(x, y):
    """Tọa độ VN-2000 (m) → mm giấy."""
    return X0 + (x - GOC[0] - CUA_SO[0]) * K, Y0 + (y - GOC[1] - CUA_SO[1]) * K


def tim(s, lech=0.0):
    """Điểm trên tim đập theo lý trình s (m) từ DP; lech > 0 về thượng lưu (bên trái hướng DP → DT)."""
    if s <= 1049.27:
        a, b, s0 = DP, C3, 0.0
    elif s <= 1050.73:
        a, b, s0 = C3, C2, 1049.27
    else:
        a, b, s0 = C2, DT, 1050.73
    L = math.dist(a, b)
    ux, uy = (b[0] - a[0]) / L, (b[1] - a[1]) / L
    t = s - s0
    return a[0] + ux * t - uy * lech, a[1] + uy * t + ux * lech


def cat(pts):
    """Cắt đường gấp khúc (m, tương đối gốc) theo cửa sổ – trả về các đoạn nằm trong."""
    x0, y0, x1, y1 = CUA_SO
    def trong(p):
        return x0 <= p[0] <= x1 and y0 <= p[1] <= y1

    def clip(p, q):
        dx, dy = q[0] - p[0], q[1] - p[1]
        t0, t1 = 0.0, 1.0
        for pp, qq in ((-dx, p[0] - x0), (dx, x1 - p[0]), (-dy, p[1] - y0), (dy, y1 - p[1])):
            if abs(pp) < 1e-12:
                if qq < 0:
                    return None
            else:
                r = qq / pp
                if pp < 0:
                    t0 = max(t0, r)
                else:
                    t1 = min(t1, r)
        if t0 > t1:
            return None
        return (p[0] + t0 * dx, p[1] + t0 * dy), (p[0] + t1 * dx, p[1] + t1 * dy)
    out, cur = [], []
    for p, q in zip(pts[:-1], pts[1:]):
        c = clip(p, q)
        if c is None:
            if len(cur) > 1:
                out.append(cur)
            cur = []
            continue
        a, b = c
        if not cur:
            cur = [a, b]
        elif math.dist(cur[-1], a) < 1e-6:
            cur.append(b)
        else:
            out.append(cur)
            cur = [a, b]
        if not trong(q):
            out.append(cur)
            cur = []
    if len(cur) > 1:
        out.append(cur)
    return out


def ve_nen(to, cua_so, x0, y0, k, khung=True):
    """Vẽ nền bình đồ trong cửa sổ cua_so (m, tương đối GOC), gốc giấy (x0, y0), k = mm/m."""
    global CUA_SO
    cu = CUA_SO
    CUA_SO = cua_so
    f = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'du_lieu', 'binh_do_tkcs.json')
    d = json.load(open(f, encoding='utf-8'))
    kieu = {'nen': ('NEN', {}), 'mai': ('NET_MANH', {}), 'tuong': ('NET_CHINH', {}), 'chan': ('NET_MANH', {}),
            'duong': ('NEN', dict(color=8)), 'nuoc': ('NUOC', {}), 'nha': ('NET_MANH', {}), 'tim': ('TRUC', {})}
    for kk, (lop, kw) in kieu.items():
        for pl in d[kk]:
            for seg in cat([tuple(q) for q in pl]):
                to.pline([(x0 + (x - cua_so[0]) * k, y0 + (y - cua_so[1]) * k) for x, y in seg], lop, **kw)
    CUA_SO = cu
    xb, yb = x0 + (cua_so[2] - cua_so[0]) * k, y0 + (cua_so[3] - cua_so[1]) * k
    if khung:
        to.rect(x0, y0, xb - x0, yb - y0, 'NET_CHINH')
    return lambda x, y: (x0 + (x - GOC[0] - cua_so[0]) * k, y0 + (y - GOC[1] - cua_so[1]) * k)


def nen(to):
    ve_nen(to, CUA_SO, X0, Y0, K)
    xa, ya = X0, Y0
    xb, yb = X0 + (CUA_SO[2] - CUA_SO[0]) * K, Y0 + (CUA_SO[3] - CUA_SO[1]) * K
    for gx in range(487000, 488700, 500):
        X = P(gx, 0)[0]
        to.line((X, ya), (X, ya + 2.5), 'NET_CHINH')
        to.line((X, yb), (X, yb - 2.5), 'NET_CHINH')
        to.text('X = %d' % gx, (X + 1.0, ya + 1.0), 1.7, 'LEFT')
    for gy in (1255500,):
        Y = P(0, gy)[1]
        to.line((xa, Y), (xa + 2.5, Y), 'NET_CHINH')
        to.line((xb, Y), (xb - 2.5, Y), 'NET_CHINH')
        to.text('Y = %d' % gy, (xa + 3.0, Y + 0.8), 1.7, 'LEFT')
    return (xa, ya, xb, yb)


def hinh_trich(to, ten, cua_so_tuyet_doi, x0, y0, tl=1000.0, ghi_chu=()):
    """Hình trích bình đồ tỷ lệ lớn: cua_so_tuyet_doi = (xmin, ymin, xmax, ymax) VN-2000."""
    k = 1000.0 / tl
    cs = (cua_so_tuyet_doi[0] - GOC[0], cua_so_tuyet_doi[1] - GOC[1], cua_so_tuyet_doi[2] - GOC[0],
          cua_so_tuyet_doi[3] - GOC[1])
    T = ve_nen(to, cs, x0, y0, k)
    yb = y0 + (cs[3] - cs[1]) * k
    to.ten_hinh(ten, x0, yb + 12.5, '1/%d' % tl, h=3.6, can='LEFT')
    # tim đập, cọc trong cửa sổ
    def trong(q):
        return x0 <= q[0] <= x0 + (cs[2] - cs[0]) * k and y0 <= q[1] <= yb
    pts = [T(*tim(ss)) for ss in [v * 2.0 for v in range(0, 551)]]
    seg = [q for q in pts if trong(q)]
    if len(seg) > 1:
        to.pline(seg, 'TRUC')
    return T, trong


def mui_ten(to, p0, p1, lop='NET_CHINH', dai=2.2, **kw):
    to.line(p0, p1, lop, **kw)
    a = math.atan2(p1[1] - p0[1], p1[0] - p0[0])
    for da in (2.7, -2.7):
        to.line(p1, (p1[0] + dai * math.cos(a + da), p1[1] + dai * math.sin(a + da)), lop, **kw)


def ghi(to, p_dau, p_chu, dong, h=2.0, rong=None):
    """Chú thích nhiều dòng (dòng đầu đậm) với đường dẫn."""
    to.line(p_dau, p_chu, 'SUA_CHUA')
    to.circle(p_dau, 0.6, 'SUA_CHUA')
    w = rong or max(len(s) for s in dong) * h * 0.5
    phai = p_chu[0] >= p_dau[0]
    x2 = p_chu[0] + (w if phai else -w)
    to.line(p_chu, (x2, p_chu[1]), 'SUA_CHUA')
    xt = p_chu[0] + (0.8 if phai else -w)
    for i, s in enumerate(dong):
        to.text(s, (xt, p_chu[1] + 0.8 - i * (h * 1.45) - (0 if i == 0 else 1.0)), h if i else h + 0.3,
                'LEFT', 'CHU_DAM' if i == 0 else 'CHU', 'VN_DAM' if i == 0 else 'VN')


def lop_phu(to):
    """Các hạng mục sửa chữa, quan trắc vẽ chồng lên nền."""
    # đập đất C23 → C2: tô đỏ dọc tim
    s_dau, s_cuoi = 63.98, 1050.73
    pts = [P(*tim(s_dau + (s_cuoi - s_dau) * k / 40)) for k in range(41)]
    to.pline(pts, 'SUA_CHUA', const_width=0.7)
    # màng khoan phụt C23 → C22 + 13,97
    pk = [P(*tim(63.98 + 50.0 * k / 10, -1.0)) for k in range(11)]
    to.pline(pk, 'SUA_CHUA', const_width=1.6, color=6)
    # cọc
    for ten, s in (('DP', 0.0), ('C23', 63.98), ('C22', 100.01), ('C20', 200.01), ('C18', 300.01), ('C16', 400.01),
                   ('C14', 500.01), ('C12', 600.01), ('C10', 700.01), ('C8', 800.01), ('C6', 900.01), ('C4', 1000.01),
                   ('C2', 1050.73), ('DT', 1100.0)):
        a, b = P(*tim(s, 22.0)), P(*tim(s, 30.0))
        to.line(a, b, 'NET_CHINH')
        to.text(ten, (b[0], b[1] + 0.8), 1.9, 'CENTER')
    # tuyến quan trắc thấm (giếng tự động)
    for ten, s in (('T9', 187.7), ('T7', 382.9), ('T5', 584.0), ('T3', 780.0), ('T1', 978.0)):
        a, b = P(*tim(s, 12.0)), P(*tim(s, -70.0))
        to.line(a, b, 'NUOC', linetype='DASHED', ltscale=0.3)
        to.text(ten, (b[0], b[1] - 3.0), 2.0, 'CENTER', 'NUOC')
        for lech in (-15.0, -35.0, -55.0):
            to.circle(P(*tim(s, lech)), 0.8, 'NUOC')
    # mốc lún: 5 trên đỉnh, 4 trên cơ hạ lưu
    for s in (200.0, 400.0, 600.0, 800.0, 980.0):
        X, Y = P(*tim(s, -1.5))
        to.pline([(X, Y + 1.2), (X - 1.0, Y - 0.6), (X + 1.0, Y - 0.6)], 'CUNG_TRUOT', dong=True)
    for s in (300.0, 500.0, 700.0, 900.0):
        X, Y = P(*tim(s, -30.0))
        to.pline([(X, Y + 1.2), (X - 1.0, Y - 0.6), (X + 1.0, Y - 0.6)], 'CUNG_TRUOT', dong=True)
    # chú thích hạng mục
    ghi(to, P(*tim(560.0, -2.5)), (P(*tim(560.0, 0))[0] - 40.0, Y0 + 170.0),
        ['ĐẬP ĐẤT C23 → C2 (L = 986 m) – BV-02, BV-03, BV-04',
         'Đắp bù, BT M250 mặt đập dày 20 cm, đỉnh +78,85; gờ chắn bánh; điện chiếu sáng',
         'Mái hạ lưu: sửa hố xói, rãnh BT chia ô 4 × 4 m, trồng cỏ; khơi thông lăng trụ',
         'Xử lý 2 hố sụt vai trái; mái thượng lưu đá xây giữ nguyên'], rong=128)
    ghi(to, P(*tim(90.0, -1.0)), (P(*tim(90.0, 0))[0] - 18.0, Y0 + 130.0),
        ['TRÀN SỰ CỐ B = 5 m – BV-05',
         'Màng khoan phụt chống thấm mang trái L = 50 m (C23 → C22 + 13,97)',
         'Lát lại tấm BTCT mái hạ lưu mang trái (1.474 m²)'], rong=112)
    ghi(to, P(486860.0, 1255535.0), (X0 + 166.0, Y0 + 32.0),
        ['TRÀN CHÍNH 9 × 2,4 m – BV-06',
         'Xử lý mặt đáy dốc (1.210 m²), khe co giãn, nạo vét kênh dẫn',
         'Giữ nguyên ngưỡng, trụ pin, tường bên, mũi phun'], rong=100)
    ghi(to, P(*tim(1033.0, 0.0)), (P(*tim(1033.0, 0))[0] - 70.0, Y0 + 205.0),
        ['CỐNG LẤY NƯỚC 1,8 × 2,0 m – BV-07',
         'Thay máy đóng mở VĐ20, gioăng, tấm trượt, giá đỡ trục vít',
         'Kiểm tra thân cống bằng camera/robot'], rong=96)
    ghi(to, P(*tim(46.4, 40.0)), (P(*tim(46.4, 0))[0] + 30.0, Y0 + 205.0),
        ['NHÀ QUẢN LÝ – trạm trung tâm SCADA',
         'Trạm đo mưa, đo mực nước hồ, 2 camera, đường quản lý BT M250 3,35 km'], rong=104)
    # hướng dòng chảy, mũi tên bắc, thước tỷ lệ
    mui_ten(to, P(*tim(300.0, 60.0)), P(*tim(300.0, 20.0)), 'NUOC', 2.5)
    to.text('HỒ CHỨA', P(*tim(300.0, 75.0)), 3.0, 'CENTER', 'NUOC', 'VN_DAM')
    to.text('HẠ LƯU', P(*tim(300.0, -110.0)), 3.0, 'CENTER', 'NUOC', 'VN_DAM')


def mui_bac(to, x, y, r=7.0):
    to.circle((x, y), r, 'NET_CHINH')
    to.pline([(x, y + r), (x - 2.6, y - r * 0.6), (x, y - r * 0.2)], 'NET_CHINH', dong=True)
    to.hatch([(x, y + r), (x - 2.6, y - r * 0.6), (x, y - r * 0.2)], 'SOLID', lop='NET_CHINH', mau_to=7)
    to.pline([(x, y + r), (x + 2.6, y - r * 0.6), (x, y - r * 0.2)], 'NET_CHINH', dong=True)
    to.text('B', (x, y + r + 1.5), 3.0, 'CENTER', 'CHU_DAM', 'VN_DAM')


def thuoc(to, x, y):
    for k in range(5):
        p = [(x + k * 8.0, y), (x + (k + 1) * 8.0, y), (x + (k + 1) * 8.0, y + 1.6), (x + k * 8.0, y + 1.6)]
        if k % 2 == 0:
            to.hatch(p, 'SOLID', lop='NET_CHINH', mau_to=7)
        to.pline(p, 'NET_CHINH', dong=True)
    for k in range(6):
        to.text('%d' % (k * 20), (x + k * 8.0, y + 2.6), 1.8, 'CENTER')
    to.text('m', (x + 43.0, y + 0.2), 1.8, 'LEFT')


def bang(to):
    cot = [30, 162]
    hang = [['Hạng mục', 'Giải pháp sửa chữa, nâng cấp được chọn (Bảng 3.2)'],
            ['Đập đất', 'Đắp bù, gia cố mặt đập BT M250 dày 20 cm, đỉnh +78,85 m; gờ chắn bánh; điện chiếu sáng;\n'
                        'sửa hố xói, hệ rãnh BT, trồng cỏ mái hạ lưu; xử lý hố sụt vai trái; khơi thông lăng trụ'],
            ['Tràn chính', 'Giữ nguyên quy mô; sửa chữa bề mặt đáy dốc, khe co giãn; nạo vét kênh dẫn; khoan neo ngưỡng D25'],
            ['Tràn sự cố', 'Màng khoan phụt XM – sét mang trái (L = 50 m, 3 hàng TKCS; đề xuất 3 hàng 15 m + 2 hàng A, B); lát lại tấm BT mái'],
            ['Cống lấy nước', 'Thay 2 máy đóng mở VĐ20; thay gioăng, tấm trượt; bổ sung giá đỡ trục vít; kiểm tra thân cống'],
            ['Quan trắc, SCADA', '14 giếng quan trắc thấm tự động (T1, T3, T5, T7, T9), 9 mốc lún, trạm đo mưa,\n'
                                 'đo mực nước, camera, trạm trung tâm'],
            ['Đường quản lý', 'Bê tông hóa 3,35 km, mặt BT M250 rộng 4 m']]
    to.bang(30.0, 318.0, cot, hang, h=1.9, cao_dong=5.0, can=['L', 'L'],
            tieu_de='TỔNG HỢP GIẢI PHÁP SỬA CHỮA, NÂNG CẤP')
    cot2 = [84, 16, 42]
    hang2 = [['Thông số', 'Đơn vị', 'Giá trị'],
             ['Cấp công trình', '–', 'II'],
             ['MNDBT / MNC', 'm', '+74,70 / +67,50'],
             ['MNLTK (P = 1%) / MNLKT (P = 0,2%)', 'm', '+77,63 / +78,54'],
             ['Dung tích ứng với MNDBT (tính lại)', '10⁶ m³', '36,41'],
             ['Đỉnh đập hiện trạng / TKCS / đồ án', 'm', '+78,35 / +78,75 / +78,85'],
             ['Đỉnh tường chắn sóng (giữ nguyên)', 'm', '+79,40'],
             ['Chiều dài / bề rộng đỉnh đập', 'm', '986 / 5,0'],
             ['Tràn chính: ngưỡng, b', 'm', '+74,70; 9 × 2,4'],
             ['Tràn sự cố: ngưỡng, b', 'm', '+71,00; 5,0'],
             ['Cống: ngưỡng, b × h, L', 'm', '+65,60; 1,8 × 2,0; 67']]
    to.bang(650.0, 318.0, cot2, hang2, h=2.0, cao_dong=5.0, tieu_de='THÔNG SỐ CHỦ YẾU')


def chu_giai(to, x, y):
    to.text('KÝ HIỆU', (x, y), 3.0, 'LEFT', 'CHU_DAM', 'VN_DAM')
    y -= 6.0
    muc = [('do', 'Đập đất sửa chữa, nâng cấp (tim đập C23 → C2)'), ('tim', 'Tim tuyến công trình'),
           ('kp', 'Màng khoan phụt chống thấm mang trái tràn sự cố'), ('tuong', 'Tường, kết cấu BT hiện trạng'),
           ('mai', 'Mái đập, mái đào (nền bình đồ TKCS)'), ('nuoc', 'Tuyến quan trắc thấm, giếng tự động'),
           ('lun', 'Mốc quan trắc lún'), ('nen', 'Địa vật, đường (nền bình đồ TKCS)')]
    for ma, ten in muc:
        xa, xb = x + 1.0, x + 11.0
        if ma == 'do':
            to.pline([(xa, y), (xb, y)], 'SUA_CHUA', const_width=0.7)
        elif ma == 'tim':
            to.line((xa, y), (xb, y), 'TRUC')
        elif ma == 'kp':
            to.pline([(xa, y), (xb, y)], 'SUA_CHUA', const_width=1.6, color=6)
        elif ma == 'tuong':
            to.line((xa, y), (xb, y), 'NET_CHINH')
        elif ma == 'mai':
            to.line((xa, y), (xb, y), 'NET_MANH')
            for k in range(4):
                to.line((xa + 1 + 3 * k, y), (xa + 1 + 3 * k, y - (2.0 if k % 2 == 0 else 1.0)), 'NET_MANH')
        elif ma == 'nuoc':
            to.line((xa, y), (xb, y), 'NUOC', linetype='DASHED', ltscale=0.3)
            to.circle((xa + 5, y), 0.8, 'NUOC')
        elif ma == 'lun':
            X = xa + 5
            to.pline([(X, y + 1.2), (X - 1.0, y - 0.6), (X + 1.0, y - 0.6)], 'CUNG_TRUOT', dong=True)
        else:
            to.line((xa, y), (xb, y), 'NEN')
        to.text(ten, (x + 14.0, y), 2.0, 'ML')
        y -= 5.6


def ghi_chu(to, x, y):
    gc = ('GHI CHÚ: \\P'
          '1. Hệ tọa độ VN-2000, cao độ Quốc gia; kích thước, cao trình ghi bằng m. Nền bình đồ trích từ hồ sơ TKCS '
          '(tỷ lệ gốc 1/1000), thể hiện lại ở tỷ lệ 1/2000. \\P'
          '2. Lý trình theo tim đập từ cọc DP; đập đất từ C23 (tường bên trái tràn sự cố, lý trình 63,98) đến C2 '
          '(1050,73), dài 986 m. Cao trình đỉnh đập sau sửa chữa +78,85 (mép thượng lưu), TKCS +78,75. \\P'
          '3. Vị trí tuyến quan trắc T1 ÷ T9 theo hồ sơ hiện trạng; vị trí mốc lún mới thể hiện sơ đồ, chuẩn xác khi '
          'lập bản vẽ thi công. \\P'
          '4. Chi tiết các hạng mục xem: BV-02 cắt dọc đập; BV-03, BV-04 đập đất; BV-05 khoan phụt (chuyên đề); '
          'BV-06 tràn chính; BV-07 cống lấy nước.')
    to.mtext(gc, (x, y), 2.0, rong=222)


def mo_rong(cs):
    cx, cy = (cs[0] + cs[2]) / 2, (cs[1] + cs[3]) / 2
    w, h = (cs[2] - cs[0]) * MO_RONG[0] / 2, (cs[3] - cs[1]) * MO_RONG[1] / 2
    return (cx - w, cy - h, cx + w, cy + h)


def trich_tsc(to):
    T, trong = hinh_trich(to, 'HÌNH TRÍCH A – KHU VỰC TRÀN SỰ CỐ, NHÀ QUẢN LÝ, ĐẦU ĐẬP TRÁI',
                          mo_rong((487370.0, 1255265.0, 487730.0, 1255465.0)), 32.0, 24.0, TL_TRICH)
    # đập đất C23 → ...
    pts = [T(*tim(63.98 + 2.0 * i)) for i in range(140)]
    to.pline([q for q in pts if trong(q)], 'SUA_CHUA', const_width=0.5)
    # 3 hàng khoan (A, B, C cách nhau 1,5 m; ở đây dịch theo tim đập, sơ đồ)
    for lech, ten in ((1.5, 'A'), (0.0, 'B'), (-1.5, 'C')):
        L = 15.0 if ten == 'B' else 50.0
        pk = [T(*tim(63.98 + L * i / 10, lech)) for i in range(11)]
        to.pline(pk, 'SUA_CHUA', color=6, const_width=0.35)
    a = T(*tim(63.98 + 25.0, 0.0))
    to.chu_thich_mui_ten(a, (a[0] + 22.0, a[1] + 26.0), 'Màng khoan phụt: 3 hàng 15 m + 2 hàng 35 m (BV-05)', h=2.0,
                         rong=92)
    for ten, ss in (('C23', 63.98), ('C22', 100.01), ('C21', 150.01), ('C20', 200.01), ('C19', 250.01)):
        p1, p2 = T(*tim(ss, 8.0)), T(*tim(ss, 14.0))
        if trong(p1):
            to.line(p1, p2, 'NET_CHINH')
            to.text(ten, (p2[0], p2[1] + 0.8), 2.0, 'CENTER')
    for ten, ss in (('T9', 187.7),):
        p1, p2 = T(*tim(ss, 10.0)), T(*tim(ss, -70.0))
        to.line(p1, p2, 'NUOC', linetype='DASHED', ltscale=0.3)
        to.text(ten, (p2[0] + 1.0, p2[1] + 1.0), 2.2, 'LEFT', 'NUOC')
    q = T(487537.0, 1255330.0)
    to.text('TRÀN SỰ CỐ', (q[0] + 3.0, q[1]), 2.4, 'LEFT', 'CHU_DAM', 'VN_DAM')
    q = T(487548.6, 1255460.0)
    to.text('HỒ CHỨA', (q[0], q[1] - 8.0), 3.0, 'CENTER', 'NUOC', 'VN_DAM')
    to.text('Lát lại tấm BTCT mái hạ lưu mang trái; xử lý hố sụt sau tường cánh sau khi khoan phụt xong', (40.0, 27.0),
            2.0, 'LEFT', 'SUA_CHUA')


def trich_cong(to):
    T, trong = hinh_trich(to, 'HÌNH TRÍCH B – KHU VỰC CỐNG LẤY NƯỚC, VAI PHẢI',
                          mo_rong((488405.0, 1255185.0, 488630.0, 1255385.0)),
                          32.0 + 360.0 * MO_RONG[0] * 1000.0 / TL_TRICH + 20.0, 24.0, TL_TRICH)
    pts = [T(*tim(800.0 + 2.0 * i)) for i in range(126)]
    to.pline([q for q in pts if trong(q) and True], 'SUA_CHUA', const_width=0.5)
    for ten, ss in (('C5', 950.01), ('C4', 1000.01), ('C3', 1049.28), ('C2', 1050.73), ('C1', 1075.01), ('DT', 1100.0)):
        p1, p2 = T(*tim(ss, 8.0)), T(*tim(ss, 14.0 if ten not in ('C3',) else 20.0))
        if trong(p1):
            to.line(p1, p2, 'NET_CHINH')
            to.text(ten, (p2[0], p2[1] + 0.8), 2.0, 'CENTER')
    p1, p2 = T(*tim(978.0, 10.0)), T(*tim(978.0, -70.0))
    to.line(p1, p2, 'NUOC', linetype='DASHED', ltscale=0.3)
    to.text('T1', (p2[0] + 1.0, p2[1] + 1.0), 2.2, 'LEFT', 'NUOC')
    a = T(*tim(1033.0, 0.0))
    to.chu_thich_mui_ten(a, (a[0] + 15.0, a[1] + 45.0), 'Cống lấy nước: thay máy đóng mở VĐ20 (BV-07)', h=2.0, rong=82)
    q = T(488500.0, 1255360.0)
    to.text('HỒ CHỨA', q, 3.0, 'CENTER', 'NUOC', 'VN_DAM')



def mat_bang_chinh(to):
    to.ten_hinh('MẶT BẰNG TỔNG THỂ CỤM CÔNG TRÌNH ĐẦU MỐI HỒ CÀ GIÂY', 24, 576, None, h=4.5, can='LEFT')
    to.text('TỶ LỆ: 1/2500', (330, 576), 3.0)
    nen(to)
    lop_phu(to)
    mui_bac(to, 795.0, 548.0)
    thuoc(to, 760.0, 331.0)


def thong_tin(to):
    bang(to)
    chu_giai(to, 232.0, 322.0)
    ghi_chu(to, 412.0, 322.0)


def hinh_trich_ab(to):
    trich_tsc(to)
    trich_cong(to)


def ve(thu_muc):
    global TL_TRICH, MO_RONG
    TL_TRICH, MO_RONG = 1250.0, (1.5, 1.6)
    to = To(1, 'MẶT BẰNG TỔNG THỂ CỤM ĐẦU MỐI\nBỐ TRÍ CÁC HẠNG MỤC SỬA CHỮA, NÂNG CẤP', '1/2000', 'HCG-BV01')
    b = to.dat(mat_bang_chinh, f=1.25, neo='TL', xy=(28.0, 826.0))
    b = to.dat(thong_tin, f=1.25, neo='TL', xy=(28.0, b[1] - 8.0))
    to.dat(hinh_trich_ab, f=1.25, neo='TL', xy=(28.0, b[1] - 10.0))
    to.luu(thu_muc + '/BV01_MatBangTongThe.dxf', thu_muc + '/BV01_MatBangTongThe.pdf',
           thu_muc + '/BV01_MatBangTongThe.png')
    return to


if __name__ == '__main__':
    import sys
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
