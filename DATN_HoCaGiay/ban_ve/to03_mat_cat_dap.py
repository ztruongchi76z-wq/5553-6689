# -*- coding: utf-8 -*-
"""BV-03: MẶT CẮT NGANG ĐẬP HIỆN TRẠNG VÀ SAU SỬA CHỮA – SƠ ĐỒ, KẾT QUẢ TÍNH THẤM VÀ ỔN ĐỊNH (mặt cắt C16)."""
import math

import numpy as np

import du_lieu_ban_ve as D
from bv_lib import To, KhungNhin, so

TL = 250            # tỷ lệ mặt cắt 1:250
XL, XR, ZDAY_VE = -95.0, 100.0, 40.0
VAT_LIEU = [  # khóa, tên, mẫu hatch, tỷ lệ, góc
    ('DD1', 'Đất đắp ĐĐ-I (khối thượng lưu, lõi, chân khay)', 'AR-SAND', 0.25, 0),
    ('DD2', 'Đất đắp ĐĐ-II (khối hạ lưu)', 'ANSI31', 2.2, 0),
    ('OK', 'Ống khói + đệm tiêu nước (cát lọc)', 'DOTS', 0.6, 0),
    ('LT', 'Lăng trụ đá thoát nước', 'GRAVEL', 0.35, 0),
    ('L4', 'Lớp 4: sét pha lẫn dăm sạn (eQ)', 'ANSI33', 2.0, 0),
    ('LI', 'Lớp I: đá cát bột kết phong hóa mạnh', 'ANSI37', 2.0, 0),
    ('LII', 'Lớp II: đá cát bột kết phong hóa vừa', 'AR-B816', 0.08, 0),
]
HMAP = {k: (m, s, g) for k, _, m, s, g in VAT_LIEU}


def da_giac(h):
    """Các vùng vật liệu (tọa độ thật)."""
    Zc, Zok, dx = h['Zc'], h['Zok'], h['dx_ok']
    z54, zd = h['Zday'], h['Zday'] + h['t_dem']
    ck = h['ck']
    return {
        'DD1': [(h['x_u3'], z54), (h['x_u2'], h['Zcu']), (h['x_u1'], h['Zcu']), (h['xcu'], Zc), (0.0, Zc),
                (1.0, Zok), (h['x_ok_u_bot'], z54)],
        'OK': [(1.0, Zok), (1.0 + dx, Zok), (h['x_ok_d_dem'], zd), (h['x_dem_end'], zd), (h['x_l0'], z54),
               (h['x_ok_u_bot'], z54)],
        'DD2': [(0.0, Zc), (h['xcd'], Zc), (h['x_b1'], h['Zcd']), (h['x_b2'], h['Zcd']), (h['x_l1'], h['Z_ltr']),
                (h['x_dem_end'], zd), (h['x_ok_d_dem'], zd), (1.0 + dx, Zok), (1.0, Zok)],
        'LT': [(h['x_l1'], h['Z_ltr']), (h['x_l2'], h['Z_ltr']), (h['x_l3'], z54), (h['x_l0'], z54),
               (h['x_dem_end'], zd)],
        'CK': list(ck),
        'L4a': [(XL, h['Z4']), (XL, z54), ck[0], ck[1]],
        'L4b': [ck[2], ck[3], (XR, z54), (XR, h['Z4'])],
        'LI': [(XL, h['ZI']), (XL, h['Z4']), (XR, h['Z4']), (XR, h['ZI'])],
        'LII': [(XL, ZDAY_VE), (XL, h['ZI']), (XR, h['ZI']), (XR, ZDAY_VE)],
    }


def mat_tren(h):
    xs = [XL, h['x_u3'], h['x_u2'], h['x_u1'], h['xcu'], h['xcd'], h['x_b1'], h['x_b2'], h['x_l1'], h['x_l2'],
          h['x_l3'], XR]
    zs = [h['Zday'], h['Zday'], h['Zcu'], h['Zcu'], h['Zc'], h['Zc'], h['Zcd'], h['Zcd'], h['Z_ltr'], h['Z_ltr'],
          h['Zday'], h['Zday']]
    return xs, zs


def ve_mat_cat(to, kn, h, tien_to, tieu_de, sau_sc=False):
    vung = da_giac(h)
    for k, pts in vung.items():
        key = {'CK': 'DD1', 'L4a': 'L4', 'L4b': 'L4'}.get(k, k)
        m, s, g = HMAP[key]
        to.hatch([kn(*p) for p in pts], m, s, goc=g)
        to.pline([kn(*p) for p in pts], 'NET_MANH', dong=True)
    xs, zs = mat_tren(h)
    to.pline(kn.pts(xs, zs), 'NET_CHINH')
    # vết cắt đáy (đường gãy)
    to.line(kn(XL, ZDAY_VE), kn(XR, ZDAY_VE), 'NET_MANH', linetype='DASHED')
    # tường chắn sóng (hiện trạng, giữ nguyên) – 0,5 m × 1,05 m
    tw = [kn(h['xcu'], h['Zc'] - 0.2), kn(h['xcu'] + 0.5, h['Zc'] - 0.2), kn(h['xcu'] + 0.5, D.TUONG_CS),
          kn(h['xcu'], D.TUONG_CS)]
    to.hatch(tw, 'SOLID', mau_to=8)
    to.pline(tw, 'NET_CHINH', dong=True)
    # mực nước thượng lưu
    for ten, z, nx in (('MNLKT', D.MN['MNLKT'], -92), ('MNLTK', D.MN['MNLTK'], -84), ('MNDBT', D.MN['MNDBT'], -76),
                       ('MNC', D.MN['MNC'], -68)):
        to.muc_nuoc(kn(XL, 0)[0], kn(D.x_mai_tl(h, z), 0)[0], kn(0, z)[1], ten, z, nhan_x=kn(nx, 0)[0])
    # đường bão hòa 3 trường hợp
    kieu = {'MNDBT': ('CONTINUOUS', 'MNDBT'), 'MNLTK': ('DASHED', 'MNLTK'), 'MNLKT': ('DASHDOT', 'MNLKT')}
    for ten, (lt, nhan) in kieu.items():
        Zw = D.MN[ten]
        pts = list(D.DBH['%s_%s' % (tien_to, ten)]) + [(D.x_mai_tl(h, Zw), Zw)]
        to.pline([kn(*p) for p in pts], 'THAM', linetype=lt, ltscale=0.6)
    # cung trượt nguy hiểm
    for ma in ('TH5', 'TH8'):
        c = D.CUNG['%s_%s' % (tien_to, ma)]
        ve_cung(to, kn, h, c)
    # cao trình
    cts = [(h['xcu'] + 0.25, D.TUONG_CS, 'L'), (h['xcd'], h['Zc'], 'R'), (h['x_u2'] + 1.5, h['Zcu'], 'L'),
           (h['x_b1'] + 7, h['Zcd'], 'R'), (h['x_l1'] + 1.5, h['Z_ltr'], 'R'), (h['x_l3'] + 4, h['Zday'], 'R'),
           (h['x_ck'], h['ck_den'], 'R'), (XR - 12, h['Z4'], 'R'), (XR - 12, h['ZI'], 'R')]
    for x, z, b in cts:
        to.cao_trinh(*kn(x, z), z, ben=b)
    # kích thước
    yk = kn(0, h['Zc'])[1] + 9
    to.dim(kn(h['xcu'], h['Zc']), kn(h['xcd'], h['Zc']), 9, TL)
    to.dim(kn(h['x_u2'], h['Zcu']), kn(h['x_u1'], h['Zcu']), 7, TL)
    to.dim(kn(h['x_b1'], h['Zcd']), kn(h['x_b2'], h['Zcd']), 7, TL)
    to.dim(kn(h['x_l1'], h['Z_ltr']), kn(h['x_l2'], h['Z_ltr']), 7, TL)
    # hệ số mái
    for (xa, za, xb, zb, nhan) in ((h['x_u1'], h['Zcu'], h['xcu'], h['Zc'], 'm = 3,0'),
                                   (h['x_u3'], h['Zday'], h['x_u2'], h['Zcu'], 'm = 3,5'),
                                   (h['xcd'], h['Zc'], h['x_b1'], h['Zcd'], 'm = 3,0'),
                                   (h['x_b2'], h['Zcd'], h['x_l1'], h['Z_ltr'], 'm = 3,5')):
        p, q = kn(xa, za), kn(xb, zb)
        goc = math.degrees(math.atan2(q[1] - p[1], q[0] - p[0]))
        if goc > 90:
            goc -= 180
        if goc < -90:
            goc += 180
        mx, my = (p[0] + q[0]) / 2, (p[1] + q[1]) / 2
        to.text(nhan, (mx, my + 1.6), 2.5, 'CENTER', rot=goc)
    # chú thích vật liệu
    z_ok = 78.6 if sau_sc else 83.0
    cg = [((-40, 62), (-35, 80.5), 'Khối ĐĐ-I (thượng lưu)'), ((6, 72), (13, z_ok), 'Ống khói thoát nước dày 1,0 m, 1:1'),
          ((50, 54.4), (70, 67.5), 'Đệm tiêu nước dày 0,8 m'), ((40, 62), (46, 73.5), 'Khối ĐĐ-II (hạ lưu)'),
          ((h['x_l2'] - 1, 57), (h['x_l2'] + 2, 66), 'Lăng trụ đá'), ((-2, 50), (-26, 46.5), 'Chân khay sâu 6 m')]
    for p, q, s in cg:
        to.chu_thich_mui_ten(kn(*p), kn(*q), s, 2.3, rong=len(s) * 1.45)
    to.text('LỚP 4', kn(-80, 50.5), 2.5, 'LEFT', 'CHU_DAM', 'VN_DAM')
    to.text('LỚP I', kn(-80, 45.0), 2.5, 'LEFT', 'CHU_DAM', 'VN_DAM')
    to.text('LỚP II', kn(-80, 41.0), 2.5, 'LEFT', 'CHU_DAM', 'VN_DAM')
    to.text('THƯỢNG LƯU', kn(-60, 56), 3.0, 'CENTER', 'CHU_DAM', 'VN_DAM')
    to.text('HẠ LƯU', kn(92, 56), 3.0, 'CENTER', 'CHU_DAM', 'VN_DAM')
    # tim đập
    to.line(kn(0, ZDAY_VE - 1), kn(0, D.TUONG_CS + 4.5), 'TRUC')
    to.text('TIM ĐẬP', kn(0, D.TUONG_CS + 4.8), 2.2, 'CENTER')
    # tên hình
    to.ten_hinh(tieu_de, kn(XL, 0)[0] + 3, kn(0, D.TUONG_CS)[1] + 14, '1/250', h=4.5, can='LEFT')
    if sau_sc:
        # các hạng mục sửa chữa (nét đỏ)
        to.pline([kn(h['xcu'] + 0.5, h['Zc']), kn(h['xcd'], h['Zc'])], 'SUA_CHUA', const_width=0.6)
        to.chu_thich_mui_ten(kn(1.8, h['Zc']), kn(9, h['Zc'] + 5.5),
                             'Đắp bù, mặt BT M250 dày 20 cm – đỉnh +78,85 (chi tiết A, BV-04)', 2.3, rong=95)
        mai = [kn(h['xcd'] + 0.3, h['Zc'] - 0.1), kn(h['x_b1'], h['Zcd']), kn(h['x_b2'], h['Zcd']),
               kn(h['x_l1'], h['Z_ltr'])]
        to.pline(mai, 'SUA_CHUA', const_width=0.45)
        to.chu_thich_mui_ten(kn(24, h['Zc'] - 7.2), kn(36, h['Zc'] - 0.5),
                             'Sửa hố xói (K ≥ 0,95), hệ rãnh BT ô 4×4 m + trồng cỏ (BV-04)', 2.3, rong=92)
        to.chu_thich_mui_ten(kn(h['x_l3'] - 1.5, 54.6), kn(h['x_l3'] - 3, 50.6),
                             'Khơi thông lăng trụ, rãnh chân đập 50×50', 2.3, rong=62)


def ve_cung(to, kn, h, c):
    xs, zs = mat_tren(h)
    t = np.linspace(math.pi * 1.02, math.pi * 1.98, 2400)
    px = c['xc'] + c['R'] * np.cos(t)
    pz = c['zc'] + c['R'] * np.sin(t)
    zt = np.interp(px, xs, zs)
    m = (pz <= zt + 1e-6) & (px > XL) & (px < XR)
    if not m.any():
        return
    idx = np.where(m)[0]
    seg = idx[(idx >= idx[0]) & (idx <= idx[-1])]
    pts = [kn(px[i], pz[i]) for i in seg]
    to.pline(pts, 'CUNG_TRUOT')
    # bán kính tới hai đầu cung (vẽ một đoạn hướng về tâm)
    for i in (seg[0], seg[-1]):
        a = kn(px[i], pz[i])
        dxv, dzv = c['xc'] - px[i], c['zc'] - pz[i]
        L = math.hypot(dxv, dzv)
        b = kn(px[i] + dxv / L * 9, pz[i] + dzv / L * 9)
        to.line(a, b, 'CUNG_TRUOT', linetype='DASHED', ltscale=0.4)
    i0 = seg[len(seg) // 2]
    p_day = kn(px[i0], pz[i0])
    nhan = '%s: Kmin = %s – tâm O(%s; %s), R = %s m' % (c['mo_ta'].split(':')[0], so(c['K'], 3), so(c['xc']),
                                                       so(c['zc']), so(c['R']))
    if c['mai'] == 'HL':
        to.chu_thich_mui_ten(p_day, (p_day[0] + 6, p_day[1] - 10), nhan, 2.3, rong=100)
    else:
        to.chu_thich_mui_ten(p_day, (p_day[0] - 6, p_day[1] - 10), nhan, 2.3, rong=100)


def mat_cat_ht(to):
    """Các khối vẽ theo bố cục A1 (tỷ lệ 1/250); đặt lên A0 phóng 1,25 → 1/200."""
    kn1 = KhungNhin(32.0, 398.0, XL, ZDAY_VE, TL)
    ve_mat_cat(to, kn1, D.mat_cat(D.DINH_HT), 'HT', 'MẶT CẮT NGANG ĐẬP HIỆN TRẠNG (CỌC C16 – KM0+400)')


def mat_cat_sc(to):
    kn2 = KhungNhin(32.0, 200.0, XL, ZDAY_VE, TL)
    ve_mat_cat(to, kn2, D.mat_cat(D.DINH_MO_HINH_SC), 'SC', 'MẶT CẮT NGANG ĐẬP SAU SỬA CHỮA – SƠ ĐỒ TÍNH',
               sau_sc=True)


def phan_duoi(to):
    # chú giải đường
    x0, y0 = 600.0, 190.0
    to.text('KÝ HIỆU', (x0, y0), 3.5, 'LEFT', 'CHU_DAM', 'VN_DAM')
    ky = [('THAM', 'CONTINUOUS', 'Đường bão hòa khi MN hồ = MNDBT'), ('THAM', 'DASHED', 'Đường bão hòa khi MN hồ = MNLTK'),
          ('THAM', 'DASHDOT', 'Đường bão hòa khi MN hồ = MNLKT'), ('CUNG_TRUOT', 'CONTINUOUS', 'Cung trượt nguy hiểm nhất'),
          ('NUOC', 'CONTINUOUS', 'Mực nước thượng lưu'), ('SUA_CHUA', 'CONTINUOUS', 'Hạng mục sửa chữa')]
    for i, (lop, lt, s) in enumerate(ky):
        y = y0 - 7 - i * 6
        to.line((x0, y), (x0 + 16, y), lop, linetype=lt, ltscale=0.6)
        to.text(s, (x0 + 19, y), 2.5, 'ML')
    yv = y0 - 7 - len(ky) * 6 - 2
    for i, (k, s, m, sc, g) in enumerate(VAT_LIEU):
        y = yv - i * 6.5
        o = [(x0, y - 2.2), (x0 + 16, y - 2.2), (x0 + 16, y + 2.2), (x0, y + 2.2)]
        to.hatch(o, m, sc, goc=g)
        to.pline(o, 'NET_MANH', dong=True)
        to.text(s, (x0 + 19, y), 2.5, 'ML')
    # bảng kết quả thấm
    cot = [40, 26, 26, 22, 20, 22]
    hang = [['Trường hợp', 'q·10⁶ hiện\ntrạng (m³/s/m)', 'q·10⁶ sau\nSC (m³/s/m)', 'J tiếp xúc\nống khói*', 'J cửa ra*',
             'Z bão hòa\ntại tim (m)']] + [list(r) for r in D.THAM]
    yb = to.bang(25, 186, cot, hang, h=2.4, cao_dong=5.6, tieu_de='BẢNG KẾT QUẢ TÍNH THẤM (sau sửa chữa, *giá trị 95%)')
    to.text('Gradien cho phép: tiếp xúc tầng lọc [J] = 2,0; cửa ra khối đắp [J] = 0,85 – đều thỏa mãn.', (25, yb - 4),
            2.3, 'LEFT')
    cot2 = [13, 66, 11, 14, 22, 18, 12, 15]
    hang2 = [['TH', 'Trường hợp tính toán', 'Mái', 'Tổ hợp', 'K hiện trạng', 'K sau SC', '[K]', 'Đánh giá']]
    hang2 += [list(r) + ['Đạt'] for r in D.ON_DINH]
    to.bang(198, 186, cot2, hang2, h=2.3, cao_dong=5.2, tieu_de='BẢNG HỆ SỐ ỔN ĐỊNH MÁI ĐẬP (BISHOP)')
    # ghi chú
    gc = ('GHI CHÚ:\\P'
          '1. Cao trình ghi bằng m, kích thước ghi bằng cm.\\P'
          '2. Mặt cắt tính toán tại cọc C16 (Km0+400), chiều cao đập lớn nhất; trục x = 0 tại tim đập.\\P'
          '3. Bề rộng chân khay, chiều dày đệm tiêu nước, mái lăng trụ là giả thiết (hồ sơ không có).\\P'
          '4. Sơ đồ tính sau sửa chữa lấy đỉnh +78,85 (đồ án), giữ hệ số mái, bề rộng đỉnh và cơ như hiện trạng.\\P'
          '5. Thấm: PTHH bão hòa – không bão hòa; ổn định: Bishop đơn giản, kiểm tra chéo Spencer (lệch ≤ 0,2%); kh = 0,05 – thiên an toàn so với agR = 0,0222g (Chương 2, 5); TH11, TH12: tổ hợp '
          'thi công (TC).')
    to.mtext(gc, (25, 112), 2.4, rong=290)


def ve(thu_muc):
    to = To(3, 'MẶT CẮT NGANG ĐẬP HIỆN TRẠNG VÀ SAU SỬA CHỮA\nSƠ ĐỒ, KẾT QUẢ TÍNH THẤM VÀ ỔN ĐỊNH', '1/200', 'HCG-BV03')
    b = to.dat(mat_cat_ht, f=1.25, neo='TL', xy=(30.0, 814.0))
    b = to.dat(mat_cat_sc, f=1.25, neo='TL', xy=(30.0, b[1] - 30.0))
    to.dat(phan_duoi, f=1.25, neo='TL', xy=(30.0, b[1] - 35.0))
    to.luu(thu_muc + '/BV03_MatCatNgangDap.dxf', thu_muc + '/BV03_MatCatNgangDap.pdf', thu_muc + '/BV03_MatCatNgangDap.png')
    return to


if __name__ == '__main__':
    import sys
    ve(sys.argv[1] if len(sys.argv) > 1 else '.')
