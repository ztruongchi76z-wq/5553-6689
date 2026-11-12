# -*- coding: utf-8 -*-
"""Chạy các trường hợp tính thấm và ổn định cho mặt cắt lòng sông (hiện trạng và sau sửa chữa)."""
import json, time
import numpy as np
import matplotlib.tri as mtri
import matplotlib.pyplot as plt
import so_lieu as S
import tham_fem as T
import on_dinh as O
import ve_hinh as VH

KQ = {}
MNLTK, MNLKT = S.MNLTK, S.MNLKT      # điều tiết lũ hai tràn mở hết cửa (thống nhất 21/10/2026)
TEN_VL = {1: 'ĐĐ-I (thân đập TL + lõi)', 3: 'ĐĐ-II (khối hạ lưu)', 4: 'Ống khói + đệm tiêu nước',
          5: 'Lăng trụ đá', 6: 'Chân khay (ĐĐ-I)', 7: 'Lớp 4 (sét pha eQ)', 8: 'Lớp I (đá PH mạnh)',
          9: 'Lớp II (đá PH vừa)', 2: 'Ống khói'}


def phan_tich_tham(mc, bt, r, ten):
    """Trích kết quả thấm. Gradient chỉ xét trong vùng bão hòa (ψ ≥ 0) – vùng không bão hòa hầu như
    không có dòng chảy (k_r rất nhỏ) nên gradient lớn cục bộ ở đó không có ý nghĩa xói ngầm."""
    xy, tri, mat = bt.xy, bt.tri, bt.mat
    gx, gz, J = bt.gradient()
    psi_e = bt.H[tri].mean(1) - xy[tri, 1].mean(1)
    bh = psi_e >= -0.05
    def elem_ke(nodes):
        m = np.isin(tri, np.asarray(nodes, int)).any(1)
        return np.where(m)[0]
    def tk(idx):
        idx = idx[bh[idx]]
        if len(idx) == 0:
            return np.nan, np.nan
        w = bt.area[idx]
        return float(np.sum(J[idx] * w) / np.sum(w)), float(np.percentile(J[idx], 95))
    drain_nodes = np.unique(tri[np.isin(mat, [2, 4])])
    e_core = np.where(np.isin(mat, [1, 6]))[0]
    tiep = e_core[np.isin(tri[e_core], drain_nodes).any(1)]
    J_tiep, J_tiep95 = tk(tiep)
    e_ra = elem_ke(bt.seep_on) if len(bt.seep_on) else np.array([], int)
    J_ra, J_ra95 = tk(e_ra)
    Jz_ra = float((-gz[e_ra][bh[e_ra]]).max()) if len(e_ra) and bh[e_ra].any() else np.nan
    J_ck, J_ck95 = tk(np.where(mat == 6)[0])
    e_vao = elem_ke(r['up']); e_vao = e_vao[np.isin(mat[e_vao], [1, 6, 7])]
    J_vao, J_vao95 = tk(e_vao)
    zmid = 0.5 * (mc.Zday + mc.Zc) - 2.0
    Tg = mtri.Triangulation(xy[:, 0], xy[:, 1], tri)
    Hi = mtri.LinearTriInterpolator(Tg, bt.H)
    xa = -2.5 - 3.0 * (mc.Zc - zmid) if zmid >= mc.Zcu else mc.x_u2 - 3.5 * (mc.Zcu - zmid)
    xb = 1.0 + (mc.Zok - zmid)
    Ha, Hb = float(Hi(xa + 0.3, zmid)), float(Hi(xb - 0.3, zmid))
    J_tb = (Ha - Hb) / (xb - xa)
    psi = bt.H - xy[:, 1]
    fig = plt.figure(); cs = plt.tricontour(Tg, psi, levels=[0.0]); plt.close(fig)
    segs = cs.allsegs[0]
    dbh = max(segs, key=len) if segs else np.zeros((0, 2))
    # điểm đường bão hòa gặp ống khói và cao độ đường bão hòa tại tim đập
    z_tim = float(np.interp(0.0, *zip(*sorted(zip(dbh[:, 0], dbh[:, 1]))))) if len(dbh) else np.nan
    out = dict(ten=ten, q=r['q_in'], q_out=r['q_out'], J_vao=J_vao, J_vao95=J_vao95, J_tiep=J_tiep,
               J_tiep95=J_tiep95, J_ra=J_ra, J_ra95=J_ra95, Jz_ra=Jz_ra, J_ck=J_ck, J_ck95=J_ck95,
               J_tb=J_tb, Ha=Ha, Hb=Hb, xa=xa, xb=xb, zmid=zmid, niter=bt.niter, z_tim=z_tim,
               dbh=dbh.tolist())
    return out


def ve_tham(mc, bt, kq, Zw, ten_file, tieu_de):
    xy, tri, mat = bt.xy, bt.tri, bt.mat
    Tg = mtri.Triangulation(xy[:, 0], xy[:, 1], tri)
    fig, ax = plt.subplots(figsize=(12.5, 4.6))
    VH.ve_vat_lieu(ax, xy, tri, mat, TEN_VL, xlim=(-110, 115), ylim=(35, 83))
    lv = np.arange(np.floor(mc.Zday), np.ceil(Zw) + 0.01, 1.0)
    cs = ax.tricontour(Tg, bt.H, levels=lv, colors='#1f4e9c', linewidths=0.5)
    ax.clabel(cs, fmt='%.0f', fontsize=6)
    d = np.array(kq['dbh'])
    if len(d):
        ax.plot(d[:, 0], d[:, 1], 'r-', lw=1.6, label='Đường bão hòa')
    xs, zs = mc.duong_mat()
    ax.plot(xs, zs, 'k-', lw=0.8)
    xw = [mc.x_trai, mc.mat_tren_inv_u(Zw) if hasattr(mc, 'mat_tren_inv_u') else mc.xcu - 3 * (mc.Zc - Zw)]
    ax.fill_between([mc.x_trai, xw[1]], [Zw, Zw], [mc.Zday, mc.Zday], color='#bfe3ff', alpha=0.35, lw=0)
    ax.plot([mc.x_trai, xw[1]], [Zw, Zw], color='#1f77b4', lw=1.0)
    ax.text(-108, Zw + 0.6, ('MN hồ %.2f' % Zw).replace('.', ','), color='#1f77b4', fontsize=8)
    if len(bt.seep_on):
        ax.plot(xy[bt.seep_on, 0], xy[bt.seep_on, 1], 'o', color='#2ca02c', ms=2.5, label='Mặt thấm ra')
    ax.set_xlabel('Khoảng cách từ tim đập x (m)'); ax.set_ylabel('Cao độ z (m)')
    _m, _k = ('%.2e' % kq['q']).split('e')
    ax.set_title(tieu_de + '  –  q = $%s\\times10^{%d}$ m³/s/m' % (_m.replace('.', '{,}'), int(_k)))
    ax.legend(loc='upper right', fontsize=7.5)
    return VH.luu(fig, ten_file)


def chay_tham(mc, Zw, ten, kdat=None, file=None, tieu_de=None):
    xy, tri, mat = T.tao_luoi(mc)
    bt, r = T.chay(mc, Zw, kdat=kdat, luoi=(xy, tri, mat))
    kq = phan_tich_tham(mc, bt, r, ten)
    kq['Zw'] = Zw; kq['nodes'] = len(xy); kq['elems'] = len(tri)
    if file:
        kq['hinh'] = ve_tham(mc, bt, kq, Zw, file, tieu_de or ten)
    return bt, kq


def ve_cung_truot(mc, bt, best, od, ten_file, tieu_de, Zw=None, Zw_ngoai=None):
    xy, tri, mat = bt.xy, bt.tri, bt.mat
    fig, ax = plt.subplots(figsize=(12.5, 4.6))
    VH.ve_vat_lieu(ax, xy, tri, mat, TEN_VL, xlim=(-110, 115), ylim=(35, 125), nhan=False)
    Tg = mtri.Triangulation(xy[:, 0], xy[:, 1], tri)
    psi = od.r.psi
    ax.contour(od.r.X, od.r.Z, np.nan_to_num(psi, nan=-1), levels=[0.0], colors='r', linewidths=1.2)
    K, (xc, zc, R) = best
    th = np.linspace(0, 2 * np.pi, 400)
    xx = xc + R * np.cos(th); zz = zc + R * np.sin(th)
    zt = mc.mat_tren(xx)
    m = zz <= zt
    ax.plot(np.where(m, xx, np.nan), np.where(m, zz, np.nan), 'k-', lw=1.6)
    ax.plot(xc, zc, 'k+', ms=8)
    for x_ in [xx[m].min(), xx[m].max()]:
        ax.plot([xc, x_], [zc, mc.mat_tren(x_)], 'k--', lw=0.5)
    if Zw_ngoai is not None:
        xw = mc.xcu - 3 * (mc.Zc - Zw_ngoai) if Zw_ngoai >= mc.Zcu else mc.x_u2 - 3.5 * (mc.Zcu - Zw_ngoai)
        ax.fill_between([mc.x_trai, xw], [Zw_ngoai] * 2, [mc.Zday] * 2, color='#bfe3ff', alpha=0.4, lw=0)
    xs, zs = mc.duong_mat(); ax.plot(xs, zs, 'k-', lw=0.8)
    ax.text(xc + 1, zc + 1, ('Kmin = %.3f' % K).replace('.', ','), fontsize=10, fontweight='bold')
    ax.set_title(tieu_de); ax.set_xlabel('x (m)'); ax.set_ylabel('z (m)')
    return VH.luu(fig, ten_file)


def mua_keo_dai(ras, mc, d_bh):
    """Bổ sung tầng bão hòa do mưa kéo dài trên mái hạ lưu: chiều sâu d_bh (m) song song mặt mái,
    dòng chảy song song mái -> cột nước áp lực = (d_bh − độ sâu)·cos²β."""
    X, Z = ras.X, ras.Z
    zt = mc.mat_tren(X[:, 0])[:, None]
    depth = zt - Z
    x = X[:, 0]
    beta = np.where((x > mc.xcd) & (x < mc.x_b1), np.arctan(1 / mc.m1d),
                    np.where((x > mc.x_b2) & (x < mc.x_l1), np.arctan(1 / mc.m2d), 0.0))[:, None]
    zone = (X > mc.xcd) & (X < mc.x_l1) & (depth >= 0) & (depth <= d_bh) & (ras.mat > 0)
    psi_rain = (d_bh - depth) * np.cos(beta) ** 2
    ras.psi = np.where(zone, np.maximum(np.nan_to_num(ras.psi, nan=-1), psi_rain), ras.psi)
    return ras


def chay_on_dinh(mc, bt, Zw_ngoai, chieu, kh=0.0, H_ap=None, mua=None, file=None, tieu_de='', vl='I'):
    ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25, H_ap=H_ap)
    if mua is not None:
        ras = mua_keo_dai(ras, mc, mua)
    od = O.OnDinh(ras, O.vat_lieu_on_dinh(vl), Zw=Zw_ngoai, kh=kh, chieu=chieu)
    if chieu > 0:
        best = od.tim((-10, 95), (70, 170), nx=28, nz=26, nr=26, x_gioi_han=(mc.xcu - 1, mc.x_phai - 2))
    else:
        best = od.tim((-95, 10), (70, 170), nx=28, nz=26, nr=26, x_gioi_han=(mc.x_trai + 2, mc.xcd + 1))
    if file:
        ve_cung_truot(mc, bt, best, od, file, tieu_de, Zw_ngoai=Zw_ngoai)
    K, ch = od.K_vong(*best[1], tra_chi_tiet=True)
    return best, od, ch


if __name__ == '__main__':
    t0 = time.time()
    trang_thai = {'HT': T.MatCat(Zc=S.Z_DINH_DAP_HT), 'SC': T.MatCat(Zc=S.Z_DINH_DAP_SC)}
    KQ['tham'] = {}; KQ['on_dinh'] = {}
    kd_tac = {2: S.DAT['DD-II']['k'], 4: S.DAT['DD-II']['k']}   # thiết bị thoát nước bị tắc
    cac_th = [('MNDBT', S.MNDBT, None), ('MNLTK', MNLTK, None), ('MNLKT', MNLKT, None),
              ('MNDBT_tac', S.MNDBT, kd_tac)]
    BT = {}
    for tt, mc in trang_thai.items():
        for ten, Zw, kd in cac_th:
            bt, kq = chay_tham(mc, Zw, ten, kdat=kd, file='tham_%s_%s.png' % (tt, ten),
                               tieu_de='%s – %s (%s)' % ('Hiện trạng' if tt == 'HT' else 'Sau sửa chữa',
                                                       {'MNDBT_tac': 'MNDBT, thoát nước bị tắc'}.get(ten, ten),
                                                       ('%.2f' % Zw).replace('.', ',')))
            kq.pop('dbh_full', None)
            KQ['tham']['%s_%s' % (tt, ten)] = {k: v for k, v in kq.items() if k != 'dbh'}
            KQ['tham']['%s_%s' % (tt, ten)]['dbh'] = kq['dbh'][::5]
            BT[(tt, ten)] = bt
            print(tt, ten, 'q=%.3e Jvao=%.3f/%.3f Jtiep=%.3f/%.3f Jra=%.3f/%.3f Jz=%.3f Jck=%.3f/%.3f Jtb=%.3f ztim=%.2f it=%d' %
                  (kq['q'], kq['J_vao'], kq['J_vao95'], kq['J_tiep'], kq['J_tiep95'], kq['J_ra'], kq['J_ra95'], kq['Jz_ra'],
                   kq['J_ck'], kq['J_ck95'], kq['J_tb'], kq['z_tim'], kq['niter']))
    print('tham xong', round(time.time() - t0, 1))
    json.dump(KQ, open('../ket_qua/tham.json', 'w'), ensure_ascii=False, indent=1, default=float)

    # ------------------------------------------------------------------ ỔN ĐỊNH
    KH = 0.05     # hệ số động đất ngang giả tĩnh (giả thiết, cần đối chiếu TCVN 9386:2012)
    D_MUA = {'HT': 3.0, 'SC': 1.0}   # chiều sâu tầng bão hòa do mưa kéo dài (giả thiết)
    TH = [  # (mã, mô tả, mái(+1 HL/-1 TL), mực nước thấm, mực nước ngoài mái TL, tổ hợp, kh, mưa, thấm ứng với H_ap)
        ('TH1', 'MNDBT, thấm ổn định', +1, 'MNDBT', None, 'CB', 0, False, None),
        ('TH2', 'MNLTK, thấm ổn định', +1, 'MNLTK', None, 'CB', 0, False, None),
        ('TH3', 'MNLKT, thấm ổn định', +1, 'MNLKT', None, 'DB', 0, False, None),
        ('TH4', 'MNDBT, thiết bị thoát nước bị tắc', +1, 'MNDBT_tac', None, 'DB', 0, False, None),
        ('TH5', 'MNDBT + động đất', +1, 'MNDBT', None, 'DB', KH, False, None),
        ('TH6', 'MNDBT + mưa kéo dài', +1, 'MNDBT', None, 'DB', 0, True, None),
        ('TH7', 'MNDBT, thấm ổn định (mái TL)', -1, 'MNDBT', S.MNDBT, 'CB', 0, False, None),
        ('TH8', 'Hạ MN từ MNDBT xuống MNC (đường bão hòa chưa kịp hạ)', -1, 'MNDBT', S.MNC, 'CB', 0, False, 'MNDBT'),
        ('TH9', 'Rút nhanh từ MNLTK xuống MNDBT', -1, 'MNLTK', S.MNDBT, 'DB', 0, False, 'MNLTK'),
        ('TH10', 'MNDBT + động đất (mái TL)', -1, 'MNDBT', S.MNDBT, 'DB', KH, False, None),
    ]
    K_CP = {'CB': 1.30, 'DB': 1.17}   # TCVN 8216:2018 Bảng 1, đập cấp II
    for tt, mc in trang_thai.items():
        for ma, mota, chieu, th_tham, Zng, toh, kh, mua, hap in TH:
            bt = BT[(tt, th_tham)]
            H_ap = None
            t1 = time.time()
            best, od, ch = chay_on_dinh(mc, bt, Zng, chieu, kh=kh, H_ap=H_ap, mua=(D_MUA[tt] if mua else None),
                                        file='od_%s_%s.png' % (tt, ma),
                                        tieu_de='%s – %s: %s' % ('Hiện trạng' if tt == 'HT' else 'Sau sửa chữa', ma, mota))
            K, (xc, zc, R) = best
            KQ['on_dinh']['%s_%s' % (tt, ma)] = dict(ma=ma, mo_ta=mota, mai='HL' if chieu > 0 else 'TL', K=K,
                                                     K_cp=K_CP[toh], to_hop=toh, xc=xc, zc=zc, R=R,
                                                     dat=bool(K >= K_CP[toh]), kh=kh)
            print(tt, ma, mota, 'K=%.3f [K]=%.2f tam(%.1f,%.1f) R=%.1f  %.1fs' % (K, K_CP[toh], xc, zc, R, time.time() - t1))
    json.dump(KQ, open('../ket_qua/tham_on_dinh.json', 'w'), ensure_ascii=False, indent=1, default=float)
    print('tong thoi gian', round(time.time() - t0, 1))
