# -*- coding: utf-8 -*-
"""
KIỂM CHỨNG BỔ SUNG THẤM – ỔN ĐỊNH (tiếp thu góp ý rà soát 11/10 – 12/10/2026)
  python3 kt_kiem_chung_12_10.py luoi      (a) tìm cung trượt trên lưới tâm LỒNG NHAU (55 × 51 tâm, 51 bán kính – chứa toàn
                                              bộ lưới cơ sở 28 × 26 × 26) + tối ưu cục bộ Nelder – Mead từ nhiều điểm
                                              xuất phát; Spencer cho cung nguy hiểm của mọi trường hợp (kể cả TH8)
  python3 kt_kiem_chung_12_10.py anis      (b) thấm bất đẳng hướng k_h/k_v = 5 cho các khối đất đắp (TCVN 8216:2018
                                              điều 11.2.2) – lưu lượng, gradien, đường bão hòa, ổn định mái
  python3 kt_kiem_chung_12_10.py gradien   (c) gradien lớn nhất (không chỉ phân vị 95 %) và vị trí, lưới cơ sở và lưới mịn
  python3 kt_kiem_chung_12_10.py cuongdo   (d) độ nhạy chỉ tiêu chống cắt của lớp chỉ có trị tiêu chuẩn (ĐĐ-II, lớp 4, lớp 5)
Mặt cắt sau sửa chữa đỉnh +78,85 m; số liệu và quy ước như chay_tham_on_dinh.py, phan_tich_bo_sung.py.
"""
import json
import os
import sys
import time

import numpy as np
import scipy.sparse as sp
from scipy.optimize import minimize

import so_lieu as S
import tham_fem as T
import on_dinh as O
import chay_tham_on_dinh as C
import phan_tich_bo_sung as P

DIR = os.path.dirname(os.path.abspath(__file__))
MNLTK, MNLKT = C.MNLTK, C.MNLKT
KD_TAC = {2: S.DAT['DD-II']['k'], 4: S.DAT['DD-II']['k']}
TH = [  # mã, mô tả, chiều, thấm ứng với, MN ngoài mái TL, tổ hợp, kh, mưa
    ('TH1', 'MNDBT, thấm ổn định', +1, 'MNDBT', None, 'CB', 0.0, False),
    ('TH2', 'MNLTK, thấm ổn định', +1, 'MNLTK', None, 'CB', 0.0, False),
    ('TH3', 'MNLKT, thấm ổn định', +1, 'MNLKT', None, 'DB', 0.0, False),
    ('TH4', 'MNDBT, thoát nước bị tắc', +1, 'MNDBT_tac', None, 'DB', 0.0, False),
    ('TH5', 'MNDBT + động đất', +1, 'MNDBT', None, 'DB', 0.05, False),
    ('TH6', 'MNDBT + mưa kéo dài', +1, 'MNDBT', None, 'DB', 0.0, True),
    ('TH7', 'MNDBT, mái TL', -1, 'MNDBT', S.MNDBT, 'CB', 0.0, False),
    ('TH8', 'Hạ MNDBT → MNC', -1, 'MNDBT', S.MNC, 'CB', 0.0, False),
    ('TH9', 'Rút nhanh MNLTK → MNDBT', -1, 'MNLTK', S.MNDBT, 'DB', 0.0, False),
    ('TH10', 'MNDBT + động đất (mái TL)', -1, 'MNDBT', S.MNDBT, 'DB', 0.05, False),
]
ZW = {'MNDBT': S.MNDBT, 'MNLTK': MNLTK, 'MNLKT': MNLKT, 'MNDBT_tac': S.MNDBT}


# --------------------------------------------------------------------------- thấm bất đẳng hướng
class ThamBatDangHuong(T.BaiToanTham):
    """k_x = r·k, k_z = k trong các vật liệu đất đắp (1 ĐĐ-I, 3 ĐĐ-II, 6 chân khay); vật liệu khác đẳng hướng.
    k (thí nghiệm mẫu) coi là k_v; k_h = r·k_v – thiên lớn về lưu lượng và đường bão hòa phía hạ lưu."""
    def __init__(self, xy, tri, mat, kdat=None, alpha=1.0, kmin=1e-4, r_hv=5.0, vl=(1, 3, 6)):
        super().__init__(xy, tri, mat, kdat=kdat, alpha=alpha, kmin=kmin)
        self.fx = np.where(np.isin(mat, vl), r_hv, 1.0)

    def lap_ma_tran(self, H):
        psi_e = H[self.tri].mean(1) - self.zc
        ke = self.ksat * self.kr(psi_e)
        Ke = (ke * self.area)[:, None, None] * (self.fx[:, None, None] * self.bx[:, :, None] * self.bx[:, None, :]
                                                + self.bz[:, :, None] * self.bz[:, None, :])
        rows = np.repeat(self.tri, 3, axis=1).ravel()
        cols = np.tile(self.tri, (1, 3)).ravel()
        K = sp.coo_matrix((Ke.ravel(), (rows, cols)), shape=(self.n, self.n)).tocsr()
        return K, ke


def chay_tham(mc, Zw, ten, kdat=None, r_hv=1.0, luoi=None):
    xy, tri, mat = luoi if luoi is not None else T.tao_luoi(mc)
    bt = ThamBatDangHuong(xy, tri, mat, kdat=kdat, r_hv=r_hv)
    up, dn, seep = T.bien(mc, xy, Zw)
    H0 = np.interp(xy[:, 0], [mc.x_u3, mc.x_l3], [Zw, mc.Zday])
    bt.giai(up, np.full(len(up), Zw), seep, H0=H0)
    r = dict(q_in=bt.luu_luong(up), q_out=-bt.luu_luong(bt.seep_on), up=up, seep=bt.seep_on)
    kq = C.phan_tich_tham(mc, bt, r, ten)
    kq['Zw'] = Zw
    return bt, kq


# --------------------------------------------------------------------------- tìm cung trượt kiểm chứng
def gioi_han(mc, chieu):
    return ((-10, 95), (70, 170), (mc.xcu - 1, mc.x_phai - 2)) if chieu > 0 else \
        ((-95, 10), (70, 170), (mc.x_trai + 2, mc.xcd + 1))


Z_DAY_MO_RONG = 36.0     # đáy cung trượt được phép xuống tới gần đáy mô hình (+35 m) – kể cả qua lớp I, lớp II


def ham_K(od, mc, xg, z_day=Z_DAY_MO_RONG):
    def f(p):
        xc, zc, R = p
        if R < max(zc - mc.mat_tren(xc) + 1.0, 3.0) or R > zc - z_day:
            return 10.0
        K, ex = od.K_vong(xc, zc, R)
        if ex is None or not (xg[0] <= ex[0] and ex[1] <= xg[1]):
            return 10.0
        return K
    return f


def tim_kiem_chung(od, mc, chieu):
    xr, zr, xg = gioi_han(mc, chieu)
    t0 = time.time()
    co_so = od.tim(xr, zr, nx=28, nz=26, nr=26, x_gioi_han=xg)                       # như đồ án
    # lưới cơ sở nhưng cho phép đáy cung xuống tới +36 m (đồ án: lưới đến +46 m, bước tinh chỉnh tới ≈ +44 m)
    mo_rong = od.tim(xr, zr, nx=28, nz=26, nr=26, x_gioi_han=xg, z_day=Z_DAY_MO_RONG)
    # lưới lồng nhau: bước tâm và bước bán kính bằng một nửa lưới (mở rộng) -> chứa mọi điểm của lưới đó
    xcs = np.linspace(*xr, 55); zcs = np.linspace(*zr, 51)
    f = ham_K(od, mc, xg)
    ung_vien = []
    for xc in xcs:
        for zc in zcs:
            Rmin = max(zc - mc.mat_tren(xc) + 1.0, 3.0); Rmax = zc - Z_DAY_MO_RONG
            if Rmax <= Rmin:
                continue
            for R in np.linspace(Rmin, Rmax, 51):
                K = f((xc, zc, R))
                if K < 10:
                    ung_vien.append((K, xc, zc, R))
    ung_vien.sort()
    K_luoi = ung_vien[0][0]
    # chọn tối đa 8 điểm xuất phát cách nhau ≥ 6 m (các vùng cực tiểu khác nhau) + điểm tốt nhất của lưới cơ sở
    xuat_phat = [co_so[1], mo_rong[1]]
    for K, xc, zc, R in ung_vien[:4000]:
        if all(np.hypot(xc - a[0], zc - a[1]) > 6.0 for a in xuat_phat):
            xuat_phat.append((xc, zc, R))
        if len(xuat_phat) >= 10:
            break
    tot = min((co_so[0], co_so[1]), (mo_rong[0], mo_rong[1]), key=lambda a: a[0])
    for p0 in xuat_phat:
        r = minimize(f, np.array(p0, float), method='Nelder-Mead',
                     options=dict(xatol=0.02, fatol=1e-6, maxiter=600, initial_simplex=None))
        if r.fun < tot[0]:
            tot = (float(r.fun), tuple(float(v) for v in r.x))
    return dict(K_co_so=float(co_so[0]), tam_co_so=[float(v) for v in co_so[1]], K_mo_rong=float(mo_rong[0]),
                tam_mo_rong=[float(v) for v in mo_rong[1]], K_luoi_long=float(K_luoi),
                K_toi_uu=tot[0], tam_toi_uu=list(tot[1]), so_diem_xuat_phat=len(xuat_phat),
                giam_pt=float(100 * (1 - tot[0] / co_so[0])), giay=time.time() - t0)


def ras_od(mc, bt, Zng, chieu, kh, mua, vl=None):
    ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
    if mua:
        ras = C.mua_keo_dai(ras, mc, 1.0)
    return O.OnDinh(ras, vl or O.vat_lieu_on_dinh('I'), Zw=Zng, kh=kh, chieu=chieu)


def che_do_luoi():
    mc = T.MatCat(Zc=S.Z_DINH_DAP_SC)
    BT = {k: C.chay_tham(mc, ZW[k], k, kdat=(KD_TAC if k == 'MNDBT_tac' else None))[0] for k in ZW}
    out = {}
    for ma, mota, ch, th, Zng, toh, kh, mua in TH:
        od = ras_od(mc, BT[th], Zng, ch, kh, mua)
        r = tim_kiem_chung(od, mc, ch)
        s = P.chi_tiet(od, *r['tam_toi_uu'])
        sp_ = P.spencer(s)
        r['spencer'] = sp_['F']; r['spencer_theta'] = sp_['theta']; r['bishop_kt'] = P.bishop(s)
        out[ma] = r
        xc, zc, R = r['tam_toi_uu']
        print('%-4s K cơ sở %.3f (đáy cung %.1f) | mở rộng %.3f | lưới lồng nhau %.3f | tối ưu %.3f (giảm %.2f%%, %d điểm XP;'
              ' tâm %.1f; %.1f R %.1f đáy %.1f) | Spencer %.3f θ=%.1f° | %.0fs'
              % (ma, r['K_co_so'], r['tam_co_so'][1] - r['tam_co_so'][2], r['K_mo_rong'], r['K_luoi_long'], r['K_toi_uu'],
                 r['giam_pt'], r['so_diem_xuat_phat'], xc, zc, R, zc - R, r['spencer'], r['spencer_theta'], r['giay']),
              flush=True)
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kiem_chung_luoi_tam.json'), 'w'), ensure_ascii=False, indent=1,
              default=float)


def che_do_anis():
    out = {}
    for tt, Zc in (('HT', S.Z_DINH_DAP_HT), ('SC', S.Z_DINH_DAP_SC)):
        mc = T.MatCat(Zc=Zc)
        luoi = T.tao_luoi(mc)
        BT = {}
        for ten in ('MNDBT', 'MNLTK', 'MNLKT', 'MNDBT_tac'):
            for r_hv in (1.0, 5.0):
                bt, kq = chay_tham(mc, ZW[ten], ten, kdat=(KD_TAC if ten == 'MNDBT_tac' else None), r_hv=r_hv, luoi=luoi)
                BT[(ten, r_hv)] = bt
                d = {k: kq[k] for k in ('q', 'q_out', 'J_tiep', 'J_tiep95', 'J_ra', 'J_ra95', 'J_ck', 'J_ck95', 'z_tim')}
                out['%s_%s_r%d' % (tt, ten, r_hv)] = d
                print('%s %-9s kh/kv=%d q=%.3e (ra %.3e) Jtiep %.2f/%.2f Jra %.3f/%.3f Jck %.2f/%.2f ztim %.2f'
                      % (tt, ten, r_hv, kq['q'], kq['q_out'], kq['J_tiep'], kq['J_tiep95'], kq['J_ra'], kq['J_ra95'],
                         kq['J_ck'], kq['J_ck95'], kq['z_tim']), flush=True)
        if tt == 'SC':
            for ma, mota, ch, th, Zng, toh, kh, mua in TH:
                for r_hv in (1.0, 5.0):
                    od = ras_od(mc, BT[(th, r_hv)], Zng, ch, kh, mua)
                    xr, zr, xg = gioi_han(mc, ch)
                    K, tam = od.tim(xr, zr, nx=28, nz=26, nr=26, x_gioi_han=xg)
                    out['SC_%s_r%d_K' % (ma, r_hv)] = float(K)
                    print('SC %-4s kh/kv=%d K=%.3f' % (ma, r_hv, K), flush=True)
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kiem_chung_bat_dang_huong.json'), 'w'), ensure_ascii=False,
              indent=1, default=float)


def gradien_max(mc, bt):
    """Gradien lớn nhất trong vùng bão hòa tại tiếp xúc khối đất – ống khói và tại cửa ra, kèm vị trí."""
    xy, tri, mat = bt.xy, bt.tri, bt.mat
    gx, gz, J = bt.gradient()
    psi_e = bt.H[tri].mean(1) - xy[tri, 1].mean(1)
    bh = psi_e >= -0.05
    drain = np.unique(tri[np.isin(mat, [2, 4])])
    e_core = np.where(np.isin(mat, [1, 6]))[0]
    tiep = e_core[np.isin(tri[e_core], drain).any(1)]
    tiep = tiep[bh[tiep]]
    i = tiep[np.argmax(J[tiep])]
    ra = np.where(np.isin(tri, bt.seep_on).any(1))[0]
    ra = ra[bh[ra]]
    j = ra[np.argmax(J[ra])] if len(ra) else None
    # gradien lớn nhất của tiếp xúc khi bỏ các phần tử ở góc (trong phạm vi 1 m quanh điểm gãy hình học)
    xg, zg = bt.xc[tiep], bt.zc[tiep]
    return dict(Jtiep_max=float(J[i]), x_tiep=float(bt.xc[i]), z_tiep=float(bt.zc[i]),
                Jtiep_99=float(np.percentile(J[tiep], 99)), Jtiep_95=float(np.percentile(J[tiep], 95)),
                Jra_max=float(J[j]) if j is not None else None, x_ra=float(bt.xc[j]) if j is not None else None,
                z_ra=float(bt.zc[j]) if j is not None else None, n_tiep=int(len(tiep)))


def che_do_gradien():
    out = {}
    for tt, Zc in (('HT', S.Z_DINH_DAP_HT), ('SC', S.Z_DINH_DAP_SC)):
        for ten in ('MNDBT', 'MNLTK', 'MNLKT', 'MNDBT_tac'):
            mc = T.MatCat(Zc=Zc)
            bt, kq = C.chay_tham(mc, ZW[ten], ten, kdat=(KD_TAC if ten == 'MNDBT_tac' else None))
            g = gradien_max(mc, bt)
            out['%s_%s' % (tt, ten)] = g
            print('%s %-9s J tiếp xúc max %.2f tại (x=%.1f; z=%.1f), P99 %.2f, P95 %.2f | J ra max %s tại x=%s z=%s'
                  % (tt, ten, g['Jtiep_max'], g['x_tiep'], g['z_tiep'], g['Jtiep_99'], g['Jtiep_95'],
                     '%.3f' % g['Jra_max'] if g['Jra_max'] is not None else '-',
                     '%.1f' % g['x_ra'] if g['x_ra'] is not None else '-', '%.1f' % g['z_ra'] if g['z_ra'] is not None else '-'),
                  flush=True)
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kiem_chung_gradien.json'), 'w'), ensure_ascii=False, indent=1,
              default=float)


def che_do_cuongdo():
    """Lớp chỉ có trị tiêu chuẩn (ĐĐ-II id 3, lớp 4 id 7): giảm c 12 %, φ 0,7° như tỷ lệ TTGH I/tiêu chuẩn của ĐĐ-I,
    lớp I trong hồ sơ; lớp 5 không có trong mặt cắt lòng sông."""
    mc = T.MatCat(Zc=S.Z_DINH_DAP_SC)
    BT = {k: C.chay_tham(mc, ZW[k], k, kdat=(KD_TAC if k == 'MNDBT_tac' else None))[0] for k in ('MNDBT', 'MNDBT_tac')}
    vl0 = O.vat_lieu_on_dinh('I')
    vl1 = dict(vl0)
    for k in (3, 7):
        g, gs, c, phi = vl0[k]
        vl1[k] = (g, gs, 0.88 * c, phi - 0.7)
    out = {}
    for ma, mota, ch, th, Zng, toh, kh, mua in TH:
        if ma not in ('TH1', 'TH4', 'TH5', 'TH6'):
            continue
        for nhan, vl in (('goc', vl0), ('giam', vl1)):
            od = ras_od(mc, BT[th], Zng, ch, kh, mua, vl=vl)
            xr, zr, xg = gioi_han(mc, ch)
            K, tam = od.tim(xr, zr, nx=28, nz=26, nr=26, x_gioi_han=xg)
            out['%s_%s' % (ma, nhan)] = float(K)
            print('%s %s K = %.3f' % (ma, nhan, K), flush=True)
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kiem_chung_cuong_do.json'), 'w'), ensure_ascii=False, indent=1,
              default=float)


if __name__ == '__main__':
    {'luoi': che_do_luoi, 'anis': che_do_anis, 'gradien': che_do_gradien, 'cuongdo': che_do_cuongdo}[sys.argv[1]]()
