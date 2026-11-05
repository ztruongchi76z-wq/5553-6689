# -*- coding: utf-8 -*-
"""
KIỂM CHỨNG TÌM CUNG TRƯỢT – BẢNG ỔN ĐỊNH SAU SỬA CHỮA (xuat_bang_on_dinh_docx.py)
Hai mặt cắt (lòng sông – chay_tham_on_dinh.py, phan_tich_bo_sung.py; sườn đồi – on_dinh_suon_doi.py), hiện trạng và sau
sửa chữa, cùng 12 trường hợp: lưới tâm × bán kính như bảng (lưới cơ sở), rồi tối ưu cục bộ Nelder – Mead từ 10 cung tốt
nhất của lưới ở các vùng tâm khác nhau (tâm cách nhau ≥ 4 m) – như kt_kiem_chung_12_10.py nhưng không dùng lưới lồng nhau.
Giữ phạm vi điểm vào – ra của cung như lưới cơ sở; đáy cung được phép sâu hơn lưới cơ sở (Zday − 8 m, bước tinh chỉnh
tới ≈ Zday − 10 m): lòng sông tới +36 m (như kt_kiem_chung_12_10.Z_DAY_MO_RONG), sườn đồi tới +49 m (đáy mô hình +48 m).
  python3 kt_cung_truot_sau_sc.py      → ket_qua/kiem_chung_cung_truot_sau_sc.json
"""
import json
import os
import time

import numpy as np
from scipy.optimize import minimize

import chay_tham_on_dinh as C
import kt_kiem_chung_12_10 as KC
import on_dinh as O
import on_dinh_suon_doi as OD
import phan_tich_bo_sung as PT
import so_lieu as S
import tham_fem as T
import tham_suon_doi as TS

DIR = os.path.dirname(os.path.abspath(__file__))
Z_DAY_SD = 49.0          # đáy cung thấp nhất khi kiểm chứng – mặt cắt sườn đồi (đáy mô hình +48 m)


def kiem_chung(od, mc, xr, zr, xg, nx, nz, nr, z_day):
    """K lưới cơ sở (như od.tim) và K sau tối ưu Nelder – Mead từ nhiều điểm xuất phát; z_day: đáy cung thấp nhất."""
    K0, c0 = od.tim(xr, zr, nx=nx, nz=nz, nr=nr, x_gioi_han=xg)

    def f(p):
        xc, zc, R = p
        if R < max(zc - float(mc.mat_tren(xc)) + 1.0, 3.0) or R > zc - z_day:
            return 10.0
        K, ex = od.K_vong(xc, zc, R)
        if ex is None or not (xg[0] <= ex[0] and ex[1] <= xg[1]):
            return 10.0
        return K

    ung_vien = []
    for xc in np.linspace(*xr, nx):
        for zc in np.linspace(*zr, nz):
            Rmin, Rmax = max(zc - float(mc.mat_tren(xc)) + 1.0, 3.0), zc - z_day
            if Rmax <= Rmin:
                continue
            for R in np.linspace(Rmin, Rmax, nr):
                K = f((xc, zc, R))
                if K < 10:
                    ung_vien.append((K, xc, zc, R))
    ung_vien.sort()
    xuat_phat = [c0]
    for K, xc, zc, R in ung_vien:
        if all(np.hypot(xc - a[0], zc - a[1]) >= 4.0 for a in xuat_phat):
            xuat_phat.append((xc, zc, R))
        if len(xuat_phat) >= 10:
            break
    tot = (K0, tuple(c0))
    for p0 in xuat_phat:
        r = minimize(f, np.array(p0, float), method='Nelder-Mead', options=dict(xatol=0.02, fatol=1e-6, maxiter=600))
        if r.fun < tot[0]:
            tot = (float(r.fun), tuple(float(v) for v in r.x))
    return dict(K_bang=float(K0), tam_bang=[float(v) for v in c0], K_toi_uu=tot[0], tam_toi_uu=list(tot[1]),
                giam_pt=float(100 * (1 - tot[0] / K0)))


def long_song(tt):
    mc = T.MatCat(Zc=S.Z_DINH_DAP_HT if tt == 'HT' else S.Z_DINH_DAP_SC)
    BT = {k: C.chay_tham(mc, KC.ZW[k], k, kdat=(KC.KD_TAC if k == 'MNDBT_tac' else None))[0] for k in KC.ZW}
    vl = O.vat_lieu_on_dinh('I')
    ds = [(ma, ch, BT[th], Zng, kh, mua, 0.0) for ma, _, ch, th, Zng, _, kh, mua in KC.TH]
    if tt == 'SC':
        ds += [('TH11', +1, BT['MNDBT'], None, 0.0, False, 20.0), ('TH12', -1, BT['MNDBT'], S.MNDBT, 0.0, False, 20.0)]
    out = {}
    for ma, ch, bt, Zng, kh, mua, q in ds:
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        if mua:
            ras = C.mua_keo_dai(ras, mc, {'HT': 3.0, 'SC': 1.0}[tt])
        od = (PT.OnDinhPhuTai(ras, vl, Zw=Zng, kh=kh, chieu=ch, q_s=q, x_qs=(mc.xcu, mc.xcd)) if q else
              O.OnDinh(ras, vl, Zw=Zng, kh=kh, chieu=ch))
        xr, zr, xg = KC.gioi_han(mc, ch)
        out[ma] = kiem_chung(od, mc, xr, zr, xg, 28, 26, 26, KC.Z_DAY_MO_RONG)
        print('lòng sông %s %-4s K bảng %.3f | tối ưu %.3f (giảm %.2f%%)'
              % (tt, ma, out[ma]['K_bang'], out[ma]['K_toi_uu'], out[ma]['giam_pt']), flush=True)
    return out


def suon_doi(tt):
    sau_sc = tt == 'SC'
    mesh = TS.luoi(sau_sc=sau_sc)
    mc = OD.MatCatSuon(sau_sc)
    vl = OD.vat_lieu(TS.NEN_CHINH)
    bts, out = {}, {}
    for ma, _, ch, Zt, kthem, Zng, _, kh, mua, q in OD.TH_C5:
        if q and not sau_sc:
            continue
        key = (Zt, kthem is not None)
        if key not in bts:
            bts[key] = TS.chay(Zt, mesh=mesh, k_them=kthem, sau_sc=sau_sc)[0]
        bt = bts[key]
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        if mua:
            ras = OD.mua_keo_dai(ras, OD.D_MUA[sau_sc])
        od = (PT.OnDinhPhuTai(ras, vl, Zw=Zng, kh=kh, chieu=ch, q_s=q, x_qs=(TS.X_DU, TS.X_DD)) if q else
              O.OnDinh(ras, vl, Zw=Zng, kh=kh, chieu=ch))
        if ch > 0:
            out[ma] = kiem_chung(od, mc, (60.0, 112.0), (79.0, 135.0), (40.0, 108.5), 26, 24, 24, Z_DAY_SD)
        else:
            out[ma] = kiem_chung(od, mc, (-2.0, 70.0), (79.0, 135.0), (1.5, 75.0), 28, 24, 24, Z_DAY_SD)
        print('sườn đồi %s %-4s K bảng %.3f | tối ưu %.3f (giảm %.2f%%)'
              % (tt, ma, out[ma]['K_bang'], out[ma]['K_toi_uu'], out[ma]['giam_pt']), flush=True)
    return out


if __name__ == '__main__':
    t0 = time.time()
    KQ = {'suon_doi': {tt: suon_doi(tt) for tt in ('HT', 'SC')}, 'long_song': {tt: long_song(tt) for tt in ('HT', 'SC')}}
    json.dump(KQ, open(os.path.join(DIR, 'ket_qua', 'kiem_chung_cung_truot_sau_sc.json'), 'w'), ensure_ascii=False,
              indent=1)
    print('Đã ghi ket_qua/kiem_chung_cung_truot_sau_sc.json  %.0f s' % (time.time() - t0))
