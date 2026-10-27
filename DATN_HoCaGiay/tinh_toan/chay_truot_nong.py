# -*- coding: utf-8 -*-
"""Cung trượt nông trong mái hạ lưu (không cắt qua nền) – so sánh hiện trạng / sau sửa chữa khi mưa kéo dài."""
import json, numpy as np
import so_lieu as S, tham_fem as T, on_dinh as O, chay_tham_on_dinh as C

out = {}
for tt, Zc in [('HT', S.Z_DINH_DAP_HT), ('SC', S.Z_DINH_DAP_SC)]:
    mc = T.MatCat(Zc=Zc)
    xy, tri, mat = T.tao_luoi(mc)
    bt, r = T.chay(mc, S.MNDBT, luoi=(xy, tri, mat))
    for d in [0.0, 0.5, 1.0, 2.0, 3.0, 4.0]:
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        if d > 0:
            ras = C.mua_keo_dai(ras, mc, d)
        od = O.OnDinh(ras, O.vat_lieu_on_dinh('I'), Zw=None, chieu=+1)
        best = od.tim((0, 80), (62, 140), nx=26, nz=24, nr=22, x_gioi_han=(mc.xcd - 1, mc.x_l3 + 3),
                      z_day=mc.Zday + 0.5)
        K, (xc, zc, R) = best
        out['%s_%.1f' % (tt, d)] = dict(K=K, xc=xc, zc=zc, R=R, d=d)
        print(tt, 'd_bh=%.1f m' % d, 'K=%.3f' % K, 'tam (%.1f, %.1f) R=%.1f' % (xc, zc, R))
        if d in (1.0, 3.0):
            C.ve_cung_truot(mc, bt, best, od, 'od_nong_%s_mua%.0f.png' % (tt, d),
                            '%s – mái hạ lưu, mưa kéo dài (tầng bão hòa %.1f m), cung trượt trong thân đập'
                            % ('Hiện trạng' if tt == 'HT' else 'Sau sửa chữa', d))
json.dump(out, open('../ket_qua/truot_nong.json', 'w'), indent=1, default=float)
