# -*- coding: utf-8 -*-
"""Kiểm chứng mô hình thấm – ổn định (mặt cắt sau sửa chữa, đỉnh +78,85):
 (a) hội tụ lưới: diện tích phần tử lớn nhất × 2; × 1; × 0,5; × 0,25 (MNLKT);
 (b) độ nhạy tham số Gardner α = 0,5; 1; 2 m⁻¹ (MNLKT);
 (c) độ nhạy hệ số động đất k_h (TH5 mái HL, TH10 mái TL) và k_h làm K = [K] = 1,17."""
import json, time, numpy as np
import so_lieu as S, tham_fem as T, on_dinh as O, chay_tham_on_dinh as C
KQ = {}
mc = T.MatCat(Zc=S.Z_DINH_DAP_SC)
goc = T.MatCat.pslg
def pslg_he_so(f):
    def p(self):
        V, segs, reg = goc(self)
        return V, segs, [(a, b, c, d * f) for a, b, c, d in reg]
    return p
for f in (2.0, 1.0, 0.5, 0.25):
    T.MatCat.pslg = pslg_he_so(f)
    t0 = time.time()
    bt, kq = C.chay_tham(mc, 78.54, 'MNLKT_f%.2f' % f)
    KQ['luoi_%.2f' % f] = dict(nodes=kq['nodes'], q=kq['q'], ztim=kq['z_tim'], Jtiep95=kq['J_tiep95'], Jra95=kq['J_ra95'],
                               Jck95=kq['J_ck95'])
    print('lưới ×%.2f nút %d q=%.4e z_tim=%.2f Jtiếp95=%.3f Jra95=%.3f %.0fs' % (f, kq['nodes'], kq['q'], kq['z_tim'],
          kq['J_tiep95'], kq['J_ra95'], time.time() - t0), flush=True)
T.MatCat.pslg = goc
for a in (0.5, 1.0, 2.0):
    xy, tri, mat = T.tao_luoi(mc)
    bt, r = T.chay(mc, 78.54, alpha=a, luoi=(xy, tri, mat))
    kq = C.phan_tich_tham(mc, bt, r, 'alpha')
    KQ['alpha_%.1f' % a] = dict(q=kq['q'], ztim=kq['z_tim'], Jtiep95=kq['J_tiep95'], Jra95=kq['J_ra95'])
    print('α=%.1f q=%.4e z_tim=%.2f Jtiếp95=%.3f Jra95=%.3f' % (a, kq['q'], kq['z_tim'], kq['J_tiep95'], kq['J_ra95']), flush=True)
bt, _ = C.chay_tham(mc, S.MNDBT, 'MNDBT')
vl = O.vat_lieu_on_dinh('I')
for kh in (0.0, 0.05, 0.075, 0.10, 0.125):
    for ch, Zng, ma in ((+1, None, 'TH5'), (-1, S.MNDBT, 'TH10')):
        ras = O.LuoiRaster(bt.xy, bt.tri, bt.mat, bt.H, mc, dx=0.25)
        od = O.OnDinh(ras, vl, Zw=Zng, kh=kh, chieu=ch)
        if ch > 0:
            best = od.tim((-10, 95), (70, 170), nx=28, nz=26, nr=26, x_gioi_han=(mc.xcu - 1, mc.x_phai - 2))
        else:
            best = od.tim((-95, 10), (70, 170), nx=28, nz=26, nr=26, x_gioi_han=(mc.x_trai + 2, mc.xcd + 1))
        KQ['%s_kh%.3f' % (ma, kh)] = dict(K=best[0], tam=best[1])
        print(ma, 'kh=%.3f K=%.3f' % (kh, best[0]), flush=True)
json.dump(KQ, open('../ket_qua/kt_hoi_tu.json', 'w'), ensure_ascii=False, indent=1, default=float)
print('xong')
