# -*- coding: utf-8 -*-
"""
ĐÁNH GIÁ BIẾN ĐỘNG LƯỢNG MƯA TỪ KHI CÔNG TRÌNH VẬN HÀNH VÀ KIỂM TRA KHẢ NĂNG THÁO LŨ KHI LŨ GIA TĂNG
(mục 4.6 – theo yêu cầu của giảng viên ngày 12/10/2026)

 (1) Chuỗi mưa năm lưu vực 1978 – 2024 [QTVH Bảng 2-9] và mưa 1 ngày lớn nhất X1max 1978 – 2020 [KTTV Bảng 3-13]:
     so sánh giai đoạn trước (1978 – 1999) và sau khi công trình vận hành (từ 2000); kiểm định xu thế Mann – Kendall
     (có hiệu chỉnh số liệu trùng) và độ dốc Sen.
 (2) Điều tiết lũ với tràn hiện trạng khi lưu lượng lũ đến tăng thêm k = 0 ÷ 30 % (giả thiết thu phóng toàn bộ tung độ
     đường quá trình, đỉnh và tổng lượng cùng tăng k), so sánh với MNLTK thiết kế ban đầu (+77,24 m, TK 1996), đỉnh đập
     sau sửa chữa (+78,85 m), đỉnh tường chắn sóng (+79,40 m) và điều kiện đỉnh đập ≥ MNLKT + 0,3 m.
 (3) Mức tăng lũ lớn nhất còn thỏa từng điều kiện; mực nước đón lũ cần hạ trước lũ để giữ MNLKT + 0,3 ≤ +78,85 m.
"""
import json
import math
import os

import numpy as np
import so_lieu as S
import dieu_tiet_lu as DT
import cao_trinh_dinh as CT

DIR = os.path.dirname(os.path.abspath(__file__))
M = json.load(open(os.path.join(DIR, 'du_lieu', 'mua_luu_vuc.json'), encoding='utf-8'))
NAM_VH = 2000                     # năm đưa công trình vào vận hành (hoàn thành 10/2000) [KTTV mục 1.4]
MNLTK_CU = 77.24                  # MNLTK thiết kế ban đầu (TK 1996, Q1% = 847 m3/s) [KTTV Bảng 1-2]
Z_DINH_SC, Z_TUONG = 78.85, 79.40


def mann_kendall(x):
    x = np.asarray(x, float)
    n = len(x)
    s = 0
    for i in range(n - 1):
        s += np.sign(x[i + 1:] - x[i]).sum()
    _, dem = np.unique(x, return_counts=True)
    var = (n * (n - 1) * (2 * n + 5) - sum(t * (t - 1) * (2 * t + 5) for t in dem)) / 18.0
    z = (s - 1) / math.sqrt(var) if s > 0 else ((s + 1) / math.sqrt(var) if s < 0 else 0.0)
    p = math.erfc(abs(z) / math.sqrt(2))          # hai phía
    doc = [(x[j] - x[i]) / (j - i) for i in range(n - 1) for j in range(i + 1, n)]
    return dict(S=int(s), Z=z, p=p, sen=float(np.median(doc)))


def thong_ke(nam, gia_tri, ten):
    nam = np.asarray(nam)
    v = np.asarray(gia_tri, float)
    a, b = v[nam < NAM_VH], v[nam >= NAM_VH]
    r = dict(ten=ten, nam_dau=int(nam.min()), nam_cuoi=int(nam.max()),
             tb_truoc=float(a.mean()), tb_sau=float(b.mean()), n_truoc=len(a), n_sau=len(b),
             cv_truoc=float(a.std(ddof=1) / a.mean()), cv_sau=float(b.std(ddof=1) / b.mean()),
             max_truoc=float(a.max()), max_sau=float(b.max()),
             thay_doi_pt=float(100 * (b.mean() / a.mean() - 1)),
             mk_toan_chuoi=mann_kendall(v), mk_sau=mann_kendall(b))
    # kiểm định Welch hai giai đoạn (xấp xỉ chuẩn)
    t = (b.mean() - a.mean()) / math.sqrt(a.var(ddof=1) / len(a) + b.var(ddof=1) / len(b))
    r['t_welch'] = float(t)
    r['p_welch'] = float(math.erfc(abs(t) / math.sqrt(2)))
    return r


def zmax(Qden, k=0.0, Z0=S.MNDBT, ket_van=False):
    tc, sc = DT.tran_hien_trang()
    r = DT.dieu_tiet([q * (1 + k) for q in Qden], tc, sc, Z0=Z0, ket_van=ket_van)
    return r


def k_gioi_han(Qden, Z_gh, lo=0.0, hi=0.6):
    """Mức tăng lũ lớn nhất để Zmax ≤ Z_gh (chia đôi)."""
    if zmax(Qden, lo)['Zmax'] > Z_gh:
        return lo
    for _ in range(18):
        mid = 0.5 * (lo + hi)
        lo, hi = (mid, hi) if zmax(Qden, mid)['Zmax'] <= Z_gh else (lo, mid)
    return lo


def Z0_can(Qden, k, Z_gh=Z_DINH_SC - 0.30):
    """Mực nước đón lũ (giữ bằng cửa van tràn sự cố) để Zmax ≤ Z_gh."""
    if zmax(Qden, k)['Zmax'] <= Z_gh:
        return S.MNDBT
    lo, hi = 71.5, S.MNDBT
    for _ in range(18):
        mid = 0.5 * (lo + hi)
        lo, hi = (mid, hi) if zmax(Qden, k, Z0=mid)['Zmax'] <= Z_gh else (lo, mid)
    return lo


if __name__ == '__main__':
    out = {}
    nam = [int(y) for y in M['mua_nam']]
    X = np.array([M['mua_nam'][str(y)] for y in nam])
    out['mua_nam'] = thong_ke(nam, X[:, 12], 'Lượng mưa năm (mm)')
    out['mua_mua'] = thong_ke(nam, X[:, 4:10].sum(1), 'Lượng mưa mùa mưa tháng 5 – 10 (mm)')
    out['mua_thang_max'] = thong_ke(nam, X[:, :12].max(1), 'Lượng mưa tháng lớn nhất trong năm (mm)')
    nx = [int(y) for y in M['x1max']]
    out['x1max'] = thong_ke(nx, [M['x1max'][str(y)] for y in nx], 'Mưa 1 ngày lớn nhất X1max (mm)')
    for k, r in out.items():
        print('%-45s %d–%d: TB trước %.1f (n=%d, Cv %.2f) | sau %.1f (n=%d, Cv %.2f) | %+.1f%% | MK Z=%.2f p=%.2f Sen %+.2f/năm'
              ' | MK từ 2000: Z=%.2f p=%.2f Sen %+.2f | Welch p=%.2f'
              % (r['ten'], r['nam_dau'], r['nam_cuoi'], r['tb_truoc'], r['n_truoc'], r['cv_truoc'], r['tb_sau'],
                 r['n_sau'], r['cv_sau'], r['thay_doi_pt'], r['mk_toan_chuoi']['Z'], r['mk_toan_chuoi']['p'],
                 r['mk_toan_chuoi']['sen'], r['mk_sau']['Z'], r['mk_sau']['p'], r['mk_sau']['sen'], r['p_welch']))
    x = np.array([M['x1max'][str(y)] for y in nx])
    out['x1max_dem'] = dict(truoc_150=int((x[np.array(nx) < NAM_VH] > 150).sum()),
                            sau_150=int((x[np.array(nx) >= NAM_VH] > 150).sum()),
                            truoc_200=int((x[np.array(nx) < NAM_VH] > 200).sum()),
                            sau_200=int((x[np.array(nx) >= NAM_VH] > 200).sum()))
    print('Số năm X1max > 150 / > 200 mm: trước', out['x1max_dem']['truoc_150'], out['x1max_dem']['truoc_200'],
          '| sau', out['x1max_dem']['sau_150'], out['x1max_dem']['sau_200'])

    # ---------------------------------------------------------------- (2) điều tiết lũ khi lũ tăng
    kq = []
    for nguon, Qd in (('BCNCKT', S.Q_KTTV), ('QTVH', S.Q_QTVH)):
        for p in ('1%', '0.2%'):
            for k in (0.0, 0.10, 0.20, 0.30):
                r = zmax(Qd[p], k)
                d = dict(nguon=nguon, P=p, k=k, Qmax=max(Qd[p]) * (1 + k), Zmax=float(r['Zmax']),
                         qxa=float(r['qmax']), q_tc=float(r['q_tc_max']), q_sc=float(r['q_sc_max']))
                kq.append(d)
                print('%-6s P=%-4s +%2d%%: Qmax %6.0f  Zmax %.2f  q xả %5.1f (tc %5.1f, sc %5.1f)  so MNLTK cũ %+.2f  so đỉnh %+.2f'
                      % (nguon, p, 100 * k, d['Qmax'], d['Zmax'], d['qxa'], d['q_tc'], d['q_sc'],
                         d['Zmax'] - MNLTK_CU, d['Zmax'] - Z_DINH_SC))
    out['dieu_tiet'] = kq
    # (3) mức tăng giới hạn (lũ BCNCKT – cơ sở thiết kế)
    # độ dềnh + sóng leo + a' với gió Đông Bắc P = 50 % (góc tới 38,8°), TCVN 8421:2010 (A.3.3, Bảng 9) – Z2 = MNLTK + DZ
    k_ne = S.GIO_HUONG.index('NE')
    r_ne = CT.tinh_truong_hop(77.63, S.GIO[50][k_ne], 3025.0, alpha=38.8)
    DZ_SONG_NE = r_ne['dh'] + r_ne['hsl'] + 0.50
    out['DZ_song_NE'] = DZ_SONG_NE
    print('Gió NE P = 50 %%: Vw = %.2f m/s, Δh + hsl + a\' = %.3f m -> Z2 = %.2f' % (r_ne['V'], DZ_SONG_NE, 77.63 + DZ_SONG_NE))
    gh = dict(
        k_P1_MNLTKcu=k_gioi_han(S.Q_KTTV['1%'], MNLTK_CU),
        k_P02_dinh_tru_03=k_gioi_han(S.Q_KTTV['0.2%'], Z_DINH_SC - 0.30),
        k_P02_dinh=k_gioi_han(S.Q_KTTV['0.2%'], Z_DINH_SC),
        k_P02_tuong=k_gioi_han(S.Q_KTTV['0.2%'], Z_TUONG),
        k_P1_dinh=k_gioi_han(S.Q_KTTV['1%'], Z_DINH_SC),
        k_P1_song_NE=k_gioi_han(S.Q_KTTV['1%'], Z_TUONG - DZ_SONG_NE),
    )
    out['k_gioi_han'] = gh
    for k_, v in gh.items():
        print('Mức tăng lũ giới hạn %-20s %.1f%%' % (k_, 100 * v))
    z0 = {}
    for k in (0.10, 0.20, 0.30):
        z0['%d' % round(100 * k)] = Z0_can(S.Q_KTTV['0.2%'], k)
        print('Lũ BCNCKT P=0,2%% +%d%%: mực nước đón lũ cần ≤ %.2f m để MNLKT + 0,3 ≤ +78,85' % (100 * k, z0['%d' % round(100 * k)]))
    out['Z0_don_lu'] = z0
    out['MNLTK_cu'] = MNLTK_CU
    json.dump(out, open(os.path.join(DIR, 'ket_qua', 'kt_mua_lu_bien_doi.json'), 'w'), ensure_ascii=False,
              indent=1, default=float)
