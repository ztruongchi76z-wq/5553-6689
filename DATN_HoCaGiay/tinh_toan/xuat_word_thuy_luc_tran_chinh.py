# -*- coding: utf-8 -*-
"""
XUẤT WORD: TÍNH TOÁN THỦY LỰC TRÀN CHÍNH THEO TRÌNH TỰ FILE MẪU “THIẾT KẾ ĐƯỜNG TRÀN THÁO LŨ” (SV gửi 05/11/2026)
Kiểm tra khả năng tháo (lưu tốc tới gần, ε, m, σn, so với kết quả điều tiết lũ – file DieuTietLu_Potapov_f1f2) và thủy
lực dốc nước (độ sâu đầu dốc, h_k, i_k, h_0, đường mặt nước 5 cấp lưu lượng, không xói, tường bên, bản đáy). Tính ở
thuy_luc_tran_chinh.py. Phần tiêu năng mũi phun của mẫu chưa tính (cần số liệu mũi phun, kênh xả hạ lưu).
  python3 xuat_word_thuy_luc_tran_chinh.py [đường dẫn .docx ra]
Mặc định: ../TAI_VE/ThuyLuc_TranChinh_TheoMau.docx (hình: ../TAI_VE/DuongMatNuoc_DocTranChinh.png)
"""
import math
import os
import sys

import docx
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

import thuy_luc_tran_chinh as T
from omml_docx import bang, can, gach, ngoac, phan_so, phuong_trinh, sub, sub_sup, sup, t
from xuat_bang_tham_docx import chu, doan, so

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'ThuyLuc_TranChinh_TheoMau.docx')
HINH = os.path.join(os.path.dirname(RA), 'DuongMatNuoc_DocTranChinh.png')
TC = T.TC


class Dem:
    """Đánh số công thức (4-…) và bảng (Bảng 4.…) liên tục trong mục."""
    def __init__(self, ct=1, bang_=3):
        self.ct, self.b = ct, bang_

    def c(self):
        s = '4-%d' % self.ct
        self.ct += 1
        return s

    def bg(self):
        s = 'Bảng 4.%d' % self.b
        self.b += 1
        return s


def tieu_de(d, chu_, cap):
    if cap == 1:
        return doan(d, chu_, co=13, dam=True, truoc=12, sau=6, can=WD_ALIGN_PARAGRAPH.LEFT)
    if cap == 2:
        return doan(d, chu_, co=13, dam=True, nghieng=True, truoc=10, sau=4, can=WD_ALIGN_PARAGRAPH.LEFT)
    return doan(d, chu_, co=13, nghieng=True, truoc=8, sau=2, can=WD_ALIGN_PARAGRAPH.LEFT)


def gach_dau(d, chu_):
    p = doan(d, '– ' + chu_, truoc=2)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    return p


def s1(x):
    return so(x, 1)


def s2(x):
    return so(x, 2)


def s3(x):
    return so(x, 3)


def ve_hinh(tl, ra):
    """Cao trình đáy dốc, mặt nước (lũ thiết kế, lũ kiểm tra) và đỉnh tường bên hiện trạng dọc dốc nước."""
    Z_dau = TC['Z_nguong'] - TC['P2']
    xs = [r['x'] for r in tl['tuong']]
    day = [Z_dau - TC['i_doc'] * x for x in xs]
    tk, kt = tl['cap'][3], tl['cap'][4]
    mn_tk, mn_kt = T.mat_nuoc_min(tk['Q'], tk['hc']), T.mat_nuoc_min(kt['Q'], kt['hc'])
    x_min = sorted(mn_kt)
    z_day = [Z_dau - TC['i_doc'] * x for x in x_min]
    plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 9})
    fig, ax = plt.subplots(figsize=(7.0, 3.6))
    ax.fill_between(x_min, [z - 0.6 for z in z_day], z_day, color='#e7e6e1', lw=0)
    ax.plot(x_min, z_day, color='#52514e', lw=1.2, label='Đáy dốc nước (i = 8 %)')
    ax.plot(x_min, [z + mn_tk[x] for x, z in zip(x_min, z_day)], color='#2a78d6', lw=2,
            label='Mặt nước lũ thiết kế, Q = %s m³/s' % s2(tk['Q']))
    ax.plot(x_min, [z + mn_kt[x] for x, z in zip(x_min, z_day)], color='#eb6834', lw=2,
            label='Mặt nước lũ kiểm tra, Q = %s m³/s' % s2(kt['Q']))
    hhk = [z + mn_kt[x] * (1 + T.yeu_to(kt['Q'], T.B_doc(x), mn_kt[x])['v'] / 100) for x, z in zip(x_min, z_day)]
    ax.plot(x_min, hhk, color='#eb6834', lw=1.2, ls=(0, (4, 2)), label='Lũ kiểm tra có kể hàm khí')
    ax.plot(xs, [z + r['ht_ht'] for z, r in zip(day, tl['tuong'])], color='#0b0b0b', lw=1.2, ls=(0, (1, 1.5)),
            marker='o', ms=3.5, label='Đỉnh tường bên hiện trạng')
    ax.axvline(TC['L_chuyen_tiep'], color='#8a8984', lw=0.6, ls=':')
    ax.text(TC['L_chuyen_tiep'] + 0.5, Z_dau + 3.6, 'hết đoạn thu hẹp\n31 → 18 m', fontsize=7.5, color='#52514e', va='top')
    ax.set_xlim(0, TC['L_doc'])
    ax.set_xlabel('Khoảng cách từ đầu dốc x (m)')
    ax.set_ylabel('Cao trình (m)')
    ax.grid(True, color='#e1e0d9', lw=0.5)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.legend(loc='upper right', fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(ra, dpi=220, facecolor='white')
    plt.close(fig)


def main():
    ts = T.doc_excel()
    kt_ = T.kha_nang_thao(ts)
    tl = T.thuy_luc_doc(ts)
    tk, kt = kt_['kq']['1%'], kt_['kq']['0.2%']
    Sb, n_k, b_k = kt_['Sb'], ts['TC_n'], ts['TC_b']
    Z_dau = TC['Z_nguong'] - TC['P2']
    dem = Dem()

    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3.0), Cm(2.0), Cm(2.5), Cm(2.5)
    st = d.styles['Normal']
    st.font.name, st.font.size = 'Times New Roman', Pt(13)
    p = lambda txt, **kw: doan(d, txt, thut=1.0, **kw)

    g = doan(d, 'Ghi chú (xoá khi dán vào thuyết minh): tính theo trình tự file mẫu “Thiết kế đường tràn tháo lũ”; thông số '
                'tràn và kết quả điều tiết lũ lấy từ file DieuTietLu_Potapov_f1f2_HoCaGiay.xlsx (ξ_{1} = %s; lũ QTVH 72 giờ; '
                'mở hết cửa khi đón lũ). Số hiệu công thức, bảng đánh liên tục từ (4-1), Bảng 4.3 – sửa lại theo thuyết minh.'
             % so(ts['TC_x1']), co=11, nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT)
    tieu_de(d, '4.2 Tính toán thủy lực tràn chính', 1)

    # ================================================================ 4.2.1 KHẢ NĂNG THÁO
    tieu_de(d, '4.2.1 Kiểm tra khả năng tháo', 2)
    tieu_de(d, '4.2.1.1 Mục đích và nhiệm vụ', 3)
    p('Tràn xả lũ có nhiệm vụ tháo nước thừa khi có lũ đến, bảo đảm an toàn cho cụm công trình đầu mối và hạn chế ngập lụt '
      'ở hạ lưu. Vì vậy cần kiểm tra khả năng tháo của tràn chính: với các hệ số tính lại theo quy phạm, tràn có tháo được '
      'lưu lượng xả lớn nhất trong kết quả điều tiết lũ thiết kế và lũ kiểm tra hay không. Theo TCVN 9147:2012, lưu lượng '
      'tháo qua tràn được xác định theo công thức:')
    phuong_trinh(d, [t('Q = '), sub(t('σ'), t('n')), t('·ε·m·Σb·'), can(t('2g')), t('·'),
                     sub_sup(t('H'), t('0'), phan_so(t('3'), t('2')))], dem.c())
    p('Trong đó: σ_{n} – hệ số ngập; ε – hệ số co hẹp bên, phụ thuộc số khoang và hình dạng mố bên, trụ pin; m – hệ số '
      'lưu lượng, phụ thuộc hình thức ngưỡng tràn và trạng thái chảy qua nó; Σb – tổng chiều rộng các khoang tràn không kể '
      'trụ pin (m); H_{0} – cột nước trên ngưỡng tràn có kể đến lưu tốc tới gần (m); g = 9,81 m/s^{2}.')
    p('Khi tính điều tiết lũ (mục 4.1, file Excel điều tiết lũ) đã dùng m = %s và ε theo TCVN 9147:2012 với cột nước '
      'H = Z − Z_{ngưỡng}, chưa kể lưu tốc tới gần. Mục này tính lại các hệ số theo trình tự quy phạm và đối chiếu khả năng '
      'tháo với kết quả điều tiết.' % so(ts['TC_m']))

    tieu_de(d, '4.2.1.2 Thông số công trình tràn chính', 3)
    gach_dau(d, 'Hình thức: ngưỡng tràn mặt cắt thực dụng thấp, có cửa van; cao trình ngưỡng +%s m (bằng MNDBT); chiều cao '
                'ngưỡng so với sân trước P_{1} = %s m, so với đáy dốc P_{2} = %s m; đỉnh rộng δ = %s m, mái thượng lưu 1:%s, '
                'mái hạ lưu 1:%s.' % (s2(ts['TC_Z']), s2(TC['P1']), s2(TC['P2']), s2(TC['delta']), so(TC['S_TL'], 1),
                                      s2(TC['S_HL'])))
    gach_dau(d, 'Khoang tràn: %d khoang × %s m, Σb = %s m; %d trụ pin: %d trụ giữa dày %s m (ξ_{1} = %s) và %d trụ chính ban '
                'đầu dày %s m (ξ_{2} = %s); tổng bề rộng tràn kể trụ pin B = %s m.'
             % (n_k, so(b_k, 1), so(Sb, 1), ts['TC_n1'] + ts['TC_n2'], ts['TC_n1'], so(ts['TC_d1'], 1), s2(ts['TC_x1']),
                ts['TC_n2'], so(ts['TC_d2'], 1), s2(ts['TC_x2']), so(ts['TC_Btong'], 1)))
    gach_dau(d, 'Kênh dẫn thượng lưu, sân trước: cao trình đáy +%s m.' % s2(kt_['Z_day']))
    gach_dau(d, 'Dốc nước: dài %s m, độ dốc i = %s %%, mặt cắt chữ nhật, bề rộng thu hẹp đều từ %s m (đầu dốc) xuống %s m '
                'trên %s m đầu, sau đó giữ %s m; cao trình đầu dốc +%s m.'
             % (so(TC['L_doc'], 0), so(TC['i_doc'] * 100, 0), so(TC['b_doc_dau'], 0), so(TC['b_doc_cuoi'], 0),
                so(TC['L_chuyen_tiep'], 0), so(TC['b_doc_cuoi'], 0), s2(Z_dau)))
    gach_dau(d, 'Tiêu năng cuối dốc: mũi phun – hố xói, sau đó là kênh xả hạ lưu.')

    tieu_de(d, '4.2.1.3 Tính toán các hệ số', 3)
    doan(d, 'a) Hình thức ngưỡng tràn và hệ số lưu lượng m', nghieng=True, truoc=4)
    p('Điều kiện để ngưỡng làm việc như tràn đỉnh rộng: (2 ÷ 3)H < δ < (8 ÷ 10)H. Với lũ thiết kế H = %s m: '
      '(2 ÷ 3)H = %s ÷ %s m > δ = %s m, nên ngưỡng không làm việc như tràn đỉnh rộng mà như đập tràn mặt cắt thực dụng '
      'thấp (P_{1} = %s m). Hệ số lưu lượng lấy m = %s – thiên an toàn so với m = 0,467 tra trong BCNCKT (như khi tính điều '
      'tiết lũ).' % (s2(tk['H']), s2(2 * tk['H']), s2(3 * tk['H']), s2(TC['delta']), s2(TC['P1']),
                                    so(ts['TC_m'])))
    doan(d, 'b) Ảnh hưởng của cột nước lưu tốc tới gần', nghieng=True, truoc=4)
    p('Theo điều 6.1.5 TCVN 9147:2012 (như mẫu): nếu Ω_{TL} > 4·Σb·H thì không xét ảnh hưởng của lưu tốc tới gần V_{0}, '
      'coi H_{0} = H; nếu Ω_{TL} < 4·Σb·H thì phải xét cột nước lưu tốc tới gần:')
    phuong_trinh(d, [sub(t('H'), t('0')), t(' = H + '), phan_so([t('α·'), sub_sup(t('V'), t('0'), t('2'))], t('2g')),
                     t(' ;   '), sub(t('V'), t('0')), t(' = '), phan_so(t('q'), sub(t('Ω'), t('TL', True))),
                     t(' ;   '), sub(t('Ω'), t('TL', True)), t(' = '), sub(t('B'), t('TL', True)), t('·'),
                     sub(t('H'), t('TL', True))], dem.c())
    p('Trong đó: Ω_{TL} – diện tích mặt cắt ướt tại mặt cắt tính toán ở thượng lưu (cửa vào kênh dẫn, đầu tường cánh); '
      'q – lưu lượng xả lớn nhất qua tràn chính (kết quả điều tiết lũ); H – cột nước trên ngưỡng ứng với q; Z_{đáy kênh} – '
      'cao trình đáy kênh dẫn vào: Z_{đáy kênh} = Z_{ngưỡng} − P_{1} = %s − %s = %s m; H_{TL} – chiều sâu nước thượng lưu '
      'tại mặt cắt tính toán: H_{TL} = Z_{MNLTK} − Z_{đáy kênh} = %s − %s = %s m; B_{TL} – bề rộng kênh thượng lưu tại mặt '
      'cắt tính toán.' % (s2(ts['TC_Z']), s2(TC['P1']), s2(kt_['Z_day']), s2(tk['Z']), s2(kt_['Z_day']), s2(tk['H_TL'])))
    p('Với góc mở rộng của tường cánh 14° (như mẫu), chiều dài tường cánh L_{s} = (3 ÷ 5)H = %s ÷ %s m (lũ kiểm tra: '
      '%s ÷ %s m) – chọn L_{s} = %s m, nằm trong khoảng của cả hai trường hợp (hồ sơ chưa có số đo tường cánh hiện trạng):'
      % (s2(tk['Ls'][0]), s2(tk['Ls'][1]), s2(kt['Ls'][0]), s2(kt['Ls'][1]), so(T.L_S, 0)))
    phuong_trinh(d, [sub(t('B'), t('TL', True)), t(' = B + 2·'), sub(t('L'), t('s')),
                     t('·tan14° = %s + 2 × %s × tan14° = %s m' % (s2(ts['TC_Btong']), so(T.L_S, 0), s2(kt_['B_TL'])))])
    phuong_trinh(d, [sub(t('Ω'), t('TL', True)), t(' = %s × %s = %s ' % (s2(kt_['B_TL']), s2(tk['H_TL']), s2(tk['Om']))),
                     sup(t('m'), t('2')),
                     t(' < 4·Σb·H = 4 × %s × %s = %s ' % (so(Sb, 1), s2(tk['H']), s2(tk['gh']))), sup(t('m'), t('2'))])
    p('nên không bỏ qua được cột nước lưu tốc tới gần. Hệ số động năng α = 1 ÷ 1,1 → α = %s:' % so(T.ALPHA_TL, 1))
    phuong_trinh(d, [sub(t('V'), t('0')), t(' = %s / %s = %s m/s ;   ' % (s2(tk['q']), s2(tk['Om']), s2(tk['V0']))),
                     sub(t('H'), t('0')), t(' = %s + ' % s2(tk['H'])),
                     phan_so([t('%s × ' % so(T.ALPHA_TL, 1)), sup(t(s2(tk['V0'])), t('2'))], t('2 × 9,81')),
                     t(' = %s m' % s2(tk['H0']))])
    p('Lũ kiểm tra tính tương tự (Bảng 4.%d): H_{TL} = %s m, Ω_{TL} = %s m^{2} < 4·Σb·H = %s m^{2}, V_{0} = %s m/s, '
      'H_{0} = %s m.' % (dem.b, s2(kt['H_TL']), s2(kt['Om']), s2(kt['gh']), s2(kt['V0']), s2(kt['H0'])))
    doan(d, 'c) Hệ số co hẹp bên ε', nghieng=True, truoc=4)
    p('Mẫu dùng ε = (ΣB − Σd)/ΣB; ở đây dùng công thức TCVN 9147:2012 có xét hình dạng mố bên và trụ pin, như trong tính '
      'toán điều tiết lũ:')
    phuong_trinh(d, [t('ε = 1 − 0,2·'), ngoac([sub(t('ξ'), t('mb', True)), t(' + '), sub(t('Σξ'), t('trụ', True))], '[', ']'),
                     t('·'), phan_so(sub(t('H'), t('0')), t('n·b')), t(' ;   '), sub(t('Σξ'), t('trụ', True)), t(' = '),
                     sub(t('n'), t('1')), sub(t('ξ'), t('1')), t(' + '), sub(t('n'), t('2')), sub(t('ξ'), t('2'))], dem.c())
    p('Với ξ_{mb} = %s (mố bên lượn tròn); Σξ_{trụ} = %d × %s + %d × %s = %s; n = %d khoang; b = %s m. Lũ thiết kế: '
      'ε = 1 − 0,2 × (%s + %s) × %s / (%d × %s) = %s; lũ kiểm tra: ε = %s.'
      % (s2(ts['TC_xmb']), ts['TC_n1'], s2(ts['TC_x1']), ts['TC_n2'], s2(ts['TC_x2']), s2(ts['TC_sx']), n_k, so(b_k, 1),
         s2(ts['TC_xmb']), s2(ts['TC_sx']), s2(tk['H0']), n_k, so(b_k, 1), s3(tk['eps']), s3(kt['eps'])))
    doan(d, 'd) Hệ số ngập σ_{n}', nghieng=True, truoc=4)
    p('Sau ngưỡng là bậc hạ P_{2} = %s m và dốc nước có i = %s %% > i_{k} (mục 4.2.2), mực nước hạ lưu không ảnh hưởng tới '
      'dòng chảy qua ngưỡng – chảy không ngập, σ_{n} = 1.' % (s2(TC['P2']), so(TC['i_doc'] * 100, 0)))

    tieu_de(d, '4.2.1.4 Kiểm tra khả năng tháo', 3)
    p('Lưu lượng tháo thực tế với lũ thiết kế:')
    phuong_trinh(d, [sub(t('Q'), t('tt', True)), t(' = %s × %s × %s × %s × ' % (so(ts['TC_s'], 0), s3(tk['eps']), so(ts['TC_m']),
                                                                               so(Sb, 1))),
                     can(t('2 × 9,81')), t(' × '), sup(t(s2(tk['H0'])), phan_so(t('3'), t('2'))),
                     t(' = %s ' % s2(tk['Qtt'])), sup(t('m'), t('3')), t('/s')])
    p('So sánh với lưu lượng xả lớn nhất qua tràn chính trong tính toán điều tiết lũ: Q_{tt} = %s m^{3}/s > q_{xả max} = '
      '%s m^{3}/s → thỏa mãn. Sai số so với kết quả điều tiết:' % (s2(tk['Qtt']), s2(tk['q'])))
    phuong_trinh(d, [t('δ = '), phan_so(ngoac([sub(t('Q'), t('tt', True)), t(' − '), sub(t('q'), t('xả max', True))], '|', '|'),
                                        sub(t('Q'), t('tt', True))),
                     t('·100 %% = %s %%' % so(tk['sai_so'], 1))], dem.c())
    hang = []
    for ten, r in (('Lũ thiết kế P = 1%', tk), ('Lũ kiểm tra P = 0,2%', kt)):
        hang.append([ten, s2(r['Z']), s2(r['H']), s2(r['H_TL']), s2(r['Om']), s2(r['gh']), s2(r['V0']), s2(r['H0']),
                     s3(r['eps']), s2(r['Qtt']), s2(r['q']), so(r['sai_so'], 1)])
    bang(d, '%s Kiểm tra khả năng tháo của tràn chính' % dem.bg(),
         ['Trường hợp', 'Z_{max} (m)', 'H (m)', 'H_{TL} (m)', 'Ω_{TL} (m^{2})', '4·Σb·H (m^{2})', 'V_{0} (m/s)',
          'H_{0} (m)', 'ε', 'Q_{tt} (m^{3}/s)', 'q_{xả max} (m^{3}/s)', 'δ (%)'], hang,
         cot=[1450, 680, 620, 660, 720, 780, 660, 640, 640, 760, 820, 560], co=9, le=50)
    p('Ở cả hai trường hợp khả năng tháo thực tế lớn hơn lưu lượng tràn chính trong kết quả điều tiết %s ÷ %s %%. Chênh '
      'lệch chỉ do khi điều tiết lũ chưa kể cột nước lưu tốc tới gần (%s ÷ %s m); nếu kể vào và tính lại điều tiết lũ thì '
      'MNLTK chỉ thấp hơn 2,4 cm, MNLKT thấp hơn 3,3 cm. Như vậy kết quả điều tiết lũ thiên an toàn và đủ chính xác, tràn '
      'chính đủ khả năng tháo lũ thiết kế và lũ kiểm tra; dùng kết quả điều tiết để tính thủy lực dốc nước.'
      % (so(min(tk['sai_so'], kt['sai_so']), 1), so(max(tk['sai_so'], kt['sai_so']), 1), s2(tk['hv']), s2(kt['hv'])), truoc=6)
    tong = {'1%': ts['P1_BinhThuong']['q'], '0.2%': ts['P02_BinhThuong']['q']}
    bang(d, '%s Bảng tổng hợp kết quả điều tiết lũ dùng để tính thủy lực' % dem.bg(),
         ['Tần suất', 'H_{tr} (m)', 'q_{xả max} tràn chính (m^{3}/s)', 'q_{xả max} cả hai tràn (m^{3}/s)', 'Mực nước lũ (m)'],
         [['P = 1%', s2(tk['H']), s2(tk['q']), s2(tong['1%']), 'MNLTK = +' + s2(tk['Z'])],
          ['P = 0,2%', s2(kt['H']), s2(kt['q']), s2(tong['0.2%']), 'MNLKT = +' + s2(kt['Z'])]],
         cot=[1500, 1300, 2100, 2100, 2071])

    # ================================================================ 4.2.2 THỦY LỰC DỐC NƯỚC
    tieu_de(d, '4.2.2 Tính toán thủy lực dốc nước', 2)
    tieu_de(d, '4.2.2.1 Tài liệu tính toán', 3)
    gach_dau(d, 'Hình thức tràn: ngưỡng thực dụng thấp có cửa van, sau ngưỡng là bậc hạ P_{2} = %s m và dốc nước.' % s2(TC['P2']))
    gach_dau(d, 'Độ dốc i = %s %%; hệ số nhám dốc nước (bê tông) n = %s.' % (so(TC['i_doc'] * 100, 0), so(T.N_NHAM, 3)))
    gach_dau(d, 'Số khoang tràn %d; chiều rộng ngưỡng tràn (chưa kể trụ pin) Σb = %s m; cao trình ngưỡng +%s m.'
             % (n_k, so(Sb, 1), s2(ts['TC_Z'])))
    gach_dau(d, 'Cao trình đầu dốc nước: Z_{đầu dốc} = Z_{ngưỡng} − P_{2} = %s − %s = %s m; dốc dài %s m, bề rộng %s m ở đầu '
                'dốc thu hẹp đều xuống %s m tại x = %s m, sau đó %s m.'
             % (s2(ts['TC_Z']), s2(TC['P2']), s2(Z_dau), so(TC['L_doc'], 0), so(TC['b_doc_dau'], 0),
                so(TC['b_doc_cuoi'], 0), so(TC['L_chuyen_tiep'], 0), so(TC['b_doc_cuoi'], 0)))
    gach_dau(d, 'Lưu lượng xả lớn nhất qua tràn chính: lũ thiết kế Q_{xả max} = %s m^{3}/s (H_{tr} = %s m, MNLTK +%s m); lũ '
                'kiểm tra %s m^{3}/s (H_{tr} = %s m, MNLKT +%s m); MNDBT +%s m.'
             % (s2(tk['q']), s2(tk['H']), s2(tk['Z']), s2(kt['q']), s2(kt['H']), s2(kt['Z']), s2(ts['Z0'])))
    p('Tính toán thủy lực nhằm xác định chiều sâu, lưu tốc dòng chảy trên dốc nước, từ đó kiểm tra khả năng làm việc của dốc, '
      'chiều cao tường bên dốc nước và điều kiện thủy lực của dòng chảy khi vào công trình tiêu năng với nhiều cấp lưu lượng '
      'khác nhau. Tính với các cấp lưu lượng 0,4Q; 0,6Q; 0,8Q; Q (lũ thiết kế) và lưu lượng lũ kiểm tra (mẫu lấy cấp lớn nhất '
      '1,2Q; ở đây lưu lượng lũ kiểm tra bằng %sQ nên lấy thay cấp 1,2Q).' % s2(kt['q'] / tk['q']), truoc=6)
    cap = tl['cap']
    bang(d, '%s Các cấp lưu lượng tính toán thủy lực dốc nước' % dem.bg(), [c['ten'] for c in cap],
         [[s2(c['Q']) + ' m^{3}/s' for c in cap]], cot=[1814] * 5)

    tieu_de(d, '4.2.2.2 Độ sâu đầu dốc', 3)
    p('Mẫu tính cho tràn đỉnh rộng nối thẳng với dốc nước nên độ sâu đầu dốc lấy bằng độ sâu phân giới (h_{đd} = h_{k}). '
      'Tràn chính Cà Giây là ngưỡng thực dụng, sau ngưỡng có bậc hạ P_{2} = %s m xuống đáy dốc, dòng chảy đổ xuống chân bậc '
      'và co hẹp; độ sâu đầu dốc là độ sâu co hẹp h_{c}, xác định từ phương trình năng lượng:' % s2(TC['P2']))
    phuong_trinh(d, [t('Q = φ·'), sub(t('B'), t('đ', True)), t('·'), sub(t('h'), t('c')), t('·'),
                     can([t('2g·'), ngoac([sub(t('E'), t('0')), t(' − '), sub(t('h'), t('c'))])]), t(' ;   '),
                     sub(t('E'), t('0')), t(' = '), sub(t('P'), t('2')), t(' + '), sub(t('H'), t('0'))], dem.c())
    p('Trong đó: φ = %s – hệ số lưu tốc; B_{đ} = %s m – bề rộng đầu dốc; H_{0} – cột nước toàn phần trên ngưỡng ứng với '
      'từng cấp lưu lượng, tính ngược từ công thức (4-1); lấy nghiệm h_{c} nhỏ (dòng xiết). So sánh với độ sâu phân giới ở '
      'đầu dốc:' % (s2(T.PHI), so(TC['b_doc_dau'], 0)))
    phuong_trinh(d, [sub(t('h'), t('k')), t(' = '), can(phan_so([t('α·'), sup(t('q'), t('2'))], t('g')), t('3')), t(' ;   q = '),
                     phan_so(t('Q'), t('B'))], dem.c())
    p('với α = 1 – hệ số động lượng; q – lưu lượng đơn vị (m^{2}/s).')
    bang(d, '%s Độ sâu đầu dốc (độ sâu co hẹp ở chân bậc, B_{đ} = %s m)' % (dem.bg(), so(TC['b_doc_dau'], 0)),
         ['Cấp', 'Q (m^{3}/s)', 'H_{0} (m)', 'E_{0} (m)', 'h_{c} (m)', 'h_{k} (m)', 'Fr = v^{2}/(g·h_{c})', 'Trạng thái'],
         [[c['ten'], s2(c['Q']), s3(c['H0']), s3(c['E0']), s3(c['hc']), s3(c['hk_dau']),
           s2((c['Q'] / (TC['b_doc_dau'] * c['hc'])) ** 2 / (T.G * c['hc'])), 'h_{c} < h_{k}: dòng xiết'] for c in cap],
         cot=[1300, 1000, 900, 900, 900, 900, 1171, 2000])

    tieu_de(d, '4.2.2.3 Định tính đường mặt nước', 3)
    p('Xác định độ dốc phân giới i_{k}: từ công thức lưu lượng của Sê-di Q = ω·C·√(R·i):')
    phuong_trinh(d, [sub(t('i'), t('k')), t(' = '), phan_so(sup(t('Q'), t('2')),
                                                          [sub_sup(t('ω'), t('k'), t('2')), t('·'), sub_sup(t('C'), t('k'), t('2')),
                                                           t('·'), sub(t('R'), t('k'))])], dem.c())
    phuong_trinh(d, [sub(t('ω'), t('k')), t(' = B·'), sub(t('h'), t('k')), t(' ;   '), sub(t('χ'), t('k')),
                     t(' = B + 2'), sub(t('h'), t('k')), t(' ;   '), sub(t('R'), t('k')), t(' = '),
                     phan_so(sub(t('ω'), t('k')), sub(t('χ'), t('k'))), t(' ;   '), sub(t('C'), t('k')), t(' = '),
                     phan_so(t('1'), t('n')), t('·'), sub_sup(t('R'), t('k'), phan_so(t('1'), t('6')))])
    p('Tính cho đoạn lăng trụ B = %s m (chiếm %s m trong %s m chiều dài dốc); ở đầu dốc B = %s m, i_{k} = %s ÷ %s, cũng '
      'rất nhỏ so với i.' % (so(TC['b_doc_cuoi'], 0), so(TC['L_doc'] - TC['L_chuyen_tiep'], 0), so(TC['L_doc'], 0),
                            so(TC['b_doc_dau'], 0), so(min(c['ik_dau']['ik'] for c in cap), 4),
                            so(max(c['ik_dau']['ik'] for c in cap), 4)))
    bang(d, '%s Kết quả tính độ dốc phân giới i_{k} (B = %s m)' % (dem.bg(), so(TC['b_doc_cuoi'], 0)),
         ['Cấp', 'Q (m^{3}/s)', 'q (m^{2}/s)', 'h_{k} (m)', 'ω_{k} (m^{2})', 'χ_{k} (m)', 'R_{k} (m)', 'C_{k} (m^{0,5}/s)',
          'i_{k}'],
         [[c['ten'], s2(c['Q']), s2(c['ik']['q']), s3(c['ik']['hk']), s2(c['ik']['w']), s2(c['ik']['chi']), s3(c['ik']['R']),
           s2(c['ik']['C']), so(c['ik']['ik'], 5)] for c in cap], cot=[1300, 1000, 900, 900, 1000, 900, 900, 1100, 1071])
    p('Xác định độ sâu dòng đều h_{0} trên dốc nước theo phương pháp đối chiếu mặt cắt lợi nhất về thủy lực (Agơrốtskin):',
      truoc=6)
    phuong_trinh(d, [t('f'), ngoac(sub(t('R'), t('ln', True))), t(' = '), phan_so([t('4'), sub(t('m'), t('0')), t('·'), can(t('i'))],
                                                                                     t('Q')),
                     t(' ;   '), sub(t('m'), t('0')), t(' = 2'), can([t('1 + '), sup(t('m'), t('2'))]), t(' − m')], dem.c())
    p('Mặt cắt chữ nhật m = 0 → 4m_{0} = 8; i = %s; n = %s. Từ f(R_{ln}) tra phụ lục 8-1 Bảng tính thủy lực được R_{ln}; lập '
      'tỷ số B/R_{ln}, tra phụ lục 8-3 được h/R_{ln}; độ sâu dòng đều h_{0} = (h/R_{ln})·R_{ln}. Ở đây R_{ln} tính trực tiếp '
      'từ Q = 4m_{0}·√i·R_{ln}^{8/3}/n và h_{0} giải trực tiếp từ công thức Sê-di – Ma-ninh, cho cùng kết quả với tra bảng.'
      % (so(TC['i_doc'], 2), so(T.N_NHAM, 3)))
    bang(d, '%s Độ sâu dòng đều h_{0} trên dốc nước (B = %s m)' % (dem.bg(), so(TC['b_doc_cuoi'], 0)),
         ['Cấp', 'Q (m^{3}/s)', 'f(R_{ln})', 'R_{ln} (m)', 'B (m)', 'B/R_{ln}', 'h/R_{ln}', 'h_{0} (m)'],
         [[c['ten'], s2(c['Q']), so(c['h0']['f'], 4), s3(c['h0']['R_ln']), so(c['h0']['B'], 0), s2(c['h0']['B_R']),
           s3(c['h0']['h_R']), s3(c['h0']['h0'])] for c in cap], cot=[1400, 1100, 1100, 1100, 900, 1100, 1100, 1171])
    h20 = {c['ten']: [r for r in c['bang'] if abs(r['x'] - TC['L_chuyen_tiep']) < 1e-6][0]['h'] for c in cap}
    assert all(c['h0']['h0'] < h20[c['ten']] < c['ik']['hk'] for c in cap)      # đường b_II trên đoạn lăng trụ
    bang(d, '%s Định tính đường mặt nước trên đoạn lăng trụ' % dem.bg(),
         ['Cấp', 'Q (m^{3}/s)', 'i', 'i_{k}', 'h_{k} (m)', 'h_{0} (m)', 'h tại x = %s m (m)' % so(TC['L_chuyen_tiep'], 0),
          'So sánh', 'Kết luận'],
         [[c['ten'], s2(c['Q']), so(TC['i_doc'], 2), so(c['ik']['ik'], 4), s3(c['ik']['hk']), s3(c['h0']['h0']),
           s3(h20[c['ten']]), 'i > i_{k}; h_{0} < h < h_{k}', 'Đường nước đổ b_{II}'] for c in cap],
         cot=[1250, 950, 650, 800, 850, 850, 1100, 1400, 1221])
    p('Như vậy i > i_{k}: dốc nước là dốc lớn, dòng chảy trên toàn dốc là dòng xiết. Trên đoạn thu hẹp (x = 0 ÷ %s m) bề rộng '
      'giảm làm chiều sâu tăng dần từ h_{c}; trên đoạn lăng trụ chiều sâu đầu đoạn lớn hơn h_{0} và nhỏ hơn h_{k} nên đường '
      'mặt nước là đường nước đổ b_{II}, chiều sâu giảm dần về h_{0}.' % so(TC['L_chuyen_tiep'], 0), truoc=6)

    tieu_de(d, '4.2.2.4 Tính toán đường mặt nước trên dốc nước', 3)
    p('Dùng phương pháp cộng trực tiếp (phương pháp sai phân): xét hai mặt cắt cách nhau ΔL, tính tuần tự từ đầu dốc đến '
      'cuối dốc. Tại mặt cắt đầu dốc đã biết h_{1} = h_{c}; giả thiết h_{2} ở mặt cắt sau, tính các yếu tố thủy lực B, ω, '
      'v, J… tại từng mặt cắt và khoảng cách giữa hai mặt cắt:')
    phuong_trinh(d, [t('ΔL = '), phan_so(t('ΔЭ'), [t('i − '), gach(t('J'))]), t(' ;   '), gach(t('J')), t(' = '),
                     phan_so([sub(t('J'), t('1')), t(' + '), sub(t('J'), t('2'))], t('2'))], dem.c())
    phuong_trinh(d, [t('ΔЭ = '), sub(t('Э'), t('2')), t(' − '), sub(t('Э'), t('1')), t(' ;   Э = h + '),
                     phan_so([t('α·'), sup(t('v'), t('2'))], t('2g')), t(' ;   J = '),
                     phan_so(sup(t('v'), t('2')), [sup(t('C'), t('2')), t('·R')])])
    p('Trên đoạn thu hẹp bề rộng thay đổi theo x nên chọn trước ΔL = %s m và thử dần h_{2} thỏa mãn (4-%d); trên đoạn lăng '
      'trụ chọn h_{2} (chia đều khoảng chiều sâu tới cuối dốc) rồi tính ΔL; bước cuối tìm h_{2} để L = L_{dốc} = %s m. '
      'Các cột của bảng: (1) thứ tự mặt cắt; (2) bề rộng dốc B; (3) chiều sâu h; (4) lưu tốc v = Q/ω; (5) tỷ năng '
      'Э = h + αv^{2}/2g; (6) ΔЭ = Э_{i+1} − Э_{i}; (7) diện tích ướt ω = B·h; (8) chu vi ướt χ = B + 2h; (9) bán kính thủy '
      'lực R = ω/χ; (10) hệ số Sê-di C = R^{1/6}/n, n = %s; (11) độ dốc thủy lực J; (12) độ dốc thủy lực trung bình '
      'J_{tb} = J̄; (13) i − J_{tb}; (14) ΔL; (15) khoảng cách cộng dồn L.' % (so(T.DL_THU_HEP, 1), dem.ct - 1, so(TC['L_doc'], 0), so(T.N_NHAM, 3)))
    COT15 = [380, 560, 600, 580, 600, 600, 640, 600, 580, 600, 660, 660, 640, 600, 640]
    for c in cap:
        hang = []
        for k, r in enumerate(c['bang']):
            hang.append(['%d' % (k + 1), s2(r['B']), s3(r['h']), s2(r['v']), s3(r['E']),
                         s3(r['dE']) if k else '', s2(r['w']), s2(r['chi']), s3(r['R']), s2(r['C']), so(r['J'], 4),
                         so(r['Jtb'], 4) if k else '', so(r['i_J'], 4) if k else '', s2(r['dL']) if k else '',
                         s2(r['x'])])
        bang(d, '%s Đường mặt nước trên dốc nước – %s, Q = %s m^{3}/s' % (dem.bg(), c['ten'], s2(c['Q'])),
             ['MC', 'B (m)', 'h (m)', 'v (m/s)', 'Э (m)', 'ΔЭ (m)', 'ω (m^{2})', 'χ (m)', 'R (m)', 'C', 'J', 'J_{tb}', 'i − J_{tb}',
              'ΔL (m)', 'L (m)'], hang, cot=COT15, co=8, cot_trai=(), le=30,
             dau=['Chiều sâu đầu dốc h_{c} = %s m; Q = %s m^{3}/s; i = %s %%; n = %s; B = %s → %s m (x = 0 ÷ %s m), %s m '
                  '(x = %s ÷ %s m).' % (s3(c['hc']), s2(c['Q']), so(TC['i_doc'] * 100, 0), so(T.N_NHAM, 3),
                                        so(TC['b_doc_dau'], 0), so(TC['b_doc_cuoi'], 0), so(TC['L_chuyen_tiep'], 0),
                                        so(TC['b_doc_cuoi'], 0), so(TC['L_chuyen_tiep'], 0), so(TC['L_doc'], 0))])
    hang = []
    for c in cap:
        r = c['bang'][-1]
        hang.append([c['ten'], s2(c['Q']), s2(r['x']), s3(r['h']), s2(r['v']), s2(r['v'] / math.sqrt(T.G * r['h']))])
    bang(d, '%s Kết quả tính chiều sâu, lưu tốc cuối dốc nước' % dem.bg(),
         ['Cấp', 'Q (m^{3}/s)', 'L (m)', 'h_{cuối dốc} (m)', 'v_{cuối dốc} (m/s)', 'Fr_{cuối dốc} = v/√(g·h)'], hang,
         cot=[1600, 1400, 1200, 1600, 1600, 1671])
    ve_hinh(tl, HINH)
    d.add_picture(HINH, width=Cm(15.5))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doan(d, 'Hình 4.1 Đường mặt nước trên dốc tràn chính và đỉnh tường bên hiện trạng', co=12, nghieng=True,
         can=WD_ALIGN_PARAGRAPH.CENTER, truoc=2, sau=6)

    tieu_de(d, '4.2.2.5 Kiểm tra điều kiện không xói bề mặt dốc', 3)
    p('Tràn tháo nước từ hồ chứa nên hàm lượng bùn cát trong nước xả nhỏ, khả năng gây bồi lắng nhỏ; chỉ cần kiểm tra điều kiện '
      'không xói bề mặt dốc:')
    phuong_trinh(d, [t('v < '), sub(ngoac(t('v'), '[', ']'), t('kx', True))], dem.c())
    p('Trong đó: v – lưu tốc trung bình trên dốc (lấy ở cuối dốc, nơi lưu tốc lớn nhất); [v]_{kx} – lưu tốc cho phép không '
      'xói của vật liệu làm dốc; theo mẫu (TCVN 4118:2012, phụ lục I, Bảng I5), với dốc nước bằng bê tông cốt thép M200 '
      '[v]_{kx} = %s m/s ứng với chiều sâu nước 0,5 m và %s m/s ứng với chiều sâu nước 1 m.' % (so(T.V_KX[0], 1), so(T.V_KX[1], 1)))
    bang(d, '%s Kết quả kiểm tra khả năng xói bề mặt dốc với các cấp lưu lượng' % dem.bg(),
         ['Q (m^{3}/s)', 'B_{cuối dốc} (m)', 'h_{cuối dốc} (m)', 'v_{cuối dốc} (m/s)', '[v]_{kx} (m/s)', 'Kiểm tra'],
         [[s2(c['Q']), so(TC['b_doc_cuoi'], 0), s3(c['bang'][-1]['h']), s2(c['bang'][-1]['v']),
           '%s ÷ %s' % (so(T.V_KX[0], 1), so(T.V_KX[1], 1)), 'Không xói' if c['bang'][-1]['v'] < T.V_KX[0] else 'Xói']
          for c in cap], cot=[1500, 1400, 1500, 1500, 1600, 1571])
    vmin, vmax = cap[0]['bang'][-1]['v'], cap[-1]['bang'][-1]['v']
    p('Lưu tốc cuối dốc %s ÷ %s m/s nhỏ hơn lưu tốc cho phép nên bề mặt dốc không bị xói. Tuy vậy với lưu tốc trên 10 m/s, '
      'bề mặt bê tông đáy dốc phải phẳng nhẵn để hạn chế xâm thực tại các chỗ gồ ghề, khe nối (mục 4.4).' % (s2(vmin), s2(vmax)),
      truoc=6)

    # ================================================================ 4.2.3 CẤU TẠO CHI TIẾT
    tieu_de(d, '4.2.3 Kiểm tra cấu tạo chi tiết dốc nước', 2)
    tieu_de(d, '4.2.3.1 Chiều cao tường bên', 3)
    p('Dòng chảy có lưu tốc lớn (v > 3 m/s) cuốn không khí vào nước, các bọt khí trộn vào lớp nước gần mặt làm chiều sâu dòng '
      'chảy trên dốc lớn hơn chiều sâu tính toán khi không có hàm khí, vì vậy chiều cao tường bên dốc nước phải tăng lên. '
      'Theo Giáo trình Thủy công tập II (chương 12), chiều sâu dòng chảy có kể hàm khí:')
    phuong_trinh(d, [sub(t('h'), t('hk', True)), t(' = h·'), ngoac([t('1 + '), phan_so(t('v'), t('100'))])], dem.c())
    p('với h – chiều sâu dòng nước tại mặt cắt tính toán (m); v – lưu tốc tại mặt cắt đó (m/s). Chiều cao tường bên yêu cầu '
      '(Giáo trình Thủy công tập II, công thức 17-31):')
    phuong_trinh(d, [sub(t('h'), t('t', True)), t(' = '), sub(t('h'), t('hk', True)), t(' + a')], dem.c())
    tw = tl['tuong']
    p('với a – độ vượt cao an toàn, chọn a = %s m. Tính với lưu lượng lớn nhất qua tràn chính (lũ kiểm tra Q = %s m^{3}/s); '
      'chiều sâu tại các mặt cắt x = 0, 5, …, %s m lấy từ đường mặt nước tính với ΔL = 0,5 m; chiều cao tường hiện trạng '
      'theo phụ lục tính toán của BCNCKT (Bảng PL1-4) đến x = %s m, sau x = %s m giả thiết bằng %s m.'
      % (s2(T.A_TUONG), s2(cap[-1]['Q']), so(TC['L_doc'], 0), so(5 * (len(TC['h_tuong']) - 1), 0),
         so(5 * (len(TC['h_tuong']) - 1), 0), s2(TC['h_tuong'][-1])))
    bang(d, '%s Kiểm tra chiều cao tường bên dốc nước (Q = %s m^{3}/s)' % (dem.bg(), s2(cap[-1]['Q'])),
         ['x (m)', 'B (m)', 'h (m)', 'v (m/s)', 'h_{hk} (m)', 'h_{t} yêu cầu (m)', 'h_{t} hiện trạng (m)',
          'Độ vượt cao h_{t} − h_{hk} (m)', 'Đánh giá'],
         [[so(r['x'], 0), s2(r['B']), s3(r['h']), s2(r['v']), s3(r['hhk']), s2(r['ht_yc']),
           s2(r['ht_ht']) + (' *' if r['gia_thiet'] else ''), s2(r['du']), 'Đạt' if r['du'] >= T.A_TUONG else 'Không đạt']
          for r in tw], cot=[700, 800, 900, 900, 1000, 1200, 1300, 1400, 871], co=9)
    doan(d, '* Giả thiết (hồ sơ chỉ ghi chiều cao tường đến x = %s m).' % so(5 * (len(TC['h_tuong']) - 1), 0), co=10,
         nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT, truoc=0)
    r_min = min(tw, key=lambda r: r['du'])
    p('Chiều cao tường bên hiện trạng lớn hơn yêu cầu ở mọi mặt cắt; độ vượt cao nhỏ nhất %s m tại x = %s m (cuối đoạn thu '
      'hẹp, nơi chiều sâu lớn nhất) lớn hơn a = %s m. Như vậy tường bên dốc nước đủ chiều cao, không cần nâng cấp.'
      % (s2(r_min['du']), so(r_min['x'], 0), s2(T.A_TUONG)), truoc=6)
    assert all(r['du'] >= T.A_TUONG for r in tw)

    tieu_de(d, '4.2.3.2 Chiều dày bản đáy dốc nước', 3)
    p('Theo công thức của V. M. Đômbrốpxki (Giáo trình Công trình tháo lũ, ĐHTL, trang 106 – như mẫu):')
    phuong_trinh(d, [t('t = (0,03 ÷ 0,035)·α·v·'), can(t('h'))], dem.c())
    bd = tl['ban_day']
    p('Trong đó: v = (v_{đầu dốc} + v_{cuối dốc})/2 – lưu tốc trung bình trên dốc; h = (h_{đầu dốc} + h_{cuối dốc})/2 – chiều '
      'sâu trung bình; α – hệ số phụ thuộc nền (nền sét α = 0,8 ÷ 1,0; nền cát và á cát α = 1,5 ÷ 2,0), lấy α = %s. Tính với '
      'lưu lượng lớn nhất Q = %s m^{3}/s:' % (so(T.ALPHA_NEN, 1), s2(bd['Q'])))
    bang(d, '%s Chiều dày bản đáy dốc nước' % dem.bg(),
         ['Q (m^{3}/s)', 'v_{đầu} (m/s)', 'v_{cuối} (m/s)', 'v_{tb} (m/s)', 'h_{đầu} (m)', 'h_{cuối} (m)', 'h_{tb} (m)',
          't tính toán (m)'],
         [[s2(bd['Q']), s2(bd['v_dau']), s2(bd['v_cuoi']), s2(bd['v_tb']), s3(bd['h_dau']), s3(bd['h_cuoi']), s3(bd['h_tb']),
           '%s ÷ %s' % (s2(bd['t'][0]), s2(bd['t'][1]))]], cot=[1150, 1100, 1100, 1100, 1100, 1100, 1100, 1321])
    p('Chiều dày bản đáy yêu cầu t = %s ÷ %s m; cần đối chiếu với chiều dày bản đáy dốc nước hiện trạng theo hồ sơ hoàn công '
      '(chưa có trong hồ sơ hiện có).' % (s2(bd['t'][0]), s2(bd['t'][1])), truoc=6)

    tieu_de(d, '4.2.4 Kết luận', 2)
    gach_dau(d, 'Tràn chính đủ khả năng tháo lũ thiết kế và lũ kiểm tra; kết quả điều tiết lũ (chưa kể lưu tốc tới gần) thiên '
                'an toàn, sai số %s ÷ %s %% so với khả năng tháo tính lại.'
             % (so(min(tk['sai_so'], kt['sai_so']), 1), so(max(tk['sai_so'], kt['sai_so']), 1)))
    gach_dau(d, 'Dòng chảy trên dốc nước là dòng xiết (i = %s > i_{k} ≈ %s); đường mặt nước trên đoạn lăng trụ là đường nước '
                'đổ b_{II}; cuối dốc h = %s ÷ %s m, v = %s ÷ %s m/s, không xói bề mặt dốc.'
             % (so(TC['i_doc'], 2), so(cap[-1]['ik']['ik'], 4), s3(cap[0]['bang'][-1]['h']), s3(cap[-1]['bang'][-1]['h']),
                s2(vmin), s2(vmax)))
    gach_dau(d, 'Tường bên dốc nước hiện trạng đủ chiều cao (độ vượt cao nhỏ nhất %s m ≥ %s m); chiều dày bản đáy yêu cầu '
                '%s ÷ %s m.' % (s2(r_min['du']), s2(T.A_TUONG), s2(bd['t'][0]), s2(bd['t'][1])))
    d.save(RA)
    print('Đã ghi', os.path.abspath(RA))


if __name__ == '__main__':
    main()
