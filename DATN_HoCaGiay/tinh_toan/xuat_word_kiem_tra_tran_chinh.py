# -*- coding: utf-8 -*-
"""
XUẤT WORD: KIỂM TRA TRÀN CHÍNH (MỤC 4.2) VÀ TRÀN SỰ CỐ (MỤC 4.3) HIỆN TRẠNG TRONG ĐỒ ÁN SỬA CHỮA. Bản 10/11/2026: thủy lực
dốc nước, mũi phun tính với 5 cấp lưu lượng (0,4Q; 0,6Q; 0,8Q; Q; Q_KT – 12/11 bỏ chữ "như file mẫu"); không nhắc hồ sơ kiểm định (SV yêu cầu – số liệu của đồ án đã độc
lập; số liệu hiện trạng ghi nguồn chung "tài liệu hiện trạng công trình"). Tính ở
kiem_tra_tran_chinh.py, kiem_tra_tran_su_co.py (tràn sự cố – SV yêu cầu 08/11/2026).
  python3 xuat_word_kiem_tra_tran_chinh.py [đường dẫn .docx ra]
Mặc định: ../TAI_VE/KiemTra_TranChinh_SuaChua.docx (hình: ../TAI_VE/KTTC_DuongMatNuoc.png, KTTC_DuongBaoHoXoi.png,
KTSC_DuongMatNuoc.png)
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

import kiem_tra_tran_chinh as K
import kiem_tra_tran_su_co as SC
import thuy_luc_tran_chinh as T
from omml_docx import bang, can, gach, ngoac, phan_so, phuong_trinh, sub, sub_sup, sup, t
from xuat_bang_tham_docx import chu, doan, so

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'KiemTra_TranChinh_SuaChua.docx')
H_MN = os.path.join(os.path.dirname(RA), 'KTTC_DuongMatNuoc.png')
H_HX = os.path.join(os.path.dirname(RA), 'KTTC_DuongBaoHoXoi.png')
H_SC = os.path.join(os.path.dirname(RA), 'KTSC_DuongMatNuoc.png')
TC = T.TC
s1, s2, s3 = (lambda x: so(x, 1)), (lambda x: so(x, 2)), (lambda x: so(x, 3))


class Dem:
    def __init__(self):
        self.ct, self.b, self.h = 1, 3, 1

    def c(self):
        self.ct += 1
        return '4-%d' % (self.ct - 1)

    def bg(self):
        self.b += 1
        return 'Bảng 4.%d' % (self.b - 1)

    def hinh(self):
        self.h += 1
        return 'Hình 4.%d' % (self.h - 1)


def tieu_de(d, chu_, cap):
    if cap == 1:
        return doan(d, chu_, co=13, dam=True, truoc=12, sau=6, can=WD_ALIGN_PARAGRAPH.LEFT)
    if cap == 2:
        return doan(d, chu_, co=13, dam=True, nghieng=True, truoc=10, sau=4, can=WD_ALIGN_PARAGRAPH.LEFT)
    return doan(d, chu_, co=13, nghieng=True, truoc=6, sau=2, can=WD_ALIGN_PARAGRAPH.LEFT)


def gach_dau(d, chu_):
    p = doan(d, '– ' + chu_, truoc=2)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    return p


def hinh(d, duong, chu_thich, rong=15.5):
    d.add_picture(duong, width=Cm(rong))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doan(d, chu_thich, co=12, nghieng=True, can=WD_ALIGN_PARAGRAPH.CENTER, truoc=2, sau=6)


def ve_mat_nuoc(kq, ra):
    Z0 = TC['Z_nguong'] - TC['P2']
    plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 9})
    fig, ax = plt.subplots(figsize=(7.0, 3.5))
    xs = [r['x'] for r in kq['tk']['bang']]
    day = [Z0 - TC['i_doc'] * x for x in xs]
    ax.fill_between(xs, [z - 0.6 for z in day], day, color='#e7e6e1', lw=0)
    ax.plot(xs, day, color='#52514e', lw=1.2, label='Đáy dốc nước')
    nho = [c for c in kq['cap'] if not c['ten'].startswith('Lũ')]
    for c, mau in zip(nho, ('#b7d3f6', '#8bb7ee', '#5b97e2')):
        ax.plot(xs, [z + r['h'] * (1 + r['v'] / 100) for z, r in zip(day, c['bang'])], color=mau, lw=1.2,
                label='%s, q = %s m³/s' % (c['ten'], s2(c['Q'])))
    for c, mau, nhan in ((kq['tk'], '#2a78d6', 'Q – lũ thiết kế'), (kq['kt'], '#eb6834', 'Lũ kiểm tra')):
        ax.plot(xs, [z + r['h'] * (1 + r['v'] / 100) for z, r in zip(day, c['bang'])], color=mau, lw=2,
                label='%s, q = %s m³/s' % (nhan, s2(c['Q'])))
    xt = [r['x'] for r in kq['doc']]
    ax.plot(xt, [Z0 - TC['i_doc'] * x + r['ht'] for x, r in zip(xt, kq['doc'])], color='#0b0b0b', lw=1.2, ls=(0, (1, 1.5)),
            marker='o', ms=3.5, label='Đỉnh tường bên (sau x = 40 m: giả thiết 2,5 m)')
    ax.set_xlim(0, TC['L_doc'])
    ax.set_xlabel('Khoảng cách từ đầu dốc x (m)')
    ax.set_ylabel('Cao trình (m)')
    ax.grid(True, color='#e1e0d9', lw=0.5)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.legend(loc='upper right', fontsize=7, frameon=False)
    fig.tight_layout()
    fig.savefig(ra, dpi=220, facecolor='white')
    plt.close(fig)


def ve_ho_xoi(kq, ra):
    """Đường bao hố xói: mỗi cấp lưu lượng một hố hình thang (mái TL 1:2,2, HL 1:1,5), độ sâu tính từ đáy lòng dẫn."""
    plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 9})
    fig, ax = plt.subplots(figsize=(7.0, 3.3))
    mau = ['#b7d3f6', '#8bb7ee', '#5b97e2', '#2a78d6', '#1f5fb0', '#13427e']
    xx = [0.01 * k for k in range(0, 5001)]
    bao = [0.0] * len(xx)
    for c, m_ in zip(kq['cap'], mau):
        p = c['mp']
        x = [p['x_tl'], p['x_tl'] + K.M_HO_TL * p['hx'], p['x_hl'] - K.M_HO_HL * p['hx'], p['x_hl']]
        y = [0, -p['hx'], -p['hx'], 0]
        ten_ = ('Q – lũ thiết kế' if c['ten'].startswith('Lũ thiết kế') else 'Lũ kiểm tra' if c['ten'].startswith('Lũ kiểm tra')
                else c['ten'])
        ax.plot(x, y, color=m_, lw=1.0, label='%s: %s m³/s' % (ten_, so(c['Q'], 0 if c['Q'] == round(c['Q']) else 2)))
        for k, xv in enumerate(xx):
            if x[0] <= xv <= x[-1]:
                d_ = min(p['hx'], (xv - x[0]) / K.M_HO_TL, (x[-1] - xv) / K.M_HO_HL)
                bao[k] = max(bao[k], d_)
    ax.plot(xx, [-b for b in bao], color='#0b0b0b', lw=2.0, label='Đường bao hố xói')
    ax.axhline(0, color='#8a8984', lw=0.8)
    ax.plot([-3, 0, 0], [0.9, 0.9, 0], color='#52514e', lw=2.5)
    ax.annotate('mũi phun (x = 0)', xy=(0, 0.9), xytext=(1.2, 1.4), fontsize=7.5, color='#52514e',
                arrowprops=dict(arrowstyle='-', color='#8a8984', lw=0.6))
    ax.set_xlim(-4, max(c['mp']['x_hl'] for c in kq['cap']) + 2)
    ax.set_ylim(-max(c['mp']['hx'] for c in kq['cap']) - 0.8, 2.0)
    ax.set_xlabel('Khoảng cách từ mũi phun về hạ lưu (m)')
    ax.set_ylabel('So với đáy lòng dẫn hạ lưu (m)')
    ax.grid(True, color='#e1e0d9', lw=0.5)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.legend(loc='upper right', fontsize=7, frameon=False, ncol=2)
    fig.tight_layout()
    fig.savefig(ra, dpi=220, facecolor='white')
    plt.close(fig)


def ve_mat_nuoc_tsc(sc, ra):
    plt.rcParams.update({'font.family': 'Liberation Serif', 'font.size': 9})
    fig, ax = plt.subplots(figsize=(7.0, 3.3))
    tk, kt = sc['ca'][0], sc['ca'][2]
    xs = [r['x'] for r in tk['bang']]
    day = [sc['ts']['SC_Z'] - sc['i'] * x for x in xs]
    ax.fill_between(xs, [z - 0.5 for z in day], day, color='#e7e6e1', lw=0)
    ax.plot(xs, day, color='#52514e', lw=1.2, label='Đáy dốc nước')
    for c, mau, nhan in ((tk, '#2a78d6', 'Lũ thiết kế'), (kt, '#eb6834', 'Lũ kiểm tra')):
        ax.plot(xs, [z + r['hhk'] for z, r in zip(day, c['bang'])], color=mau, lw=2,
                label='%s, Q = %s m³/s (có kể hàm khí)' % (nhan, s2(c['Q'])))
    can_ = [max(c['bang'][k]['hhk'] for c in sc['ca']) + T.A_TUONG for k in range(len(xs))]
    ax.plot(xs, [z + h for z, h in zip(day, can_)], color='#0b0b0b', lw=1.2, ls=(0, (1, 1.5)),
            label='Đỉnh tường cần (h_hk + a, trường hợp bất lợi nhất)')
    ax.set_xlim(0, sc['L'])
    ax.set_xlabel('Khoảng cách từ đầu dốc x (m)')
    ax.set_ylabel('Cao trình (m)')
    ax.grid(True, color='#e1e0d9', lw=0.5)
    for s_ in ('top', 'right'):
        ax.spines[s_].set_visible(False)
    ax.legend(loc='upper right', fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(ra, dpi=220, facecolor='white')
    plt.close(fig)


def muc_tran_su_co(d, dem, kq):
    """Mục 4.3 – kiểm tra tràn sự cố hiện trạng (cùng trình tự với tràn chính)."""
    ts = kq['ts']
    sc = SC.tinh(ts)
    ca = sc['ca']
    tk, kt = ca[0], ca[2]
    p = lambda txt, **kw: doan(d, txt, thut=1.0, **kw)
    B = sc['B']
    ty = [100 * c['ty_le'] for c in ca]
    q_tc = [ts[s]['qtc'] / TC['b_doc_cuoi'] for s in ('P1_BinhThuong', 'P02_BinhThuong')]
    boi = [tk['q'] / q_tc[0], kt['q'] / q_tc[1]]
    tieu_de(d, '4.3 Kiểm tra tràn sự cố hiện trạng', 1)
    p('Tràn sự cố đặt ở vai phải đập: 1 khoang b = %s m có cửa van phẳng, ngưỡng +%s m – thấp hơn ngưỡng tràn chính %s m, sau '
      'ngưỡng không có bậc hạ; dốc nước lăng trụ rộng %s m, dài %s m, i = %s %% (QTVH Bảng 13). Theo quy trình vận hành, cửa van '
      'tràn sự cố được mở ngay khi lũ về, nên khi xả lũ thiết kế và lũ kiểm tra tràn sự cố xả %s ÷ %s %% tổng lưu lượng xả của hồ; '
      'lưu lượng đơn vị trên dốc %s ÷ %s m^{3}/s/m, gấp %s ÷ %s lần ở cuối dốc tràn chính. Đồ án kiểm tra tràn sự cố theo cùng '
      'trình tự với tràn chính: khả năng '
      'xả, thủy lực dốc nước (chiều cao tường cần, xói). Phần tiêu năng cuối dốc chỉ nêu thông số dòng chảy vì hồ sơ chưa có số liệu '
      'công trình tiêu năng.'
      % (so(B, 0), s2(ts['SC_Z']), s2(ts['TC_Z'] - ts['SC_Z']), so(B, 0), so(sc['L'], 0), so(sc['i'] * 100, 0), so(min(ty), 0),
         so(max(ty), 0), s2(min(c['q'] for c in ca)), s2(max(c['q'] for c in ca)), s1(min(boi)), s1(max(boi))))

    tieu_de(d, '4.3.1 Khả năng xả', 2)
    p('Lưu lượng qua tràn sự cố tính theo (4-1) với Σb = b = %s m, σ_{n} = 1; hệ số co hẹp bên của tràn một khoang (TCVN '
      '9147:2012):' % so(B, 0))
    phuong_trinh(d, [t('ε = 1 − 0,2·'), sub(t('ξ'), t('mb', True)), t('·'), phan_so(sub(t('H'), t('0')), t('b')), t(' ;   '),
                     phan_so(sub(t('H'), t('0')), t('b')), t(' ≤ 1')], dem.c())
    hb = [c['H_b'] for c in ca]
    p('với ξ_{mb} = %s – mố bên vuông góc (bình đồ cửa vào: tường thẳng, đầu tường quay vuông góc, không có tường cánh cong); '
      'm = %s. Khi xả lũ H_{0}/b = %s ÷ %s > 1 nên ε = %s, ε·m = %s. Sau ngưỡng không có bậc hạ nên đầu dốc là mặt cắt khống '
      'chế, dòng chảy qua mặt cắt này ở trạng thái phân giới; theo công thức độ sâu co hẹp (4-%d) với P = 0, lưu lượng lớn nhất '
      'đạt được khi h = 2E_{0}/3, ứng với hệ số lưu lượng giới hạn:'
      % (s2(ts['SC_xmb']), s2(ts['SC_m']), s3(min(hb)), s3(max(hb)), s3(tk['eps']), s3(tk['m0']), kq['ct_hc']))
    phuong_trinh(d, [sub(t('m'), t('max', True)), t(' = φ·'), phan_so(t('2'), t('3')), t('·'), can(phan_so(t('1'), t('3'))),
                     t(' = %s·%s = %s' % (s2(T.PHI), s3(2 / 3 * math.sqrt(1 / 3)), s3(sc['m_max'])))], dem.c())
    p('Hệ số đồ án dùng (ε·m = %s) sát giới hạn này và gần bằng hệ số của QTVH 2026 (m = 0,37), phù hợp điều kiện chảy phân giới '
      'ở đầu dốc. Lưu lượng tràn sự cố khi xả lũ cho trong Bảng %s.'
      % (s3(tk['m0']), '4.%d' % dem.b))
    bang(d, '%s Lưu lượng tràn sự cố khi xả lũ' % dem.bg(),
         ['Trường hợp', 'Z_{max} (m)', 'H_{0} (m)', 'H_{0}/b', 'ε', 'ε·m', 'Q tràn sự cố (m^{3}/s)', 'q xả của hồ (m^{3}/s)',
          'Tỷ lệ (%)'],
         [[c['ten'], s2(c['Z']), s3(c['H']), s3(c['H_b']), s3(c['eps']), s3(c['m0']), s2(c['Q']), s2(c['q_tong']),
           s1(100 * c['ty_le'])] for c in ca],
         cot=[2471, 800, 800, 750, 700, 700, 1000, 1000, 850], co=10)

    tieu_de(d, '4.3.2 Thủy lực dốc nước', 2)
    p('Độ sâu phân giới, độ dốc phân giới và độ sâu dòng đều trên dốc lăng trụ b = %s m:' % so(B, 0))
    phuong_trinh(d, [sub(t('h'), t('k')), t(' = '), can(phan_so([t('α·'), sup(t('q'), t('2'))], t('g')), t('3')), t(' ;   '),
                     sub(t('i'), t('k')), t(' = '),
                     phan_so(sup(t('Q'), t('2')), [sub_sup(t('ω'), t('k'), t('2')), t('·'), sub_sup(t('C'), t('k'), t('2')), t('·'),
                                                   sub(t('R'), t('k'))]),
                     t(' ;   Q = '), sub(t('ω'), t('0')), t('·'), sub(t('C'), t('0')), t('·'), can([sub(t('R'), t('0')), t('·i')])],
                 dem.c())
    bang(d, '%s Độ sâu phân giới, độ dốc phân giới, độ sâu dòng đều của dốc tràn sự cố' % dem.bg(),
         ['Trường hợp', 'Q (m^{3}/s)', 'q (m^{3}/s/m)', 'h_{k} (m)', 'i_{k}', 'h_{0} (m)', 'h đầu dốc 0,99·h_{k} (m)'],
         [[c['ten'], s2(c['Q']), s2(c['q']), s3(c['ik']['hk']), so(c['ik']['ik'], 5), s3(c['h0']['h0']), s3(c['dau']['h'])]
          for c in ca], cot=[2871, 950, 950, 950, 950, 950, 1450], co=10)
    p('i = %s %% lớn hơn nhiều i_{k} nên dốc nước tràn sự cố là dốc dốc, đầu dốc là mặt cắt khống chế; trên dốc dòng chảy xiết tăng '
      'tốc, đường mặt nước là đường nước hạ b_{II}, độ sâu giảm dần từ h_{k} về h_{0} nhưng chưa đạt h_{0} trên chiều dài %s m. '
      'Đường mặt nước tính theo (4-%d), chọn trước Δl = %s m, thử dần h như dốc tràn chính, bắt đầu từ 0,99·h_{k} (tại đúng h_{k} '
      'phương trình sai phân không xác định); độ sâu có kể hàm khí theo (4-%d); chiều cao tường cần = h_{hk} + a, a = %s m.'
      % (so(sc['i'] * 100, 0), so(sc['L'], 0), kq['ct_dmn'], so(SC.DL, 0), kq['ct_hk'], s2(T.A_TUONG)))
    hang = []
    for k in range(0, len(tk['bang']), 2):
        a, b = tk['bang'][k], kt['bang'][k]
        can_ = max(c['bang'][k]['hhk'] for c in ca) + T.A_TUONG
        hang.append([so(a['x'], 0), s3(a['h']), s3(a['hhk']), s2(a['v']), s3(b['h']), s3(b['hhk']), s2(b['v']), s2(b['Fr']), s2(can_)])
    b_tsc = dem.bg()
    bang(d, '%s Kết quả thủy lực dốc nước tràn sự cố (lũ thiết kế Q = %s m^{3}/s; lũ kiểm tra Q = %s m^{3}/s)'
         % (b_tsc, s2(tk['Q']), s2(kt['Q'])),
         ['x (m)', 'h TK (m)', 'h_{hk} TK (m)', 'v TK (m/s)', 'h KT (m)', 'h_{hk} KT (m)', 'v KT (m/s)', 'Fr KT', 'Tường cần (m)'],
         hang, cot=[800, 1000, 1050, 1000, 1000, 1050, 1000, 900, 1271], co=10)
    doan(d, 'Tường cần lấy theo trường hợp bất lợi nhất trong 4 trường hợp ở Bảng %s (kể cả khi 1 cửa tràn chính không mở). Bảng '
            'tính đầy đủ (Δl = 2 m) ở cuối mục.' % ('4.%d' % (dem.b - 3)), co=10, nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT, truoc=0)
    can_dau = max(c['dau']['ht_can'] for c in ca)
    can_cuoi = max(c['cuoi']['ht_can'] for c in ca)
    v_max = max(c['v_max'] for c in ca)
    p('Tường bên dốc nước tràn sự cố cần cao tối thiểu %s m ở đầu dốc và %s m ở cuối dốc (Hình %s). Hồ sơ hiện có chưa ghi chiều '
      'cao tường bên dốc nước tràn sự cố (chỉ có bình đồ tổng thể, không có bản vẽ tràn); các trị số này dùng để đối chiếu khi '
      'khảo sát lập bản vẽ thi công. Lưu tốc lớn nhất %s m/s nhỏ hơn lưu tốc cho '
      'phép không xói %s m/s (bê tông M200, như dốc tràn chính) – dốc nước không xói theo lý thuyết; nửa cuối dốc có lưu tốc trên '
      '10 m/s nên khi lát lại tấm, sửa bề mặt cần bảo đảm độ phẳng như dốc tràn chính (mục 4.4).'
      % (s2(can_dau), s2(can_cuoi), '4.%d' % dem.h, s2(v_max), so(SC.V_KX, 1)), truoc=6)
    ve_mat_nuoc_tsc(sc, H_SC)
    hinh(d, H_SC, '%s Đường mặt nước (có kể hàm khí) trên dốc tràn sự cố và đỉnh tường cần' % dem.hinh())

    tieu_de(d, '4.3.3 Tiêu năng cuối dốc', 2)
    cu = [c['cuoi'] for c in ca]
    p('Hồ sơ hiện có chưa ghi hình thức, kích thước công trình tiêu năng cuối dốc tràn sự cố, cao trình và mực nước hạ lưu nên chưa '
      'kiểm tra được như mũi phun tràn chính (mục 4.2.3). Thông số dòng chảy cuối dốc (Bảng %s) dùng làm căn cứ kiểm tra khi có số '
      'liệu: lưu lượng đơn vị %s ÷ %s m^{3}/s/m, lưu tốc %s ÷ %s m/s, Fr = %s ÷ %s – năng lượng dòng chảy lớn hơn cuối dốc tràn chính. '
      'Cần khảo sát hiện trạng kênh xả, hố xói hạ lưu tràn sự cố và quan trắc sau mỗi lần xả lũ lớn.'
      % ('4.%d' % dem.b, s2(min(c['q'] for c in ca)), s2(max(c['q'] for c in ca)), s2(min(x['v'] for x in cu)),
         s2(max(x['v'] for x in cu)), s2(min(x['Fr'] for x in cu)), s2(max(x['Fr'] for x in cu))))
    bang(d, '%s Thông số dòng chảy cuối dốc tràn sự cố' % dem.bg(),
         ['Trường hợp', 'Q (m^{3}/s)', 'q (m^{3}/s/m)', 'h (m)', 'v (m/s)', 'Fr = v^{2}/(g·h)', 'E (m)', 'h_{hk} (m)'],
         [[c['ten'], s2(c['Q']), s2(c['q']), s3(c['cuoi']['h']), s2(c['cuoi']['v']), s2(c['cuoi']['Fr']), s2(c['cuoi']['E']),
           s3(c['cuoi']['hhk'])] for c in ca], cot=[2871, 850, 850, 850, 850, 1000, 850, 950], co=10)

    tieu_de(d, '4.3.4 Kết luận', 2)
    gach_dau(d, 'Khả năng xả: tràn sự cố xả %s m^{3}/s (lũ thiết kế) và %s m^{3}/s (lũ kiểm tra), chiếm %s ÷ %s %% lưu lượng xả của '
                'hồ; hệ số ε·m = %s (mố bên vuông góc, H_{0}/b ≤ 1) không vượt giới hạn %s của ngưỡng không có bậc hạ.'
             % (s2(tk['Q']), s2(kt['Q']), so(min(ty), 0), so(max(ty), 0), s3(tk['m0']), s3(sc['m_max'])))
    gach_dau(d, 'Dốc nước: dòng xiết từ h_{k} ở đầu dốc; tường bên cần cao tối thiểu %s m (đầu dốc) ÷ %s m (cuối dốc) – kiến nghị '
                'đo kiểm tra chiều cao tường thực tế khi khảo sát lập bản vẽ thi công; lưu tốc lớn nhất %s m/s, không xói theo lý thuyết.'
             % (s2(can_dau), s2(can_cuoi), s2(v_max)))
    gach_dau(d, 'Tiêu năng cuối dốc: chưa có hồ sơ để kiểm tra; khảo sát kênh xả, hố xói hạ lưu và quan trắc sau lũ lớn.')
    gach_dau(d, 'Sửa chữa tràn sự cố: chống thấm mang trái (chương 8), lát lại tấm mái bị gãy và xử lý hố sụt sau tường cánh (mục '
                '4.4); không thay đổi quy mô tràn.')

    tieu_de(d, 'Bảng tính đường mặt nước trên dốc tràn sự cố (dùng cho %s)' % b_tsc, 2)
    COT15 = [380, 560, 600, 580, 600, 600, 640, 600, 580, 600, 660, 660, 640, 600, 640]
    for c in (tk, kt):
        hang = []
        for k, r in enumerate(c['bang']):
            hang.append(['%d' % (k + 1), so(r['x'], 0), s3(r['h']), s2(r['v']), s3(r['E']), s3(r['dE']) if k else '', s2(r['w']),
                         s2(r['chi']), s3(r['R']), s2(r['C']), so(r['J'], 5), so(r['Jtb'], 5) if k else '',
                         so(r['i_J'], 5) if k else '', s2(r['dL']) if k else '', s3(r['hhk'])])
        bang(d, '%s Đường mặt nước dốc tràn sự cố – %s, Q = %s m^{3}/s' % (dem.bg(), c['ten'].lower(), s2(c['Q'])),
             ['MC', 'x (m)', 'h (m)', 'v (m/s)', 'E (m)', 'ΔE (m)', 'ω (m^{2})', 'χ (m)', 'R (m)', 'C', 'J', 'J_{tb}', 'i − J_{tb}',
              'Δl (m)', 'h_{hk} (m)'], hang, cot=COT15, co=8, cot_trai=(), le=30,
             dau=['Độ sâu đầu dốc 0,99·h_{k} = %s m (h_{k} = %s m); b = %s m; i = %s %%; n = %s; Δl chọn = %s m.'
                  % (s3(c['dau']['h']), s3(c['ik']['hk']), so(B, 0), so(sc['i'] * 100, 0), so(T.N_NHAM, 3), so(SC.DL, 0))])


def main():
    kq = K.tinh()
    ts, tk, kt, doc_, cap = kq['ts'], kq['tk'], kq['kt'], kq['doc'], kq['cap']
    Sb = ts['TC_n'] * ts['TC_b']
    dem = Dem()
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3.0), Cm(2.0), Cm(2.5), Cm(2.5)
    st = d.styles['Normal']
    st.font.name, st.font.size = 'Times New Roman', Pt(13)
    p = lambda txt, **kw: doan(d, txt, thut=1.0, **kw)

    doan(d, 'Ghi chú (xoá khi dán vào thuyết minh): mục 4.2 (tràn chính) và 4.3 (tràn sự cố); thủy lực dốc nước tính với 5 cấp '
            'lưu lượng; thông số tràn và kết quả điều tiết lấy từ file DieuTietLu_Potapov_f1f2. '
            'Số hiệu công thức, bảng, hình đánh liên tục từ (4-1), Bảng 4.3, Hình 4.1 – sửa lại theo thuyết minh.', co=11,
         nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT)
    tieu_de(d, '4.2 Kiểm tra tràn chính hiện trạng', 1)
    p('Tràn chính là công trình đang vận hành; đồ án sửa chữa không thiết kế lại hình thức, kích thước tràn mà kiểm tra các điều '
      'kiện làm việc của tràn hiện có với lũ cập nhật (QTVH 2026): khả năng xả lũ, thủy lực dốc nước với 5 cấp lưu lượng (tường '
      'bên, xói bề mặt), tiêu năng mũi phun – hố xói. Điều kiện không tràn đỉnh đập đã kiểm tra ở mục '
      '2.3, ổn định ngưỡng tràn ở mục 4.5. Các nội dung của thiết kế mới (chọn hình thức, bề rộng tràn, kênh dẫn, chiều dày bản '
      'đáy, kích thước công trình tiêu năng) không tính lại – Bảng %s.' % ('4.%d' % dem.b))
    bang(d, '%s Nội dung tính toán tràn chính trong đồ án sửa chữa' % dem.bg(),
         ['Hạng mục', 'Nội dung tính toán', 'Mục'],
         [['Hình thức, kích thước, bố trí', 'Không tính lại – mô tả hiện trạng (bình đồ, ảnh hiện trạng)', '4.2.1'],
          ['Khả năng xả, điều tiết lũ', 'Điều tiết với lũ QTVH 2026; hệ số co hẹp theo hình dạng mố, trụ hiện có; quan hệ Q ~ Z',
           '4.1; 4.2.1'],
          ['Lưu tốc tới gần', 'Bỏ qua (H_{0} = H) – thiên an toàn', '4.2.1'],
          ['Thủy lực dốc nước', '5 cấp lưu lượng 0,4Q; 0,6Q; 0,8Q; Q; lũ kiểm tra: h_{c}, h_{k}, i_{k}, h_{0}, đường mặt nước',
           '4.2.2'],
          ['Tường bên, bề mặt dốc', 'Kiểm tra tường hiện có với lũ thiết kế, lũ kiểm tra; kiểm tra xói; sửa chữa bề mặt dốc',
           '4.2.2; 4.4'],
          ['Tiêu năng', 'Kiểm tra mũi phun – hố xói hiện có với 5 cấp lưu lượng (TCVN 8420:2010)', '4.2.3'],
          ['Không tràn đỉnh đập', 'Cao trình đỉnh đập, tường chắn sóng', '2.3'],
          ['Ổn định ngưỡng tràn', 'Kiểm tra ổn định trượt, lật', '4.5']],
         cot=[2600, 5300, 1171], co=10, cot_trai=(0, 1))

    # ================================================================ 4.2.1
    tieu_de(d, '4.2.1 Khả năng xả lũ của ngưỡng tràn', 2)
    p('Tràn chính là đập tràn thực dụng có cửa van: %d khoang × %s m (Σb = %s m), ngưỡng +%s m; %d trụ giữa dày %s m và %d trụ '
      'chính dày %s m; sau ngưỡng có bậc hạ P_{2} = %s m xuống đầu dốc nước dài %s m, i = %s %%, bề rộng thu hẹp %s → %s m trên %s m '
      'đầu; cuối dốc là mũi phun. Lưu lượng tháo qua tràn (TCVN 9147:2012):'
      % (ts['TC_n'], so(ts['TC_b'], 1), so(Sb, 1), s2(ts['TC_Z']), ts['TC_n1'], so(ts['TC_d1'], 1), ts['TC_n2'],
         so(ts['TC_d2'], 1), s2(TC['P2']), so(TC['L_doc'], 0), so(TC['i_doc'] * 100, 0), so(TC['b_doc_dau'], 0),
         so(TC['b_doc_cuoi'], 0), so(TC['L_chuyen_tiep'], 0)))
    phuong_trinh(d, [t('Q = '), sub(t('σ'), t('n')), t('·ε·m·Σb·'), can(t('2g')), t('·'),
                     sub_sup(t('H'), t('0'), phan_so(t('3'), t('2')))], dem.c())
    phuong_trinh(d, [t('ε = 1 − 0,2·'), ngoac([sub(t('ξ'), t('mb', True)), t(' + Σξ')], '[', ']'), t('·'),
                     phan_so(t('1'), t('n')), t('·'), phan_so(sub(t('H'), t('0')), t('b')), t(' ;   '),
                     phan_so(sub(t('H'), t('0')), t('b')), t(' ≤ 1 ;   Σξ = '), sub(t('n'), t('1')), sub(t('ξ'), t('1')), t(' + '),
                     sub(t('n'), t('2')), sub(t('ξ'), t('2'))], dem.c())
    p('Trong đó: σ_{n} = 1 – sau ngưỡng là dốc nước chảy xiết (i > i_{k}) nên chảy không ngập; m = %s – đập tràn thực dụng '
      'thấp; ξ_{mb} = %s – mố bên lượn tròn (bình đồ cửa vào: tường cánh thượng lưu cong, nối tiếp gần như tiếp tuyến với mố '
      'bên); Σξ = %d × %s + %d × %s = %s – cả 8 trụ pin có đầu phía thượng lưu lượn tròn (bình đồ hiện trạng; đuôi trụ phía hạ '
      'lưu vuông hay bo tròn không ảnh hưởng vì chảy không ngập); khi H_{0}/b > 1 thì lấy H_{0}/b = 1. Bỏ qua lưu tốc tới gần, '
      'H_{0} = H = Z − Z_{ngưỡng}: lưu tốc tới gần làm tăng cột nước toàn phần và khả năng xả nên bỏ qua là thiên an toàn (mực '
      'nước lũ tính ra cao hơn). Đây cũng là công thức dùng khi điều tiết lũ (mục 4.1), nên khả năng xả của ngưỡng không cần kiểm '
      'tra lại bằng một phép tính riêng. Các hệ số tổng hợp trong Bảng %s.' % (so(ts['TC_m']), s2(ts['TC_xmb']), ts['TC_n1'],
                                                                       s2(ts['TC_x1']), ts['TC_n2'], s2(ts['TC_x2']),
                                                                       s2(ts['TC_sx']), '4.%d' % dem.b))
    e_tk, e_kt = T.eps(tk['H'], ts), T.eps(kt['H'], ts)
    khoang = lambda a, b, f: f(a) if f(a) == f(b) else '%s ÷ %s' % (f(a), f(b))
    e_lon = max(e_tk, e_kt)
    bang(d, '%s Hệ số tính khả năng xả của tràn chính' % dem.bg(),
         ['Đại lượng', 'Giá trị', 'Ghi chú'],
         [['Bề rộng tràn nước Σb (m)', so(Sb, 1), '9 khoang × 2,4 m, có cửa van'],
          ['Hệ số lưu lượng m', so(ts['TC_m'], 3), 'Đập tràn thực dụng thấp'],
          ['Hệ số hình dạng mố bên ξ_{mb}', s2(ts['TC_xmb']), 'Mố bên lượn tròn (bình đồ cửa vào)'],
          ['Hệ số hình dạng trụ pin ξ_{1} = ξ_{2}', s2(ts['TC_x1']), '8 trụ đầu tròn phía thượng lưu (bình đồ)'],
          ['Hệ số co hẹp bên ε', khoang(e_tk, e_kt, s3), 'Khi xả lũ H_{0}/b > 1 → lấy H_{0}/b = 1'],
          ['m_{0} = σ_{n}·ε·m', khoang(e_tk * ts['TC_m'], e_kt * ts['TC_m'], s3), 'σ_{n} = 1 – chảy không ngập'],
          ['Lưu tốc tới gần V_{0}', 'Bỏ qua', 'H_{0} = H – thiên an toàn']],
         cot=[3000, 1600, 4471], co=10, cot_trai=(0, 2))
    p('Bình đồ hiện trạng (mặt bằng nâng cấp cụm công trình đầu mối) cho cao độ sân lát đá trước ngưỡng +73,7 ÷ +74,2 m, ngưỡng '
      '+74,70 m, đáy ngay sau ngưỡng +73,30 m (đường chân ngưỡng cách dãy trụ pin ≈ 4,0 m; mái hạ lưu ngưỡng dài ≈ 3,0 m), tức '
      'P_{1} ≈ 0,7 m, P_{2} ≈ 1,4 m. Đồ án dùng P_{1} = %s m, P_{2} = %s m, δ = %s m.'
      % (s2(TC['P1']), s2(TC['P2']), s2(TC['delta'])), truoc=6)
    qz = kq['qz']
    bang(d, '%s Quan hệ lưu lượng – mực nước của công trình tháo lũ' % dem.bg(),
         ['Z (m)', 'H tràn chính (m)', 'ε (9 khoang)', 'q tràn chính 9 khoang (m^{3}/s)', 'q tràn chính 8 khoang (m^{3}/s)',
          'q tràn sự cố (m^{3}/s)', 'q tổng – bình thường (m^{3}/s)'],
         [[s2(r['Z']), s2(r['H']), s3(r['eps']), s2(r['q9']), s2(r['q8']), s2(r['qsc']), s2(r['q_bt'])] for r in qz],
         cot=[1000, 1150, 1150, 1450, 1450, 1350, 1521], co=10, cot_trai=())
    tr = [('P = 1%, bình thường', 'P1_BinhThuong'), ('P = 1%, 1 cửa tràn chính không mở', 'P1_SuCo'),
          ('P = 0,2%, bình thường', 'P02_BinhThuong'), ('P = 0,2%, 1 cửa tràn chính không mở', 'P02_SuCo')]
    bang(d, '%s Kết quả điều tiết lũ (lũ QTVH 2026)' % dem.bg(),
         ['Trường hợp', 'Z_{max} (m)', 'q tràn chính (m^{3}/s)', 'q tràn sự cố (m^{3}/s)', 'q tổng (m^{3}/s)',
          'Đỉnh đập sau SC − Z_{max} (m)'],
         [[ten, s2(ts[k]['Z']), s2(ts[k]['qtc']), s2(ts[k]['qsc']), s2(ts[k]['qtc'] + ts[k]['qsc']),
           s2(ts['Z_dinh_SC'] - ts[k]['Z'])] for ten, k in tr],
         cot=[2771, 1150, 1300, 1300, 1150, 1400], co=10)
    p('Với lũ QTVH 2026 (Q_{max} 1%% = 921 m^{3}/s, thu phóng trận lũ trạm Đá Bàn) và quan hệ Z ~ F ~ W theo bình đồ lòng hồ '
      '2022, tràn chính và tràn sự cố tháo được lũ thiết kế và lũ kiểm tra; MNLTK = +%s m, MNLKT = +%s m thấp hơn đỉnh đập sau '
      'sửa chữa +%s m (mục 2.3).' % (s2(tk['MNL']), s2(kt['MNL']), s2(ts['Z_dinh_SC'])), truoc=6)

    # ================================================================ 4.2.2
    tieu_de(d, '4.2.2 Thủy lực dốc nước', 2)
    p('Thủy lực dốc nước được tính với 5 cấp lưu lượng: 0,4Q; 0,6Q; 0,8Q; Q và Q_{KT}, trong đó Q = %s m^{3}/s là lưu lượng '
      'xả lớn nhất qua tràn chính khi lũ thiết kế (P = 1%%), Q_{KT} = %s m^{3}/s (= %sQ) là lưu lượng xả lớn nhất khi lũ kiểm '
      'tra (P = 0,2%%). Tràn có cửa van nên khi vận hành, lưu lượng qua tràn có thể nhận mọi giá trị từ 0 đến Q_{KT}: các cấp '
      'nhỏ dùng để kiểm tra trạng thái chảy trên dốc và tiêu năng (hố xói thường lấn gần mũi phun nhất khi lưu lượng nhỏ), cấp '
      'lớn nhất dùng để kiểm tra chiều cao tường bên và lưu tốc lớn nhất. Cột nước trên ngưỡng H ứng với từng cấp giải ngược '
      'từ (4-1), (4-2).' % (s2(tk['Q']), s2(kt['Q']), s2(kt['Q'] / tk['Q'])))
    p('Độ sâu co hẹp ở đầu dốc nước (chân bậc hạ sau ngưỡng) tính theo công thức tổng quát, giải bằng thử dần:')
    phuong_trinh(d, [t('Q = φ·'), sub(t('ω'), t('c')), t('·'), can([t('2g·'), ngoac([sub(t('E'), t('0')), t(' − '),
                                                                                       sub(t('h'), t('c'))])]),
                     t(' ;   '), sub(t('ω'), t('c')), t(' = '), sub(t('b'), t('c')), t('·'), sub(t('h'), t('c')), t(' ;   '),
                     sub(t('E'), t('0')), t(' = '), sub(t('P'), t('2')), t(' + H')], dem.c())
    kq['ct_hc'] = dem.ct - 1
    p('với φ = %s; b_{c} = %s m (bề rộng đầu dốc); E_{0} = P_{2} + H (bỏ qua V_{0}). Độ sâu phân giới, độ dốc phân giới và độ sâu '
      'dòng đều tính cho đoạn lăng trụ B = %s m:' % (s2(T.PHI), so(TC['b_doc_dau'], 0), so(TC['b_doc_cuoi'], 0)))
    phuong_trinh(d, [sub(t('h'), t('k')), t(' = '), can(phan_so([t('α·'), sup(t('q'), t('2'))], t('g')), t('3')), t(' ;   '),
                     sub(t('i'), t('k')), t(' = '),
                     phan_so(sup(t('Q'), t('2')), [sub_sup(t('ω'), t('k'), t('2')), t('·'), sub_sup(t('C'), t('k'), t('2')), t('·'),
                                                   sub(t('R'), t('k'))]),
                     t(' ;   Q = '), sub(t('ω'), t('0')), t('·'), sub(t('C'), t('0')), t('·'), can([sub(t('R'), t('0')), t('·i')])],
                 dem.c())

    def nhan(c):
        if c['ten'].startswith('Lũ thiết kế'):
            return 'Q (lũ thiết kế P = 1%)'
        if c['ten'].startswith('Lũ kiểm tra'):
            return 'Lũ kiểm tra P = 0,2%% (%sQ)' % s2(c['Q'] / tk['Q'])
        return c['ten']
    b_cap = dem.bg()
    bang(d, '%s Các cấp lưu lượng và thông số dòng chảy trên dốc nước' % b_cap,
         ['Cấp', 'Q (m^{3}/s)', 'H (m)', 'MNL (m)', 'h_{c} (m)', 'h_{k} (m)', 'i_{k}', 'h_{0} (m)', 'h cuối dốc (m)',
          'v cuối dốc (m/s)'],
         [[nhan(c), s2(c['Q']), s3(c['H']), s2(c['MNL']), s3(c['hc']), s3(c['ik']['hk']), so(c['ik']['ik'], 5), s3(c['h0']['h0']),
           s3(c['bang'][-1]['h']), s2(c['bang'][-1]['v'])] for c in cap],
         cot=[1871, 800, 700, 800, 700, 700, 850, 700, 850, 1100], co=10)
    p('Ở mọi cấp lưu lượng i = %s %% lớn hơn nhiều i_{k} (%s ÷ %s) và h_{0} < h_{k}: dốc nước là dốc dốc, dòng chảy trên dốc là '
      'dòng xiết; độ sâu đầu dốc h_{c} nhỏ hơn h_{k} của đầu dốc nên sau bậc hạ dòng chảy đã ở trạng thái xiết. Trên đoạn thu hẹp '
      '%s → %s m độ sâu tăng do bề rộng giảm, trên đoạn lăng trụ độ sâu giảm dần về h_{0} (đường nước hạ b_{II}) nhưng chưa đạt '
      'h_{0} ở cuối dốc. Đường mặt nước tính bằng phương pháp sai phân theo phương trình:'
      % (so(TC['i_doc'] * 100, 0), so(min(c['ik']['ik'] for c in cap), 4), so(max(c['ik']['ik'] for c in cap), 4),
         so(TC['b_doc_dau'], 0), so(TC['b_doc_cuoi'], 0)), truoc=6)
    phuong_trinh(d, [phan_so(t('ΔE'), t('Δl')), t(' = i − '), gach(t('J')), t(' ;   '), gach(t('J')), t(' = '),
                     phan_so([sub(t('J'), t('i')), t(' + '), sub(t('J'), t('i+1'))], t('2')), t(' ;   J = '),
                     sup(ngoac(phan_so(t('v'), [t('C'), can(t('R'))])), t('2')), t(' ;   C = '), phan_so(t('1'), t('n')), t('·'),
                     sup(t('R'), phan_so(t('1'), t('6')))], dem.c())
    kq['ct_dmn'] = dem.ct - 1
    p('Dốc có bề rộng thay đổi nên chia dốc thành các đoạn có chiều dài chọn trước (%s m trên đoạn thu hẹp, %s m trên đoạn lăng '
      'trụ), giả thiết độ sâu cuối đoạn, tính lại Δl theo (4-%d) và thay đổi h tới khi Δl tính bằng Δl chọn; n = %s (bê tông). Bảng '
      'tính đường mặt nước của 5 cấp lưu lượng ở cuối mục. Độ sâu có kể hàm khí:'
      % (so(TC['L_chuyen_tiep'] / T.N_BUOC_LT, 1), so((TC['L_doc'] - TC['L_chuyen_tiep']) / T.N_BUOC_LT, 0), dem.ct - 1,
         so(T.N_NHAM, 3)))
    phuong_trinh(d, [sub(t('h'), t('hk', True)), t(' = h·'), ngoac([t('1 + '), phan_so(t('v'), t('100'))])], dem.c())
    kq['ct_hk'] = dem.ct - 1
    p('Chiều cao tường bên kiểm tra với lũ thiết kế và lũ kiểm tra (Bảng %s); các cấp lưu lượng nhỏ hơn có mặt nước thấp hơn nên '
      'không khống chế chiều cao tường.' % ('4.%d' % dem.b))
    b_doc = dem.bg()
    bang(d, '%s Kết quả kiểm tra thủy lực dốc nước (lũ thiết kế q = %s m^{3}/s; lũ kiểm tra q = %s m^{3}/s)'
         % (b_doc, s2(tk['Q']), s2(kt['Q'])),
         ['x (m)', 'B (m)', 'h TK (m)', 'h_{hk} TK (m)', 'v TK (m/s)', 'h KT (m)', 'h_{hk} KT (m)', 'v KT (m/s)',
          'Tường bên (m)', 'Độ vượt cao TK (m)', 'Độ vượt cao KT (m)', 'Tường cần h_{hk} + a (m)'],
         [[so(r['x'], 0), s2(r['B']), s3(r['h_tk']), s3(r['hhk_tk']), s2(r['v_tk']), s3(r['h_kt']), s3(r['hhk_kt']),
           s2(r['v_kt']), s2(r['ht']) + ('*' if r['gia_thiet'] else ''), s2(r['du_tk']), s2(r['du_kt']), s2(r['ht_can'])]
          for r in doc_], cot=[560, 680, 700, 760, 700, 700, 760, 700, 760, 880, 880, 991], co=9, le=40)
    doan(d, '* Chiều cao tường bên x = 0 ÷ 40 m theo tài liệu hiện trạng công trình; sau x = 40 m tạm lấy 2,5 m như đoạn '
            'x = 10 ÷ 40 m (không có bản vẽ tràn). Độ vượt cao = chiều cao tường − h_{hk}; tường cần = h_{hk} lớn nhất + a, a = %s m.'
            % s2(T.A_TUONG), co=10, nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT, truoc=0)
    du_min = min(doc_, key=lambda r: r['du_kt'])
    v_max = max(max(r['v'] for r in c['bang']) for c in cap)
    p('Mặt nước có kể hàm khí thấp hơn đỉnh tường bên ở mọi mặt cắt; độ vượt cao nhỏ nhất %s m tại x = %s m với lũ kiểm tra '
      '(lớn hơn 0,5 m) – tường bên dốc nước đủ chiều cao. Lưu tốc lớn nhất trên dốc ở cả 5 cấp lưu lượng là %s m/s (cuối dốc, '
      'lũ kiểm tra), nhỏ hơn lưu tốc cho phép không xói của dốc bê tông M200 – đồ án lấy [v]_{kx} = %s m/s, nhỏ hơn trị số '
      '19,2 ÷ 21,1 m/s theo TCVN 4118:2012 (phụ lục I, Bảng I5), thiên an toàn – nên dốc nước bảo đảm điều '
      'kiện không xói theo lý thuyết; tuy vậy bề mặt dốc không phẳng nhẵn, nước xả mang cát sạn nên nửa cuối dốc đã bị xói nhẹ '
      '(khảo sát hiện trạng) – cần sửa chữa bề mặt dốc (mục 4.4). Hình %s thể hiện đường mặt nước của 5 cấp lưu lượng.'
      % (s2(du_min['du_kt']), so(du_min['x'], 0), s2(v_max), so(K.V_KX, 1), '4.%d' % dem.h), truoc=6)
    gt = [r for r in doc_ if r['gia_thiet']]
    H_ANH = 2.0                                          # tường cuối dốc ước lượng theo ảnh hiện trạng (m)
    thieu = [r for r in gt if H_ANH - max(r['hhk_tk'], r['hhk_kt']) < T.A_TUONG]
    p('Đoạn x = %s ÷ %s m hồ sơ chưa ghi chiều cao tường bên; tài liệu hiện có chỉ gồm bình đồ tổng thể cụm đầu mối, không có '
      'bản vẽ tràn, nên đồ án tạm lấy chiều cao tường đoạn này bằng %s m như đoạn x = 10 ÷ 40 m. Tường cần tối thiểu ở đoạn này '
      '%s ÷ %s m (mặt nước giảm dần về cuối dốc), thấp hơn chiều cao giả thiết %s m – với giả thiết này tường đủ chiều cao. '
      'Ảnh hiện trạng cho thấy tường cuối dốc có thể chỉ cao khoảng %s m so với đáy (ước lượng theo bề rộng dốc 18 m): ở cuối '
      'dốc độ vượt cao lũ kiểm tra là %s m (%s 0,5 m)%s. Vì vậy kiến nghị khi khảo sát lập bản vẽ thi công đo kiểm tra chiều '
      'cao tường đoạn x = 40 ÷ 60 m; chỗ nào thấp hơn tường cần tối thiểu (cột cuối %s) thì nâng tường cho đủ.'
      % (so(gt[0]['x'], 0), so(gt[-1]['x'], 0), s2(gt[0]['ht']), s2(min(r['ht_can'] for r in gt)), s2(max(r['ht_can'] for r in gt)),
         s2(gt[0]['ht'] - max(r['ht_can'] for r in gt)), so(H_ANH, 1), s2(H_ANH - gt[-1]['hhk_kt']),
         'lớn hơn' if H_ANH - gt[-1]['hhk_kt'] >= T.A_TUONG else 'nhỏ hơn',
         (', nhưng nếu đoạn x = %s ÷ %s m tường cũng chỉ cao khoảng %s m thì độ vượt cao còn %s ÷ %s m, thiếu vài cm'
          % (so(thieu[0]['x'], 0), so(thieu[-1]['x'], 0), so(H_ANH, 1), s2(min(H_ANH - r['hhk_kt'] for r in thieu)),
             s2(max(H_ANH - r['hhk_kt'] for r in thieu)))) if thieu else '', b_doc))
    L_BD, Z_CUOI_BD = 61.87, 68.90                      # bình đồ: chân ngưỡng → cuối mũi phun (SV đo 09/11/2026); TP11
    i_tb = (TC['Z_nguong'] - TC['P2'] - Z_CUOI_BD) / L_BD
    dn_tb = K.do_nhay_doc(ts, i_tb)
    p('Trên bình đồ hiện trạng, từ chân ngưỡng (đầu dốc +73,30 m) đến cuối mũi phun (+68,90 m) theo tim tràn dài %s m: phù hợp '
      'dốc nước dài %s m với i = %s %% (QTVH; đáy cuối dốc +%s m) và mũi phun dài khoảng %s m nâng lên khoảng %s m. Nếu coi cả '
      'đoạn là dốc đều thì i = %s %%; tính lại với độ dốc này: lưu tốc cuối dốc lớn nhất %s m/s, độ vượt cao tường bên nhỏ nhất '
      '%s m – kết luận không đổi.'
      % (s2(L_BD), so(TC['L_doc'], 0), so(TC['i_doc'] * 100, 0), s2(TC['Z_nguong'] - TC['P2'] - TC['i_doc'] * TC['L_doc']),
         s1(L_BD - TC['L_doc']), s1(Z_CUOI_BD - (TC['Z_nguong'] - TC['P2'] - TC['i_doc'] * TC['L_doc'])), s1(i_tb * 100),
         s1(dn_tb['v_max']), s2(dn_tb['du_min'])))
    ve_mat_nuoc(kq, H_MN)
    hinh(d, H_MN, '%s Đường mặt nước (có kể hàm khí) trên dốc tràn chính – 5 cấp lưu lượng – và đỉnh tường bên' % dem.hinh())

    # ================================================================ 4.2.3
    tieu_de(d, '4.2.3 Tiêu năng mũi phun – hố xói', 2)
    p('Ảnh hiện trạng (đáy dốc nước, mũi phun; mũi phun, hố xói và kênh dẫn) cho thấy cuối dốc là mũi phun có răng: một hàng '
      'khoảng 8 ÷ 9 răng bê tông xen khe ngang hết bề rộng dốc, bước khoảng 2 m, răng cao khoảng 1,0 ÷ 1,35 m (bình đồ: hai điểm '
      '+70,24 / +70,25 ở hai đầu hàng răng, đáy cuối dốc trên tim +68,90 m). Hạ lưu là lòng dẫn đá tự nhiên, cây bụi mọc nhiều, '
      'chưa thấy hố xói sâu sát mũi phun. Răng chia luồng phun, tăng tiêu hao năng lượng nên hố xói thực tế nhỏ hơn khi tính '
      'theo công thức cho mũi phun trơn dưới đây – kết quả thiên an toàn. Kiểm tra theo TCVN 8420:2010 với 5 cấp lưu lượng như '
      'thủy lực dốc nước. '
      'Chiều dài phóng xa của luồng chảy khỏi mũi phun:')
    phuong_trinh(d, [t('L = k·'), sup(t('φ'), t('2')), t('·'), sub(t('Z'), t('H')), t('·sin2α·'),
                     ngoac([t('1 + '), can([t('1 + '), phan_so([sub(t('Z'), t('0')), t(' − '), sub(t('Z'), t('H'))],
                                                                 [sup(t('φ'), t('2')), t('·'), sub(t('Z'), t('H')), t('·'),
                                                                  sup(t('sin'), t('2')), t('α')])])], '[', ']'),
                     t(' ;   φ = '), phan_so(sub(t('v'), t('H')), can([t('2g·'), sub(t('Z'), t('H'))]))], dem.c())
    p('Trong đó: Z_{H} – chênh lệch mực nước thượng lưu với mũi phun; Z_{0} – chênh lệch mực nước thượng lưu và hạ lưu; v_{H} – '
      'lưu tốc trên mũi phun (lấy bằng lưu tốc cuối dốc); α = %s° – góc hắt; k = %s (Fr_{H} < 30 ÷ 35). Với φ tính như trên, '
      'φ^{2}·Z_{H} = v_{H}^{2}/2g nên L đúng là quỹ đạo vật ném xiên với lưu tốc thực ở mũi phun (thiên an toàn khi xét hố xói '
      'lấn về phía mũi phun). Khoảng cách tới điểm '
      'sâu nhất của hố xói và góc đổ của luồng chảy vào mặt nước hạ lưu:' % (so(K.ALPHA_MP, 0), so(K.K_MP, 0)))
    phuong_trinh(d, [sub(t('L'), t('1')), t(' = L + '), phan_so(t('t'), [t('tg'), sub(t('α'), t('1'))]), t(' ;   tg'),
                     sub(t('α'), t('1')), t(' = '),
                     can([t('tg'), sup(t('α'), t('2')), t(' + '),
                          phan_so([t('2g·'), ngoac([t('0,5·'), sub(t('h'), t('H')), t('·cosα + a − h')])],
                                  [sub_sup(t('v'), t('H'), t('2')), t('·co'), sup(t('s'), t('2')), t('α')])])], dem.c())
    p('với a − h = Z_{mũi} − MNHL – chiều cao mũi phun trên mặt nước hạ lưu. Chiều sâu xói lòng dẫn đá nứt nẻ, chiều sâu hố xói '
      'dưới đáy lòng dẫn và chiều dài hố (mái thượng lưu 1:2,2, hạ lưu 1:1,5):')
    phuong_trinh(d, [t('t = '), ngoac([t('0,1 + 0,45·'), can(t('Fr'), t('4'))]), t('·'), can(t('Fr')), t('·'),
                     can(phan_so(t('h'), t('d')), t('4')), t('·'), sub(t('h'), t('H')), t(' ;   Fr = '),
                     phan_so(sub_sup(t('v'), t('H'), t('2')), [t('g·'), sub(t('h'), t('H'))])], dem.c())
    phuong_trinh(d, [sub(t('h'), t('x')), t(' = t − h + c ;   '), sub(t('b'), t('x')), t(' = 2,5·h + 3,7·'), sub(t('h'), t('x')),
                     t(' ;   '), sub(t('x'), t('TL', True)), t(' = '), sub(t('L'), t('1')), t(' − 1,25·h − 2,2·'),
                     sub(t('h'), t('x'))], dem.c())
    p('Trong đó: h – chiều sâu nước hạ lưu; d = c = %s m – kích thước khối đá nền tương đương (đá nứt nẻ); x_{TL} – khoảng cách '
      'từ mũi phun tới mép thượng lưu hố xói ở cao trình đáy lòng dẫn (điểm sâu nhất ở giữa đáy hố dài 2,5h). Cao trình mũi phun '
      '+%s m (bình đồ hiện trạng: cuối dốc tại tim tràn +68,90 m); góc hắt α, kích thước khối đá d = c, hệ số mái hố xói và quan '
      'hệ lưu lượng – chiều sâu, mực nước hạ lưu kênh dẫn (Bảng %s) theo tài liệu hiện trạng công trình, nội suy theo Q.'
      % (s2(K.D_DA), s2(K.Z_MUI), '4.%d' % dem.b))
    bang(d, '%s Quan hệ lưu lượng – mực nước hạ lưu kênh dẫn sau mũi phun' % dem.bg(),
         ['Q (m^{3}/s)'] + [so(q, 0) for q in K.KD_Q],
         [['Chiều sâu nước hạ lưu h (m)'] + [s2(h) for h in K.KD_h],
          ['Mực nước hạ lưu MNHL (m)'] + [s2(z) for z in K.KD_MNHL]],
         cot=[2771] + [1050] * len(K.KD_Q), co=10)
    bang(d, '%s Kết quả tính toán hố xói sau mũi phun' % dem.bg(),
         ['Q (m^{3}/s)', 'MNL (m)', 'h_{H} (m)', 'v_{H} (m/s)', 'Fr_{H}', 'φ', 'Z_{H} (m)', 'Z_{0} (m)', 'h (m)', 'L (m)', 'α_{1} (°)',
          't (m)', 'h_{x} (m)', 'L_{1} (m)', 'b_{x} (m)', 'x_{TL} (m)'],
         [[s2(c['Q']), s2(c['MNL']), s3(c['mp']['hH']), s2(c['mp']['vH']), s2(c['mp']['Fr']), s3(c['mp']['phi']),
           s2(c['mp']['ZH']), s2(c['mp']['Z0']), s2(c['mp']['h']), s2(c['mp']['L']), so(c['mp']['a1'], 1), s2(c['mp']['t']),
           s2(c['mp']['hx']), s2(c['mp']['L1']), s2(c['mp']['bx']), s2(c['mp']['x_tl'])] for c in cap],
         cot=[640, 620, 540, 540, 500, 520, 540, 540, 500, 540, 520, 520, 540, 580, 580, 580], co=8, cot_trai=(), le=30)
    x_min = min(c['mp']['x_tl'] for c in cap)
    p('Mép thượng lưu của các hố xói cách mũi phun %s ÷ %s m (x_{TL} > 0): chân mũi phun nằm ngoài đường bao hố xói (Hình %s), '
      'tiêu năng mũi phun bảo đảm an toàn. Hố xói sâu nhất %s m dưới đáy lòng dẫn, điểm sâu '
      'nhất cách mũi phun %s m (lũ kiểm tra). Không sửa chữa hạng mục tiêu năng; quan trắc địa hình hố xói sau mỗi trận lũ lớn.'
      % (s2(x_min), s2(max(c['mp']['x_tl'] for c in cap)), '4.%d' % dem.h, s2(kt['mp']['hx']), s2(kt['mp']['L1'])), truoc=6)
    ve_ho_xoi(kq, H_HX)
    hinh(d, H_HX, '%s Đường bao hố xói sau mũi phun tràn chính' % dem.hinh())

    # ================================================================ 4.2.4
    tieu_de(d, '4.2.4 Kết luận', 2)
    gach_dau(d, 'Khả năng xả: với lũ QTVH 2026, MNLTK = +%s m, MNLKT = +%s m (1 cửa tràn chính không mở: +%s / +%s m), thấp hơn '
                'đỉnh đập sau sửa chữa +%s m; hệ số co hẹp ε = %s (mố bên lượn tròn, 8 trụ pin đầu tròn).'
             % (s2(tk['MNL']), s2(kt['MNL']), s2(ts['P1_SuCo']['Z']), s2(ts['P02_SuCo']['Z']), s2(ts['Z_dinh_SC']), s3(e_lon)))
    gach_dau(d, 'Dốc nước (5 cấp lưu lượng): dòng xiết, tường bên đủ chiều cao (độ vượt cao nhỏ nhất %s m); lưu tốc lớn nhất %s '
                'm/s, không xói theo lý thuyết '
                'nhưng bề mặt nửa cuối dốc đã mài mòn – sửa chữa bề mặt bằng lớp vữa, bê tông chống mài mòn (mục 4.4).'
             % (s2(du_min['du_kt']), s2(v_max)))
    gach_dau(d, 'Tiêu năng mũi phun có răng – hố xói: chân mũi phun nằm ngoài đường bao hố xói (x_{TL} ≥ %s m, tính theo mũi phun '
                'trơn – thiên an toàn); không sửa chữa kết cấu; phát quang cây bụi, dọn đá lăn trong kênh dẫn hạ lưu trước mùa lũ, '
                'quan trắc hố xói sau lũ lớn.' % s2(x_min))
    gach_dau(d, 'Số liệu: P_{1}, P_{2}, cao độ đầu và cuối dốc, chiều dài từ chân ngưỡng đến cuối mũi phun (61,87 m) đã đối chiếu '
                'bình đồ hiện trạng; chiều cao tường bên đoạn x = 45 ÷ 60 m tạm lấy theo giả thiết 2,5 m (không có bản vẽ tràn). Kiến '
                'nghị khi khảo sát lập bản vẽ thi công: đo chiều cao tường đoạn x = 40 ÷ 60 m (ảnh hiện trạng cho thấy tường cuối '
                'dốc có thể chỉ khoảng 2,0 m – chỗ nào thấp hơn tường cần tối thiểu thì nâng tường) và kích thước răng mũi phun.')

    # ================================================================ PHỤ LỤC: bảng đường mặt nước
    tieu_de(d, 'Bảng tính đường mặt nước trên dốc nước – 5 cấp lưu lượng (dùng cho %s, %s)' % (b_cap, b_doc), 2)
    COT15 = [380, 560, 600, 580, 600, 600, 640, 600, 580, 600, 660, 660, 640, 600, 640]
    for c in cap:
        hang = []
        for k, r in enumerate(c['bang']):
            hang.append(['%d' % (k + 1), s2(r['B']), s3(r['h']), s2(r['v']), s3(r['E']), s3(r['dE']) if k else '', s2(r['w']),
                         s2(r['chi']), s3(r['R']), s2(r['C']), so(r['J'], 4), so(r['Jtb'], 4) if k else '',
                         so(r['i_J'], 4) if k else '', s2(r['dL']) if k else '', s2(r['x'])])
        bang(d, '%s Đường mặt nước – %s, q = %s m^{3}/s' % (dem.bg(), nhan(c), s2(c['Q'])),
             ['MC', 'B (m)', 'h (m)', 'v (m/s)', 'E (m)', 'ΔE (m)', 'ω (m^{2})', 'χ (m)', 'R (m)', 'C', 'J', 'J_{tb}', 'i − J_{tb}',
              'Δl (m)', 'L (m)'], hang, cot=COT15, co=8, cot_trai=(), le=30,
             dau=['Độ sâu đầu dốc h_{c} = %s m; i = %s %%; n = %s; B = %s → %s m (x = 0 ÷ %s m), sau đó %s m.'
                  % (s3(c['hc']), so(TC['i_doc'] * 100, 0), so(T.N_NHAM, 3), so(TC['b_doc_dau'], 0), so(TC['b_doc_cuoi'], 0),
                     so(TC['L_chuyen_tiep'], 0), so(TC['b_doc_cuoi'], 0))])
    muc_tran_su_co(d, dem, kq)
    d.save(RA)
    print('Đã ghi', os.path.abspath(RA))


if __name__ == '__main__':
    main()
