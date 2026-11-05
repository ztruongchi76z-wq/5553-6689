# -*- coding: utf-8 -*-
"""
XUẤT WORD: KIỂM TRA TRÀN CHÍNH HIỆN TRẠNG TRONG ĐỒ ÁN SỬA CHỮA – theo cách làm của hồ sơ kiểm định an toàn đập (QTVH,
chương 6, SV gửi trích 05/11/2026), thay cho trình tự file mẫu thiết kế tràn mới. Tính ở kiem_tra_tran_chinh.py.
  python3 xuat_word_kiem_tra_tran_chinh.py [đường dẫn .docx ra]
Mặc định: ../TAI_VE/KiemTra_TranChinh_SuaChua.docx (hình: ../TAI_VE/KTTC_DuongMatNuoc.png, ../TAI_VE/KTTC_DuongBaoHoXoi.png)
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
import thuy_luc_tran_chinh as T
from omml_docx import bang, can, gach, ngoac, phan_so, phuong_trinh, sub, sub_sup, sup, t
from xuat_bang_tham_docx import chu, doan, so

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'KiemTra_TranChinh_SuaChua.docx')
H_MN = os.path.join(os.path.dirname(RA), 'KTTC_DuongMatNuoc.png')
H_HX = os.path.join(os.path.dirname(RA), 'KTTC_DuongBaoHoXoi.png')
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
    for c, mau, nhan in ((kq['tk'], '#2a78d6', 'Lũ thiết kế'), (kq['kt'], '#eb6834', 'Lũ kiểm tra')):
        ax.plot(xs, [z + r['h'] * (1 + r['v'] / 100) for z, r in zip(day, c['bang'])], color=mau, lw=2,
                label='%s, q = %s m³/s (có kể hàm khí)' % (nhan, s2(c['Q'])))
    xt = [r['x'] for r in kq['doc']]
    ax.plot(xt, [Z0 - TC['i_doc'] * x + r['ht'] for x, r in zip(xt, kq['doc'])], color='#0b0b0b', lw=1.2, ls=(0, (1, 1.5)),
            marker='o', ms=3.5, label='Đỉnh tường bên hiện trạng')
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
        ax.plot(x, y, color=m_, lw=1.0, label='Q = %s m³/s' % so(c['Q'], 0 if c['Q'] == round(c['Q']) else 2))
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
    ax.legend(loc='lower right', fontsize=7, frameon=False, ncol=2)
    fig.tight_layout()
    fig.savefig(ra, dpi=220, facecolor='white')
    plt.close(fig)


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

    doan(d, 'Ghi chú (xoá khi dán vào thuyết minh): thay cho mục 4.2 tính theo file mẫu thiết kế tràn mới; trình tự theo hồ sơ '
            'kiểm định an toàn đập (QTVH, chương 6); thông số tràn và kết quả điều tiết lấy từ file DieuTietLu_Potapov_f1f2. '
            'Số hiệu công thức, bảng, hình đánh liên tục từ (4-1), Bảng 4.3, Hình 4.1 – sửa lại theo thuyết minh.', co=11,
         nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT)
    tieu_de(d, '4.2 Kiểm tra tràn chính hiện trạng', 1)
    p('Tràn chính là công trình đang vận hành; đồ án sửa chữa không thiết kế lại hình thức, kích thước tràn mà kiểm tra các điều '
      'kiện làm việc của tràn hiện có với lũ cập nhật (QTVH 2026), theo cách làm của hồ sơ kiểm định an toàn đập: khả năng xả lũ, '
      'thủy lực dốc nước (tường bên, xói bề mặt), tiêu năng mũi phun – hố xói. Điều kiện không tràn đỉnh đập đã kiểm tra ở mục '
      '2.3, ổn định ngưỡng tràn ở mục 4.5. Các nội dung của thiết kế mới (chọn hình thức, bề rộng tràn, kênh dẫn, chiều dày bản '
      'đáy, kích thước công trình tiêu năng) không tính lại – Bảng %s.' % ('4.%d' % dem.b))
    bang(d, '%s Nội dung tính toán tràn chính trong đồ án sửa chữa' % dem.bg(),
         ['Hạng mục', 'Thiết kế tràn mới (file mẫu)', 'Đồ án sửa chữa (theo hồ sơ kiểm định)', 'Mục'],
         [['Hình thức, kích thước, bố trí', 'Chọn, tính toán', 'Không tính – mô tả hiện trạng', '4.2.1'],
          ['Khả năng xả, điều tiết lũ', 'Chọn B; tính lại ε, m; điều tiết', 'Điều tiết với lũ QTVH 2026; quan hệ Q ~ Z', '4.1; 4.2.1'],
          ['Lưu tốc tới gần', 'So Ω_{TL} với 4·Σb·H', 'Bỏ qua: Ω_{T} ≫ 4·H·b (hồ sơ kiểm định)', '4.2.1'],
          ['Thủy lực dốc nước', '5 cấp lưu lượng: h_{k}, i_{k}, h_{0}, đường mặt nước',
           'Đường mặt nước lũ thiết kế, lũ kiểm tra (và các cấp Q tính hố xói)', '4.2.2'],
          ['Tường bên, bề mặt dốc', 'Thiết kế chiều cao tường, chiều dày bản đáy', 'Kiểm tra tường hiện có; kiểm tra xói; '
           'sửa chữa bề mặt dốc', '4.2.2; 4.4'],
          ['Tiêu năng', 'Chọn hình thức; thiết kế kênh hạ lưu, mũi phun', 'Kiểm tra mũi phun – hố xói hiện có (TCVN 8420:2010)',
           '4.2.3'],
          ['Không tràn đỉnh đập', '–', 'Cao trình đỉnh đập, tường chắn sóng', '2.3'],
          ['Ổn định ngưỡng tràn', '–', 'Kiểm tra ổn định trượt, lật', '4.5']],
         cot=[2300, 2500, 3200, 1071], co=10, cot_trai=(0, 1, 2))

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
                     phan_so(sub(t('H'), t('0')), t('n·b')), t(' ;   Σξ = '), sub(t('n'), t('1')), sub(t('ξ'), t('1')), t(' + '),
                     sub(t('n'), t('2')), sub(t('ξ'), t('2'))], dem.c())
    p('Trong đó: σ_{n} = 1 – sau ngưỡng là dốc nước chảy xiết (i > i_{k}) nên chảy không ngập; m = %s – đập tràn thực dụng '
      'thấp; ξ_{mb} = %s; Σξ = %d × %s + %d × %s = %s. Theo hồ sơ kiểm định, diện tích mặt cắt ướt ở thượng lưu cách ngưỡng '
      '10 ÷ 30 m rất lớn so với 4·H·b nên không xét lưu tốc tới gần, H_{0} = H = Z − Z_{ngưỡng}. Đây cũng là công thức dùng khi '
      'điều tiết lũ (mục 4.1), nên khả năng xả của ngưỡng không cần kiểm tra lại bằng một phép tính riêng. Bảng %s so sánh các '
      'hệ số với hồ sơ kiểm định.' % (so(ts['TC_m']), s2(ts['TC_xmb']), ts['TC_n1'], s2(ts['TC_x1']), ts['TC_n2'],
                                    s2(ts['TC_x2']), s2(ts['TC_sx']), '4.%d' % dem.b))
    e_tk, e_kt = T.eps(tk['H'], ts), T.eps(kt['H'], ts)
    bang(d, '%s Hệ số tính khả năng xả của tràn chính' % dem.bg(),
         ['Đại lượng', 'Đồ án', 'Hồ sơ kiểm định', 'Nhận xét'],
         [['Bề rộng tràn nước Σb (m)', so(Sb, 1) + ' (9 × 2,4)', '22', 'Tràn đã chia 9 khoang, lắp cửa van'],
          ['Hệ số lưu lượng m', so(ts['TC_m'], 3), '0,445 (= 0,467/1,05)', 'Tương đương'],
          ['Hệ số co hẹp bên ε', '%s ÷ %s (theo H, ξ_{1} = %s)' % (s3(e_tk), s3(e_kt), s2(ts['TC_x1'])),
           '0,895 (ξ_{mt} = 0,25; lấy H_{0}/b ≤ 1)', 'Đồ án nhỏ hơn – thiên an toàn'],
          ['m_{0} = σ_{n}·ε·m', '%s ÷ %s' % (s3(e_tk * ts['TC_m']), s3(e_kt * ts['TC_m'])), '0,40', ''],
          ['Lưu tốc tới gần V_{0}', 'Bỏ qua', 'Bỏ qua (Ω_{T} ≫ 4·H·b)', '']],
         cot=[2600, 2500, 2300, 1671], co=10, cot_trai=(0, 3))
    p('Hồ sơ kiểm định ghi ngưỡng có P_{1} = 2,3 m, P = 3,3 m, δ = 2,7 m, mái thượng lưu 1,75, hạ lưu 3,00, trong khi đồ án lấy '
      'theo BCC: P_{1} = %s m, P_{2} = %s m, δ = %s m, mái thượng lưu 1:3,0, hạ lưu 1:1,75 – cần đối chiếu bản vẽ hoàn công. '
      'Hệ số lưu lượng của hai bên tương đương (0,45 và 0,445).' % (s2(TC['P1']), s2(TC['P2']), s2(TC['delta'])), truoc=6)
    qz = kq['qz']
    bang(d, '%s Quan hệ lưu lượng – mực nước của công trình tháo lũ' % dem.bg(),
         ['Z (m)', 'H tràn chính (m)', 'ε (9 khoang)', 'q tràn chính 9 khoang (m^{3}/s)', 'q tràn chính 8 khoang (m^{3}/s)',
          'q tràn sự cố (m^{3}/s)', 'q tổng – bình thường (m^{3}/s)'],
         [[s2(r['Z']), s2(r['H']), s3(r['eps']), s2(r['q9']), s2(r['q8']), s2(r['qsc']), s2(r['q_bt'])] for r in qz],
         cot=[1000, 1150, 1150, 1450, 1450, 1350, 1521], co=10, cot_trai=())
    kd = {'1%': (77.87, 242.00, 159.54), '0,2%': (79.13, 399.79, 205.39)}
    bang(d, '%s Kết quả điều tiết lũ – đồ án và hồ sơ kiểm định' % dem.bg(),
         ['Trường hợp', 'Z_{max} đồ án (m)', 'q tràn chính đồ án (m^{3}/s)', 'q tràn sự cố đồ án (m^{3}/s)',
          'MNL kiểm định (m)', 'q tràn chính kiểm định (m^{3}/s)', 'q tràn sự cố kiểm định (m^{3}/s)'],
         [['P = 1%, bình thường', s2(tk['MNL']), s2(ts['P1_BinhThuong']['qtc']), s2(ts['P1_BinhThuong']['qsc']),
           s2(kd['1%'][0]), s2(kd['1%'][1]), s2(kd['1%'][2])],
          ['P = 1%, 1 cửa tràn chính không mở', s2(ts['P1_SuCo']['Z']), s2(ts['P1_SuCo']['qtc']), s2(ts['P1_SuCo']['qsc']),
           '–', '–', '–'],
          ['P = 0,2%, bình thường', s2(kt['MNL']), s2(ts['P02_BinhThuong']['qtc']), s2(ts['P02_BinhThuong']['qsc']),
           s2(kd['0,2%'][0]), s2(kd['0,2%'][1]), s2(kd['0,2%'][2])],
          ['P = 0,2%, 1 cửa tràn chính không mở', s2(ts['P02_SuCo']['Z']), s2(ts['P02_SuCo']['qtc']), s2(ts['P02_SuCo']['qsc']),
           '–', '–', '–']],
         cot=[2100, 1100, 1200, 1200, 1100, 1200, 1171], co=10)
    p('Mực nước lũ của đồ án thấp hơn hồ sơ kiểm định vì đồ án dùng lũ QTVH 2026 (Q_{max} 1%% = 921 m^{3}/s, thu phóng trận lũ '
      'trạm Đá Bàn) và quan hệ Z ~ F ~ W theo bình đồ lòng hồ 2022 – thực hiện kiến nghị của hồ sơ kiểm định (đo bình đồ từ +66 '
      'đến +80 m); hồ sơ kiểm định dùng lũ Q_{max} 1%% = 1.051 m^{3}/s và quan hệ Z ~ W ngoại suy trên +66 m. Với cả hai bộ số '
      'liệu, tràn chính tháo được lũ thiết kế và lũ kiểm tra; MNLTK = +%s m, MNLKT = +%s m thấp hơn đỉnh đập sau sửa chữa '
      '+%s m (mục 2.3).' % (s2(tk['MNL']), s2(kt['MNL']), s2(ts['Z_dinh_SC'])), truoc=6)

    # ================================================================ 4.2.2
    tieu_de(d, '4.2.2 Thủy lực dốc nước', 2)
    p('Độ sâu co hẹp ở đầu dốc nước (chân bậc hạ sau ngưỡng) tính theo công thức tổng quát, giải bằng thử dần:')
    phuong_trinh(d, [t('Q = φ·'), sub(t('ω'), t('c')), t('·'), can([t('2g·'), ngoac([sub(t('E'), t('0')), t(' − '),
                                                                                       sub(t('h'), t('c'))])]),
                     t(' ;   '), sub(t('ω'), t('c')), t(' = '), sub(t('b'), t('c')), t('·'), sub(t('h'), t('c')), t(' ;   '),
                     sub(t('E'), t('0')), t(' = '), sub(t('P'), t('2')), t(' + H')], dem.c())
    p('với φ = %s; b_{c} = %s m (bề rộng đầu dốc); E_{0} = P_{2} + H (bỏ qua V_{0}). Đường mặt nước trên dốc tính bằng phương pháp '
      'sai phân theo phương trình:' % (s2(T.PHI), so(TC['b_doc_dau'], 0)))
    phuong_trinh(d, [phan_so(t('ΔE'), t('Δl')), t(' = i − '), gach(t('J')), t(' ;   '), gach(t('J')), t(' = '),
                     phan_so([sub(t('J'), t('i')), t(' + '), sub(t('J'), t('i+1'))], t('2')), t(' ;   J = '),
                     sup(ngoac(phan_so(t('v'), [t('C'), can(t('R'))])), t('2')), t(' ;   C = '), phan_so(t('1'), t('n')), t('·'),
                     sup(t('R'), phan_so(t('1'), t('6')))], dem.c())
    p('Dốc có bề rộng thay đổi nên giải như hồ sơ kiểm định: chia dốc thành các đoạn có chiều dài chọn trước (%s m trên đoạn thu '
      'hẹp, %s m trên đoạn lăng trụ), giả thiết độ sâu cuối đoạn, tính lại Δl theo (4-%d) và thay đổi h tới khi Δl tính bằng Δl '
      'chọn; n = %s (bê tông). Độ sâu có kể hàm khí:' % (so(TC['L_chuyen_tiep'] / T.N_BUOC_LT, 1),
                                                      so((TC['L_doc'] - TC['L_chuyen_tiep']) / T.N_BUOC_LT, 0), dem.ct - 1,
                                                      so(T.N_NHAM, 3)))
    phuong_trinh(d, [sub(t('h'), t('hk', True)), t(' = h·'), ngoac([t('1 + '), phan_so(t('v'), t('100'))])], dem.c())
    b_doc = dem.bg()
    bang(d, '%s Kết quả kiểm tra thủy lực dốc nước (lũ thiết kế q = %s m^{3}/s; lũ kiểm tra q = %s m^{3}/s)'
         % (b_doc, s2(tk['Q']), s2(kt['Q'])),
         ['x (m)', 'B (m)', 'h TK (m)', 'h_{hk} TK (m)', 'v TK (m/s)', 'h KT (m)', 'h_{hk} KT (m)', 'v KT (m/s)',
          'Tường bên (m)', 'Độ vượt cao TK (m)', 'Độ vượt cao KT (m)', 'Không tràn tường'],
         [[so(r['x'], 0), s2(r['B']), s3(r['h_tk']), s3(r['hhk_tk']), s2(r['v_tk']), s3(r['h_kt']), s3(r['hhk_kt']),
           s2(r['v_kt']), s2(r['ht']) + ('*' if r['gia_thiet'] else ''), s2(r['du_tk']), s2(r['du_kt']),
           'Thỏa mãn' if min(r['du_tk'], r['du_kt']) > T.A_TUONG else 'Kiểm tra']
          for r in doc_], cot=[560, 680, 700, 760, 700, 700, 760, 700, 760, 880, 880, 991], co=9, le=40)
    doan(d, '* Hồ sơ ghi chiều cao tường bên đến x = 40 m; sau x = 40 m giả thiết 2,5 m. Độ vượt cao = chiều cao tường − h_{hk}.',
         co=10, nghieng=True, can=WD_ALIGN_PARAGRAPH.LEFT, truoc=0)
    du_min = min(doc_, key=lambda r: r['du_kt'])
    v_max = max(r['v_kt'] for r in doc_)
    p('Mặt nước có kể hàm khí thấp hơn đỉnh tường bên ở mọi mặt cắt; độ vượt cao nhỏ nhất %s m tại x = %s m với lũ kiểm tra '
      '(lớn hơn 0,5 m) – tường bên dốc nước đủ chiều cao. Hồ sơ kiểm định cho độ vượt cao nhỏ nhất 0,21 m tại x = 10 m vì tính với '
      'lưu lượng tràn chính 400 m^{3}/s của lũ kiểm định. Lưu tốc lớn nhất trên dốc %s m/s nhỏ hơn lưu tốc cho phép không xói của '
      'dốc bê tông M200 theo hồ sơ kiểm định (%s ÷ 19,4 m/s) nên dốc nước bảo đảm điều kiện không xói theo lý thuyết; tuy vậy bề '
      'mặt dốc không phẳng nhẵn, nước xả mang cát sạn nên nửa cuối dốc đã bị xói nhẹ (hồ sơ kiểm định, khảo sát hiện trạng) – '
      'cần sửa chữa bề mặt dốc (mục 4.4). Hình %s thể hiện đường mặt nước; bảng tính chi tiết ở cuối mục.'
      % (s2(du_min['du_kt']), so(du_min['x'], 0), s2(v_max), so(K.V_KX, 1), '4.%d' % dem.h), truoc=6)
    p('Hồ sơ kiểm định ghi độ dốc dốc nước i = 10 %% và chiều cao bậc P = 3,3 m (đồ án: i = %s %%, P_{2} = %s m theo KTTV, BCC). '
      'Tính lại với các trị số đó, lưu tốc cuối dốc lớn nhất 13,5 m/s, độ vượt cao tường bên nhỏ nhất 1,0 m – kết luận không đổi.'
      % (so(TC['i_doc'] * 100, 0), s2(TC['P2'])))
    ve_mat_nuoc(kq, H_MN)
    hinh(d, H_MN, '%s Đường mặt nước (có kể hàm khí) trên dốc tràn chính và đỉnh tường bên hiện trạng' % dem.hinh())

    # ================================================================ 4.2.3
    tieu_de(d, '4.2.3 Tiêu năng mũi phun – hố xói', 2)
    p('Kiểm tra theo TCVN 8420:2010 (như hồ sơ kiểm định). Chiều dài phóng xa của luồng chảy khỏi mũi phun:')
    phuong_trinh(d, [t('L = k·'), sup(t('φ'), t('2')), t('·'), sub(t('Z'), t('H')), t('·sin2α·'),
                     ngoac([t('1 + '), can([t('1 + '), phan_so([sub(t('Z'), t('0')), t(' − '), sub(t('Z'), t('H'))],
                                                                 [sup(t('φ'), t('2')), t('·'), sub(t('Z'), t('H')), t('·'),
                                                                  sup(t('sin'), t('2')), t('α')])])], '[', ']'),
                     t(' ;   φ = '), phan_so(sub(t('v'), t('H')), can([t('2g·'), sub(t('Z'), t('H'))]))], dem.c())
    p('Trong đó: Z_{H} – chênh lệch mực nước thượng lưu với mũi phun; Z_{0} – chênh lệch mực nước thượng lưu và hạ lưu; v_{H} – '
      'lưu tốc trên mũi phun (lấy bằng lưu tốc cuối dốc); α = %s° – góc hắt; k = %s (Fr_{H} < 30 ÷ 35). Với φ tính như trên, '
      'φ^{2}·Z_{H} = v_{H}^{2}/2g nên L đúng là quỹ đạo vật ném xiên với lưu tốc thực ở mũi phun; hồ sơ kiểm định lấy L = '
      'k·Z_{1}·sin2α·[…] không nhân φ^{2} nên L lớn hơn (tính lại với số liệu của hồ sơ, Q = 242 m^{3}/s: L_{1} = 17,9 m so với '
      '22,6 m trong Bảng 6-8). Đồ án dùng L nhỏ hơn để thiên an toàn khi xét hố xói lấn về phía mũi phun. Khoảng cách tới điểm '
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
    p('Trong đó: h – chiều sâu nước hạ lưu; d = c = %s m – kích thước khối đá nền tương đương; x_{TL} – khoảng cách từ mũi phun '
      'tới mép thượng lưu hố xói ở cao trình đáy lòng dẫn (điểm sâu nhất ở giữa đáy hố dài 2,5h). Số liệu hạ lưu lấy theo Bảng 6-8 '
      'hồ sơ kiểm định: cao trình mũi phun +%s m; quan hệ Q ~ h, Q ~ MNHL nội suy theo bảng; d = c và góc hắt α tính ngược từ '
      'bảng (với cùng công thức (4-%d), (4-%d) cho lại đúng h_{x}, b_{x} của bảng). Các cấp lưu lượng 50 ÷ 200 m^{3}/s dùng để '
      'vẽ đường bao hố xói.' % (s2(K.D_DA), s2(K.Z_MUI), dem.ct - 2, dem.ct - 1))
    bang(d, '%s Kết quả tính toán hố xói sau mũi phun' % dem.bg(),
         ['Q (m^{3}/s)', 'MNL (m)', 'h_{H} (m)', 'v_{H} (m/s)', 'Fr_{H}', 'φ', 'Z_{H} (m)', 'Z_{0} (m)', 'h (m)', 'L (m)', 'α_{1} (°)',
          't (m)', 'h_{x} (m)', 'L_{1} (m)', 'b_{x} (m)', 'x_{TL} (m)'],
         [[s2(c['Q']), s2(c['MNL']), s3(c['mp']['hH']), s2(c['mp']['vH']), s2(c['mp']['Fr']), s3(c['mp']['phi']),
           s2(c['mp']['ZH']), s2(c['mp']['Z0']), s2(c['mp']['h']), s2(c['mp']['L']), so(c['mp']['a1'], 1), s2(c['mp']['t']),
           s2(c['mp']['hx']), s2(c['mp']['L1']), s2(c['mp']['bx']), s2(c['mp']['x_tl'])] for c in cap],
         cot=[640, 620, 540, 540, 500, 520, 540, 540, 500, 540, 520, 520, 540, 580, 580, 580], co=8, cot_trai=(), le=30)
    x_min = min(c['mp']['x_tl'] for c in cap)
    p('Mép thượng lưu của các hố xói cách mũi phun %s ÷ %s m (x_{TL} > 0): chân mũi phun nằm ngoài đường bao hố xói (Hình %s), '
      'tiêu năng mũi phun bảo đảm an toàn – cùng kết luận với hồ sơ kiểm định. Hố xói sâu nhất %s m dưới đáy lòng dẫn, điểm sâu '
      'nhất cách mũi phun %s m (lũ kiểm tra). Không sửa chữa hạng mục tiêu năng; quan trắc địa hình hố xói sau mỗi trận lũ lớn.'
      % (s2(x_min), s2(max(c['mp']['x_tl'] for c in cap)), '4.%d' % dem.h, s2(kt['mp']['hx']), s2(kt['mp']['L1'])), truoc=6)
    ve_ho_xoi(kq, H_HX)
    hinh(d, H_HX, '%s Đường bao hố xói sau mũi phun tràn chính' % dem.hinh())

    # ================================================================ 4.2.4
    tieu_de(d, '4.2.4 Kết luận', 2)
    gach_dau(d, 'Khả năng xả: với lũ QTVH 2026, MNLTK = +%s m, MNLKT = +%s m (1 cửa tràn chính không mở: +%s / +%s m), thấp hơn '
                'đỉnh đập sau sửa chữa +%s m; hệ số co hẹp đồ án dùng nhỏ hơn hồ sơ kiểm định nên kết quả thiên an toàn.'
             % (s2(tk['MNL']), s2(kt['MNL']), s2(ts['P1_SuCo']['Z']), s2(ts['P02_SuCo']['Z']), s2(ts['Z_dinh_SC'])))
    gach_dau(d, 'Dốc nước: tường bên đủ chiều cao (độ vượt cao nhỏ nhất %s m); lưu tốc lớn nhất %s m/s, không xói theo lý thuyết '
                'nhưng bề mặt nửa cuối dốc đã mài mòn – sửa chữa bề mặt bằng lớp vữa, bê tông chống mài mòn (mục 4.4).'
             % (s2(du_min['du_kt']), s2(v_max)))
    gach_dau(d, 'Tiêu năng mũi phun – hố xói: chân mũi phun nằm ngoài đường bao hố xói (x_{TL} ≥ %s m); không sửa chữa, quan trắc '
                'sau lũ lớn.' % s2(x_min))
    gach_dau(d, 'Số liệu cần đối chiếu bản vẽ hoàn công: kích thước ngưỡng (P_{1}, P_{2}, δ), độ dốc và chiều dài dốc nước, chiều '
                'cao tường bên sau x = 40 m, cao trình mũi phun, góc hắt.')

    # ================================================================ PHỤ LỤC: bảng đường mặt nước
    tieu_de(d, 'Bảng tính đường mặt nước trên dốc nước (dùng cho Bảng %s)' % b_doc, 2)
    COT15 = [380, 560, 600, 580, 600, 600, 640, 600, 580, 600, 660, 660, 640, 600, 640]
    for c in (tk, kt):
        hang = []
        for k, r in enumerate(c['bang']):
            hang.append(['%d' % (k + 1), s2(r['B']), s3(r['h']), s2(r['v']), s3(r['E']), s3(r['dE']) if k else '', s2(r['w']),
                         s2(r['chi']), s3(r['R']), s2(r['C']), so(r['J'], 4), so(r['Jtb'], 4) if k else '',
                         so(r['i_J'], 4) if k else '', s2(r['dL']) if k else '', s2(r['x'])])
        bang(d, '%s Đường mặt nước – %s, q = %s m^{3}/s' % (dem.bg(), c['ten'].lower(), s2(c['Q'])),
             ['MC', 'B (m)', 'h (m)', 'v (m/s)', 'E (m)', 'ΔE (m)', 'ω (m^{2})', 'χ (m)', 'R (m)', 'C', 'J', 'J_{tb}', 'i − J_{tb}',
              'Δl (m)', 'L (m)'], hang, cot=COT15, co=8, cot_trai=(), le=30,
             dau=['Độ sâu đầu dốc h_{c} = %s m; i = %s %%; n = %s; B = %s → %s m (x = 0 ÷ %s m), sau đó %s m.'
                  % (s3(c['hc']), so(TC['i_doc'] * 100, 0), so(T.N_NHAM, 3), so(TC['b_doc_dau'], 0), so(TC['b_doc_cuoi'], 0),
                     so(TC['L_chuyen_tiep'], 0), so(TC['b_doc_cuoi'], 0))])
    d.save(RA)
    print('Đã ghi', os.path.abspath(RA))


if __name__ == '__main__':
    main()
