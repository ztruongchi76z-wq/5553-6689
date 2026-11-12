# -*- coding: utf-8 -*-
"""
XUẤT WORD: BIỆN LUẬN CHƯƠNG 3 – GIỮ NGUYÊN HAY MỞ RỘNG TRÀN XẢ LŨ CHÍNH; GIỮ NGUYÊN CỐNG LẤY NƯỚC (cống ngầm)
Theo gợi ý GVHD (SV gửi 02/11/2026): so sánh với mực nước lũ của thiết kế ban đầu, điều chỉnh dòng chảy (vận hành cửa van)
để đỉnh lũ không vượt đỉnh đập; cống ngầm chôn sâu, còn tốt theo BCNCKT nên giữ nguyên.
Mục 3.3.1 không nhắc BCNCKT (SV yêu cầu 07/11/2026): chỉ so với thiết kế ban đầu; Bảng 3.1 theo mẫu bảng so sánh các phương
án trong thuyết minh của SV (= sheet PhuongAn, mục C của file điều tiết lũ).
Số liệu (bản 09/11/2026 – trụ pin đầu tròn ξ₁ = ξ₂ = 0,45, mố bên tràn chính ξmb = 0,70, tràn sự cố 1,00, ε lấy H/b ≤ 1):
  - TK 1996: MNLTK +77,24 (Q1% = 847 m³/s, tràn 25 m).
  - Đồ án: lũ QTVH 72 giờ, tràn chính 9 × 2,4 m + tràn sự cố 5 m, hồ ở MNDBT khi lũ đến (DieuTietLu_Potapov_f1f2 –
    VH_mo = 1: 77,457 / 78,388; 1 cửa không mở 77,505 / 78,455; VH_mo = 2: 77,517 / 78,429; VH_mo = 2 + 1 cửa:
    77,567 / 78,495 – LibreOffice tính lại file với ô VH_mo = 2).
  - Phương án so sánh (cùng chương trình, mở hết cửa): PA1 10 × 2,4 = 24 m: 77,411 / 78,329 (1 cửa không mở 77,457 /
    78,388); PA2 25 m 3 khoang: 77,368 / 78,271; qmax lũ P = 1% / 0,2%: 309,1 / 435,9; 322,7 / 457,5; 338,0 / 477,3 m³/s.
  - Dự trữ: lũ P = 0,2% tăng 4,6% thì MNLKT + 0,3 m = +78,85; tăng 13,4% thì MNLKT chạm +78,85 (điều tiết Python cùng
    công thức, kiểm tra lại Zmax của file Excel sai khác ≤ 0,3 mm).
  - Cống: BTCT 1,8 × 2,0 m, dài 67 m, đặt trên lớp đá II; đỉnh cống tại tim đập ≈ +67,10; γ đất đắp 20,6 kN/m³ (Chương 6);
    Q ở MNC = 7,31 > 6,54 m³/s.
  python3 xuat_word_bien_luan_chuong3.py [đường dẫn .docx ra]
Mặc định: ../TAI_VE/BienLuan_Chuong3_TranChinh_Cong.docx
"""
import os
import sys

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from xuat_bang_tham_docx import bang, doan, so

DIR = os.path.dirname(os.path.abspath(__file__))
RA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, '..', 'TAI_VE', 'BienLuan_Chuong3_TranChinh_Cong.docx')

Z_DINH_TK, Z_DINH_SC, Z_TUONG = 78.50, 78.85, 79.40
MNLTK_1996 = 77.24
# Zmax (m) mở hết cửa từ đầu trận lũ: lũ P = 1% bình thường, P = 1% 1 cửa không mở, P = 0,2% bình thường, P = 0,2% 1 cửa
# không mở (PA2 không có cửa van – None); qmax (m³/s) P = 1%, 0,2% bình thường
PHUONG_AN = [
    dict(ten='Hiện trạng 9 × 2,4 m', Sb=21.6, cua='9 cửa', Z=(77.4573, 77.5051, 78.3881, 78.4545), q=(309.1404, 435.9204)),
    dict(ten='PA1 – xây thêm 1 khoang (Σb = 24 m)', Sb=24.0, cua='10 cửa', Z=(77.4109, 77.4573, 78.3289, 78.3881),
         q=(322.6693, 457.4919)),
    dict(ten='PA2 – mở rộng về 25 m (3 khoang tự do)', Sb=25.0, cua='Không (chảy tự do)', Z=(77.3684, None, 78.2711, None),
         q=(337.9516, 477.2518)),
]
HT = (PHUONG_AN[0]['Z'][0], PHUONG_AN[0]['Z'][2])          # tràn hiện trạng: MNLTK, MNLKT
VAN_HANH = [('Kéo hết cửa van từ đầu trận lũ', 78.3881),
            ('Giữ MNDBT, khi lũ đến lớn hơn khả năng xả mới kéo hết cửa', 78.4291),
            ('Kéo hết cửa, 1 cửa tràn chính không mở được', 78.4545),
            ('Giữ MNDBT và 1 cửa tràn chính không mở được', 78.4952)]
K_DU_TRU = (4.6, 13.4)                      # % tăng lũ P = 0,2%: MNLKT + 0,3 m = +78,85; MNLKT = +78,85
GAMMA, Z_DINH_CONG = 20.6, 67.10


def tieu_de(d, chu_):
    return doan(d, chu_, dam=True, nghieng=True, truoc=12, sau=6, can=WD_ALIGN_PARAGRAPH.LEFT)


def main():
    d = docx.Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3.0), Cm(2.0), Cm(2.5), Cm(2.5)
    st = d.styles['Normal']
    st.font.name, st.font.size = 'Times New Roman', Pt(13)
    p = lambda txt: doan(d, txt, thut=1.0)

    # ------------------------------------------------------------------ tràn xả lũ chính
    tieu_de(d, '3.3.1 Tràn xả lũ chính – giữ nguyên hay mở rộng')
    p('Theo thiết kế ban đầu (năm 1996), tràn chính rộng 25 m, gồm 3 khoang chảy tự do. Trong quá trình khai thác, đơn vị '
      'quản lý đã chia tràn thành 9 khoang × 2,4 m (Σb = 21,6 m) và lắp cửa van để tích thêm nước cuối mùa lũ; đồng thời '
      'lũ thiết kế được cập nhật tăng (Q_{1%} từ 847 m^{3}/s năm 1996 lên 921 m^{3}/s theo QTVH 2026). Cả hai thay đổi đều '
      'làm giảm khả năng tháo lũ so với thiết kế ban đầu, vì vậy cần xem xét giữ nguyên quy mô tràn chính hay phải mở rộng. '
      'Tiêu chí đánh giá: mực nước lũ so với mực nước lũ của thiết kế ban đầu; đỉnh lũ không vượt đỉnh đập, '
      'với điều kiện của TCVN 8216:2018 điều 10.1.5: đỉnh đập (không kể tường chắn sóng) ≥ MNLKT + 0,3 m, đỉnh tường chắn '
      'sóng ≥ cao trình đỉnh yêu cầu Z_{đ}.')
    doan(d, 'a) So sánh với mực nước lũ của thiết kế ban đầu', dam=True, nghieng=False, truoc=6, sau=2,
         can=WD_ALIGN_PARAGRAPH.LEFT)
    p('Hai phương án mở rộng tràn chính được tính cùng chương trình điều tiết lũ, cùng điều kiện vận hành với tràn hiện '
      'trạng: PA1 – xây thêm 1 khoang 2,4 m có cửa van (Σb = 24 m); PA2 – đập bỏ 6 trụ phân khoang, trở về 3 khoang chảy tự do '
      'như thiết kế ban đầu (Σb = 25 m). Kết quả so sánh cho trong Bảng 3.1.')
    A, B, C = PHUONG_AN
    z2 = lambda z: '–' if z is None else so(z)
    hang = [['Bề rộng thông nước Σb'] + [so(x['Sb']) for x in PHUONG_AN] + ['m'],
            ['Cửa van tràn chính'] + [x['cua'] for x in PHUONG_AN] + [''],
            ['MNLTK – Zmax lũ P = 1%, vận hành bình thường'] + [z2(x['Z'][0]) for x in PHUONG_AN] + ['m'],
            ['Zmax lũ P = 1%, sự cố 1 cửa không mở'] + [z2(x['Z'][1]) for x in PHUONG_AN] + ['m'],
            ['MNLKT – Zmax lũ P = 0,2%, vận hành bình thường'] + [z2(x['Z'][2]) for x in PHUONG_AN] + ['m'],
            ['Zmax lũ P = 0,2%, sự cố 1 cửa không mở'] + [z2(x['Z'][3]) for x in PHUONG_AN] + ['m'],
            ['q_{max} lũ P = 1%, vận hành bình thường'] + [so(x['q'][0], 1) for x in PHUONG_AN] + ['m^{3}/s'],
            ['q_{max} lũ P = 0,2%, vận hành bình thường'] + [so(x['q'][1], 1) for x in PHUONG_AN] + ['m^{3}/s'],
            ['MNLTK thấp hơn hiện trạng', '–'] + [so(100 * (A['Z'][0] - x['Z'][0]), 1) for x in (B, C)] + ['cm'],
            ['MNLKT thấp hơn hiện trạng', '–'] + [so(100 * (A['Z'][2] - x['Z'][2]), 1) for x in (B, C)] + ['cm'],
            ['Đỉnh đập sau SC − MNLKT'] + [so(Z_DINH_SC - x['Z'][2]) for x in PHUONG_AN] + ['m'],
            ['Đỉnh đập sau SC − Zmax lớn nhất (kể cả sự cố 1 cửa)']
            + [so(Z_DINH_SC - max(z for z in x['Z'][2:] if z is not None)) for x in PHUONG_AN] + ['m']]
    bang(d, 'Bảng 3.1 Bảng so sánh các phương án',
         ['Chỉ tiêu', 'Hiện trạng 9 × 2,4 m', 'PA1 – xây thêm 1 khoang (Σb = 24 m)', 'PA2 – mở rộng về 25 m (3 khoang tự do)',
          'Đơn vị'], hang, cot=[3671, 1300, 1600, 1700, 800])
    doan(d, 'Ghi chú: các phương án dùng đường quá trình lũ QTVH 2026 (72 giờ), hồ ở MNDBT khi lũ đến, cửa van tràn chính và '
            'tràn sự cố (5 m, ngưỡng +71,00 m) kéo lên hoàn toàn khi đón lũ; hệ số co hẹp bên: mố bên tràn chính lượn tròn '
            'ξ_{mb} = 0,70, trụ pin đầu tròn ξ = 0,45, mố bên tràn sự cố vuông góc ξ_{mb} = 1,00, lấy H/b ≤ 1. PA2 không có '
            'cửa van nên không có trường hợp sự cố cửa.', co=11, nghieng=True, truoc=2)
    zt, zk, z_sc = A['Z'][0], A['Z'][2], A['Z'][3]
    p('Nhận xét:')
    p('(1) Với tràn hiện trạng, MNLTK +%s m, cao hơn MNLTK thiết kế ban đầu (+%s m) %s m. Chênh lệch không lớn, do lũ thiết kế '
      'cập nhật lớn hơn (+8,7%%) và tràn chính thực tế hẹp hơn thiết kế ban đầu (21,6 so với 25 m). Thiết kế ban đầu không có '
      'lũ kiểm tra nên chỉ so sánh được MNLTK; MNLKT của tràn hiện trạng là +%s m.'
      % (so(zt), so(MNLTK_1996), so(zt - MNLTK_1996), so(zk)))
    p('(2) MNLKT +%s m vẫn thấp hơn đỉnh đập thiết kế +%s m và đỉnh tường chắn sóng +%s m: khi các cửa van được kéo hết, lũ '
      'kiểm tra không tràn qua đỉnh đập. Tuy nhiên, khi 1 cửa tràn chính không mở được, mực nước lũ lớn nhất lên tới +%s m, chỉ '
      'thấp hơn đỉnh đập thiết kế %s m. Đồng thời độ vượt cao của đỉnh đập so với MNLKT chưa thỏa mãn, vì đỉnh đập phải ≥ '
      'MNLKT + 0,3 = +%s m.' % (so(zk), so(Z_DINH_TK), so(Z_TUONG), so(z_sc), so(Z_DINH_TK - z_sc), so(zk + 0.3)))
    p('(3) Mở rộng tràn chỉ hạ MNLTK %s ÷ %s cm và hạ MNLKT %s cm (thêm 1 khoang) đến %s cm (mở rộng về 25 m). Ngay cả khi mở '
      'rộng về 25 m, MNLKT + 0,3 m = +%s m vẫn cao hơn đỉnh đập +%s m, nghĩa là vẫn phải nâng đỉnh đập. Ngược lại, chỉ cần nâng '
      'đỉnh đập thêm %s m (lên +%s m, mục 3.2) đã thỏa điều kiện với dự trữ %s m mà không phải mở rộng tràn; kể cả khi 1 cửa '
      'tràn chính không mở, đỉnh đập vẫn cao hơn mực nước lũ lớn nhất %s m.'
      % (so(100 * (zt - B['Z'][0]), 1), so(100 * (zt - C['Z'][0]), 1), so(100 * (zk - B['Z'][2]), 1),
         so(100 * (zk - C['Z'][2]), 1), so(C['Z'][2] + 0.3), so(Z_DINH_TK), so(Z_DINH_SC - Z_DINH_TK), so(Z_DINH_SC),
         so(Z_DINH_SC - zk - 0.3), so(Z_DINH_SC - z_sc)))
    p('(4) Mở rộng tràn phải phá dỡ và làm lại mố, trụ pin, tường bên, cầu giao thông, cửa van và có thể cả dốc nước dài 60 m, '
      'mũi phun – các hạng mục đang còn tốt; thi công trên công trình đang khai thác, phải hạ mực nước hồ, kinh phí lớn hơn '
      'nhiều so với đắp nâng đỉnh đập. Vì vậy mở rộng tràn chính không hợp lý về kinh tế – kỹ thuật.')

    doan(d, 'b) Điều chỉnh dòng chảy để đỉnh lũ không vượt đỉnh đập', dam=True, truoc=6, sau=2, can=WD_ALIGN_PARAGRAPH.LEFT)
    p('Khả năng tháo nêu trên chỉ đạt được khi các cửa van được kéo lên hoàn toàn khi đón lũ (tràn làm việc như đập tràn chảy '
      'tự do). Vì vậy dòng chảy qua tràn phải được điều chỉnh bằng quy định vận hành cửa van để đỉnh lũ không vượt đỉnh đập. '
      'Mực nước lũ kiểm tra ứng với các cách vận hành cho trong Bảng 3.2.')
    hang = [[ten, '+' + so(z), '+' + so(z + 0.3), 'dư ' + so(Z_DINH_SC - z - 0.3) + ' m'] for ten, z in VAN_HANH]
    bang(d, 'Bảng 3.2 Mực nước lũ kiểm tra (P = 0,2%) theo cách vận hành cửa van',
         ['Cách vận hành khi đón lũ', 'MNLKT (m)', 'MNLKT + 0,3 m', 'So với đỉnh đập +78,85 m'], hang,
         cot=[4471, 1500, 1500, 1600])
    p('Mọi trường hợp đều thỏa điều kiện đỉnh đập sau sửa chữa +%s m ≥ MNLKT + 0,3 m, kể cả khi một cửa tràn chính không mở '
      'được. Với cách vận hành kéo hết cửa, tràn hiện trạng còn dự trữ cho lũ kiểm tra tăng khoảng %s%% trước khi vi phạm điều '
      'kiện MNLKT + 0,3 m và khoảng %s%% trước khi mực nước lũ chạm đỉnh đập +%s m. Đồ án đề xuất bổ sung vào quy trình vận '
      'hành: (1) trong mùa lũ, khi có lũ về phải kéo hết các cửa van tràn chính và tràn sự cố, không đóng cửa để tích nước khi '
      'đang có lũ; (2) trước mùa lũ kiểm tra, bảo dưỡng và vận hành thử toàn bộ cửa van, máy đóng mở, bố trí nguồn điện dự '
      'phòng và phương án đóng mở thủ công để không để xảy ra kẹt cửa; (3) khi có dự báo mưa lớn diện rộng trên lưu vực, xả '
      'trước qua tràn sự cố (ngưỡng +71,00 m, thấp hơn MNDBT 3,7 m) để tạo thêm dung tích đón lũ; (4) cập nhật tính toán '
      'thủy văn định kỳ – chỉ khi lũ kiểm tra cập nhật tăng vượt mức dự trữ nêu trên mới cần xét mở rộng tràn.'
      % (so(Z_DINH_SC), so(K_DU_TRU[0], 1), so(K_DU_TRU[1], 0), so(Z_DINH_SC)))

    doan(d, 'c) Kết luận', dam=True, truoc=6, sau=2, can=WD_ALIGN_PARAGRAPH.LEFT)
    p('Giữ nguyên quy mô tràn xả lũ chính (9 khoang × 2,4 m, ngưỡng +74,70 m), kết hợp nâng đỉnh đập lên +%s m và quy định '
      'vận hành cửa van nêu trên. Tràn chính chỉ sửa chữa các hư hỏng: nạo vét kênh dẫn thượng lưu bị bồi lấp; xử lý mài mòn '
      'bề mặt đáy dốc nước; vệ sinh, thay vật liệu chèn khe co giãn; thông các lỗ thoát nước dưới bản đáy; bảo dưỡng cửa van '
      'và máy đóng mở.' % so(Z_DINH_SC))

    # ------------------------------------------------------------------ cống lấy nước
    tieu_de(d, '3.4 Giải pháp cho cống lấy nước (cống ngầm)')
    h_ht, h_sc = Z_DINH_TK - Z_DINH_CONG, Z_DINH_SC - Z_DINH_CONG
    dq = GAMMA * (h_sc - h_ht)
    p('Cống lấy nước là cống hộp BTCT chảy không áp, kích thước 1,8 × 2,0 m, dài 67 m, ngưỡng +65,60 m, đặt trên nền đá lớp II '
      'và nằm sâu dưới thân đập: tại tim đập, đỉnh cống ở khoảng +67,10 m, chiều cao đất đắp phía trên khoảng %s m. Theo hồ sơ '
      'BCNCKT, kết cấu cống còn tốt, chưa có dấu hiệu thấm hay lún dọc tuyến cống; tháp van, cửa van phẳng, bể tiêu năng và '
      'cụm chia nước hạ lưu còn tốt; chỉ hai máy đóng mở đã hư hỏng, đơn vị quản lý phải sửa chữa thường xuyên để vận hành tạm.'
      % so(h_ht, 1))
    p('Đồ án chọn giữ nguyên cống vì các lý do sau: (1) cống còn làm việc tốt và đủ khả năng tháo – ở MNC +67,50 m cống tháo '
      'được 7,31 m^{3}/s, lớn hơn lưu lượng thiết kế 6,54 m^{3}/s (Chương 6); (2) cống chôn sâu dưới thân đập, muốn thay mới '
      'phải đào hố móng xuyên thân đập hoặc làm cống mới: khối lượng đào đắp lớn, phải hạ mực nước hồ và ngừng cấp nước tưới, '
      'đồng thời tạo mặt tiếp xúc mới giữa đất đắp và kết cấu – vị trí dễ phát sinh thấm tiếp xúc, xói ngầm, là nguy cơ lớn '
      'nhất đối với đập đất; (3) nâng đỉnh đập từ +%s lên +%s m chỉ làm chiều cao đất đắp trên cống tăng từ %s lên %s m, áp '
      'lực đất đứng lên nắp cống tăng khoảng %s kPa (≈ %s%%), không làm thay đổi đáng kể điều kiện chịu lực của thân cống '
      '(Chương 6).' % (so(Z_DINH_TK), so(Z_DINH_SC), so(h_ht), so(h_sc), so(dq, 1), so(100 * dq / (GAMMA * h_ht), 0)))
    p('Nội dung sửa chữa giới hạn ở phần cơ khí và công tác kiểm tra: thay hai máy đóng mở đã hư hỏng bằng máy vít VĐ20; '
      'thay gioăng kín nước và tấm trượt cửa van bằng vật liệu ma sát thấp; bổ sung giá đỡ trung gian cho trục vít; kiểm tra '
      'bên trong thân cống bằng camera (robot) và đo cao độ đáy cống tại cửa vào, cửa ra trong đợt ngừng cấp nước giữa hai vụ; '
      'bổ sung quan trắc thấm tại cửa ra cống. Kết luận: giữ nguyên kết cấu cống lấy nước, chỉ sửa chữa phần cơ khí và kiểm '
      'tra định kỳ.')
    d.save(RA)
    print('Đã ghi', os.path.abspath(RA))


if __name__ == '__main__':
    main()
