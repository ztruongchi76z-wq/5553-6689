# -*- coding: utf-8 -*-
"""
XUẤT WORD: BIỆN LUẬN CHƯƠNG 3 – GIỮ NGUYÊN HAY MỞ RỘNG TRÀN XẢ LŨ CHÍNH; GIỮ NGUYÊN CỐNG LẤY NƯỚC (cống ngầm)
Theo gợi ý GVHD (SV gửi 02/11/2026): so sánh với mực nước lũ của thiết kế ban đầu, điều chỉnh dòng chảy (vận hành cửa van)
để đỉnh lũ không vượt đỉnh đập; cống ngầm chôn sâu, còn tốt theo BCNCKT nên giữ nguyên.
Số liệu:
  - TK 1996: MNLTK +77,24 (Q1% = 847 m³/s, tràn 25 m) [KTTV Bảng 1-2]; BCNCKT 2022: +77,43 / +78,25 (tràn 25 m).
  - Đồ án: lũ QTVH 72 giờ, tràn chính 9 × 2,4 m + tràn sự cố 5 m, hồ ở MNDBT khi lũ đến (DieuTietLu_Potapov_f1f2 –
    VH_mo = 1: 77,430 / 78,389; 1 cửa không mở 77,477 / 78,458; VH_mo = 2: 78,433; VH_mo = 2 + 1 cửa: 78,504).
  - Phương án so sánh (cùng chương trình, mở hết cửa): 10 × 2,4 = 24 m: 77,384 / 78,327; 25 m 3 khoang: 77,339 / 78,251.
  - Dự trữ: lũ P = 0,2% tăng 4,5% thì MNLKT + 0,3 m = +78,85; tăng 12,9% thì MNLKT chạm +78,85.
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
BCNCKT = (77.43, 78.25)
HT = (77.430, 78.389)                       # tràn hiện trạng, mở hết cửa từ đầu trận lũ
PA = [('Đồ án – thêm 1 khoang', '10 × 2,4 = 24 m', (77.384, 78.327)),
      ('Đồ án – mở rộng về 25 m', '25 m, 3 khoang chảy tự do', (77.339, 78.251))]
VAN_HANH = [('Kéo hết cửa van từ đầu trận lũ', 78.389),
            ('Giữ MNDBT, khi lũ đến lớn hơn khả năng xả mới kéo hết cửa', 78.433),
            ('Kéo hết cửa, 1 cửa tràn chính không mở được', 78.458),
            ('Giữ MNDBT và 1 cửa tràn chính không mở được', 78.504)]
K_DU_TRU = (4.5, 12.9)                      # % tăng lũ P = 0,2%: MNLKT + 0,3 m = +78,85; MNLKT = +78,85
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
      'Tiêu chí đánh giá: mực nước lũ so với mực nước lũ của thiết kế ban đầu và hồ sơ BCNCKT; đỉnh lũ không vượt đỉnh đập, '
      'với điều kiện của TCVN 8216:2018 điều 10.1.5: đỉnh đập (không kể tường chắn sóng) ≥ MNLKT + 0,3 m, đỉnh tường chắn '
      'sóng ≥ cao trình đỉnh yêu cầu Z_{đ}.')
    doan(d, 'a) So sánh với mực nước lũ của thiết kế ban đầu và hồ sơ BCNCKT', dam=True, nghieng=False, truoc=6, sau=2,
         can=WD_ALIGN_PARAGRAPH.LEFT)
    p('Mực nước lũ của các nguồn tài liệu và của các phương án tràn (tính cùng chương trình điều tiết, cùng điều kiện vận '
      'hành) được tổng hợp trong Bảng 3.1.')
    hang = [['Thiết kế ban đầu (1996)', '847 / –', '25 m, 3 khoang chảy tự do', '+' + so(MNLTK_1996), '–'],
            ['Hồ sơ BCNCKT (2022)', '914 / 1.239', '25 m (theo thiết kế ban đầu)', '+' + so(BCNCKT[0]), '+' + so(BCNCKT[1])],
            ['Đồ án – tràn hiện trạng', '921 / 1.239', '9 × 2,4 = 21,6 m', '+' + so(HT[0]), '+' + so(HT[1])]]
    hang += [[ten, '921 / 1.239', b, '+' + so(z[0]), '+' + so(z[1])] for ten, b, z in PA]
    bang(d, 'Bảng 3.1 Mực nước lũ theo các nguồn tài liệu và các phương án tràn chính',
         ['Nguồn – trường hợp', 'Q_{max} lũ P = 1% / 0,2% (m^{3}/s)', 'Tràn chính', 'MNLTK (m)', 'MNLKT (m)'], hang,
         cot=[2671, 1600, 2400, 1200, 1200])
    doan(d, 'Ghi chú: các trường hợp của đồ án dùng đường quá trình lũ QTVH 2026 (72 giờ), hồ ở MNDBT khi lũ đến, cửa van tràn '
            'chính và tràn sự cố (5 m, ngưỡng +71,00 m) kéo lên hoàn toàn khi đón lũ.', co=11, nghieng=True, truoc=2)
    lech_kt = HT[1] - BCNCKT[1]
    p('Nhận xét: (1) với tràn hiện trạng, MNLTK +%s m bằng trị số của hồ sơ BCNCKT và cao hơn MNLTK thiết kế ban đầu %s m; '
      'MNLKT +%s m cao hơn BCNCKT %s m. Chênh lệch nhỏ, do lũ thiết kế cập nhật lớn hơn (+8,7%%) và tràn chính thực tế hẹp hơn '
      'thiết kế ban đầu (21,6 so với 25 m). (2) MNLKT +%s m vẫn thấp hơn đỉnh đập thiết kế +%s m và đỉnh tường chắn sóng +%s m: '
      'khi các cửa van được kéo hết, lũ kiểm tra không tràn qua đỉnh đập; điều chưa thỏa mãn là độ vượt cao của đỉnh đập so '
      'với MNLKT – đỉnh đập phải ≥ MNLKT + 0,3 = +%s m. (3) Mở rộng tràn chỉ hạ MNLKT %s m (thêm 1 khoang) đến %s m (mở rộng '
      'về 25 m); ngay cả khi mở rộng về 25 m, MNLKT + 0,3 m = +%s m vẫn cao hơn đỉnh đập +%s m, nghĩa là vẫn phải nâng đỉnh '
      'đập. Ngược lại, chỉ cần nâng đỉnh đập thêm %s m (lên +%s m, mục 3.2) đã thỏa điều kiện với dự trữ %s m mà không phải '
      'mở rộng tràn. (4) Mở rộng tràn phải phá dỡ và làm lại mố, trụ pin, tường bên, cầu giao thông, cửa van và có thể cả '
      'dốc nước dài 60 m, mũi phun – các hạng mục đang còn tốt; thi công trên công trình đang khai thác, phải hạ mực nước hồ, '
      'kinh phí lớn hơn nhiều so với đắp nâng đỉnh đập. Vì vậy mở rộng tràn chính không hợp lý về kinh tế – kỹ thuật.'
      % (so(HT[0]), so(HT[0] - MNLTK_1996), so(HT[1]), so(lech_kt), so(HT[1]), so(Z_DINH_TK), so(Z_TUONG),
         so(HT[1] + 0.3), so(HT[1] - PA[0][2][1]), so(HT[1] - PA[1][2][1]), so(PA[1][2][1] + 0.3), so(Z_DINH_TK),
         so(Z_DINH_SC - Z_DINH_TK), so(Z_DINH_SC), so(Z_DINH_SC - HT[1] - 0.3)))

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
