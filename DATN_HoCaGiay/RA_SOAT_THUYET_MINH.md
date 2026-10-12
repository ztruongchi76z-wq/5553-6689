# Rà soát thuyết minh ĐATN hồ Cà Giây – 10/10/2026 (cập nhật đợt 6: 12/10/2026)

Tệp được rà soát: `thuyet_minh/ban_09-10-2026/DATN_TranQuocThanh_HoCaGiay.docx` (bản 09/10, 88 trang).
Bản sau rà soát: `thuyet_minh/DATN_TranQuocThanh_HoCaGiay.docx` và `.pdf` (89 trang: bìa + tờ nhiệm vụ,
i–xii, 1–74).

Cách rà soát: đọc toàn văn (8 chương, kết luận, 37 tài liệu tham khảo, 5 phụ lục); tính tay lại các phép
tính chính; kiểm tra tự động đánh số hình/bảng/công thức, thứ tự trích dẫn, viết tắt, cấu trúc tệp Word;
xem từng trang PDF; đối chiếu hướng dẫn trình bày của Trường (`HuongDanTB_HPTN_(26-2-2020).docx`).

## 0a. Đợt 6 – 12/10/2026 (bản hiện hành; chi tiết: `thuyet_minh/nhat_ky_tiep_thu_gop_y_6.txt`)

Thuyết minh `thuyet_minh/DATN_TranQuocThanh_HoCaGiay.docx` / `.pdf` – **122 trang**, mục lục khớp số trang. Dựng bằng
`thuyet_minh/tiep_thu_gop_y_6.py` (đầu vào là bản trước đợt 6: `git show 0276cd6:DATN_HoCaGiay/thuyet_minh/DATN_TranQuocThanh_HoCaGiay.docx`)
rồi `chen_muc_luc.py`. Nguồn góp ý: (a) yêu cầu của cô Nguyễn Thị Phương Mai (Zalo); (b) báo cáo rà soát GPT 11/10
(`BaoCao_RaSoat_DATN_HoCaGiay_20261011.docx`); (c) 4 tiêu chuẩn SV tải lên Drive 04_TieuChuan và Hugo/TCVN (QCVN 04-05:2022,
TCVN 8421:2010, TCVN 9151:2012, TCVN 4116:2023); (d) tệp "Lời tư vấn" và 2 ảnh bảng phân cấp (TCXDVN 285:2002, QCVN 04-05:2012).

**Yêu cầu của GVHD → mục 4.6 (mới) và 5.6 (mới)**
- 4.6.1 Mưa lưu vực 1978–1999 so với 2000–2024: mưa năm −0,1 %, mùa mưa −2,1 %, tháng lớn nhất −12,3 %, X1max −4,1 %;
  Mann – Kendall, Sen, Welch không có ý nghĩa (p = 0,06 ÷ 0,93); Cv tăng (0,20 → 0,26) – mưa cực trị phân tán hơn
  (Bảng 4.7, Hình 4.2, Phụ lục 7 – Bảng PL12). Dẫn nghiên cứu BĐKH [31] Lê Thị Hải Yến và cs. (2017)
  (https://vjol.info.vn/DHTL/article/view/30969); TLTK đánh số lại [31] ÷ [41] → [32] ÷ [42].
- 4.6.2 MNLTK hiện nay +77,63 m so với TK 1996 +77,24 m: +0,20 m do lũ cập nhật (847 → 914 m³/s), +0,12 m do tràn thực tế
  21,6 m (thiết kế 25 m). Lũ tăng k = 10/20/30 % (Bảng 4.8, Hình 4.3): lũ kiểm tra chạm đỉnh +78,85 khi k ≈ 9 % (BCNCKT)
  ÷ 12 % (QTVH), chạm tường khi ≈ 25 %; sóng NE với lũ thiết kế chạm tường khi ≈ 12 %; điều kiện MNLKT + 0,3 m hết dự trữ
  khi k > 0,4 %. Hạ mực nước đón lũ cố định không khả thi; xả trước qua tràn sự cố về +73,84 m mất 27 ÷ 44 giờ.
- 4.6.3 Bảng 4.9 (công trình tháo lũ trước/sau sửa chữa); 5.6 Bảng 5.7 (đập: đỉnh, thấm, gradien, ổn định, lớp mặt, τ).

**Lỗi thực chất phát hiện khi đối chiếu tiêu chuẩn mới (đã sửa)**
- QCVN 04-05:2022 công thức (1), Bảng B.2: trượt bê tông trên đá có γc = 0,95 → [K] ngưỡng tràn 1,211 / 1,089 (bản trước
  1,15 / 1,035); c cần 7,6 / 10,4 kPa; lực neo cần 48 ÷ 65 kN/thanh (28 ÷ 38 % Ntt). Ký hiệu γn, γlc, γc thay Kn, nc.
  Thêm: phân loại kẹt van theo 2.6.3.10 (trả lời GPT B1), lũ vượt kiểm tra 0,1 % (2.6.3.5), bồi lắng 75 năm (Bảng 9),
  tiêu chí chiều cao đập (nền B → cấp II), lịch sử phân cấp.
- TCVN 8421:2010: gió phải quy đổi về mặt nước Vw = kl·Vl (A.3.3); bảng kα trong chương trình cũ sai (Bảng 9) →
  Z2 gió Đông Bắc +78,97 → **+79,07 m** (< tường +79,40, dự trữ 0,33 m); đỉnh đập +78,85 m không đổi.
- TCVN 9151:2012 điều 7.2.3: sơ đồ đập tràn khi H < 1,15h (bản trước ghi 1,2 ÷ 1,4h); H0 = 2,34 m = 1,17h → vùng chuyển
  tiếp, yêu cầu +67,94 ÷ +67,97 m; Fr cuối cống 4,9 ÷ 5,7 < 10 → độ đầy giới hạn 0,90, không bắt buộc cấp khí.
- TCVN 4116:2023: ξ = 0,11 ÷ 0,14 ≪ ξR; kiểm tra bề rộng khe nứt (9.2); cường độ thực tế (7.15); chống ăn mòn neo (7.4, 6.3.3).
- TCVN 8216:2018 điều 11.4.4 (+ QCVN 04-05:2022 điều 2.6.2.2.1a): dự phòng lún ≥ 0,5 %H → **0,13 m** (bản trước 2 cm):
  đất đắp thi công +78,73, nghiệm thu ≥ +78,98 (mép thượng lưu) / +78,89 (mép hạ lưu tấm); đắp bù **2.210 m³**;
  Bảng 7.1 dòng 9 ≈ 23.600 m³ (trừ phần nâng đỉnh của TKCS).

**Góp ý GPT 11/10**: tiếp thu A1, A2, A4 (làm rõ εm = ε·m; độ nhạy: đỉnh +78,85 đạt khi m ≥ 0,441), A5 ÷ A11, B1 ÷ B10
(bất đẳng hướng kh/kv = 5 – Bảng PL10; tìm lại cung trượt bằng miền mở rộng, lưới lồng, Nelder – Mead, Spencer – Bảng PL11;
ô rãnh theo đường chéo 5,66 m: τ = 2,23 Pa; trục vít μ = 1,0 → **02 giá đỡ dẫn hướng**, L ≤ 2 m, Kođ ≥ 2,97).
Không tiếp thu: **A3** – giả thiết "9 cửa van tràn chính" không đúng hồ sơ (tràn chính chảy tự do, không cửa van – QTVH,
BCNCKT, KTTV); **A12** – trang nhiệm vụ có chữ ký là việc của SV.

**Tệp "Lời tư vấn" (12/10)**: ý "tư vấn phân cấp sai, dùng lũ 1,5 % / 0,5 %" **không đúng với hồ sơ** – BCNCKT (mục
4.1.1), TKCS (1.2.3), báo cáo KTTV (1.1.9) và QTVH 2026 đều xác định cấp II, lũ 1 % / 0,2 % (QTVH thêm 0,1 %). Cấp III / IV
chỉ là cấp cũ theo TCVN 5060-90 (QTVH ghi III, BCNCKT ghi IV – không thống nhất); 1,5 % / 0,5 % là tần suất của cấp III theo
TCXDVN 285:2002 / QCVN 04-05:2012 – không hồ sơ nào dùng. Sai sót thật ở bước phân cấp của BCNCKT: xét H = 25,4 m với nền
loại A (thực tế 24,75 m, nền B – kết luận cấp II không đổi) và dùng hệ số an toàn ổn định mái của cấp III (1,25 / 1,13).
Hai ảnh bảng phân cấp được dùng để viết đoạn "lịch sử phân cấp" ở mục 1.1.3 và câu hỏi dự phòng trên slide. Các ý 1 ÷ 5 và
lỗi nhỏ còn lại trong tệp đã có trong thuyết minh (mục 2.1 ÷ 2.5, Kết luận – kiến nghị).

**Bản vẽ, slide**: BV-02 (Bảng 2.4, 2.5, ghi chú 3), BV-04 (khối lượng, ghi chú 3, 6, 8), BV-06 ([K], neo, chống ăn mòn),
BV-07 (02 giá đỡ, Bảng 6.3, biểu đồ Kođ, TCVN 9151) xuất lại; gói `TAI_VE/BanVe_A0_DATN_HoCaGiay.zip` dựng lại. Slide: thêm
slide 8 (mưa, lũ gia tăng) và 15 (hiệu quả trước/sau); sửa 5, 11, 12, 13, 14, 20, 21 và slide dự phòng (22 + 1 slide);
gói `TAI_VE/Slide_BaoVe_DATN_HoCaGiay.zip` dựng lại (link slide trên Claude không còn khớp).

**Cần hỏi GVHD**: (1) dự phòng lún 0,13 m theo 11.4.4 cho đập cũ đã cố kết (tính lún chỉ 0,5 ÷ 1 cm) – giữ hay giảm;
(2) vị trí khống chế cao trình đỉnh (mép thượng lưu, chân tường); (3) mức độ Chương 6 (cống) – giữ ở mức kiểm tra;
(4) lời cam đoan và khai báo dùng AI; (5) hệ số lưu lượng tràn chính m = 0,45 (nếu 0,40: MNLKT +78,62); (6) màng 2 hàng
chỉ áp dụng sau phụt thí nghiệm; (7) chọn địa hình A (kl = 1,10) khi quy đổi gió.

## 0. Bản 11/10/2026 – đã tiếp thu vào tệp Word (chi tiết: `thuyet_minh/nhat_ky_tiep_thu_gop_y.txt`)

Đã xử lý toàn bộ mục 3 (trừ ý 6, 7 – việc của SV) và mục 6 ý 1 ÷ 13, cùng góp ý astra/gem/"Góp ý bổ sung":
- Đỉnh đập theo TCVN 8216:2018 mục 10.1.5 (gió NE: Z₂ = +78,97 < tường +79,40; đỉnh ≥ MNLKT + 0,3 = +78,84);
  +78,85 là cao trình tối thiểu sau lún, dự phòng lún 2 cm; cấu tạo mặt đập khép kín (đất +78,60 + đá dăm 5 cm + BT 20 cm).
- Mô hình thấm, ổn định sau sửa chữa chạy lại với +78,85 (Bảng 5.1, 5.2, Hình 5.1 ÷ 5.4, PL1 ÷ PL10); thêm TH11, TH12
  (thi công, [K] = 1,24); kiểm tra chéo Spencer, lưới tâm mịn, hội tụ lưới PTHH (Phụ lục 6).
- k_h = 0,05 có căn cứ: a_gR = 0,0222g (TCVN 9386:2012 Phụ lục H, Chợ Lầu) → thiên an toàn; độ nhạy: đạt khi k_h ≤ 0,09.
- Bảng 4.2 phân loại tình huống kẹt van; neo ngưỡng D25 thành hạng mục thiết kế (mục 4.5.3, Bảng 4.6, Bảng 7.1);
  khung kín cống hộp tham khảo (Bảng 6.4); kịch bản độ dốc cống 3‰; 02 máy đóng mở.
- Chương 8: 2 hàng kề nhau A, B (không phải A, C), đáy màng theo mục 9.2.2.9 (mặt lớp II − 3 m), Bảng 8.7, Hình 8.6.
- P = m/n so sánh; sửa câu dung tích hữu ích; ngưỡng cảnh báo quan trắc có căn cứ; trình tự thi công mùa mưa;
  TLTK: TCVN 9386-1/-5:2025, TCVN 4116:2023, QCVN 02:2022/BXD, Spencer (1967); "LỜI CẢM ƠN"; khai báo dùng AI (mục 1.7).
- Không làm: khái toán chi tiết (thiếu đơn giá 2026 – chỉ dẫn vốn 48,69 tỷ theo BCNCKT), đưa link GitHub vào thuyết minh
  (kho đang công khai – nên chuyển Private).
- Bản vẽ BV-01 ÷ BV-07 đã cập nhật đồng bộ (rà lại chiều 11/10: BV-02 thêm gió Đông Bắc, kết luận đỉnh đập theo
  điều 10.1.5; BV-06 số Bảng 4.3, 4.5 + dòng có neo; BV-07 đường bão hòa +65,6 ÷ +73,0; BV-05 mục 8.6, ≈ 0,8 Lu);
  gói `TAI_VE/BanVe_A0_DATN_HoCaGiay.zip` dựng lại. Slide bảo vệ đã cập nhật
  (slide 8, 9, 10, 12, 13, 16, 17, 18, 19, 21: K = 1,349 / 1,496; neo ngưỡng; màng 2 hàng A, B, đáy theo 9.2.2.9;
  hình thấm sau sửa chữa q = 8,05·10⁻⁶); gói `TAI_VE/Slide_BaoVe_DATN_HoCaGiay.zip` dựng lại.

### Đợt 2 (tối 11/10/2026) – góp ý `Gpt_gop_y.txt` (chi tiết: `thuyet_minh/nhat_ky_tiep_thu_gop_y_2.txt`)

Cả 6 điểm kỹ thuật đều có cơ sở, đã sửa (script `thuyet_minh/tiep_thu_gop_y_2.py`, thuyết minh 102 trang):
1. Dự phòng lún: tách cao trình thiết kế sau lún (+78,85 / +78,76) và cao độ nghiệm thu trước lún (≥ +78,87 / ≥ +78,78;
   đất đắp đến +78,62); đắp bù 1.610 → 1.700 m³ (Bảng 7.1, BV-04, BV-02).
2. Neo D25: nêu rõ neo thụ động, cơ chế ma sát – cắt (EN 1992-1-1 mục 6.2.5 [40]); N_tt = min(thép 171,8; bám dính
   vữa – đá 191 theo FHWA [41]; thép – vữa 314) kN; lực cần 40 ÷ 57 kN = 23 ÷ 33% N_tt, bám dính cần 0,08 ÷ 0,12 MPa,
   độ mở 0,1 ÷ 0,2 mm; K = 2,23 / 1,81 là trạng thái giới hạn; thí nghiệm kéo nhổ trước thi công đại trà.
3. Tràn sự cố: (4-1) vô nghiệm với P = 1% (Q_max = 138,25 < 138,6 m³/s) → đầu dốc lấy h_k = 4,28 / 4,76 m
   (`tinh_toan/kt_tran_su_co.py`); cuối dốc h = 2,56 ÷ 2,91 m; tường yêu cầu 5,5 → 3,7 m – cần đo tường thực tế.
4. Cống: điều kiện cửa vào chưa ngập H ≤ (1,2 ÷ 1,4)h_c = 2,4 ÷ 2,8 m; H0 = 2,34 m vẫn thỏa; sơ đồ chảy qua lỗ cho +67,97 m.
5. Chống xói: kết luận có điều kiện (τ = 1,4 ÷ 2,2 Pa so với [τ] đất chưa phủ ≈ 1,8 ÷ 3,6 Pa) – trồng cỏ vầng/thảm xơ
   dừa đến khi cỏ phủ kín.
6. Màng: thêm Bảng 8.8 η theo K lớp II (`tinh_toan/kt_lop_II_lu.py`) – với 5 Lu η chỉ còn 26 ÷ 32%; giữ kết luận hiệu
   quả khi q ≤ 1 Lu.

Sửa nhỏ: TH2 "chưa đạt"; mục 2.3 dẫn Bảng 4.2; "6 điểm phần trăm"; [14] → TCVN 10778:2024 (chưa đọc được toàn văn bản
2024 – số liệu tổn thất thấm (1 ÷ 3)%/tháng nay dẫn [13]); "chưa tìm thấy cung bất lợi hơn trong phạm vi khảo sát";
ghi chú slide 6 (dung tích hữu ích 29,60 > 28,83 triệu m³). Bản vẽ BV-02, 04, 05, 06, 07 và slide (bản 16) đồng bộ.
Không làm được: trang "Nhiệm vụ đồ án" là trang để đóng bản có chữ ký của Bộ môn, GVHD – SV tự chèn.

### Đợt 3 (đêm 11/10/2026) – rà soát độc lập sau đợt 2 (chi tiết: `thuyet_minh/nhat_ky_tiep_thu_gop_y_3.txt`)

Kiểm chứng lại 6 điểm GPT và các sửa nhỏ của đợt 2: **đều đã sửa đúng**. Đã chạy lại `dieu_tiet_lu.py`,
`chay_tham_on_dinh.py`, `kt_tran_su_co.py`, `kt_lop_II_lu.py` – các số trong Bảng 2.3, 2.6, 2.7, 5.1, 5.2, 8.8 và mục 4.3
khớp chương trình; tính tay lại Bảng 4.5, 4.6, 5.4, 5.5, 5.6, 6.2, 6.3, 6.4, 7.1 – khớp. Các chỗ còn lệch đã sửa
(script `thuyet_minh/tiep_thu_gop_y_3.py`, thuyết minh 103 trang, mục lục 219 dòng, 0 lỗi số trang):

| # | Vị trí | Vấn đề | Đã sửa |
|---|---|---|---|
| 1 | Mục 6.5 | Đường bão hòa tại tim đập ghi +65,7 ÷ +73,4 m (số của mô hình cũ) | +65,6 ÷ +73,0 m theo Bảng 5.1 |
| 2 | Mục 6.4 | "chia hai nhịp L ≤ 3 m (Kođ tăng lên 3,4 ÷ 6,5)" – Bảng 6.3 có trường hợp 2,69 | 3,4 ÷ 6,5 khi f ≤ 0,4; f = 0,5 ở MNLTK vẫn 2,7 > 2 |
| 3 | Bảng 1.1 | Chiều cao đập lớn nhất 24,75 m (trị số TKCS, đỉnh +78,75) | 24,35 / 24,85 m (hiện trạng / sau sửa chữa, đáy +54,0) |
| 4 | Mục 4.5.1 | Áp lực đẩy ngược ghi theo "h_c" nhưng chương trình dùng 0 / 1,0 / 1,2 m | Ghi rõ chiều sâu nước tại chân ngưỡng như chương trình (thiên an toàn) |
| 5 | Mục 5.1.1 | Độ lún bổ sung 0,5 ÷ 1 cm chưa có căn cứ | Thêm ước tính: Δσ ≈ 11 kPa, ∫Δσ·dz ≈ 100 kPa·m, E_oed ≈ 10 ÷ 20 MPa |
| 6 | Mục 8.5.4 | "tăng η thêm 3 ÷ 5%" (đợt 2 mới sửa một chỗ "6%") | "3 ÷ 5 điểm phần trăm" (cả slide 16) |
| 7 | Mục 8.6, Bảng 8.8 | "1 Lu < q ≤ 5 Lu thì η chỉ còn 26 ÷ 42%" – sai ở đầu khoảng (gần 1 Lu η ≈ 50%) | Thêm dòng K = 1,3·10⁻⁵ cm/s (1 Lu) vào Bảng 8.8 (`kt_lop_II_lu.py` chạy lại): η = 50,2 ÷ 51,5% (3 hàng), 43,8 ÷ 45,4% (2 hàng); viết lại câu |
| 8 | Mục 8.6, 8.7 | Kết quả mô hình chưa nêu: với lớp II 5 Lu, ở MNLKT dòng thấm **vẫn ra trên mái hạ lưu** (+74,95 m) dù có màng | Bổ sung vào mục 8.6, kết luận chuyên đề, ghi chú slide 17 |
| 9 | Bảng 8.7 | Ghi chú "không kể hố phụt thí nghiệm" nhưng 51 hố của TKCS gồm cả 10 hố thí nghiệm; tổng khoan 506 m ≠ Bảng 7.1 (480 m) | Ghi rõ cách tính, chênh 5 ÷ 7% so với Bảng 2-2 TKCS, chỉ dùng so sánh tương đối |
| 10 | TLTK | Thứ tự trích dẫn lần đầu không đúng IEEE sau các lần thêm tài liệu: [38] QCVN 02 nêu ở mục 2.1 không ghi số; [39] Spencer, [40], [41], [35] xuất hiện trước số nhỏ hơn | Ghi số trích dẫn ở lần nêu đầu, **đánh số lại 41 tài liệu theo thứ tự xuất hiện** (hướng dẫn trình bày của Trường: IEEE) – bảng đổi số ở dưới |
| 11 | TLTK [12], Spencer | Tên TCVN 9386-1, -5:2025 ghi "Thiết kế công trình chịu động đất…"; tác giả Spencer ghi "S." | "Thiết kế kết cấu chịu động đất – Phần 1: Quy định chung, tác động động đất và quy định cho nhà" và "– Phần 5: Nền móng, kết cấu chắn và các vấn đề địa kỹ thuật"; "E. Spencer" |
| 12 | Nhỏ | Mục 2.2.1 "0,45 m" (đúng 0,44); 7.3 tăng 490 m³ chưa kể dự phòng lún; "-0,00" ở Bảng 1.2; chữ hoa sau dấu ";" (7.6); "(4) Chất lượng màng" dính vào đoạn (3); thụt lề, cỡ chữ lẫn ở 2 đoạn; "≈ 1 Lu" ở kết luận (đúng ≈ 0,8 Lu); Bảng 4.5, 6.3 số bị xuống dòng | Đã sửa hết |
| 13 | PDF | LibreOffice 24.2 làm mất một số chỉ số trên ở cuối dòng (vd. "kG/cm²" thành "kG/cm", "1·10⁻⁵" tách dòng) | Xuất PDF bằng LibreOffice 26.2 – đã đối chiếu toàn văn Word ↔ PDF, không còn mất ký tự |

**Bảng đổi số tài liệu tham khảo (số cũ → số mới):** [1]…[12] giữ nguyên; [38]→[13]; [13]…[21]→[14]…[22]; [39]→[23];
[22]…[26]→[24]…[28]; [40]→[29]; [41]→[30]; [27]…[32]→[31]…[36]; [35]→[37]; [33]→[38]; [34]→[39]; [36]→[40]; [37]→[41].
Các số tài liệu nêu trong mục 0 – 6 của tệp này là **số cũ**.

Không sửa (để SV, GVHD quyết): trang danh mục bảng biểu còn 1 dòng (Bảng PL9) sang trang xii khi xuất bằng LibreOffice –
Word dàn trang lại khi cập nhật trường (F9), xem lại khi in. Bộ bản vẽ A0 không có số liệu nào bị ảnh hưởng (đã đối chiếu
BV-01 ÷ BV-07), không dựng lại. **Đính chính:** đối chiếu kỹ lại sau đó (toàn bộ chữ trích từ DXF) thấy BV-04, BV-05,
BV-06 còn lệch – đã sửa, xem mục ngay dưới.

### Đối chiếu bản vẽ, slide với thuyết minh bản cuối (khuya 11/10/2026)

Đọc toàn bộ chữ trên 7 tờ A0 (trích từ DXF) và 21 slide, so với thuyết minh 103 trang. BV-01, BV-02, BV-03, BV-07 và
các slide còn lại khớp. Đã sửa:

| # | Vị trí | Lệch | Đã sửa |
|---|---|---|---|
| 1 | BV-04, bảng khối lượng | Ghi "(Bảng 7.1)" nhưng mục 3, 4 đảo thứ tự (nilon / đá dăm) và thiếu mục 9 (đào, đắp xử lý hố xói ≈ 25.000 m³), nên số thứ tự 9 ÷ 14 lệch Bảng 7.1 | Đúng thứ tự Bảng 7.1, đủ 15 mục; ghi chú 8: mục 9 lấy theo TKCS, chuẩn xác bằng đo đạc hố xói |
| 2 | BV-04, ghi chú 6 | "1,4 ÷ 2,2 Pa, Bảng 5.5" – Bảng 5.5 chỉ có i = 115 mm/h (1,82 Pa), khoảng 1,4 ÷ 2,2 Pa nằm trong lời văn | Dẫn "mục 5.5.1" |
| 3 | BV-05, ghi chú 3 và dòng "Điều kiện áp dụng" | Phương án hợp lý hóa ghi điều kiện "ép nước lớp II q ≤ 5 Lu"; thuyết minh (mục 8.6, 8.7 sau đợt 3) và slide 17 yêu cầu q ≤ 1 Lu – 5 Lu chỉ là tiêu chí cao trình đáy màng | q ≤ 1 Lu; 1 < q ≤ 5 Lu thì tính lại η theo K đo, cân nhắc giữ 3 hàng; chú thích dưới bảng thêm: với 5 Lu, ở MNLKT dòng thấm vẫn ra mái hạ lưu (+74,95 m) |
| 4 | BV-06, mặt cắt 1-1 | hhk = 1,31 / 1,78 m | 1,30 / 1,77 m như mục 4.2 (Bảng 4.4); độ vượt cao 0,73 m giữ nguyên |
| 5 | Slide 10 | Sơ đồ ghi "Đắp bù + BT M250 dày 20 cm", thiếu lớp đá dăm đệm 5 cm | Thêm dòng "+ đá dăm đệm 5 cm" |
| 6 | Slide 12 | "(≈ 36 thanh, Bảng 4.6)" – Bảng 4.6 không có số thanh | "(36 thanh, mục 4.5.3, Bảng 7.1)" |
| 7 | Slide 18 | Dòng tiến độ "Sửa đáy dốc tràn, nạo vét kênh dẫn" thiếu khoan neo ngưỡng (Bảng 7.2 có) | "Sửa đáy dốc, neo ngưỡng tràn, kênh dẫn" |
| 8 | Slide 21, câu 2 | Dẫn "Bảng 2.2" nhưng các số theo m/n (40,9 / 88,6 / 84,1%) nằm trong lời văn | "Mục 2.2.3" |
| 9 | Thuyết minh mục 5.5.3 | "cỏ Vetiver hoặc cỏ lá tre" – lệch mục 5.5.1, 7.4, Kết luận và BV-04 (trồng cỏ bằng vầng hoặc phủ thảm xơ dừa) | "cỏ bản địa (cỏ gà, cỏ lá tre) bằng vầng hoặc gieo hạt có phủ thảm xơ dừa cố định bằng ghim (mục 5.5.1)" – script `thuyet_minh/tiep_thu_gop_y_4.py`; mục lục, số trang không đổi (103 trang, 0 lỗi) |

Bản vẽ dựng lại bằng `ban_ve/to04…`, `to05…`, `to06…`; bộ PDF gộp và hai gói zip trong `TAI_VE/` đã thay. Slide trực
tuyến trên Claude (link trong HANDOFF mục 2) không sửa được từ phiên này – danh sách chỗ cần sửa tay khi xuất PowerPoint
từ link ghi trong `DOC_TRUOC.txt` của gói slide.

### Đối chiếu TCVN 10778:2024 và TCVN 4116:2023 (khuya 11/10/2026 – SV tải bản đầy đủ hai tiêu chuẩn lên Drive)

Hai tiêu chuẩn đã có trong danh mục tài liệu ([15], [34]) nhưng trước đây chưa có văn bản để đối chiếu nội dung. Đã đọc
phần quy định (TCVN 10778:2024 mục 1 ÷ 11, Phụ lục B.3; TCVN 4116:2023 mục 4 ÷ 7) và sửa (script
`thuyet_minh/tiep_thu_gop_y_5.py`, nhật ký `nhat_ky_tiep_thu_gop_y_5.txt`; thuyết minh 104 trang, mục lục 0 lỗi; số thứ tự
tài liệu tham khảo không đổi):

| # | Tiêu chuẩn, điều | Thuyết minh trước đây | Đã sửa |
|---|---|---|---|
| 1 | TCVN 10778:2024 công thức (6), mục 6.3 – dòng chảy tối thiểu Q_tt trong cân bằng nước | Nhu cầu nước lấy theo tổng của QTVH (tưới + sinh hoạt), chưa có Q_tt (QTVH mục 3.4.2: Q_tt = 0,027 m³/s, 0,85 triệu m³/năm) | Tính lại có Q_tt (`tinh_toan/kt_dong_chay_toi_thieu.py`): mức bảo đảm không đổi (40,0 / 86,7 / 82,2 %), ΣW thiếu 159,7 / 8,9 / 14,6 triệu m³, MNDBT cần cho TH1A +74,62 → **+74,68 m** (< +74,70 m) – thêm vào mục 2.2.3, slide 6 |
| 2 | Công thức (14), (9) – P = m/(n + 1), số năm đủ nước m ≥ (n + 1)·P/100 | P = m/(n + 1) dẫn [14] | Ghi rõ là công thức (14) của TCVN 10778:2024; điều kiện (9): cần ≥ 39/44 năm – TH1A vừa đạt 39 năm; QTVH dùng m/n là khác tiêu chuẩn (mục 2.2.2, slide 21) |
| 3 | Mục 7.1 ÷ 7.3 – mực nước chết thỏa đồng thời điều kiện bồi lắng và lấy nước | Chỉ kiểm tra khả năng lấy nước | Thêm điều kiện bồi lắng: bùn cát 70 năm ≈ 0,48 triệu m³ (QTVH mục 3.5.4) ≪ dung tích dưới ngưỡng cống 3,52 triệu m³ (mục 2.2.1) |
| 4 | Bảng 1 – tổn thất thấm; mục 8.2.1.2 – điều tiết toàn chuỗi khi chuỗi ≥ 15 năm; mục 10 – điều tiết lũ, sai số lặp ≤ 5 cm | Không dẫn | Dẫn chiếu (k_t = 1 %/tháng nằm trong 0,5 ÷ 1,5 %/tháng; chuỗi 44 năm; sai số lặp ≪ 5 cm) – mục 2.2.2, 2.3.1; Bảng 2.1 thêm dòng 12; danh mục viết tắt MNLTK, MNLKT ghi tên theo tiêu chuẩn |
| 5 | TCVN 4116:2023 mục 6.3.3 – lớp bê tông bảo vệ ≥ 60 mm và ≥ 2d ở mặt tiếp xúc trực tiếp với nước | Khung cống tính với a = 5 cm | a = 7 cm (60 mm + Ø20/2); tính lại `bo_sung_giai_tich.py` – Bảng 6.4: A_s sau sửa chữa 8,4 ÷ 16,7 cm²/m (trước 7,9 ÷ 15,7), văn bản "8 ÷ 17 cm²/m (5Ø16 ÷ 6Ø20)", BV-07 ghi chú 7; mức tăng do nâng đỉnh vẫn ≈ 4 % |
| 6 | Mục 5.1.2 – chỉ tiêu bê tông là cấp độ bền chịu nén B (và mác chống thấm W) | Ghi mác M250, M200 | Ghi B20 (M250), B15 (M200) trong lời văn, Bảng 3.1, 3.2, 7.1, BV-01 ÷ BV-04, slide 9, 10, 18; Bảng 2.1 thêm dòng 11 |
| 7 | Mục 5.1.9 – vùng có yêu cầu chống mài mòn, xâm thực: tối thiểu B25, W8 | Vữa xi măng – polime ≥ 40 MPa (đạt cường độ) | Thêm yêu cầu mác chống thấm ≥ W8 (mục 4.4, BV-06) |

Đã khớp, không sửa: thép CB300-V R_s = 260 MPa, CB400-V R_s = 350 MPa (TCVN 4116:2023 Bảng 11); K_n, n_c theo QCVN 04-05
(mục 7.1 dẫn chiếu); phương trình cân bằng nước và cách tính bốc hơi phụ thêm; điều tiết lũ bắt đầu từ MNDBT.
Lưu ý cho SV: công thức tính A_s là ước tính tham khảo (bỏ qua hệ số điều kiện làm việc γ_s2 = 1,1 – thiên an toàn); nếu
GVHD yêu cầu tính khung cống đầy đủ theo TCVN 4116:2023 mục 8.2 thì cần hồ sơ hoàn công kết cấu cống.

## 1. Đã sửa trong tệp Word

| # | Vị trí | Vấn đề | Đã sửa |
|---|---|---|---|
| 1 | Mục lục, Danh mục hình, Danh mục bảng | Trong tệp Word chỉ có dòng "Nhấn chuột phải… Update Field / F9", không có mục nào, không có số trang (chỉ bản PDF mới có) | Điền sẵn 205 dòng (132 đề mục + 32 hình + 41 bảng) có số trang; mỗi dòng là liên kết tới đề mục và vẫn cập nhật được bằng F9 trong Word. Kiểm tra lại: 0 dòng sai số trang |
| 2 | Công thức (8-1) | Có ký tự lỗi `\{}H\`, `\{}Q\`; dấu `]` đứng riêng làm chỉ số trên – LibreOffice cắt mất nửa công thức, Word hiện sai | Dựng lại: [K_e] = k_e·A_e·[B]ᵀ[B] ; [K]·{H} = {Q} |
| 3 | Công thức (2-1) | Dấu gạch trên của Z̄, V̄ nằm tách riêng → hiện ô vuông | Ghép dấu gạch trên vào Z, V |
| 4 | Mục 5.1.2 | "Kiểm tra thấm và ổn định … ở mục 5.3 và 5.4" (5.4 là trượt nông) | "… ở mục 5.2 và 5.3" |
| 5 | Mục 4.3 (tràn sự cố) | "h ≈ h_k = 4,24 m (P = 1%) và 4,29 m (P = 0,2%)": với Q = 162,7 m³/s, b = 5 m thì h_k = 4,76 m | Ghi đúng: độ sâu đầu dốc là nghiệm dòng xiết của (4-1): 4,24 m (h_k = 4,28 m) và 4,29 m (h_k = 4,76 m) |
| 6 | Toàn bài | 59/73 hình, bảng không được nhắc số trong lời văn (mẫu Trường: "Khi đề cập đến các hình và bảng biểu, phải nêu rõ số") | Bổ sung dẫn số ở 28 đoạn và 3 câu dẫn mới; nay 73/73 hình, bảng đều được dẫn |
| 7 | Bảng PL2 | Cột V lệch Bảng PL1 tới 0,07·10⁶ m³ (vd. +75,0 m: 38,32 so với 38,25) | Thêm ghi chú: chương trình tích phân trên lưới 0,01 m, Bảng PL1 dùng (1-1) bước 1 m; sai khác ≤ 0,2 %, không ảnh hưởng điều tiết lũ |
| 8 | Danh mục từ viết tắt | Thiếu BT (dùng 13 lần), TB (11), CB/ĐB, SC, PA, MN, ng.đ; "TH" sai thứ tự A, B, C | 19 → 26 mục, xếp lại theo A, B, C |
| 9 | Bìa chính, bìa phụ | "BỘ NÔNG NGHIỆP VÀ MÔI TRƯỜNG" bị xuống dòng; tên đề tài để chữ "ĐỒNG" một mình một dòng | Đổi độ rộng ô (7 + 9 cm); ngắt tên đề tài 3 dòng cân đối |
| 10 | Danh mục hình, bảng | Các dòng sát nhau | Cách đoạn 6 pt (mục "Các danh mục" của hướng dẫn trình bày) |

Nhật ký chi tiết: `thuyet_minh/nhat_ky_hieu_chinh.txt`; danh sách mục lục kèm số trang: `thuyet_minh/muc_luc_so_trang.txt`.

## 2. Đã kiểm tra – đạt

- **Đánh số trang đúng mẫu Trường:** không đánh số bìa chính, bìa phụ, tờ nhiệm vụ; i, ii, … từ LỜI CAM ĐOAN;
  1, 2, … từ Chương 1 liên tục đến hết phụ lục; số trang ở chân trang, giữa, cách mép dưới 0,5 cm; không dùng
  đầu trang. Lề 2,5 / 2,5 / 3,0 / 2,0 cm; khổ A4.
- Đánh số chương, mục, hình, bảng, công thức liên tục, không thiếu số; chú thích bảng ở trên, hình ở dưới.
- 37 tài liệu tham khảo đều được trích, số thứ tự đúng thứ tự xuất hiện lần đầu (IEEE).
- Phụ lục 8 trang (giới hạn 30 trang).
- Tính tay lại, khớp với thuyết minh: Z1, Z2, Z3 và Δh, gD/V² (Bảng 2.4); ε, Q tràn chính (Bảng 4.2);
  ổn định ngưỡng tràn và lực dính cần thiết 6,4 / 9,1 kPa (Bảng 4.4); lực đóng mở, hệ số ổn định trục vít
  (Bảng 6.3); Fr, độ sâu liên hiệp sau cống; áp lực đất lên cống; tổng lượng thấm toàn tuyến (15.894 m³/tháng);
  J_th lớp 5; T = B_b + 0,6A; tỷ số hiệu quả màng (Bảng 8.2–8.6); khối lượng đất đắp thêm của phương án Đ3.
- **Đối chiếu TCVN 8216:2018 (bản gốc gửi ngày 10/10):** các trị số đồ án dùng đều đúng – [K] đập cấp II = 1,30 (cơ bản) /
  1,17 (đặc biệt) (điều 5.2.2, Bảng 1; cấp III mới là 1,25 / 1,13, nên nhận xét "BCNCKT dùng 1,13" là đúng);
  a = 0,7 / 0,5 / 0,2 m (điều 10.1.4, Bảng 3); gió P = 4 % ở MNDBT, 50 % ở MNLTK, ở MNLKT không xét sóng leo
  (Bảng 4); có tường chắn sóng thì đỉnh đập ≥ MNLKT + 0,3 m và tường cao hơn đỉnh đập không quá 1,2 m (điều 10.1.5:
  +78,85 − đỉnh tường +79,40 → 0,55 m, đạt); K bộ phận chống thấm ≤ 1·10⁻⁵ cm/s (điều 7.2.2); [J] cửa ra 0,85 cho đất
  á sét, cấp II (điều 11.2.3, Bảng 6); màng phụt trong thân đập dày ≥ 1/10 cột nước, chỉ dùng cho đập cấp II trở
  xuống (điều 10.4.2, mục 5); màng phụt vữa xi măng – sét trong nền đất T = H/J với J = 3 ÷ 4 (điều 9.2.3.2, công
  thức 4 – khớp với [Jm] ≈ 3,6 mà TKCS dùng ở mục 8.2.2).
- Cấu trúc tệp Word: trường, bookmark, liên kết cân đối; thứ tự thẻ trong rPr/pPr đúng lược đồ (tránh lỗi
  "unreadable content" khi mở bằng Word).

## 3. Cần sinh viên kiểm tra hoặc hỏi GVHD trước khi nộp (chưa sửa)

1. **Ghi số điều TCVN 8216:2018 vào cột "Căn cứ" của Bảng 2.1** (các trị số đã đúng, xem mục 2): dòng 3 – Bảng 4;
   dòng 4 – Bảng 3; dòng 5 – Bảng 1 (điều 5.2.2); dòng 8 – điều 7.2.2; dòng 9 – Bảng 6 (0,85). Trị số 2,0 (tiếp xúc tầng
   lọc) không có trong TCVN 8216 (Bảng 7 chỉ cho gradien tới hạn J_th, [J] = J_th/[K]) – cần ghi rõ nguồn (TCVN 8422
   hoặc hồ sơ [3]). Bảng 1 còn có điều kiện **"Thi công và sửa chữa": [K] = 1,24 (cấp II)** – nên có một câu ở mục
   3.2 hoặc 7.1 nói rõ giai đoạn thi công sửa chữa (đào xử lý hố xói mái hạ lưu, tập kết vật liệu trên đỉnh) được
   kiểm tra theo trị số này hoặc vì sao không chi phối.
2. **Hướng gió tính sóng** – chương trình chỉ tính hướng Bắc (V4% = 13,8 m/s; D = 2.970 m), trong khi hướng Đông,
   Tây có V4% = 20,9 / 20,3 m/s. Cần bổ sung bảng đà gió theo các hướng thổi vào đập (đo trên bình đồ lòng hồ) và
   một câu giải thích vì sao hướng Bắc là bất lợi nhất. Hội đồng rất hay hỏi chỗ này.
3. **Độ dốc đáy cống** – thuyết minh dùng i = 3 % (QTVH), tệp số liệu gốc ghi 3 ‰. Nếu bản vẽ cắt dọc cống ghi 3 ‰
   thì với Q = 10 m³/s độ sâu dòng đều h₀ = 1,96 m gần bằng chiều cao cống 2,0 m – không còn chảy không áp, mục 6.3
   phải viết lại. Cần đối chiếu bản vẽ cống.
4. **Độ dư cao trình đỉnh đập chỉ 1 cm** (+78,85 so với yêu cầu +78,84). Điều 11.4.4 TCVN 8216:2018 yêu cầu với đập
   cao trên 20 m phải kể chiều cao dự phòng lún ≥ 0,5 %H vào cao trình đỉnh khi nghiệm thu (với H ≈ 25 m là ≈ 0,12 m).
   Điều này viết cho đập xây mới; đập Cà Giây đã lún ổn định hơn 20 năm và lớp đắp bù chỉ dày 0,3 ÷ 0,5 m, nên có thể
   lập luận giữ +78,85 m – nhưng phải ghi rõ trong thuyết minh và bản vẽ: +78,85 m là cao trình **tối thiểu sau
   lún**, dung sai cao độ mặt bê tông chỉ cho phép dương, lớp đất đắp bù đắp phòng lún. Nếu GVHD yêu cầu dự phòng thì
   chọn +78,90 m (phải sửa đồng bộ Chương 3, 5, 7, kết luận và bản vẽ BV-03, BV-04).
5. Các giả thiết đã ghi ở mục 6 của `HANDOFF.md` (k_h = 0,05; chỉ tiêu lớp II; bản đáy ngưỡng 1 m, f = 0,60;
   m = 0,35 cửa vào cống; G cửa van 20 kN; d chân ren 72 mm; mưa 30 phút = 30 % X_ngày; góc mũi phun 25°…).
6. Bìa: cách ghi "KHOA CÔNG TRÌNH", "ĐỒ ÁN TỐT NGHIỆP KỸ SƯ", tên Bộ chủ quản, "TP. HỒ CHÍ MINH" – xác nhận với
   Bộ môn (Phân hiệu).
7. Tờ nhiệm vụ ĐATN: đóng bản gốc có chữ ký vào trang thứ 3.
8. **Tình huống kẹt van tràn sự cố với lũ P = 0,1 % (QTVH)**: Bảng 2.3 cho Zmax = +79,59 m, **cao hơn đỉnh tường
   chắn sóng +79,40 m**, nhưng mục 4.1 (đoạn "Trường hợp cửa van tràn sự cố bị kẹt…") chỉ bàn trường hợp P = 0,2 %
   (+79,14 m). Nên bổ sung một câu nêu trường hợp này (vượt tần suất kiểm tra của đập cấp II, là sự cố chồng sự cố)
   để nhấn mạnh yêu cầu chạy thử cửa van, nguồn điện dự phòng và quy trình mở cửa thủ công trước mùa lũ.
9. **Số máy đóng mở cống**: mục 3.4 và bảng tổng hợp giải pháp ghi "thay 2 máy đóng mở VĐ20", Bảng 7.1 ghi 2 bộ,
   trong khi mô tả cống chỉ có một cửa van phẳng trong tháp van. Cần đối chiếu TKCS: nếu tháp van có 2 cửa (cửa sửa
   chữa + cửa vận hành) thì ghi rõ ở mục 1.6.4 và Bảng 6.1; nếu không thì sửa số lượng.
10. **Ổn định trượt ngưỡng tràn chính (Bảng 4.4)**: khi c = 0, K = 0,82 / 0,65 < [K]. Kết luận "bảo đảm ổn định"
   đang dựa vào lực dính bê tông – đá chưa được kiểm chứng. Nên đổi cách viết thành có điều kiện (đạt nếu c ≥ 9,1 kPa,
   phải xác minh bằng hồ sơ hoàn công hoặc khoan lấy lõi) và hỏi GVHD có đưa neo D25 thành hạng mục bắt buộc
   (bổ sung vào Bảng 7.1) hay không.
11. Cách ghi chuỗi dòng chảy: "44 năm (1981 – 2025)" là đúng vì là năm thủy văn 1981-1982 ÷ 2024-2025 (44 năm),
   nhưng dễ bị hiểu là 45 năm dương lịch – nên ghi "44 năm thủy văn (1981-1982 ÷ 2024-2025)" ở mục 1.4.2 và 2.2.2.
12. **Lời cam đoan và việc dùng AI.** `00_START_note` (06/10) ghi nguyên tắc "Claude giảng phương pháp, kiểm tra bài
   và chỉ lỗi, không viết hộ", nhưng bản thuyết minh hiện tại do trợ lý AI soạn ở phiên 09/10 (kể cả các câu "các
   bài toán được tác giả lập trình…", "khối lượng … được tác giả tính…"), trong khi Lời cam đoan khẳng định đồ án là
   của bản thân tác giả. Sinh viên nên tìm hiểu quy định của Trường, Bộ môn về sử dụng AI và khai báo nếu được yêu
   cầu, đồng thời tự chạy lại, hiểu và giải thích được từng con số – hội đồng sẽ hỏi.
13. **Khối lượng so sánh phương án màng trên BV-05 chưa có trong thuyết minh.** Bảng "So sánh phương án" của BV-05
   (51 → 39 hố; tổng khoan 506 → 381 m, còn 300 m nếu nâng được đáy màng; tổng phụt 431 → 323 / 243 m) là ước tính
   hình học theo bố trí hố trên bản vẽ, còn mục 8.6 chỉ ghi "đoạn 35 m giảm khoảng 1/3 khối lượng". Nếu muốn nêu các
   con số này khi bảo vệ thì bổ sung một bảng vào mục 8.6 (và tự tính lại); nếu không, chỉ dùng cách nói của mục 8.6.
14. Gradien tiếp xúc ống khói: hiện trạng phân vị 95 % ≤ 1,19 (mục 2.5), sau sửa chữa ≤ 0,93 (mục 5.3) – khi trả lời
   hội đồng nói rõ đang dùng trạng thái nào (slide 8 dùng 1,19 là giá trị lớn hơn).

## 4. Lưu ý khi dùng tệp

- Mở bằng Microsoft Word: nếu hỏi cập nhật trường, chọn **Yes** để Word tính lại số trang theo cách dàn trang của
  Word (bản PDF do LibreOffice dàn trang với phông Liberation Serif cùng kích thước chữ với Times New Roman; vài mục
  có thể lệch ±1 trang nếu máy thiếu phông). Chọn No vẫn thấy đủ mục lục.
- Các sửa trên được làm trực tiếp trong tệp Word bằng `thuyet_minh/hieu_chinh_docx.py` và
  `thuyet_minh/chen_muc_luc.py`. Mã dựng cũ (`build.py`, `c1.py`…`c8.py` trong gói zip ngày 09/10) **chưa** có các
  sửa này – từ nay nên sửa thẳng trong Word; nếu buộc phải dựng lại bằng `build.py` thì chạy tiếp hai chương trình:

  ```bash
  python3 hieu_chinh_docx.py DATN_TranQuocThanh_HoCaGiay.docx b1.docx
  python3 chen_muc_luc.py b1.docx DATN_TranQuocThanh_HoCaGiay.docx DATN_TranQuocThanh_HoCaGiay.pdf
  ```
  (cần LibreOffice có gói `libreoffice-math` – thiếu gói này công thức sẽ không hiện trong PDF).

## 5. Nhận xét về các góp ý trong thư mục "Gop y" (Gemini, GPT, Qwen – 10/10)

Các góp ý được đối chiếu với số liệu gốc và TCVN 8216:2018 trước khi đưa vào mục 3:

| Góp ý | Đánh giá | Căn cứ |
|---|---|---|
| Chuỗi 1981–2025 là 45 năm, phải tính lại P = m/(n+1) | **Sai** – chuỗi có 44 năm thủy văn (1981-1982 ÷ 2024-2025); chỉ cần ghi rõ (mục 3, ý 11) | `tinh_toan/du_lieu/can_bang_nuoc.json` |
| Kẹt van TSC: mực nước vượt tường chắn sóng | **Đúng một phần** – chỉ với lũ P = 0,1 % (QTVH): +79,59 > +79,40 (mục 3, ý 8) | Bảng 2.3 |
| Đỉnh +78,85 chỉ dư 1 cm | **Đúng, cần giải trình** (mục 3, ý 4) | TCVN 8216:2018 điều 10.1.5, 11.4.4 |
| Ngưỡng tràn chính phụ thuộc lực dính, nên neo bắt buộc | **Đúng về nguy cơ**; việc chọn neo bắt buộc do GVHD quyết (mục 3, ý 10) | Bảng 4.4 |
| Độ dốc cống 3 % hay 3 ‰ | **Đúng** – trùng mục 3, ý 3 | — |
| Bảng 7.1 ghi 2 bộ máy VĐ20 cho 1 cống | **Đúng, cần giải thích** (mục 3, ý 9) | Mục 3.4, 6.1, Bảng 7.1 |
| [K] = 1,17 hay 1,13; a, tần suất gió | Đồ án **đúng** | TCVN 8216:2018 Bảng 1, 3, 4 |
| "Bảng PL5 ghi trang 73 nhưng bảng kéo sang trang 74" | **Không phải lỗi** – danh mục ghi trang có tên bảng | — |
| k_h = 0,05 cần căn cứ; TCVN 9386 loại trừ đập lớn | Đúng – đồ án đã ghi là giả thiết; hỏi GVHD tiêu chuẩn động đất áp dụng cho đập | mục 2.1 |
| Thêm bài toán kiểm chứng chương trình Python, lưu đồ thuật toán | Nên làm nếu còn thời gian; đồ án đã có đối chứng với tư vấn (thấm 6,22 so với 5,71·10⁻⁶ m³/s/m; điều tiết lũ lệch < 0,1 m) | mục 2.3, 2.5 |
| Khái toán kinh phí | Chỉ làm nếu nhiệm vụ ĐATN yêu cầu (đề cương coi dự toán là chuyên đề tùy chọn); cần đơn giá Lâm Đồng 2026 | — |
| Báo cáo "Gpt deep research", "Qwen deep research" | Phần lớn là nội dung chung, nhiều chỗ không khớp công trình (ống áp lực, tháp điều áp, tràn xây bậc…) – không dùng | — |


## 6. Góp ý mới trong "Hugo/Gop y" (Tổng hợp, GPT plus, Báo cáo phản biện – chiều 10/10)

Đã đối chiếu từng ý với bản PDF 89 trang. Các ý **đúng** dưới đây chưa sửa trong thuyết minh – SV tự sửa và hỏi GVHD
những chỗ ghi "GVHD quyết".

| # | Góp ý | Kiểm chứng | Việc cần làm |
|---|---|---|---|
| 1 | Cao trình cấu tạo đỉnh đập không khép kín | **Đúng về câu chữ.** Mục 5.1.3 ghi: đất đắp đến +78,60, BT dày 20 cm, đỉnh +78,85, dốc ngang 2 % về hạ lưu – thiếu lớp **cấp phối đá dăm đắp bù** dưới lớp nilon (chi tiết mặt đập TKCS, Hình 3.1: +78,75 = +78,50 + đá dăm ≈ 0,05 + BT 0,20). Bản vẽ BV-04 đã vẽ: đất +78,60 + đá dăm 0,05 + BT 0,20 = +78,85 tại mép thượng lưu (sát tường), dốc 2 % về hạ lưu còn +78,76 ở mép tấm (cách tường 4,3 m). Lưu ý TKCS làm mặt dốc 2 % về hai phía (đỉnh +78,75 cách tường 2,1 m), đồ án đổi sang dốc một phía | Bổ sung lớp đá dăm vào mục 3.2.1, 5.1.3 và khối lượng. Điều kiện đỉnh đập ≥ MNLKT + 0,3 = +78,84 chỉ thỏa ở mép thượng lưu (mép hạ lưu +78,76) – kiểm tra tại mép nào, có cần nâng thêm không: **GVHD quyết** |
| 2 | Mô hình thấm, ổn định sau sửa chữa và mô hình chuyên đề vẫn dùng đỉnh +78,75 | **Đúng** (mục 5.1.2, 5.2, 8.3.1) | Chạy lại với +78,85 (sửa một số liệu trong `tinh_toan/so_lieu.py`) hoặc nêu định lượng sai khác, thay cho câu "không ảnh hưởng" |
| 3 | Mức bảo đảm P = m/(n + 1) chưa có căn cứ | **Đúng một phần** – công thức tần suất kinh nghiệm dùng phổ biến, nhưng thuyết minh chưa dẫn nguồn | Dẫn nguồn cho công thức (ví dụ giáo trình điều tiết hồ chứa [13] hoặc giáo trình thủy văn công trình); nếu dùng m/n: 18/44 = 40,9 %, 39/44 = 88,6 %, 37/44 = 84,1 % – kết luận không đổi (TH1 không đạt, TH1A đạt, TH2 vẫn < 85 %) |
| 4 | Z–W tại +75 m: Bảng 1.2 ghi 38,25, PL2 ghi 38,32 triệu m³ | **Đúng, đã tìm ra nguyên nhân**: Bảng 1.2 cộng dồn bước 1 m (từ +74), PL2 bước 0,5 m (qua +74,5, F = 6,13 km²); công thức chóp cụt cho hai kết quả lệch 0,07 triệu m³ (0,2 %) | Dùng một bảng cho mọi tính toán (nên dùng bước 0,5 m) và ghi chú; kiểm tra chương trình điều tiết lũ đang đọc bảng nào |
| 5 | Mục 6.5: tải đất lên cống tăng 4,4 % "nằm trong hệ số an toàn" | **Đúng** – mức tăng tải không chứng minh khả năng chịu lực khi không có bản vẽ kết cấu | Chỉ nêu mức tăng tải; chuyển thành kiến nghị kiểm tra kết cấu khi có hồ sơ hoàn công, kết quả kiểm tra camera |
| 6 | TCVN 4116:1985 (tài liệu [30]) đã hết hiệu lực | **Đúng** – đã có TCVN 4116:2023 "Công trình thủy lợi – Kết cấu bê tông và bê tông cốt thép thủy công – Yêu cầu thiết kế" thay thế (tra trên cơ sở dữ liệu VSQI, luatvietnam) | Thay [30] và câu viện dẫn ở mục 6.5 |
| 7 | Kẹt van tràn sự cố: so với mặt đập +78,85 và tường +79,40 | **Đúng** – +79,14 (BCNCKT) và +79,18 (QTVH) đã cao hơn mặt đập; +79,59 (QTVH, P = 0,1 %) cao hơn cả tường. Thuyết minh (mục 4.1) xếp vào "sự cố chồng sự cố", nhưng ngầm coi tường chắn sóng giữ được nước | Bổ sung sau Bảng 2.3 một bảng phân loại: tổ hợp thiết kế / kiểm tra / khảo sát rủi ro; không coi tường giữ nước khi chưa kiểm tra chịu lực, chống thấm; giữ biện pháp vận hành (kiểm tra van, nguồn dự phòng). Đây là câu hội đồng dễ hỏi nhất (xem slide dự phòng) |
| 8 | Thi công tháng 1 – 6 nhưng mùa mưa bắt đầu tháng 5 (mục 1.4.1) | **Đúng** | Thêm biện pháp cho tháng 5 – 6 (thoát nước, che phủ, ứng phó lũ sớm) hoặc dời hạng mục mái hạ lưu, mặt đập lên trước |
| 9 | Ngưỡng cảnh báo "vượt giá trị dự báo trên 1,5 m" (mục 7.5) chưa có căn cứ | **Đúng** | Đổi thành: lập đường cơ sở sau sửa chữa theo mực nước hồ, cảnh báo theo độ lệch, tốc độ tăng, lưu lượng và độ đục nước thấm |
| 10 | Mưa 30 phút = 30 % mưa ngày (mục 5.5) là giả thiết mạnh | **Đúng** | Thêm độ nhạy: theo (5-2) τ tỷ lệ với i^0,6, nên với 20 % / 40 % mưa ngày τ của ô 4 × 4 m là ≈ 1,4 / 2,2 Pa (so với 1,8 Pa) – SV tự tính lại và ghi vào mục 5.5 |
| 11 | Kết luận động đất, ngưỡng tràn, thân cống viết như kết luận chắc chắn | **Đúng** | Viết "đạt với giả thiết k_h = 0,05", "đạt có điều kiện c ≥ 9,1 kPa", "chưa kiểm tra chịu lực" (trùng mục 3, ý 5, 10) |
| 12 | "LỜI CÁM ƠN" | Nên thống nhất "LỜI CẢM ƠN" | Sửa tiêu đề trang |
| 13 | Thêm Tóm tắt 200 – 300 từ; bảng kiểm chứng chương trình Python; bảng tổng hợp cuối Chương 2 | Hướng dẫn 2020 của Trường không bắt buộc Tóm tắt | Làm nếu GVHD đồng ý; bảng kiểm chứng nên làm (đã có số liệu đối chứng ở mục 2.3, 2.5) |
| 14 | Spencer/Morgenstern–Price, Monte Carlo, HEC-RAS/PLAXIS, công khai GitHub, thay cống, neo D25 bắt buộc | Không ưu tiên | Không đưa vào trước khi xong các ý 1 – 11 |
| 15 | Tệp "10. ĐC TTN 27.06.2023 - K64" là đề cương **Thực tập** tốt nghiệp | **Đúng** – không có quy định khổ, số tờ bản vẽ ĐATN | Hỏi GVHD xác nhận khổ giấy, số tờ (SV nhớ 6 – 8 tờ A0 – bộ bản vẽ đã chuyển sang A0) |

**Câu hỏi hội đồng dễ hỏi** (đã đưa vào slide dự phòng): vì sao chuỗi 1981 – 2025 là 44 năm; vì sao P = m/(n + 1);
+78,60 + 0,20 sao ra +78,85; bỏ lực dính thì ngưỡng không đạt sao kết luận ổn định; k_h lấy ở đâu; vì sao chọn đỉnh
+78,85 khi kẹt van cho mực nước cao hơn.
