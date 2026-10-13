# Chương trình tính toán (Python 3) – bản 12/10/2026

Mô hình sau sửa chữa dùng đỉnh +78,85 m (`so_lieu.py`, `Z_DINH_DAP_SC`). Kết quả JSON trong `ket_qua/`.

| Tệp | Nội dung |
|---|---|
| so_lieu.py, dieu_tiet_lu.py, cao_trinh_dinh.py (đợt 6: Vw = kl·Vl theo TCVN 8421:2010 A.3.3, kα Bảng 9), tham_fem.py, on_dinh.py, chay_tham_on_dinh.py, chay_truot_nong.py, cong_va_tran_ket_cau.py, chuyen_de_khoan_phut.py, ve_hinh.py | Chương trình gốc (Bảng PL5) |
| phan_tich_bo_sung.py | Spencer, lưới tâm mịn, tải trọng thi công (TH11, TH12) |
| bo_sung_giai_tich.py | Hướng gió (Vw quy đổi), [K] ngưỡng tràn = γn·γlc/γc (γc = 0,95), độ nhạy m tràn, cống i = 3‰, neo ngưỡng, khung kín cống (a = 7 cm: lớp bảo vệ ≥ 60 mm theo TCVN 4116:2023 mục 6.3.3 – Bảng 6.4), mưa, m/n |
| kt_hoi_tu.py | Hội tụ lưới PTHH, tham số α, độ nhạy k_h |
| kt_hai_hang.py, kt_dinh.py | Màng 2 hàng A–B, đáy màng theo TCVN 8216:2018 mục 9.2.2.9, ảnh hưởng đỉnh +78,85 |
| kt_tran_su_co.py | Thủy lực tràn sự cố: độ sâu đầu dốc = h_k (P = 0), đường mặt nước, tường yêu cầu (ket_qua/kt_tran_su_co.json) |
| kt_lop_II_lu.py | Hiệu quả màng theo K lớp II 1e-5; 1,3e-5 (≈ 1 Lu); 3e-5; 6,5e-5; 1e-4 cm/s (tiêu chí 5 Lu) – Bảng 8.8 (ket_qua/kt_lop_II_lu.json) |
| kt_dong_chay_toi_thieu.py | Điều tiết toàn chuỗi có kể dòng chảy tối thiểu Q_tt = 0,027 m³/s (TCVN 10778:2024 công thức 6), mức bảo đảm, MNDBT cần cho P = 85 % – mục 2.2.3 (ket_qua/kt_dong_chay_toi_thieu.json); cần du_lieu/can_bang_nuoc.json đi kèm can_bang_nuoc.py |
| kt_cap_nuoc_bo_sung.py | Mức bảo đảm theo mục tiêu (TCVN 10778:2024 mục 8.2.2.2), sinh hoạt ưu tiên, số tháng MN < +67,94 m (ket_qua/kt_cap_nuoc_bo_sung.json) |
| xuat_excel_dieu_tiet_toan_chuoi.py | Xuất `../TAI_VE/HuongDan_DieuTiet_ToanChuoi_44nam.xlsx` – file Excel hướng dẫn điều tiết toàn chuỗi (toàn công thức, thuật toán như kt_dong_chay_toi_thieu.dieu_tiet_chuoi; kiểm tra lại bằng LibreOffice: TH1A 39/44, TH2 37/44, TH1 18/44) |
| xuat_excel_dieu_tiet_nam_lap_bang.py | Xuất `../TAI_VE/DieuTietNam_LapBang_CoTonThat.xlsx` – điều tiết năm thiết kế P = 85% theo phương pháp lập bảng có kể tổn thất (tích sớm, lặp đến khi Vhi hội tụ ≤ 1%); kiểm tra LibreOffice: TH1A 27,85 / TH2 28,81 / TH1 36,19 triệu m³ |
| kt_mua_lu_bien_doi.py | Mưa 1978–2024 (Mann – Kendall, Sen, Welch – du_lieu/mua_luu_vuc.json), điều tiết khi lũ tăng k, mức tăng giới hạn, mực nước đón lũ – mục 4.6 (ket_qua/kt_mua_lu_bien_doi.json) |
| kt_bo_sung_ket_cau.py | Neo ngưỡng ([K] kể γc = 0,95 – QCVN 04-05:2022; phía bê tông, nhổ khối đá), ổn định trục vít theo sơ đồ gối (μ = 1,0), τ theo đường chéo ô rãnh (ket_qua/kt_bo_sung_ket_cau.json) |
| kt_kiem_chung_12_10.py | Gradien lớn nhất và vị trí, bất đẳng hướng kh/kv = 5, tìm lại cung trượt (miền mở rộng, lưới lồng, Nelder – Mead, Spencer) – Bảng PL10, PL11 (ket_qua/kiem_chung_*.json) |

Số liệu `du_lieu/can_bang_nuoc.json` (cân bằng nước 44 năm) và `du_lieu/mua_luu_vuc.json` (mưa năm, X1max 1978 ÷ 2024 – Bảng PL12) đã có trong kho (12/10). Các tệp gốc còn lại (can_bang_nuoc.py, thuy_luc_tran.py, mai_ha_luu.py, tong_hop.py, xuat_excel.py) nằm trong thư mục Drive "Hugo/Tinh toan bo sung (chua co tren GitHub)" – chép vào đây để chạy.
Chạy: `python3 chay_tham_on_dinh.py` (cần numpy, scipy, triangle, matplotlib).
