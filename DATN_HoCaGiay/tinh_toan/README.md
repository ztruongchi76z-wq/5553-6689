# Chương trình tính toán (Python 3) – bản 11/10/2026

Mô hình sau sửa chữa dùng đỉnh +78,85 m (`so_lieu.py`, `Z_DINH_DAP_SC`). Kết quả JSON trong `ket_qua/`.

| Tệp | Nội dung |
|---|---|
| so_lieu.py, dieu_tiet_lu.py, cao_trinh_dinh.py, tham_fem.py, on_dinh.py, chay_tham_on_dinh.py, chay_truot_nong.py, cong_va_tran_ket_cau.py, chuyen_de_khoan_phut.py, ve_hinh.py | Chương trình gốc (Bảng PL5) |
| phan_tich_bo_sung.py | Spencer, lưới tâm mịn, tải trọng thi công (TH11, TH12) |
| bo_sung_giai_tich.py | Hướng gió, độ nhạy m tràn, cống i = 3‰, neo ngưỡng, khung kín cống (a = 7 cm: lớp bảo vệ ≥ 60 mm theo TCVN 4116:2023 mục 6.3.3 – Bảng 6.4), mưa, m/n |
| kt_hoi_tu.py | Hội tụ lưới PTHH, tham số α, độ nhạy k_h |
| kt_hai_hang.py, kt_dinh.py | Màng 2 hàng A–B, đáy màng theo TCVN 8216:2018 mục 9.2.2.9, ảnh hưởng đỉnh +78,85 |
| kt_tran_su_co.py | Thủy lực tràn sự cố: độ sâu đầu dốc = h_k (P = 0), đường mặt nước, tường yêu cầu (ket_qua/kt_tran_su_co.json) |
| kt_lop_II_lu.py | Hiệu quả màng theo K lớp II 1e-5; 1,3e-5 (≈ 1 Lu); 3e-5; 6,5e-5; 1e-4 cm/s (tiêu chí 5 Lu) – Bảng 8.8 (ket_qua/kt_lop_II_lu.json) |
| kt_dong_chay_toi_thieu.py | Điều tiết toàn chuỗi có kể dòng chảy tối thiểu Q_tt = 0,027 m³/s (TCVN 10778:2024 công thức 6), mức bảo đảm, MNDBT cần cho P = 85 % – mục 2.2.3 (ket_qua/kt_dong_chay_toi_thieu.json); cần du_lieu/can_bang_nuoc.json đi kèm can_bang_nuoc.py |

Các tệp gốc còn lại (can_bang_nuoc.py, thuy_luc_tran.py, mai_ha_luu.py, tong_hop.py, xuat_excel.py và du_lieu/can_bang_nuoc.json) nằm trong thư mục Drive "Hugo/Tinh toan bo sung (chua co tren GitHub)" – chép vào đây (giữ thư mục du_lieu/) để chạy.
Chạy: `python3 chay_tham_on_dinh.py` (cần numpy, scipy, triangle, matplotlib).
