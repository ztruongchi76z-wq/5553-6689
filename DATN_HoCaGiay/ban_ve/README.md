# Bộ bản vẽ A0 – ĐATN hồ chứa nước Cà Giây

7 tờ A0 (1189 × 841 mm), khung tên theo mẫu của Trường, mỗi tờ có tệp DXF (AutoCAD 2018, mở và sửa được bằng
AutoCAD/ZWCAD) và PDF. Bản in gộp: `xuat/BoBanVe_A0_DATN_HoCaGiay.pdf` (có mục lục theo tờ).

| Tờ | Tệp | Nội dung | Tỷ lệ |
|---|---|---|---|
| BV-01 | `BV01_MatBangTongThe` | Mặt bằng tổng thể cụm đầu mối, bố trí các hạng mục sửa chữa; hình trích khu vực tràn sự cố và khu vực cống; bảng tổng hợp giải pháp, thông số chủ yếu | 1/2000, 1/1000 |
| BV-02 | `BV02_CatDocDap` | Cắt dọc tim đập (từ lý trình 40 m đến DT): mặt đập hiện trạng theo cọc, đỉnh TKCS +78,75 và đồ án +78,85, địa tầng, chân khay, hố khoan, màng khoan phụt; bảng cao trình đỉnh đập (Bảng 2.4, 2.5) | ngang 1/1000, đứng 1/100 |
| BV-03 | `BV03_MatCatNgangDap` | Mặt cắt ngang đập hiện trạng và sau sửa chữa; đường bão hòa, cung trượt nguy hiểm; bảng thấm, ổn định | 1/200 |
| BV-04 | `BV04_ChiTietMatDap_Ranh` | Chi tiết mặt đập sau sửa chữa (cao trình hai mép tấm), gờ chắn bánh, các loại rãnh, mặt bằng khai triển hệ rãnh mái hạ lưu, xử lý hố xói; khối lượng | 1/20, 1/25, 1/250 |
| BV-05 | `BV05_KhoanPhut_ChuyenDe` | Chuyên đề: cắt dọc tuyến màng, mặt cắt tính thấm, mặt bằng bố trí hố khoan phương án hợp lý hóa, sơ đồ hố khoan, bảng kết quả, độ nhạy, so sánh phương án, mực nước dự báo giếng đo áp | ngang 1/100 – đứng 1/50, 1/100, 1/200 |
| BV-06 | `BV06_TranChinh` | Tràn chính: cắt dọc với đường mặt nước P = 1 %, 0,2 %; mặt bằng; mặt cắt 1-1; mặt cắt ngưỡng và sơ đồ ổn định; chi tiết xử lý mặt dốc, khe co giãn | ngang 1/200 – đứng 1/100, 1/200, 1/50, 1/25, 1/10 |
| BV-07 | `BV07_CongLayNuoc` | Cống lấy nước: đường mặt nước trong thân cống, mặt cắt ngang và tải trọng, sơ đồ cửa van – trục vít – máy VĐ20 có giá đỡ trung gian, biểu đồ khả năng tháo và lực đóng mở | ngang 1/200 – đứng 1/50, 1/20 |

## Dựng lại

```bash
cd ban_ve
python3 to01_mat_bang.py xuat      # tương tự to02 ... to07
```

Cần Python 3 với `ezdxf` (≥ 1.4) và `matplotlib`. Chữ dùng kiểu chữ `VN` (Arial – có sẵn trên Windows); PDF xem trước
dựng bằng phông Liberation Sans cùng kích thước.

Cách dựng tờ A0 (`bv_lib.py`): mỗi hình chiếu, bảng, ghi chú là một khối do một hàm vẽ ra (theo bố cục A1 cũ), rồi
được đặt lên tờ bằng `To.dat(ham, f=..., neo=..., xy=...)` – tịnh tiến và phóng đồng dạng hệ số f. f chỉ chọn sao cho
tỷ lệ mới là tỷ lệ chuẩn (1/2500 × 1,25 → 1/2000; 1/250 → 1/200; 1/25 → 1/20; 1/1000 × 2 → 1/500…); chữ "TỶ LỆ" trong
khối được đổi tự động, kích thước vẫn ghi số thật. Muốn dời một hình: sửa tọa độ `xy` trong hàm `ve()` của tờ đó.

## Nguồn số liệu, giới hạn

- Số liệu tính toán lấy từ thuyết minh và `ket_qua/*.json` của đồ án (Chương 2 – 8).
- Bình đồ (BV-01), cắt dọc đỉnh đập (BV-02), cắt dọc tuyến màng khoan phụt (BV-05) số hóa từ bản vẽ TKCS dạng DWG
  (`du_lieu/binh_do_tkcs.json`, `du_lieu_cat_doc_dap.py`, `du_lieu_ban_ve.py`); sai số số hóa ≈ 0,1 ÷ 0,2 m.
- Các bộ phận không có kích thước trong hồ sơ (tháp van, bể tiêu năng, mũi phun, chiều dày bản đáy, nắp cống…)
  vẽ **sơ đồ** và ghi "giả thiết" trên bản vẽ – phải đối chiếu hồ sơ hoàn công trước khi dùng.
- Mặt đập (BV-04): đất đắp +78,60 + cấp phối đá dăm đệm 5 cm + BT M250 20 cm = +78,85 ở mép thượng lưu, dốc ngang
  2 % → +78,76 ở mép tấm. Thuyết minh mục 5.1.3 chưa ghi lớp đá dăm – xem `RA_SOAT_THUYET_MINH.md` mục 6.
- Phương án hợp lý hóa màng khoan phụt (BV-05) là đề xuất có điều kiện của chuyên đề, chỉ áp dụng khi phụt thí nghiệm
  đạt yêu cầu.
- Sinh viên cần tự kiểm tra lại từng con số, chỉnh sửa theo góp ý của GVHD; tên người hướng dẫn, người thực hiện và
  ngày tháng trong khung tên sửa trong `bv_lib.py` (hàm `ve_khung`).
