# Hồ sơ Trường THCS Trần Phú: dữ liệu tra cứu và kiểm tra

Thư mục này là môi trường làm việc chính cho việc tái cấu trúc hồ sơ (11/2026): dữ liệu vòng đời từng đơn, chương trình kiểm tra tự động, chương trình xuất tài liệu tra cứu và kết quả nghiệm thu. Tài liệu đọc được nằm trên Google Drive, trong hai bộ thư mục riêng tư của người lập (bộ làm việc và bộ chuẩn bị công khai); kho này không chứa tệp gốc.

## Nguyên tắc

- Kho đang công khai, nên dữ liệu ở đây không có họ tên, địa chỉ, số điện thoại, số CCCD, mã thư mục hay mã tệp Drive. Người gửi đơn được gọi là “người lập”; cán bộ, giáo viên được gọi theo chức danh.
- Bảng nối mã tài liệu với tệp Drive được giữ riêng, ngoài kho.
- Bản đã nộp cơ quan không bị sửa. Điểm sai, điểm chưa chặt được ghi trong `data/kiem_toan.tsv`.
- Mỗi yêu cầu trong đơn có hai trạng thái độc lập: thủ tục (đã có văn bản xử lý chưa) và nội dung (đã được đáp ứng đến đâu).

## Cấu trúc

| Đường dẫn | Nội dung |
|---|---|
| `data/ho_so.tsv` | 25 hồ sơ (22 đã gửi, 3 đã soạn ngày 01/11/2026): ngày đề, ngày gửi, mã bưu gửi, ngày nhận, chứng từ, văn bản xử lý |
| `data/yeu_cau.tsv` | 150 yêu cầu trong các đơn, mỗi yêu cầu có trạng thái thủ tục và trạng thái nội dung |
| `data/van_ban.tsv` | 16 văn bản của nhà trường và cơ quan nhà nước, chức năng pháp lý |
| `data/van_de.tsv` | 14 vấn đề tra cứu (VD01–VD14) |
| `data/luan_diem.tsv` | 14 luận điểm, tách sự kiện và suy luận |
| `data/kiem_toan.tsv` | 30 điểm rà soát: A sai cần sửa, B chưa chặt, C phát hiện mới, D đã vững |
| `data/can_cu.tsv` | 19 căn cứ pháp lý và mức độ đã đối chiếu nguyên văn |
| `data/quan_he.tsv` | quan hệ giữa đơn và văn bản (chuyển, trả lời, đôn đốc) |
| `data/tai_lieu.tsv` | 117 tệp của thư mục nguồn, vị trí mới, quyết định công khai (P, T, D, X) |
| `data/toan_ven.tsv` | mã băm của các tệp ảnh đã kiểm tra |
| `scripts/kiem_tra.py` | kiểm tra tính toàn vẹn dữ liệu, tính lại mốc thời hạn theo ngày làm việc, đối chiếu các con số 16 và 18 |
| `scripts/xuat_ban.py` | sinh tài liệu tra cứu (Markdown cho kho, HTML cho Google Docs) |
| `tra-cuu/` | bản Markdown của tài liệu 01, 02, 04, 05 |
| `bao-cao/` | báo cáo kiểm tra tự động và biên bản nghiệm thu |

## Chạy

```
python3 scripts/kiem_tra.py            # ghi bao-cao/KIEM_TRA_TU_DONG.md; mã thoát 1 nếu có lỗi
python3 scripts/xuat_ban.py md tra-cuu # sinh lại tài liệu tra cứu dạng Markdown
python3 scripts/xuat_ban.py html <thư_mục_ra> <thư_mục_bản_đồ_riêng>
```

Chế độ `html` cần bảng nối mã tài liệu với tệp Drive (giữ riêng). Tài liệu trên Drive là ảnh chụp dữ liệu ngày 02/11/2026; khi dữ liệu thay đổi, sinh lại và cập nhật bản trên Drive.

## Ký hiệu công khai trong `tai_lieu.tsv`

- `P`: đưa vào bộ công khai nguyên bản (đã rà soát).
- `T`: không công khai ảnh gốc; dùng bản chép đã che thông tin cá nhân.
- `D`: chờ người lập quyết định.
- `X`: chỉ trong bộ làm việc.
