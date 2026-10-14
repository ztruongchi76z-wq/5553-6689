# Giáo án Tin học 6 tích hợp năng lực số (NLS) và AI – năm học 2026–2027

Soạn theo Phụ lục I – Kế hoạch dạy học môn Tin học khối 6 (Tổ Toán – Tin, Trường THCS Trần Phú), bộ sách Kết nối tri thức với cuộc sống.

- `HK1/` – Bài 1–8 (mỗi bài một file) và file tổng hợp học kì I.
- `HK2/` – Bài 9–17 (mỗi bài một file) và file tổng hợp học kì II.
- `GA_Tin6_NLS_AI_*.zip` – bản nén để tải về (HK1, HK2, cả năm).
- `_ma_nguon/` – mã nguồn tạo file Word: nội dung từng bài nằm trong `lessons/baiXX.js`.
  - Tạo lại file: `node build_all.js 1 <thư_mục_ra>` (HK1) hoặc `node build_all.js 2 <thư_mục_ra>` (HK2); cần gói npm `docx`.
  - Đối chiếu số tiết, mã NLS/AI với KHDH: `node check_khdh.js`.
