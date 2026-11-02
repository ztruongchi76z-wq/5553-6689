# Biên bản nghiệm thu tái cấu trúc hồ sơ (02/11/2026)

Kết quả: sáu phép thử tra cứu đều đạt trên tài liệu thật trong Google Drive. Kiểm tra tự động không có lỗi. Bộ làm việc và bộ chuẩn bị công khai vẫn ở chế độ riêng tư; thư mục nguồn không bị thay đổi. Việc chia sẻ công khai và gửi thư do người lập quyết định.

## Phạm vi đã làm

| Hạng mục | Kết quả |
|---|---|
| Bộ làm việc riêng tư | 117/117 tệp của thư mục nguồn được sao chép phía máy chủ, đối chiếu đủ; 5 tài liệu tra cứu và 1 tài liệu “việc cần người lập quyết định” |
| Bộ chuẩn bị công khai | 77 tệp đã rà soát (gồm 15 chứng từ gửi, nhận); 6 bản chép văn bản cơ quan thay họ tên giáo viên, chuyên viên, người bị tố cáo bằng chức danh; 5 tài liệu tra cứu (00, 01, 02, 04, 05) |
| Chưa đưa vào bộ công khai | 21 tệp chờ quyết định (D); 5 tệp chỉ công khai dạng bản chép (T); 14 tệp chỉ dùng trong bộ làm việc (X) |
| Quyết định của người lập ngày 02/11/2026 | Thông tin cá nhân của người lập được công khai; chỉ yêu cầu che mặt học sinh. 22 tệp trước đó bị giữ lại vì thông tin của người lập đã được đưa vào bộ công khai và gắn liên kết trong tài liệu tra cứu |
| Quyền chia sẻ | Bộ làm việc, bộ công khai: chỉ chủ sở hữu. Thư mục nguồn: giữ nguyên như trước (bất kỳ ai có đường liên kết được xem) |

## Sáu phép thử

1. **Một lá đơn (HS21, đơn gửi Chánh Thanh tra tỉnh ngày 06/10/2026).** Tài liệu 01, mục HS21, cho đủ: ngày đề 06/10, gửi 08/10 (mã bưu gửi), nhận 15g08 ngày 12/10, văn bản trả lời Công văn 3292 (liên kết ảnh hai trang), 9 yêu cầu: 4 đáp ứng một phần, 5 chưa giải quyết. Tìm toàn văn trên Drive theo mã bưu gửi của đơn này trả về đúng tài liệu 01. Đạt.
2. **Vấn đề “Hồ sơ tố cáo ngày 09/7/2026”.** Tài liệu 02, mục VD06: câu hỏi trung tâm, chứng cứ (chứng từ nhận, bản bổ sung), văn bản liên quan (Phiếu chuyển 2329, Văn bản 2644, Công văn 3292; Thông báo 5553 và Công văn 6689 không nhắc hồ sơ), 35 yêu cầu liên quan, trong đó 32 chưa có kết quả. Chuỗi 09/7 được trình bày tách khỏi chuỗi đơn tháng 5. Tìm toàn văn “VD06” trả về tài liệu 02. Đạt.
3. **Văn bản “Công văn 3292”.** Tài liệu 05: cơ quan, người ký theo chức danh, nơi gửi, trích yếu, chức năng pháp lý, hồ sơ liên quan, liên kết hai trang ảnh; tài liệu 01 mục HS21 ánh xạ từng yêu cầu với nội dung Công văn 3292. Tìm toàn văn số hiệu “3292/TTr-PKNTC.1” trả về tài liệu 05. Đạt.
4. **Một luận điểm truy về chứng cứ.** Tài liệu 04, LD01 (“xương sống cầu giàn” mâu thuẫn với trang “thêm thanh chéo thành 2 tam giác”): liên kết thẳng tới ảnh trang bài giảng 06.03 và 06.01 trong bộ công khai, kèm điểm rà soát KT10 (nhóm D, đã vững). Đạt.
5. **Người phản biện.** Tài liệu 04, phần 2, liệt kê sẵn các điểm yếu: KT07 (nhận định “không đơn nào đề ngày 18/5” chưa đủ chứng cứ), KT08 (mệnh đề góc nhọn của thang nâng chỉ đúng khi góc tại khớp nhỏ hơn 90°), KT11 (thang nâng là ứng dụng đúng của hình bình hành), KT14 (“đảo chủ ngữ” không vững), KT21, KT22 (cáo buộc vượt chứng cứ, suy diễn động cơ). Đạt.
6. **Tính độc lập của bản sao.** Mã tệp của bản làm việc, bản công khai và tệp nguồn không trùng nhau (0 trùng ở cả ba cặp). 371 liên kết (sau cập nhật ngày 02/11: 353 liên kết trong bốn tài liệu 00, 01, 02, 04 được kiểm tra lại) trong các tài liệu tra cứu công khai đều trỏ tới tệp hoặc thư mục của bộ công khai; không liên kết nào trỏ về bộ làm việc hay thư mục nguồn. Quyền của hai bộ mới chỉ gồm chủ sở hữu. Đạt.

## Kiểm tra tự động

`scripts/kiem_tra.py`: 0 lỗi, 0 cảnh báo; 6 hồ sơ chưa có chứng từ nhận trong thư mục (HS01, HS02, HS03, HS05, HS07, HS22). Xem `KIEM_TRA_TU_DONG.md`.

## Giới hạn

- Bản chép văn bản cơ quan dựa trên nhận dạng chữ từ ảnh; có thể còn lỗi nhỏ, cần đối chiếu ảnh gốc khi trích dẫn chính thức.
- Tài liệu trên Drive là ảnh chụp dữ liệu ngày 02/11/2026, không tự đồng bộ với kho.
- Chưa đối chiếu nguyên văn: Luật 136/2025/QH15, Nghị định 155/2026/NĐ-CP, Luật 121/2025/QH15, hiệu lực Thông tư 05/2021/TT-TTCP năm 2026, quy định của Luật Nhà giáo 2025 về phát tán thông tin.
- Chưa xem: hai video (02.4, 03.6), hai ảnh tin nhắn (02.2, 02.3); chưa rà soát toàn văn các phụ lục 07.1–07.8, 12.10, 17.17a.
