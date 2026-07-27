# Checklist nghiệm thu SỸ LAND Creator bản thương mại

Tài liệu này chia việc nghiệm thu thành hai giai đoạn để chỉ thuê GPU khi mã
nguồn, dữ liệu và cấu hình đã sẵn sàng. Không nhập khóa bí mật vào tài liệu,
GitHub Issue, Pull Request hoặc cửa sổ chat.

## A. Hoàn thiện trước khi thuê máy chủ

- [x] Giao diện Creator nằm trong website SỸ LAND hiện có.
- [x] Dùng chung phiên đăng nhập Supabase của SỸ LAND.
- [x] Nhập liên kết hoặc tải video lên; chọn AI, Whisper, giọng đọc và phụ đề.
- [x] Chặn URL không phải HTTPS, host giả mạo và nguồn không được hỗ trợ.
- [x] Chặn video riêng tư, cần đăng nhập, hội viên, giới hạn tuổi và phát trực tiếp.
- [x] Không nhận cookie nền tảng hoặc API key AI từ trình duyệt.
- [x] Tách API CPU, worker GPU, Redis và scheduler dọn tệp.
- [x] Mỗi tài khoản chỉ đọc dữ liệu của chính mình bằng Supabase RLS.
- [x] Giữ, sử dụng và hoàn hạn mức phút bằng RPC có khóa giao dịch.
- [x] Liên kết tải MP4 có chữ ký và thời hạn ngắn.
- [x] Thông báo lỗi công khai không chứa API key, đường dẫn hoặc log nội bộ.
- [x] Kiểm tra cấu trúc bản dịch AI trước khi tạo giọng đọc.
- [x] Lưu đúng trạng thái quyền sở hữu: verified/declared/unverified.
- [x] Có công cụ preflight phát hiện biến thiếu và giá trị mẫu.
- [x] Có CI build website, kiểm tra bảo mật, unit test và Docker Compose.
- [ ] Kiểm tra migration `SUPABASE_CREATOR.sql` trên bản sao dự án Supabase.
- [ ] Chốt hạn mức phút và thời lượng tối đa của từng gói thương mại.
- [ ] Chuẩn bị 20 video kiểm thử có quyền sử dụng hợp lệ.
- [ ] Chuẩn bị nội dung điều khoản sử dụng Creator và chính sách lưu video 24 giờ.

## B. Thực hiện sau cùng khi thuê GPU

- [ ] Bật MFA cho tài khoản nhà cung cấp GPU.
- [ ] Nạp ngân sách thử nghiệm đã được chủ tài khoản phê duyệt.
- [ ] Tạo máy chủ staging; chưa kết nối website sản xuất.
- [ ] Tự nhập khóa Supabase, OpenAI/Gemini và signing secret trên máy chủ.
- [ ] Chạy `python scripts/check-creator-env.py .env.creator` và đạt preflight.
- [ ] Khởi động Redis, API, worker GPU và scheduler.
- [ ] `/health` trả `ok=true`.
- [ ] `/ready` xác nhận storage, database và queue đều hoạt động.
- [ ] Kiểm thử tải lên MP4/MOV/WebM/MKV và giới hạn dung lượng.
- [ ] Kiểm thử link công khai từ YouTube, TikTok, Douyin và Bilibili.
- [ ] Kiểm thử video tiếng Trung, Hàn và Thái.
- [ ] Kiểm thử đủ bốn lựa chọn giọng Việt và ba cỡ phụ đề.
- [ ] Kiểm thử hủy tác vụ, hoàn phút và xóa tệp hết hạn.
- [ ] Kiểm thử hai tài khoản để bảo đảm không xem chéo dữ liệu.
- [ ] Đo thời gian, chi phí AI và GPU trên 20 video đại diện.
- [ ] Chỉ sau khi nghiệm thu mới cấu hình `VITE_SYLAND_CREATOR_API_URL`.
- [ ] Chỉ merge `main` và triển khai website khi được phê duyệt riêng.

## Điều kiện được phép mở bán

Creator chỉ được mở bán khi toàn bộ mục A và B đã hoàn thành, CI đang xanh,
không có khóa bí mật trong Git, `/ready` đạt, kiểm thử phân quyền hai tài khoản
đạt và có số liệu chi phí thực tế để chốt hạn mức/gói.
