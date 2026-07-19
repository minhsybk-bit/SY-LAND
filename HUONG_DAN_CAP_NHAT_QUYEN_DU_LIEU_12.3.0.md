# Cập nhật SỸ LAND 12.3.0 — Quyền dữ liệu cá nhân

## Nội dung đã hoàn thiện

- Người dùng đăng nhập có thể tải bản sao dữ liệu tài khoản dạng JSON.
- Bản xuất gồm tài khoản, bản quyền, thiết bị, yêu cầu tư vấn, phiếu hỗ trợ và yêu cầu quyền dữ liệu.
- Hồ sơ Word, PDF và Excel xử lý cục bộ không bị tải lên hoặc đưa vào bản xuất.
- Người dùng có thể gửi yêu cầu xóa tài khoản và nhận mã theo dõi.
- Admin thấy toàn bộ yêu cầu, đổi trạng thái và nhập phản hồi cho người dùng.
- Row Level Security bảo đảm người dùng thường chỉ thấy yêu cầu của chính mình.
- Không xóa tài khoản tự động từ trình duyệt; bước xóa cuối cùng cần admin xác minh.

## Các bước cập nhật một lần

1. Tải toàn bộ tệp và thư mục trong gói ZIP lên thư mục gốc kho `SY-LAND`, giữ nguyên đường dẫn.
2. Mở Supabase → **SQL Editor** → **New query**.
3. Mở tệp `supabase/SUPABASE_PRIVACY_REQUESTS.sql`, sao chép toàn bộ nội dung vào SQL Editor và chọn **Run**.
4. Mở GitHub → **Actions** → chờ quy trình `Deploy SỸ LAND to GitHub Pages` có dấu tích xanh.
5. Mở website, nhấn `Ctrl + F5`, đăng nhập và kéo đến mục **Quyền dữ liệu cá nhân** để kiểm tra.

## Kiểm thử nhanh

1. Tài khoản người dùng: chọn **Tải tệp JSON**, kiểm tra trình duyệt tải tệp `SYLAND_DU_LIEU_CA_NHAN_YYYY-MM-DD.json`.
2. Tài khoản người dùng: gửi yêu cầu xóa, kiểm tra xuất hiện mã dạng `PRV-YYMM-xxxxx`.
3. Tài khoản admin: tải lại trang, kiểm tra thấy các yêu cầu và thay đổi được trạng thái.
4. Tài khoản người dùng: tải lại trang, kiểm tra trạng thái và phản hồi admin đã cập nhật.

> Lưu ý: Không đặt `service_role` key trong mã website, GitHub hoặc phần mềm người dùng.
