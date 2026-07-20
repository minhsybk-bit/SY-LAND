# Cập nhật SỸ LAND 12.4.0 — Thanh toán QR và tự động cấp mã

## Quy trình hoạt động

1. Người dùng đăng nhập và chọn gói Cá nhân hoặc Văn phòng.
2. Website tạo mã đơn, nội dung chuyển khoản và QR đúng số tiền.
3. Người dùng chuyển khoản rồi chọn **Tôi đã chuyển khoản**.
4. Admin kiểm tra giao dịch thực tế trong tài khoản ngân hàng.
5. Admin chọn **Xác nhận tiền vào**.
6. Supabase tự tạo mã bản quyền, đúng gói, thời hạn và số thiết bị.
7. Mã xuất hiện trong lịch sử đơn của người dùng và bảng bản quyền.

## Bộ chọn gói kinh tế

- Người dùng nhập số người, số hồ sơ/tháng và ngân sách dự kiến.
- Hệ thống đề xuất gói Cá nhân, Văn phòng hoặc yêu cầu báo giá riêng.
- Thanh toán tháng: linh hoạt, không giảm giá.
- Thanh toán 06 tháng: giảm 10% tổng giá.
- Thanh toán 12 tháng: giảm 20% tổng giá.
- Giá và mức giảm được Supabase tính lại; không tin số tiền sửa từ trình duyệt.

## Cấu hình ngân hàng bắt buộc

Thông tin đã cấu hình:

- Ngân hàng: BIDV.
- Mã BIN VietQR: `970418`.
- Số tài khoản: `3950549732`.
- Chủ tài khoản: `NGUYEN MINH SY`.

Mở `app/payment-config.ts` và điền:

- `bankBin`: mã BIN VietQR của ngân hàng.
- `bankName`: tên ngân hàng hiển thị.
- `accountNumber`: số tài khoản nhận tiền.
- `accountName`: tên chủ tài khoản viết in hoa, không dấu.

Khi chưa điền đủ, nút tạo QR tự khóa để tránh chuyển nhầm tiền.

## Cài đặt một lần

1. Tải toàn bộ nội dung gói lên thư mục gốc kho GitHub `SY-LAND`, giữ nguyên đường dẫn.
2. Mở Supabase → **SQL Editor** → **New query**.
3. Sao chép toàn bộ `supabase/SUPABASE_PAYMENTS.sql` vào SQL Editor và chọn **Run**.
4. Chờ GitHub Actions triển khai có dấu xanh rồi mở website và nhấn `Ctrl + F5`.

## Kiểm thử trước khi nhận tiền thật

1. Đăng nhập tài khoản người dùng, tạo một đơn Cá nhân.
2. Không cần chuyển tiền khi thử; chọn **Tôi đã chuyển khoản**.
3. Đăng nhập admin, kiểm tra đúng mã đơn và số tiền rồi chọn **Xác nhận tiền vào**.
4. Đăng nhập lại người dùng, kiểm tra mã bản quyền xuất hiện.
5. Sau thử nghiệm, khóa mã thử trong khu vực quản trị bản quyền.

## Kiểm soát an toàn

- Trình duyệt không có quyền tự xác nhận đã nhận tiền.
- Giá, gói, thời hạn và số thiết bị được kiểm tra lại ở Supabase.
- Người dùng không thể sửa số tiền hoặc tự điền mã bản quyền.
- Đơn đã cấp mã không thể đổi ngược trạng thái.
- Không đưa khóa `service_role` vào website hoặc phần mềm người dùng.

> Phiên bản này là đối soát có admin xác nhận. Muốn xác nhận hoàn toàn tự động theo biến động số dư cần tích hợp API/webhook chính thức của ngân hàng hoặc cổng thanh toán, không dùng tài khoản Internet Banking cá nhân trong mã website.
