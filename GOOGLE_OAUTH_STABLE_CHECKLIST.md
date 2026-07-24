# Cấu hình Google OAuth ổn định cho SỸ LAND

Mã website dùng phiên OAuth do Supabase quản lý. Cấu hình dưới đây là phần
máy chủ bắt buộc phải khớp chính xác; không đưa `Client Secret` lên GitHub.

## 1. Google Cloud

Trong OAuth Client loại **Web application**:

- Authorized JavaScript origin:
  `https://minhsybk-bit.github.io`
- Authorized redirect URI:
  `https://oremxodzeikeydxdwllf.supabase.co/auth/v1/callback`

Không dùng đường dẫn `/SY-LAND/` trong JavaScript origin. Không dùng URL
GitHub Pages làm redirect URI của Google; Google trả về Supabase trước.

## 2. Supabase

Vào **Authentication → Providers → Google**:

1. Bật Google provider.
2. Dán đúng toàn bộ Google `Client ID`.
3. Tạo một `Client Secret` mới trong Google Cloud và sao chép ngay.
4. Dán secret mới vào Supabase rồi lưu.
5. Sau khi đăng nhập thử thành công, vô hiệu hóa secret cũ.

Trong **Authentication → URL Configuration**:

- Site URL: `https://minhsybk-bit.github.io/SY-LAND/`
- Redirect URLs:
  - `https://minhsybk-bit.github.io/SY-LAND/`
  - `https://minhsybk-bit.github.io/SY-LAND`

## 3. Dữ liệu tài khoản và gói

Chạy toàn bộ tệp `SUPABASE_STABLE_ACCOUNT_SYNC.sql` một lần trong SQL Editor.
Tệp này:

- tạo profile còn thiếu cho tài khoản cũ, mới và Google;
- gắn bản quyền lịch sử vào đúng `auth.users.id`;
- cung cấp một hàm quyền dùng chung cho website và phần mềm Windows;
- ngăn bộ nhớ đệm của tài khoản trước bị dùng cho tài khoản vừa đăng nhập.

## 4. Kiểm thử bắt buộc

Thực hiện bằng cửa sổ ẩn danh:

1. Đăng nhập Google bằng tài khoản người dùng thường.
2. Kiểm tra Trung tâm tài khoản hiển thị đúng gói.
3. Đăng xuất hoàn toàn.
4. Đăng nhập tài khoản quản trị `minhsybk@gmail.com`.
5. Kiểm tra hiển thị **Quản trị viên**, **100%**, **Không giới hạn**.
6. Mở phần mềm Windows, đăng nhập cùng hai tài khoản theo thứ tự trên và bấm
   **Đồng bộ ngay**; kết quả phải giống website và không lấy chéo gói.

Nếu Google báo `Unable to exchange external code`, nguyên nhân máy chủ thường
là Client Secret Google trong Supabase không còn hợp lệ hoặc redirect URI
không khớp. Mã nguồn website không thể tự thay secret này.
