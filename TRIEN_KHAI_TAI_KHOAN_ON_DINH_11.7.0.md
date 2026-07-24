# Triển khai tài khoản ổn định SỸ LAND 11.7.0

## 1. Cập nhật cơ sở dữ liệu một lần

Trong Supabase Dashboard, mở **SQL Editor**, dán toàn bộ nội dung
`SUPABASE_STABLE_ACCOUNT_SYNC.sql`, rồi chọn **Run**.

Kết quả cuối phải trả về các cột `auth_users`, `profiles`,
`licenses_linked`, `legacy_licenses_unlinked`, `admins`.

- `auth_users` phải bằng `profiles`.
- `legacy_licenses_unlinked` nên bằng `0`. Nếu lớn hơn `0`, các mã đó đang
  dùng email không trùng với tài khoản Auth và cần quản trị viên sửa email.
- Tài khoản `minhsybk@gmail.com` phải có vai trò `admin`.

Từ bản này, web và Windows chỉ dùng RPC
`get_my_syland_entitlements()` làm nguồn xác định gói. Không sửa hạn mức riêng
trong phần mềm.

## 2. Bật Google trong Supabase

Vào **Authentication → Providers → Google**:

1. Bật Google.
2. Nhập đúng Client ID và Client Secret còn hiệu lực từ Google Cloud.
3. Chọn Save.

Trong Google Cloud OAuth Client, Authorized redirect URI phải là:

`https://oremxodzeikeydxdwllf.supabase.co/auth/v1/callback`

Không nhập GitHub Pages hoặc localhost vào ô callback của Google Cloud.

## 3. Cho phép URL trả về

Trong **Supabase → Authentication → URL Configuration**:

- Site URL: `https://minhsybk-bit.github.io/SY-LAND/`
- Redirect URLs:
  - `https://minhsybk-bit.github.io/SY-LAND/`
  - `https://minhsybk-bit.github.io/SY-LAND/**`
  - `http://127.0.0.1:8765/callback`

URL cuối dành cho nút **Tiếp tục bằng Google** trên phần mềm Windows. Nó chỉ
trả phiên đăng nhập về chính máy đang chạy phần mềm.

## 4. Kiểm thử bắt buộc

1. Đăng nhập web bằng email/mật khẩu, ghi lại gói và hạn mức.
2. Đăng xuất, đăng nhập bằng Google với đúng email đó; gói phải giống bước 1.
3. Đăng nhập Windows bằng email/mật khẩu; chọn **Đồng bộ ngay**; gói phải giống web.
4. Đăng xuất Windows, chọn **Tiếp tục bằng Google**; gói phải giống bước 3.
5. Đăng nhập tài khoản khác; không được nhìn thấy đơn, mã hoặc gói của tài khoản trước.
6. Đăng nhập `minhsybk@gmail.com`; cả web và Windows phải hiện Quản trị viên,
   100% công cụ và không giới hạn.

Nếu Google báo `Unable to exchange external code`, lỗi nằm ở Client Secret
Google đang lưu trong Supabase. Tạo secret mới trong Google Cloud, sao chép ngay,
thay secret cũ trong Supabase rồi Save.
