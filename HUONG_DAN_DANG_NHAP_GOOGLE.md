# Cấu hình đăng nhập Google và email một lần cho SỸ LAND

Website và phần mềm Windows dùng chung một dự án Supabase Auth. Người dùng tạo
tài khoản bằng Google, email/mật khẩu hoặc liên kết email đều được nhận diện bằng
`auth.users.id`; dữ liệu nghiệp vụ được tách bằng Row Level Security (RLS).

## 1. Tạo thông tin OAuth trên Google Cloud

1. Mở Google Cloud Console và chọn hoặc tạo dự án SỸ LAND.
2. Vào **Google Auth Platform** → **Clients** → tạo **OAuth client ID** loại
   **Web application**.
3. Trong **Authorized redirect URIs**, thêm:

   `https://PROJECT_REF.supabase.co/auth/v1/callback`

4. Thay `PROJECT_REF` bằng mã dự án Supabase đang dùng.
5. Sao chép **Client ID** và **Client secret**.

## 2. Bật Google trong Supabase

1. Mở Supabase Dashboard → **Authentication** → **Providers** → **Google**.
2. Bật Google, nhập Client ID và Client secret rồi lưu.
3. Vào **Authentication** → **URL Configuration**.
4. Đặt **Site URL**:

   `https://minhsybk-bit.github.io/SY-LAND/`

5. Thêm các **Redirect URLs**:

   - `https://minhsybk-bit.github.io/SY-LAND/**`
   - `http://127.0.0.1:8765/callback`

Redirect `127.0.0.1` chỉ dùng để trình duyệt trả phiên đăng nhập về đúng phần
mềm Windows đang chạy trên máy người dùng.

## 3. Bật đăng nhập qua email

Trong **Authentication** → **Providers** → **Email**:

- Bật Email.
- Giữ xác nhận email để hạn chế tài khoản giả.
- Cấu hình SMTP riêng trước khi phát hành thương mại để thư ổn định và mang
  thương hiệu SỸ LAND.

Website hỗ trợ:

- Email + mật khẩu.
- Liên kết đăng nhập một lần gửi qua email.
- Quên mật khẩu.

## 4. Nguyên tắc độc lập tài khoản

- Không dùng `service_role key` trên website hoặc trong bộ cài.
- Client chỉ dùng Project URL và anon public key.
- Mỗi hồ sơ, đơn thanh toán, thông báo và thiết bị gắn với `auth.uid()`.
- Giấy phép gắn với email đã xác thực, mà Supabase không cho hai tài khoản Auth
  dùng trùng một email trong cùng dự án.
- RLS chỉ cho người dùng đọc dữ liệu của chính mình; quản trị viên mới được đọc
  và đối soát toàn hệ thống.
- Khi đăng xuất, website xóa dữ liệu phiên và dữ liệu tạm của tài khoản trước
  khỏi giao diện trước khi tài khoản khác đăng nhập.

## 5. Kiểm tra trước khi phát hành

1. Tạo tài khoản Google A và tài khoản Google B.
2. Tạo một đơn thanh toán bằng A.
3. Đăng xuất hoàn toàn rồi đăng nhập B.
4. Xác nhận B không nhìn thấy đơn, thông báo, giấy phép hoặc thiết bị của A.
5. Đăng nhập admin và xác nhận admin nhìn thấy cả hai để đối soát.
6. Trên Windows, nhấn **Tiếp tục bằng Google**, đăng nhập A và kiểm tra đúng gói
   của A được nhận.

