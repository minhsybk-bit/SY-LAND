# SỸ LAND 12.5.0 — Bộ chọn gói theo nhu cầu

## Tính năng mới

- Người dùng nhập số người, số hồ sơ/tháng và ngân sách.
- Tự đề xuất gói Cá nhân, Văn phòng hoặc báo giá Đơn vị.
- Hiển thị giá và quyền lợi ngay tại khu vực chọn gói.
- Chu kỳ tháng: không giảm giá.
- Chu kỳ 06 tháng: giảm 10%.
- Chu kỳ 12 tháng: giảm 20%.
- Tổng tiền và số tiền tiết kiệm được hiển thị trước khi tạo QR.
- Supabase tự tính lại giá để chống sửa số tiền từ trình duyệt.

## Bảng giá đang cấu hình

| Gói | Giá tháng | 06 tháng | 12 tháng |
|---|---:|---:|---:|
| Cá nhân | 199.000đ | 1.074.600đ | 1.910.400đ |
| Văn phòng | 1.490.000đ | 8.046.000đ | 14.304.000đ |
| Đơn vị | Báo giá riêng | Báo giá riêng | Báo giá riêng |

## Cập nhật

1. Tải toàn bộ nội dung gói lên thư mục gốc GitHub `SY-LAND`.
2. Chạy lại toàn bộ `supabase/SUPABASE_PAYMENTS.sql` trong Supabase SQL Editor để nâng cấp chu kỳ thanh toán.
3. Chờ GitHub Actions có dấu xanh.
4. Mở website và nhấn `Ctrl + F5`.
5. Kiểm tra các trường số người, hồ sơ và ngân sách; thử cả chu kỳ 1/6/12 tháng.

> Giá có thể điều chỉnh sau khi có dữ liệu dùng thử thực tế. Không nên giảm giá sâu trước khi xác định chi phí hỗ trợ, lưu trữ, OCR và chăm sóc khách hàng.
