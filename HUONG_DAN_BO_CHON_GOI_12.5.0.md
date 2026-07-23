# SỸ LAND 12.5.0 — Bộ chọn gói theo nhu cầu

## Tính năng mới

- Người dùng nhập số người, số hồ sơ/tháng và ngân sách.
- Tự đề xuất gói Go, Plus, Pro hoặc Văn phòng theo số người và số thửa/lượt.
- Hiển thị giá và quyền lợi ngay tại khu vực chọn gói.
- Chu kỳ tháng: không giảm giá.
- Chu kỳ 06 tháng: giảm 10%.
- Chu kỳ 12 tháng: giảm 20%.
- Tổng tiền và số tiền tiết kiệm được hiển thị trước khi tạo QR.
- Supabase tự tính lại giá để chống sửa số tiền từ trình duyệt.

## Bảng giá đang cấu hình

| Gói | Giá tháng | 06 tháng | 12 tháng |
|---|---:|---:|---:|
| Go | 99.000đ | 534.600đ | 950.400đ |
| Plus | 199.000đ | 1.074.600đ | 1.910.400đ |
| Pro | 399.000đ | 2.154.600đ | 3.830.400đ |
| Văn phòng | 298.000đ/tài khoản | Giảm 10% | Giảm 20% |

## Quyền sử dụng đồng bộ

| Gói | Quyền công cụ | Giới hạn mỗi lượt |
|---|---:|---:|
| Dùng thử | 25% | 50 thửa/tệp |
| Go | 40% | 80 thửa/tệp |
| Plus | 70% | 140 thửa/tệp |
| Pro | 100% | 200 thửa/tệp |
| Văn phòng 2–4 tài khoản | Như Plus | 140 thửa/tệp |
| Văn phòng từ 5 tài khoản | 100% | Không giới hạn |
| Quản trị viên | 100% | Không giới hạn |

## Cập nhật

1. Tải toàn bộ nội dung gói lên thư mục gốc GitHub `SY-LAND`.
2. Chạy lại toàn bộ `SUPABASE_SCHEMA.sql`, sau đó `SUPABASE_PAYMENTS.sql` trong Supabase SQL Editor để đồng bộ số tài khoản, hạn mức và cấp mã.
3. Chờ GitHub Actions có dấu xanh.
4. Mở website và nhấn `Ctrl + F5`.
5. Kiểm tra các trường số người, hồ sơ và ngân sách; thử cả chu kỳ 1/6/12 tháng.

> Giá có thể điều chỉnh sau khi có dữ liệu dùng thử thực tế. Không nên giảm giá sâu trước khi xác định chi phí hỗ trợ, lưu trữ, OCR và chăm sóc khách hàng.
