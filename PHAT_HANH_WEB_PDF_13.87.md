# SỸ LAND Web — bố cục và quy trình PDF build 13.87

## Thay đổi giao diện

- Bộ công cụ được phân nhóm: Tách, Chỉnh sửa, Chuyển đổi, Kiểm tra và Tối ưu.
- Mặc định mở nhóm Chỉnh sửa để tránh hiển thị quá nhiều công cụ cùng lúc.
- Công cụ đang chọn có tiêu đề, trạng thái phát hành và khu vực thiết lập riêng.
- Bố cục thống nhất: chọn tệp bên trái, xem trước và thiết lập bên phải.
- Giao diện đáp ứng cho máy tính, máy tính bảng và điện thoại.
- Thay nhãn “Bản thử nghiệm công khai” bằng “Phiên bản chính thức”.

## Quy trình chung của mọi công cụ PDF

1. Chọn và kiểm tra tệp nguồn.
2. Tạo ảnh xem trước cục bộ.
3. Thiết lập phạm vi/phương án xử lý.
4. Theo dõi tiến độ và xuất tệp kết quả mới.

Mọi công cụ đều hiển thị xem trước nguồn, thanh tiến độ và nút Tạm dừng/Hủy khi tác vụ đang chạy. Các vòng xử lý dài kiểm tra lệnh dừng tại điểm an toàn giữa các trang hoặc giữa các tệp.

## Nguyên tắc an toàn

- Tệp PDF/ảnh được xử lý trong trình duyệt, không tự tải lên máy chủ.
- Không ghi đè tệp gốc.
- Hủy tác vụ trước khi tải kết quả sẽ không thay đổi tệp nguồn.
- Trang trắng hoặc trang lỗi chỉ là cảnh báo; người dùng phải xem trước và xác nhận.

## Tệp cần đưa lên thư mục gốc của GitHub

- `pdf-toolkit.tsx`
- `globals.css`
- `page.tsx`
- `system-check.tsx`

Sau khi commit, chờ GitHub Actions hoàn tất cả `build` và `deploy`, rồi tải lại website bằng `Ctrl + F5`.
