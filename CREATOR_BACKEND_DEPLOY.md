# Triển khai backend thương mại SỸ LAND Creator

Backend này không chạy trên GitHub Pages. GitHub Pages chỉ hiển thị giao diện.
API có thể chạy trên CPU nhỏ; worker xử lý video cần GPU NVIDIA, ổ lưu trữ dùng
chung và có thể bật/tắt hoặc tự co giãn theo hàng đợi.

## Kiến trúc

- `api`: xác thực phiên Supabase, kiểm tra nguồn, nhận tệp và trả tiến độ.
- `worker`: tải nguồn công khai, chạy Whisper, dịch, Edge TTS và FFmpeg.
- `redis`: hàng đợi để website không bị treo trong lúc dựng video.
- `scheduler`: mỗi giờ xóa MP4 đã hết thời hạn lưu.
- `Supabase`: tài khoản, dự án, tác vụ và sổ hạn mức theo gói.
- `/data/creator`: vùng lưu video riêng, không phục vụ công khai.

`Dockerfile.creator-api` là image CPU nhẹ cho API luôn hoạt động.
`Dockerfile.creator` là image CUDA cho worker GPU. Compose mẫu chạy cả hai trên
một máy để thử nghiệm; khi vận hành thương mại có thể tách worker sang dịch vụ GPU
theo nhu cầu nhưng vẫn dùng chung Redis và vùng lưu trữ.

Video riêng tư hoặc cần đăng nhập bị chặn. Hệ thống không nhận cookie nền tảng,
không vượt quyền truy cập và không tự khẳng định bản quyền. Kết quả đánh giá nguồn
chỉ là tín hiệu quản trị rủi ro.

## 1. Chuẩn bị Supabase

Mở **Supabase → SQL Editor**, chạy toàn bộ tệp `SUPABASE_CREATOR.sql` sau
`SUPABASE_SCHEMA.sql`. Tệp này tạo:

- bảng dự án, tác vụ và nhật ký phút;
- hạn mức mặc định cho Dùng thử, Go, Plus, Pro và các gói cũ;
- hàm giữ/hoàn hạn mức có khóa giao dịch để tránh dùng vượt khi bấm nhiều lần;
- RLS: người dùng chỉ đọc tác vụ của chính họ, trình duyệt không được tạo job.

Có thể đổi số phút mà không sửa mã:

```sql
update public.creator_plan_limits
set monthly_minutes = 300, max_video_minutes = 20
where plan = 'Plus';
```

## 2. Cấu hình bí mật trên máy chủ

Sao chép `.env.creator.example` thành `.env.creator`, rồi điền:

- `SUPABASE_URL`, `SUPABASE_ANON_KEY`: tại Supabase Project Settings → API.
- `SUPABASE_SERVICE_ROLE_KEY`: chỉ đặt trên máy chủ; tuyệt đối không đưa lên web.
- `OPENAI_API_KEY` hoặc `GEMINI_API_KEY`: khóa của tài khoản doanh nghiệp SỸ LAND;
  khách hàng không phải nhập khóa.
- `DOWNLOAD_SIGNING_SECRET`: chạy `openssl rand -hex 32` để tạo.
- `OUTPUT_RETENTION_HOURS`: số giờ giữ MP4 kết quả, mặc định 24 giờ.
- `PUBLIC_API_URL`: tên miền HTTPS của API.
- `ALLOWED_ORIGINS`: `https://minhsybk-bit.github.io`.

Không gửi các khóa này qua chat và không commit `.env.creator` lên GitHub.

## 3. Khởi động

Để thử nghiệm all-in-one, trên máy chủ đã cài Docker, NVIDIA Driver và NVIDIA
Container Toolkit:

```bash
docker compose -f docker-compose.creator.yml build
docker compose -f docker-compose.creator.yml up -d
docker compose -f docker-compose.creator.yml ps
```

Kiểm tra:

```bash
curl https://creator-api.ten-mien-cua-ban.vn/health
```

Kết quả đúng có `"ok": true`. Đặt Nginx/Caddy/Cloudflare phía trước cổng 8000
để cấp HTTPS. Không mở Redis ra Internet.

Kiểm tra đầy đủ trước khi nhận lưu lượng:

```bash
curl https://creator-api.ten-mien-cua-ban.vn/ready
```

Ba mục `storage`, `database`, `queue` đều phải là `true`.

## 4. Nối website

Trong GitHub repository:

1. Mở **Settings → Secrets and variables → Actions**.
2. Tạo secret `VITE_SYLAND_CREATOR_API_URL`.
3. Giá trị là `https://creator-api.ten-mien-cua-ban.vn`.
4. Chạy lại workflow GitHub Pages.

Website chỉ nhận địa chỉ API. Mọi khóa AI và `service_role` vẫn ở máy chủ.

## 5. Vận hành thương mại

- Worker mặc định `--concurrency=1` để một GPU xử lý một video tại một thời điểm.
- Giữ API CPU hoạt động liên tục; worker GPU có thể scale từ 0 lên 1 theo độ dài
  hàng đợi để không trả tiền GPU khi không có video.
- Tăng số worker khi đã có nhiều GPU; không tăng concurrency tùy ý trên một GPU.
- Sao lưu Supabase và volume `creator_media`.
- Thiết lập tác vụ xóa video đầu vào/kết quả theo chính sách lưu trữ đã công bố.
- Website có nút hủy; worker hoàn số phút đã giữ và xóa tệp tạm khi nhận trạng
  thái hủy.
- Theo dõi chi phí AI, thời gian GPU, hàng đợi Redis và lỗi FFmpeg.
- Trước khi bán chính thức, kiểm thử ít nhất 20 video đại diện cho từng nền tảng
  và từng ngôn ngữ nguồn.

## API chính

- `POST /v1/video/inspect`
- `POST /v1/video/jobs`
- `POST /v1/video/jobs/upload`
- `GET /v1/video/jobs/{id}`
- `POST /v1/video/jobs/{id}/cancel`
- `GET /v1/video/jobs/{id}/download` với chữ ký hết hạn

Swagger dành cho quản trị kỹ thuật ở `/docs`. Mọi endpoint nghiệp vụ đều dùng
JWT tài khoản SỸ LAND hiện có.
