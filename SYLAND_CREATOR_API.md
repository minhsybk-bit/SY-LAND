# Hợp đồng máy chủ SỸ LAND Creator

Website GitHub Pages chỉ cung cấp giao diện. Whisper, tải video, dịch, lồng tiếng
và FFmpeg phải chạy trên máy chủ riêng có hàng đợi xử lý.

Mã triển khai tham chiếu nằm trong `creator_api/`; cách đóng gói và vận hành nằm
trong `CREATOR_BACKEND_DEPLOY.md`.

Biến cấu hình khi build website:

```text
VITE_SYLAND_CREATOR_API_URL=https://creator-api.ten-mien-cua-ban.vn
```

API phải xác thực Supabase JWT gửi trong header:

```text
Authorization: Bearer <access_token>
```

## 1. Kiểm tra nguồn video

`POST /v1/video/inspect`

```json
{"url":"https://www.youtube.com/watch?v=..."}
```

Kết quả:

```json
{
  "platform": "YouTube",
  "visibility": "public",
  "ownership": "unverified",
  "license": "youtube",
  "risk": "medium",
  "canProcess": true,
  "sourceTitle": "Tên video",
  "sourceCreator": "Tên kênh",
  "reasons": ["Video công khai", "Chưa xác minh quyền sử dụng"]
}
```

Máy chủ phải trả `canProcess=false`, `risk=blocked` khi nguồn riêng tư, yêu cầu
đăng nhập nhưng người dùng chưa cấp OAuth, bị cấm truy cập hoặc không đáp ứng
quy tắc xử lý.

## 2. Tạo tác vụ từ link

`POST /v1/video/jobs`

```json
{
  "url": "https://...",
  "settings": {
    "provider": "openai",
    "voice": "female-north",
    "whisperModel": "small",
    "backgroundVolume": 10,
    "subtitleSize": "medium",
    "rightsConfirmed": true
  },
  "inspection": {}
}
```

## 3. Tạo tác vụ từ file

`POST /v1/video/jobs/upload`

`multipart/form-data` gồm:

- `video`: MP4/MOV/WebM/MKV.
- `settings`: chuỗi JSON có cấu trúc như trên.

## 4. Theo dõi tiến độ

`GET /v1/video/jobs/{job_id}`

```json
{
  "id": "uuid",
  "status": "transcribing",
  "progress": 35,
  "message": "Whisper đang nhận diện lời thoại"
}
```

Khi hoàn thành:

```json
{
  "id": "uuid",
  "status": "completed",
  "progress": 100,
  "message": "Đã hoàn thành",
  "outputUrl": "https://signed-url..."
}
```

`outputUrl` phải là URL tải xuống có chữ ký và thời hạn ngắn. Không công khai
bucket chứa video.

## 5. Hủy tác vụ

`POST /v1/video/jobs/{job_id}/cancel`

Chỉ chủ tài khoản được hủy tác vụ của mình. Worker kiểm tra trạng thái hủy giữa
các công đoạn, hoàn lại số phút đã giữ và xóa tệp tạm. Tác vụ đã hoàn thành hoặc
thất bại không bị thay đổi.

## 6. Giám sát máy chủ

- `GET /health`: kiểm tra tiến trình API còn hoạt động.
- `GET /ready`: kiểm tra đồng thời vùng lưu trữ, Supabase và Redis.

`/ready` trả HTTP 503 nếu một thành phần chưa sẵn sàng để hệ thống giám sát không
chuyển lưu lượng vào máy chủ lỗi.

## Yêu cầu bảo mật

- Không nhận OpenAI/Gemini API Key từ trình duyệt.
- Khóa AI chỉ nằm trong secret của máy chủ.
- Xác thực Supabase JWT cho mọi endpoint.
- Kiểm tra quyền gói và hạn mức phút ở máy chủ.
- Không tin `inspection` hoặc `rightsConfirmed` do trình duyệt gửi; máy chủ phải
  kiểm tra lại.
- Không tải hoặc vượt đăng nhập video riêng tư.
- Quét loại MIME và giới hạn dung lượng trước khi đưa file vào hàng đợi.
- Xóa file tạm và URL tải xuống theo thời hạn.
- Scheduler xóa MP4 kết quả sau `OUTPUT_RETENTION_HOURS`.
- Ghi nhật ký quyền sử dụng, trạng thái và mức tiêu thụ; không ghi API Key.
