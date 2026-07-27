from __future__ import annotations


def public_error(error: Exception) -> str:
    text = str(error).lower()
    known = {
        "creator_quota_exceeded": "Bạn đã dùng hết số phút Creator của tháng này.",
        "creator_video_too_long_for_plan": "Video vượt thời lượng tối đa của gói hiện tại.",
        "creator_plan_not_configured": "Gói tài khoản chưa được cấu hình cho Creator.",
    }
    for marker, message in known.items():
        if marker in text:
            return message
    if "private" in text or "login" in text or "sign in" in text:
        return "Nguồn riêng tư hoặc cần đăng nhập không được xử lý."
    if "whisper không nhận diện" in text:
        return "Không nhận diện được lời thoại rõ ràng trong video."
    if "không tải được video công khai" in text:
        return "Không tải được video công khai từ nền tảng."
    if "ffmpeg/ffprobe" in text:
        return "Máy chủ xử lý video đang thiếu thành phần cần thiết."
    if "ai trả về" in text or "bản dịch" in text:
        return "AI trả về bản dịch chưa hợp lệ. Vui lòng thử lại."
    if "api_key" in text or "chưa cấu hình" in text:
        return "Nhà cung cấp AI đang tạm thời chưa sẵn sàng."
    return "Không thể xử lý video này. Vui lòng thử lại hoặc sử dụng video khác."
