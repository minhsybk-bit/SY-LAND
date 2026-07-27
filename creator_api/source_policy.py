from __future__ import annotations

from urllib.parse import urlparse


PLATFORMS: dict[str, tuple[str, ...]] = {
    "YouTube": ("youtube.com", "youtu.be"),
    "TikTok": ("tiktok.com",),
    "Douyin": ("douyin.com",),
    "Bilibili": ("bilibili.com", "b23.tv"),
    "Instagram": ("instagram.com",),
    "Facebook": ("facebook.com", "fb.watch"),
}


class SourceUrlError(ValueError):
    pass


def platform_for_host(host: str) -> str | None:
    normalized = host.lower().rstrip(".")
    for platform, domains in PLATFORMS.items():
        if any(normalized == domain or normalized.endswith(f".{domain}") for domain in domains):
            return platform
    return None


def parse_source_url(raw_url: str) -> tuple[str, str]:
    parsed = urlparse(raw_url)
    try:
        port = parsed.port
    except ValueError as error:
        raise SourceUrlError("Cổng mạng trong liên kết không hợp lệ.") from error
    if (
        parsed.scheme != "https"
        or not parsed.hostname
        or parsed.username
        or parsed.password
        or port not in {None, 443}
    ):
        raise SourceUrlError("Chỉ chấp nhận liên kết HTTPS công khai từ nền tảng được hỗ trợ.")
    platform = platform_for_host(parsed.hostname)
    if not platform:
        raise SourceUrlError("Nền tảng này chưa nằm trong danh sách nguồn được hỗ trợ.")
    return parsed.hostname, platform
