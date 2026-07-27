from __future__ import annotations

import asyncio
import ipaddress
import socket

from fastapi import HTTPException, status
import yt_dlp

from .config import Settings
from .models import Inspection
from .source_policy import SourceUrlError, parse_source_url


def validate_source_url(raw_url: str) -> tuple[str, str]:
    try:
        return parse_source_url(raw_url)
    except SourceUrlError as error:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=str(error),
        ) from error


async def reject_private_resolution(host: str) -> None:
    loop = asyncio.get_running_loop()
    try:
        results = await loop.run_in_executor(None, socket.getaddrinfo, host, 443)
    except socket.gaierror as error:
        raise HTTPException(status_code=422, detail="Không phân giải được địa chỉ nguồn video.") from error
    for result in results:
        address = ipaddress.ip_address(result[4][0])
        if not address.is_global:
            raise HTTPException(status_code=422, detail="Địa chỉ nguồn video không an toàn.")


def _extract_metadata(url: str, settings: Settings) -> dict:
    options = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "noplaylist": True,
        "playlist_items": "1",
        "socket_timeout": settings.yt_dlp_socket_timeout,
        "retries": 1,
        "extract_flat": False,
        "cookiefile": None,
        "username": None,
        "password": None,
        "geo_bypass": False,
    }
    with yt_dlp.YoutubeDL(options) as downloader:
        return downloader.extract_info(url, download=False)


async def inspect_public_source(url: str, settings: Settings) -> Inspection:
    host, platform = validate_source_url(url)
    await reject_private_resolution(host)
    try:
        info = await asyncio.to_thread(_extract_metadata, url, settings)
    except yt_dlp.utils.DownloadError as error:
        text = str(error).lower()
        private_markers = ("private", "login", "sign in", "members-only", "age-restricted", "friends")
        reason = (
            "Video riêng tư hoặc cần đăng nhập; SỸ LAND không vượt quyền truy cập."
            if any(marker in text for marker in private_markers)
            else "Không thể đọc nguồn công khai hoặc nền tảng đang giới hạn truy cập."
        )
        return Inspection(
            platform=platform,
            visibility="private" if "private" in text else "restricted",
            risk="blocked",
            canProcess=False,
            reasons=[reason],
        )

    if info.get("_type") in {"playlist", "multi_video"}:
        return Inspection(
            platform=platform,
            visibility="restricted",
            risk="blocked",
            canProcess=False,
            reasons=["Chỉ hỗ trợ một video, không xử lý danh sách phát."],
        )

    availability = str(info.get("availability") or "public").lower()
    blocked = availability in {"private", "premium_only", "subscriber_only", "needs_auth", "unlisted"}
    age_limit = int(info.get("age_limit") or 0)
    live_status = str(info.get("live_status") or "")
    duration = float(info["duration"]) if info.get("duration") is not None else None
    license_name = str(info.get("license") or "Theo điều khoản nền tảng")
    reasons = ["Đã đọc được metadata công khai từ nền tảng."]
    if not info.get("license"):
        reasons.append("Không tìm thấy giấy phép tái sử dụng trong metadata.")
    reasons.append("Không thể tự động chứng minh người đăng nhập sở hữu video.")

    if blocked or age_limit >= 18 or live_status in {"is_live", "is_upcoming"}:
        reasons.insert(0, "Nguồn riêng tư, hạn chế độ tuổi, hội viên hoặc phát trực tiếp không được xử lý.")
        return Inspection(
            platform=platform,
            visibility="private" if availability == "private" else "restricted",
            risk="blocked",
            canProcess=False,
            sourceTitle=str(info.get("title") or ""),
            sourceCreator=str(info.get("uploader") or info.get("channel") or ""),
            durationSeconds=duration,
            sourceId=str(info.get("id") or ""),
            license=license_name,
            reasons=reasons,
        )

    if duration and duration > settings.max_video_seconds:
        return Inspection(
            platform=platform,
            visibility="public",
            risk="blocked",
            canProcess=False,
            sourceTitle=str(info.get("title") or ""),
            sourceCreator=str(info.get("uploader") or info.get("channel") or ""),
            durationSeconds=duration,
            sourceId=str(info.get("id") or ""),
            license=license_name,
            reasons=[f"Video vượt giới hạn {settings.max_video_seconds // 60} phút của hệ thống."],
        )

    return Inspection(
        platform=platform,
        visibility="public",
        risk="medium",
        canProcess=True,
        sourceTitle=str(info.get("title") or ""),
        sourceCreator=str(info.get("uploader") or info.get("channel") or ""),
        durationSeconds=duration,
        sourceId=str(info.get("id") or ""),
        license=license_name,
        reasons=reasons,
    )
