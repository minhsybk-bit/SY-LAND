from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import urlparse


PLACEHOLDER_MARKERS = (
    "YOUR_",
    "REPLACE_",
    "TEN-MIEN-CUA-BAN",
    "EXAMPLE.COM",
)
REQUIRED_KEYS = (
    "ENVIRONMENT",
    "PUBLIC_API_URL",
    "ALLOWED_ORIGINS",
    "SUPABASE_URL",
    "SUPABASE_ANON_KEY",
    "SUPABASE_SERVICE_ROLE_KEY",
    "REDIS_URL",
    "CREATOR_STORAGE_DIR",
    "DOWNLOAD_SIGNING_SECRET",
)


def parse_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            raise ValueError(f"Dòng {number} không đúng dạng TEN=GIA_TRI.")
        key, value = line.split("=", 1)
        key = key.strip()
        if not re.fullmatch(r"[A-Z][A-Z0-9_]*", key):
            raise ValueError(f"Tên biến ở dòng {number} không hợp lệ.")
        values[key] = value.strip().strip("\"'")
    return values


def _is_placeholder(value: str) -> bool:
    upper = value.upper()
    return not value or any(marker in upper for marker in PLACEHOLDER_MARKERS)


def _is_https_url(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme == "https" and bool(parsed.netloc) and not parsed.username


def validate_environment(values: dict[str, str]) -> list[str]:
    errors: list[str] = []
    for key in REQUIRED_KEYS:
        if not values.get(key):
            errors.append(f"Thiếu biến bắt buộc {key}.")

    environment = values.get("ENVIRONMENT", "").lower()
    if environment not in {"staging", "production"}:
        errors.append("ENVIRONMENT phải là staging hoặc production trên máy chủ.")

    for key in ("PUBLIC_API_URL", "SUPABASE_URL"):
        value = values.get(key, "")
        if value and not _is_https_url(value):
            errors.append(f"{key} phải là địa chỉ HTTPS hợp lệ.")

    supabase_url = values.get("SUPABASE_URL", "")
    if supabase_url and not supabase_url.rstrip("/").endswith(".supabase.co"):
        errors.append("SUPABASE_URL phải dùng tên miền dự án *.supabase.co.")

    origins = [item.strip() for item in values.get("ALLOWED_ORIGINS", "").split(",") if item.strip()]
    if not origins:
        errors.append("ALLOWED_ORIGINS phải có ít nhất một website HTTPS.")
    for origin in origins:
        if origin == "*" or not _is_https_url(origin):
            errors.append("ALLOWED_ORIGINS không được dùng * và mọi origin phải là HTTPS.")
            break

    for key in (
        "PUBLIC_API_URL",
        "SUPABASE_URL",
        "SUPABASE_ANON_KEY",
        "SUPABASE_SERVICE_ROLE_KEY",
        "DOWNLOAD_SIGNING_SECRET",
    ):
        value = values.get(key, "")
        if value and _is_placeholder(value):
            errors.append(f"{key} vẫn đang dùng giá trị mẫu.")

    signing_secret = values.get("DOWNLOAD_SIGNING_SECRET", "")
    if signing_secret and len(signing_secret) < 32:
        errors.append("DOWNLOAD_SIGNING_SECRET phải có ít nhất 32 ký tự.")

    if _is_placeholder(values.get("OPENAI_API_KEY", "")) and _is_placeholder(
        values.get("GEMINI_API_KEY", "")
    ):
        errors.append("Phải cấu hình ít nhất một trong OPENAI_API_KEY hoặc GEMINI_API_KEY.")

    for key, minimum, maximum in (
        ("SIGNED_URL_TTL_SECONDS", 60, 86400),
        ("MAX_UPLOAD_BYTES", 1024, 5 * 1024 * 1024 * 1024),
        ("MAX_VIDEO_SECONDS", 10, 7200),
        ("OUTPUT_RETENTION_HOURS", 1, 720),
    ):
        raw = values.get(key)
        if not raw:
            continue
        try:
            number = int(raw)
        except ValueError:
            errors.append(f"{key} phải là số nguyên.")
            continue
        if not minimum <= number <= maximum:
            errors.append(f"{key} phải nằm trong khoảng {minimum}–{maximum}.")

    return errors


def check_env_file(path: Path) -> list[str]:
    if not path.is_file():
        return [f"Không tìm thấy tệp cấu hình {path}."]
    try:
        return validate_environment(parse_env_file(path))
    except (OSError, UnicodeError, ValueError) as error:
        return [str(error)]
