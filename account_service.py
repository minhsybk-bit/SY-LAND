"""Xác thực tài khoản SỸ LAND dùng chung qua Supabase Auth.

Chỉ dùng URL dự án và anon key (khóa công khai dành cho client). Tuyệt đối không
đưa service_role key vào phần mềm. Website và ứng dụng Windows cùng gọi một Auth
project nên người dùng có thể đăng ký ở một nơi và đăng nhập ở nơi còn lại.
"""

import json
import os
from pathlib import Path

import requests


class AccountError(RuntimeError):
    pass


def _config_paths():
    root = Path(__file__).resolve().parents[2]
    return [root / "config" / "supabase_config.json", Path.cwd() / "config" / "supabase_config.json"]


def load_config():
    url = os.environ.get("SYLAND_SUPABASE_URL", "").strip().rstrip("/")
    anon_key = os.environ.get("SYLAND_SUPABASE_ANON_KEY", "").strip()
    if url and anon_key:
        return {"url": url, "anon_key": anon_key}
    for path in _config_paths():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            url = str(data.get("url", "")).strip().rstrip("/")
            anon_key = str(data.get("anon_key", "")).strip()
            if url and anon_key and "YOUR_" not in anon_key:
                return {"url": url, "anon_key": anon_key}
        except (OSError, ValueError, TypeError):
            continue
    raise AccountError("Chưa cấu hình máy chủ tài khoản chung. Hãy cấu hình Supabase trong config/supabase_config.json.")


def _request(method, path, payload=None, token=None, timeout=20):
    config = load_config()
    headers = {"apikey": config["anon_key"], "Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    try:
        response = requests.request(method, f'{config["url"]}/auth/v1{path}', headers=headers, json=payload, timeout=timeout)
    except requests.RequestException as exc:
        raise AccountError("Không kết nối được máy chủ tài khoản. Hãy kiểm tra Internet.") from exc
    try:
        data = response.json()
    except ValueError:
        data = {}
    if not response.ok:
        detail = data.get("msg") or data.get("message") or data.get("error_description") or "Yêu cầu xác thực thất bại."
        raise AccountError(str(detail))
    return data


def sign_in(email, password):
    email = (email or "").strip().lower()
    if not email or not password:
        raise AccountError("Hãy nhập email và mật khẩu.")
    return _request("POST", "/token?grant_type=password", {"email": email, "password": password})


def sign_up(email, password, full_name=""):
    email = (email or "").strip().lower()
    if not email or "@" not in email:
        raise AccountError("Email chưa hợp lệ.")
    if len(password or "") < 8:
        raise AccountError("Mật khẩu cần ít nhất 8 ký tự.")
    return _request("POST", "/signup", {"email": email, "password": password, "data": {"full_name": (full_name or "").strip()}})


def configured():
    try:
        load_config()
        return True
    except AccountError:
        return False
