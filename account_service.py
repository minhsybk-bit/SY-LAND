"""Xác thực tài khoản SỸ LAND dùng chung qua Supabase Auth.

Chỉ dùng URL dự án và anon key (khóa công khai dành cho client). Tuyệt đối không
đưa service_role key vào phần mềm. Website và ứng dụng Windows cùng gọi một Auth
project nên người dùng có thể đăng ký ở một nơi và đăng nhập ở nơi còn lại.
"""

import json
import os
from datetime import date
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


def _rest_request(method, path, payload=None, token=None, timeout=20):
    """Gọi Supabase REST/RPC bằng cùng phiên đăng nhập với website."""
    config = load_config()
    if not token:
        raise AccountError("Phiên đăng nhập không hợp lệ.")
    headers = {
        "apikey": config["anon_key"],
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    try:
        response = requests.request(
            method, f'{config["url"]}/rest/v1{path}', headers=headers, json=payload, timeout=timeout
        )
    except requests.RequestException as exc:
        raise AccountError("Không đồng bộ được gói quyền. Hãy kiểm tra Internet.") from exc
    try:
        data = response.json()
    except ValueError:
        data = {}
    if not response.ok:
        detail = data.get("message") or data.get("details") or data.get("hint") or "Không đồng bộ được gói quyền."
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


def guest_entitlements():
    """Quyền dùng khi chưa đăng ký/đăng nhập, giống chính sách trên website."""
    return {
        "ok": True,
        "role": "guest",
        "plan": "Dùng thử",
        "feature_percent": 25,
        "max_parcels_per_run": 5,
        "max_file_size_mb": 10,
        "max_total_upload_mb": 25,
        "max_uses_per_day": 3,
        "full_access": False,
        "recommended_upgrade": "Go",
    }


def sync_entitlements(access_token, email, device_hash, device_name="Windows PC", app_version=""):
    """Đăng ký/đồng bộ thiết bị và nhận quyền gói hiện hành từ Supabase."""
    result = _rest_request(
        "POST",
        "/rpc/activate_syland_device",
        {
            "p_email": (email or "").strip().lower(),
            "p_device_hash": (device_hash or "").strip(),
            "p_device_name": (device_name or "Windows PC").strip(),
            "p_app_version": (app_version or "").strip(),
        },
        token=access_token,
    )
    if not result.get("ok"):
        raise AccountError(result.get("message") or "Tài khoản chưa có bản quyền hoạt động.")
    return result


def _usage_path():
    root = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "SYLAND"
    root.mkdir(parents=True, exist_ok=True)
    return root / "guest_usage.json"


def consume_guest_use():
    """Đếm lượt khách cục bộ. Gọi ngay trước khi bắt đầu một tác vụ."""
    policy = guest_entitlements()
    path = _usage_path()
    today = date.today().isoformat()
    try:
        usage = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        usage = {}
    used = int(usage.get("count", 0)) if usage.get("date") == today else 0
    limit = policy["max_uses_per_day"]
    if used >= limit:
        raise AccountError("Đã hết 3 lượt trải nghiệm hôm nay. Hãy đăng ký hoặc đăng nhập SỸ LAND.")
    used += 1
    path.write_text(json.dumps({"date": today, "count": used}, ensure_ascii=False), encoding="utf-8")
    return {"used": used, "remaining": max(0, limit - used), **policy}


def configured():
    try:
        load_config()
        return True
    except AccountError:
        return False
