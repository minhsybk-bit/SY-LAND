"""Xác thực tài khoản SỸ LAND dùng chung qua Supabase Auth.

Chỉ dùng URL dự án và anon key (khóa công khai dành cho client). Tuyệt đối không
đưa service_role key vào phần mềm. Website và ứng dụng Windows cùng gọi một Auth
project nên người dùng có thể đăng ký ở một nơi và đăng nhập ở nơi còn lại.
"""

import json
import os
import queue
import threading
import webbrowser
from datetime import date
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import urlencode

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


def sign_out(auth_session=None):
    """Thu hồi phiên Supabase; không xóa cấu hình hoặc mã bản quyền ngoại tuyến."""
    access_token = str((auth_session or {}).get("access_token") or "").strip()
    if not access_token:
        return {"ok": True, "remote": False}
    try:
        _request("POST", "/logout", {}, token=access_token, timeout=10)
        return {"ok": True, "remote": True}
    except AccountError:
        return {"ok": True, "remote": False}


def sign_in_with_google(timeout=180):
    """Đăng nhập Google qua trình duyệt, trả token về 127.0.0.1 của máy hiện tại."""
    config = load_config()
    callback_url = "http://127.0.0.1:8765/callback"
    result_queue = queue.Queue(maxsize=1)

    class OAuthHandler(BaseHTTPRequestHandler):
        def log_message(self, _format, *_args):
            return

        def _send(self, status, content, content_type="text/html; charset=utf-8"):
            body = content.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            if self.path.split("?", 1)[0] != "/callback":
                self._send(404, "Không tìm thấy.")
                return
            self._send(200, """<!doctype html><html lang="vi"><meta charset="utf-8">
<title>Đăng nhập SỸ LAND</title><main><h1>SỸ LAND</h1>
<p id="status">Đang hoàn tất đăng nhập Google…</p></main><script>
const values=Object.fromEntries(new URLSearchParams(location.hash.slice(1)));
fetch('/token',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(values)})
.then(r=>{if(!r.ok)throw new Error();document.getElementById('status').textContent='Đăng nhập thành công. Có thể đóng cửa sổ này.'})
.catch(()=>document.getElementById('status').textContent='Không hoàn tất được đăng nhập.');
</script></html>""")

        def do_POST(self):
            if self.path != "/token":
                self._send(404, "Không tìm thấy.", "text/plain; charset=utf-8")
                return
            try:
                length = min(int(self.headers.get("Content-Length", "0")), 20000)
                result_queue.put_nowait(json.loads(self.rfile.read(length).decode("utf-8")))
                self._send(204, "", "text/plain; charset=utf-8")
            except Exception:
                self._send(400, "Dữ liệu không hợp lệ.", "text/plain; charset=utf-8")

    class LoopbackServer(HTTPServer):
        allow_reuse_address = True

    try:
        server = LoopbackServer(("127.0.0.1", 8765), OAuthHandler)
    except OSError as exc:
        raise AccountError("Không mở được cổng Google 8765. Hãy đóng phiên SỸ LAND khác.") from exc
    threading.Thread(target=server.serve_forever, daemon=True).start()
    webbrowser.open(f'{config["url"]}/auth/v1/authorize?' + urlencode({
        "provider": "google", "redirect_to": callback_url,
    }))
    try:
        values = result_queue.get(timeout=max(30, int(timeout)))
    except queue.Empty as exc:
        raise AccountError("Đăng nhập Google đã hết thời gian chờ.") from exc
    finally:
        server.shutdown()
        server.server_close()
    if values.get("error") or values.get("error_description"):
        raise AccountError(str(values.get("error_description") or values.get("error")))
    access_token = str(values.get("access_token") or "")
    if not access_token:
        raise AccountError("Google chưa trả về phiên hợp lệ.")
    user = _request("GET", "/user", token=access_token)
    return {"access_token": access_token, "refresh_token": str(values.get("refresh_token") or ""), "user": user}


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
