from __future__ import annotations

from fastapi import Depends, Header, HTTPException, status
import httpx

from .config import Settings, get_settings
from .models import UserIdentity


async def authenticated_user(
    authorization: str | None = Header(default=None),
    settings: Settings = Depends(get_settings),
) -> UserIdentity:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Hãy đăng nhập tài khoản SỸ LAND.")

    token = authorization[7:].strip()
    if not token:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Phiên đăng nhập không hợp lệ.")

    try:
        async with httpx.AsyncClient(timeout=10) as client:
            response = await client.get(
                f"{settings.supabase_url.rstrip('/')}/auth/v1/user",
                headers={
                    "apikey": settings.supabase_anon_key,
                    "Authorization": f"Bearer {token}",
                },
            )
    except httpx.HTTPError as error:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Chưa thể xác thực tài khoản. Hãy thử lại sau.",
        ) from error

    if response.status_code != 200:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Phiên đăng nhập đã hết hạn.")

    payload = response.json()
    user_id = str(payload.get("id") or "")
    email = str(payload.get("email") or "").strip().lower()
    if not user_id or not email:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Tài khoản thiếu thông tin xác thực.")
    return UserIdentity(id=user_id, email=email)
