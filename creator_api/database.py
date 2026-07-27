from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

import httpx

from .config import Settings


STATUS_MESSAGES = {
    "queued": "Đang chờ máy chủ",
    "downloading": "Đang tiếp nhận video",
    "transcribing": "Whisper đang nhận diện lời thoại",
    "translating": "AI đang dịch sang tiếng Việt",
    "dubbing": "Đang tạo giọng đọc tiếng Việt",
    "rendering": "Đang dựng video và phụ đề",
    "completed": "Đã hoàn thành",
    "failed": "Xử lý thất bại",
    "cancelled": "Đã hủy",
}


class Database:
    def __init__(self, settings: Settings):
        self.base = f"{settings.supabase_url.rstrip('/')}/rest/v1"
        self.headers = {
            "apikey": settings.supabase_service_role_key,
            "Authorization": f"Bearer {settings.supabase_service_role_key}",
            "Content-Type": "application/json",
        }

    def _request(self, method: str, path: str, **kwargs: Any) -> httpx.Response:
        request_headers = {**self.headers, **kwargs.pop("headers", {})}
        response = httpx.request(
            method,
            f"{self.base}/{path.lstrip('/')}",
            headers=request_headers,
            timeout=20,
            **kwargs,
        )
        response.raise_for_status()
        return response

    def insert(self, table: str, payload: dict) -> dict:
        response = self._request(
            "POST",
            table,
            json=payload,
            headers={**self.headers, "Prefer": "return=representation"},
        )
        rows = response.json()
        if not rows:
            raise RuntimeError(f"Không tạo được dữ liệu {table}.")
        return rows[0]

    def create_job(
        self,
        *,
        user_id: str,
        title: str,
        source_type: str,
        source_url: str | None,
        source_platform: str,
        source_visibility: str,
        rights_evidence: dict,
        settings: dict,
        input_object_key: str | None = None,
    ) -> dict:
        ownership_status = str(rights_evidence.get("ownership") or "unverified")
        if ownership_status not in {"verified", "declared", "unverified"}:
            ownership_status = "unverified"
        project = self.insert(
            "creator_projects",
            {
                "user_id": user_id,
                "title": title,
                "source_type": source_type,
                "source_url": source_url,
                "source_platform": source_platform,
                "source_visibility": source_visibility,
                "ownership_status": ownership_status,
                "rights_status": rights_evidence.get("risk", "medium"),
                "rights_evidence": rights_evidence,
                "rights_confirmed_at": datetime.now(timezone.utc).isoformat(),
                "status": "ready",
            },
        )
        return self.insert(
            "creator_jobs",
            {
                "project_id": project["id"],
                "user_id": user_id,
                "provider": settings["provider"],
                "voice": settings["voice"],
                "whisper_model": settings["whisperModel"],
                "background_volume": settings["backgroundVolume"],
                "subtitle_size": settings["subtitleSize"],
                "input_object_key": input_object_key,
            },
        )

    def get_owned_job(self, job_id: str, user_id: str) -> dict | None:
        response = self._request(
            "GET",
            "creator_jobs",
            params={
                "id": f"eq.{job_id}",
                "user_id": f"eq.{user_id}",
                "select": "*",
                "limit": "1",
            },
        )
        rows = response.json()
        return rows[0] if rows else None

    def get_job_with_project(self, job_id: str) -> dict | None:
        response = self._request(
            "GET",
            "creator_jobs",
            params={
                "id": f"eq.{job_id}",
                "select": "*,creator_projects(*)",
                "limit": "1",
            },
        )
        rows = response.json()
        return rows[0] if rows else None

    def ping(self) -> None:
        self._request(
            "GET",
            "creator_jobs",
            params={"select": "id", "limit": "1"},
        )

    def get_usage_summary(self, user_id: str) -> dict:
        response = self._request(
            "POST",
            "rpc/creator_usage_summary",
            json={"p_user_id": user_id},
        )
        payload = response.json()
        if not isinstance(payload, dict):
            raise RuntimeError("Không đọc được hạn mức Creator.")
        return payload

    def update_job(self, job_id: str, **changes: Any) -> None:
        changes["updated_at"] = datetime.now(timezone.utc).isoformat()
        self._request("PATCH", "creator_jobs", params={"id": f"eq.{job_id}"}, json=changes)

    def update_project(self, project_id: str, status: str) -> None:
        self._request(
            "PATCH",
            "creator_projects",
            params={"id": f"eq.{project_id}"},
            json={"status": status, "updated_at": datetime.now(timezone.utc).isoformat()},
        )

    def cancel_owned_job(self, job_id: str, user_id: str) -> dict | None:
        job = self.get_owned_job(job_id, user_id)
        if not job:
            return None
        if job["status"] in {"completed", "failed", "cancelled"}:
            return job
        now = datetime.now(timezone.utc).isoformat()
        self.update_job(
            job_id,
            status="cancelled",
            progress=int(job.get("progress") or 0),
            completed_at=now,
            error_code=None,
            error_message=None,
        )
        self.update_project(job["project_id"], "cancelled")
        job.update(
            {
                "status": "cancelled",
                "completed_at": now,
                "error_code": None,
                "error_message": None,
            }
        )
        return job

    def list_expired_outputs(self, completed_before: str) -> list[dict]:
        response = self._request(
            "GET",
            "creator_jobs",
            params={
                "status": "eq.completed",
                "completed_at": f"lt.{completed_before}",
                "output_object_key": "not.is.null",
                "select": "id,user_id,output_object_key",
                "limit": "500",
            },
        )
        return list(response.json())

    def reserve_minutes(self, job_id: str, minutes: float) -> None:
        self._request(
            "POST",
            "rpc/creator_reserve_minutes",
            json={"p_job_id": job_id, "p_minutes": round(minutes, 3)},
        )

    def consume_minutes(self, job_id: str, minutes: float) -> None:
        self._request(
            "POST",
            "rpc/creator_consume_minutes",
            json={"p_job_id": job_id, "p_minutes": round(minutes, 3)},
        )

    def refund_minutes(self, job_id: str, minutes: float) -> None:
        self._request(
            "POST",
            "rpc/creator_refund_minutes",
            json={"p_job_id": job_id, "p_minutes": round(minutes, 3)},
        )
