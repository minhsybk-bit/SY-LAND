from __future__ import annotations

import asyncio
import json
import tempfile
import uuid
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import ValidationError

from .config import Settings, get_settings
from .database import Database, STATUS_MESSAGES
from .media_validation import valid_video_signature
from .models import CreateUrlJobRequest, InspectRequest, JobResponse, JobSettings, UsageSummary, UserIdentity
from .pipeline import probe_duration
from .security import authenticated_user
from .signing import signed_download_url, valid_download_signature
from .sources import inspect_public_source
from .tasks import process_video_task


app = FastAPI(
    title="SỸ LAND Creator API",
    version="0.2.0",
    docs_url="/docs",
    redoc_url=None,
)
settings = get_settings()
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.origins,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["Authorization", "Content-Type"],
)

ALLOWED_UPLOADS = {
    "video/mp4": ".mp4",
    "video/quicktime": ".mov",
    "video/webm": ".webm",
    "video/x-matroska": ".mkv",
}


def require_provider(settings: JobSettings, server: Settings) -> None:
    if settings.provider == "openai" and not server.openai_api_key:
        raise HTTPException(status_code=503, detail="Dịch vụ OpenAI chưa được quản trị viên cấu hình.")
    if settings.provider == "gemini" and not server.gemini_api_key:
        raise HTTPException(status_code=503, detail="Dịch vụ Gemini chưa được quản trị viên cấu hình.")


def enqueue_job(job: dict, database: Database) -> None:
    try:
        result = process_video_task.delay(job["id"])
        database.update_job(job["id"], queue_task_id=result.id)
        job["queue_task_id"] = result.id
    except Exception as error:
        database.update_job(
            job["id"],
            status="failed",
            progress=100,
            error_code="QUEUE_UNAVAILABLE",
            error_message="Hàng đợi xử lý đang tạm gián đoạn.",
        )
        raise HTTPException(status_code=503, detail="Hàng đợi xử lý đang tạm gián đoạn.") from error


def job_response(job: dict, settings: Settings) -> JobResponse:
    output_url = None
    if job["status"] == "completed" and job.get("output_object_key"):
        output_url = signed_download_url(settings, job["id"], job["user_id"])
    message = str(job.get("error_message") or STATUS_MESSAGES.get(job["status"], "Đang xử lý"))
    return JobResponse(
        id=job["id"],
        status=job["status"],
        progress=int(job.get("progress") or 0),
        message=message,
        outputUrl=output_url,
    )


@app.get("/health")
def health() -> dict:
    return {"ok": True, "service": "syland-creator-api", "version": app.version}


@app.get("/ready")
def readiness(current_settings: Settings = Depends(get_settings)):
    checks = {"storage": False, "database": False, "queue": False}
    try:
        with tempfile.NamedTemporaryFile(dir=current_settings.creator_storage_dir, prefix=".ready-"):
            checks["storage"] = True
    except OSError:
        pass
    try:
        Database(current_settings).ping()
        checks["database"] = True
    except Exception:
        pass
    try:
        from redis import Redis

        client = Redis.from_url(current_settings.redis_url, socket_timeout=3)
        checks["queue"] = bool(client.ping())
        client.close()
    except Exception:
        pass
    if not all(checks.values()):
        raise HTTPException(
            status_code=503,
            detail={"ok": False, "service": "syland-creator-api", "checks": checks},
        )
    return {"ok": True, "service": "syland-creator-api", "checks": checks}


@app.post("/v1/video/inspect")
async def inspect_video(
    request: InspectRequest,
    user: UserIdentity = Depends(authenticated_user),
    current_settings: Settings = Depends(get_settings),
):
    del user
    return await inspect_public_source(str(request.url), current_settings)


@app.get("/v1/video/usage", response_model=UsageSummary, response_model_by_alias=True)
def get_creator_usage(
    user: UserIdentity = Depends(authenticated_user),
    current_settings: Settings = Depends(get_settings),
):
    return Database(current_settings).get_usage_summary(user.id)


@app.post("/v1/video/jobs", response_model=JobResponse, response_model_by_alias=True)
async def create_url_job(
    request: CreateUrlJobRequest,
    user: UserIdentity = Depends(authenticated_user),
    current_settings: Settings = Depends(get_settings),
):
    if not request.settings.rights_confirmed:
        raise HTTPException(status_code=422, detail="Bạn phải xác nhận quyền sử dụng video.")
    require_provider(request.settings, current_settings)
    inspection = await inspect_public_source(str(request.url), current_settings)
    if not inspection.can_process:
        raise HTTPException(status_code=422, detail="Nguồn video riêng tư hoặc bị giới hạn không được xử lý.")

    database = Database(current_settings)
    job = database.create_job(
        user_id=user.id,
        title=inspection.source_title or "Video từ liên kết",
        source_type="url",
        source_url=str(request.url),
        source_platform=inspection.platform,
        source_visibility=inspection.visibility,
        rights_evidence=inspection.model_dump(by_alias=True),
        settings=request.settings.model_dump(by_alias=True),
    )
    enqueue_job(job, database)
    return job_response(job, current_settings)


@app.post("/v1/video/jobs/upload", response_model=JobResponse, response_model_by_alias=True)
async def create_upload_job(
    video: UploadFile = File(...),
    settings_json: str = Form(..., alias="settings"),
    user: UserIdentity = Depends(authenticated_user),
    current_settings: Settings = Depends(get_settings),
):
    try:
        job_settings = JobSettings.model_validate(json.loads(settings_json))
    except (json.JSONDecodeError, ValidationError) as error:
        raise HTTPException(status_code=422, detail="Thiết lập xử lý không hợp lệ.") from error
    if not job_settings.rights_confirmed:
        raise HTTPException(status_code=422, detail="Bạn phải xác nhận quyền sử dụng video.")
    require_provider(job_settings, current_settings)
    extension = ALLOWED_UPLOADS.get(str(video.content_type or "").lower())
    if not extension:
        raise HTTPException(status_code=415, detail="Chỉ hỗ trợ MP4, MOV, WebM hoặc MKV.")

    job_id = str(uuid.uuid4())
    job_dir = current_settings.creator_storage_dir / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    destination = job_dir / f"upload{extension}"
    total = 0
    try:
        with destination.open("xb") as handle:
            while chunk := await video.read(1024 * 1024):
                total += len(chunk)
                if total > current_settings.max_upload_bytes:
                    raise HTTPException(status_code=413, detail="Video vượt quá 500 MB.")
                handle.write(chunk)
    except Exception:
        destination.unlink(missing_ok=True)
        job_dir.rmdir()
        raise
    finally:
        await video.close()

    if not valid_video_signature(destination, extension):
        destination.unlink(missing_ok=True)
        job_dir.rmdir()
        raise HTTPException(status_code=415, detail="Nội dung tệp không đúng định dạng video đã khai báo.")
    try:
        duration = await asyncio.to_thread(probe_duration, destination)
    except Exception as error:
        destination.unlink(missing_ok=True)
        job_dir.rmdir()
        raise HTTPException(status_code=422, detail="Video bị lỗi hoặc không đọc được.") from error
    if duration > current_settings.max_video_seconds:
        destination.unlink(missing_ok=True)
        job_dir.rmdir()
        raise HTTPException(
            status_code=422,
            detail=f"Video vượt giới hạn {current_settings.max_video_seconds // 60} phút của hệ thống.",
        )

    database = Database(current_settings)
    try:
        job = database.create_job(
            user_id=user.id,
            title=Path(video.filename or "Video tải lên").name[:240],
            source_type="upload",
            source_url=None,
            source_platform="upload",
            source_visibility="unknown",
            rights_evidence={
                "risk": "medium",
                "ownership": "declared",
                "reasons": ["Người dùng khai báo sở hữu hoặc đã được phép sử dụng tệp tải lên."],
            },
            settings=job_settings.model_dump(by_alias=True),
            input_object_key=str(destination.relative_to(current_settings.creator_storage_dir)),
        )
    except Exception:
        destination.unlink(missing_ok=True)
        job_dir.rmdir()
        raise

    if job["id"] != job_id:
        final_dir = current_settings.creator_storage_dir / job["id"]
        job_dir.rename(final_dir)
        database.update_job(
            job["id"],
            input_object_key=str((final_dir / destination.name).relative_to(current_settings.creator_storage_dir)),
        )
    enqueue_job(job, database)
    return job_response(job, current_settings)


@app.get("/v1/video/jobs/{job_id}", response_model=JobResponse, response_model_by_alias=True)
def get_job(
    job_id: uuid.UUID,
    user: UserIdentity = Depends(authenticated_user),
    current_settings: Settings = Depends(get_settings),
):
    job = Database(current_settings).get_owned_job(str(job_id), user.id)
    if not job:
        raise HTTPException(status_code=404, detail="Không tìm thấy tác vụ.")
    return job_response(job, current_settings)


@app.post("/v1/video/jobs/{job_id}/cancel", response_model=JobResponse, response_model_by_alias=True)
def cancel_job(
    job_id: uuid.UUID,
    user: UserIdentity = Depends(authenticated_user),
    current_settings: Settings = Depends(get_settings),
):
    database = Database(current_settings)
    job_id_text = str(job_id)
    existing = database.get_owned_job(job_id_text, user.id)
    if not existing:
        raise HTTPException(status_code=404, detail="Không tìm thấy tác vụ.")
    if existing["status"] in {"completed", "failed", "cancelled"}:
        return job_response(existing, current_settings)

    cancelled = database.cancel_owned_job(job_id_text, user.id)
    task_id = str(existing.get("queue_task_id") or "")
    if task_id:
        try:
            process_video_task.app.control.revoke(task_id, terminate=False)
        except Exception:
            pass

    if existing["status"] == "queued" and existing.get("input_object_key"):
        input_path = (current_settings.creator_storage_dir / existing["input_object_key"]).resolve()
        root = current_settings.creator_storage_dir.resolve()
        if root in input_path.parents:
            input_path.unlink(missing_ok=True)
            try:
                input_path.parent.rmdir()
            except OSError:
                pass
    return job_response(cancelled or existing, current_settings)


@app.get("/v1/video/jobs/{job_id}/download")
def download_result(
    job_id: uuid.UUID,
    user: str = Query(...),
    expires: int = Query(...),
    signature: str = Query(...),
    current_settings: Settings = Depends(get_settings),
):
    job_id_text = str(job_id)
    if not valid_download_signature(current_settings, job_id_text, user, expires, signature):
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Liên kết tải xuống không hợp lệ hoặc đã hết hạn.")
    job = Database(current_settings).get_owned_job(job_id_text, user)
    if not job or job["status"] != "completed" or not job.get("output_object_key"):
        raise HTTPException(status_code=404, detail="Không tìm thấy video kết quả.")
    output = (current_settings.creator_storage_dir / job["output_object_key"]).resolve()
    root = current_settings.creator_storage_dir.resolve()
    if root not in output.parents or not output.is_file():
        raise HTTPException(status_code=404, detail="Tệp kết quả không còn tồn tại.")
    return FileResponse(
        output,
        media_type="video/mp4",
        filename=f"SYLAND_CREATOR_{job_id_text[:8]}.mp4",
        headers={"Cache-Control": "private, no-store"},
    )
