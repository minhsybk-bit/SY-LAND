from __future__ import annotations

import traceback
from datetime import datetime, timedelta, timezone
from pathlib import Path

from .celery_app import celery_app
from .config import get_settings
from .database import Database
from .errors import public_error
from .pipeline import download_video, probe_duration, process_media


class JobCancelled(RuntimeError):
    pass


@celery_app.task(name="creator.process_video", bind=True, max_retries=1)
def process_video_task(self, job_id: str) -> None:
    settings = get_settings()
    database = Database(settings)
    job = database.get_job_with_project(job_id)
    if not job:
        return

    job_dir = (settings.creator_storage_dir / job_id).resolve()
    storage_root = settings.creator_storage_dir.resolve()
    if storage_root not in job_dir.parents:
        raise RuntimeError("Đường dẫn tác vụ không an toàn.")
    job_dir.mkdir(parents=True, exist_ok=True)
    reserved_minutes = 0.0

    def ensure_active() -> None:
        current = database.get_owned_job(job_id, job["user_id"])
        if not current or current["status"] == "cancelled":
            raise JobCancelled("Tác vụ đã được người dùng hủy.")

    def progress(status: str, percent: int) -> None:
        ensure_active()
        database.update_job(job_id, status=status, progress=percent)

    try:
        ensure_active()
        database.update_project(job["project_id"], "processing")
        project = job.get("creator_projects") or {}
        if isinstance(project, list):
            project = project[0] if project else {}
        source_key = job.get("input_object_key")
        if source_key:
            source = (settings.creator_storage_dir / source_key).resolve()
            if storage_root not in source.parents or not source.is_file():
                raise RuntimeError("Không tìm thấy tệp video tải lên.")
        else:
            progress("downloading", 10)
            source = download_video(str(project.get("source_url") or ""), job_dir, settings)

        duration = probe_duration(source)
        reserved_minutes = max(0.001, duration / 60)
        database.reserve_minutes(job_id, reserved_minutes)
        database.update_job(job_id, duration_seconds=duration, started_at=datetime.now(timezone.utc).isoformat())

        output = job_dir / "SYLAND_CREATOR.mp4"
        process_media(source, output, job, settings, progress)
        ensure_active()
        relative_output = str(output.relative_to(settings.creator_storage_dir))
        database.consume_minutes(job_id, reserved_minutes)
        database.update_job(
            job_id,
            status="completed",
            progress=100,
            output_object_key=relative_output,
            completed_at=datetime.now(timezone.utc).isoformat(),
            error_code=None,
            error_message=None,
        )
        database.update_project(job["project_id"], "completed")
    except JobCancelled:
        (job_dir / "SYLAND_CREATOR.mp4").unlink(missing_ok=True)
        if reserved_minutes:
            try:
                database.refund_minutes(job_id, reserved_minutes)
            except Exception:
                pass
        database.update_job(
            job_id,
            status="cancelled",
            completed_at=datetime.now(timezone.utc).isoformat(),
            error_code=None,
            error_message=None,
        )
        database.update_project(job["project_id"], "cancelled")
    except Exception as error:
        (job_dir / "SYLAND_CREATOR.mp4").unlink(missing_ok=True)
        if reserved_minutes:
            try:
                database.refund_minutes(job_id, reserved_minutes)
            except Exception:
                pass
        database.update_job(
            job_id,
            status="failed",
            progress=100,
            error_code=error.__class__.__name__,
            error_message=public_error(error),
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
        database.update_project(job["project_id"], "failed")
        traceback.print_exc()
    finally:
        temporary_files = []
        for pattern in ("tts_*", "source.*", "upload.*", "input.*"):
            temporary_files.extend(job_dir.glob(pattern))
        temporary_files.extend(
            job_dir / name
            for name in ("dub.wav", "background.wav", "transcript.json", "subtitles.ass")
        )
        for path in temporary_files:
            if path.is_file() and path.name != "SYLAND_CREATOR.mp4":
                try:
                    path.unlink()
                except OSError:
                    pass


@celery_app.task(name="creator.cleanup_expired_outputs")
def cleanup_expired_outputs_task() -> dict[str, int]:
    settings = get_settings()
    database = Database(settings)
    cutoff = datetime.now(timezone.utc) - timedelta(hours=settings.output_retention_hours)
    storage_root = settings.creator_storage_dir.resolve()
    deleted = 0
    skipped = 0
    for job in database.list_expired_outputs(cutoff.isoformat()):
        output = (storage_root / str(job.get("output_object_key") or "")).resolve()
        if storage_root not in output.parents or output.name != "SYLAND_CREATOR.mp4":
            skipped += 1
            continue
        try:
            output.unlink(missing_ok=True)
            database.update_job(job["id"], output_object_key=None)
            try:
                output.parent.rmdir()
            except OSError:
                pass
            deleted += 1
        except OSError:
            skipped += 1
    return {"deleted": deleted, "skipped": skipped}
