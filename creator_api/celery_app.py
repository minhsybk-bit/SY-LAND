from celery import Celery

from .config import get_settings


settings = get_settings()
celery_app = Celery(
    "syland_creator",
    broker=settings.redis_url,
    backend=settings.redis_url,
    include=["creator_api.tasks"],
)
celery_app.conf.update(
    task_serializer="json",
    result_serializer="json",
    accept_content=["json"],
    task_track_started=True,
    worker_prefetch_multiplier=1,
    task_acks_late=True,
    task_reject_on_worker_lost=True,
    timezone="Asia/Ho_Chi_Minh",
    beat_schedule={
        "cleanup-expired-creator-outputs": {
            "task": "creator.cleanup_expired_outputs",
            "schedule": 3600.0,
        },
    },
)
