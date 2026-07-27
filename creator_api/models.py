from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, HttpUrl


class ApiModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True)


class UserIdentity(ApiModel):
    id: str
    email: str


class InspectRequest(ApiModel):
    url: HttpUrl


class Inspection(ApiModel):
    platform: str
    visibility: Literal["public", "unlisted", "private", "restricted", "unknown"]
    ownership: Literal["verified", "declared", "unverified"] = "unverified"
    license: str = "Chưa xác định"
    risk: Literal["low", "medium", "high", "blocked"]
    can_process: bool = Field(alias="canProcess")
    source_title: str = Field(default="", alias="sourceTitle")
    source_creator: str = Field(default="", alias="sourceCreator")
    duration_seconds: float | None = Field(default=None, alias="durationSeconds")
    source_id: str = Field(default="", alias="sourceId")
    reasons: list[str]


class JobSettings(ApiModel):
    provider: Literal["openai", "gemini"]
    voice: Literal["female-north", "male-north", "female-south", "male-south"]
    whisper_model: Literal["base", "small", "medium"] = Field(default="small", alias="whisperModel")
    background_volume: int = Field(default=10, ge=0, le=30, alias="backgroundVolume")
    subtitle_size: Literal["small", "medium", "large"] = Field(default="medium", alias="subtitleSize")
    rights_confirmed: bool = Field(default=False, alias="rightsConfirmed")


class CreateUrlJobRequest(ApiModel):
    url: HttpUrl
    settings: JobSettings
    inspection: dict | None = None


JobStatus = Literal[
    "queued",
    "downloading",
    "transcribing",
    "translating",
    "dubbing",
    "rendering",
    "completed",
    "failed",
    "cancelled",
]


class JobResponse(ApiModel):
    id: str
    status: JobStatus
    progress: int = Field(ge=0, le=100)
    message: str
    output_url: str | None = Field(default=None, alias="outputUrl")


class UsageSummary(ApiModel):
    plan: str
    monthly_limit_minutes: float = Field(alias="monthlyLimitMinutes", ge=0)
    used_minutes: float = Field(alias="usedMinutes", ge=0)
    remaining_minutes: float = Field(alias="remainingMinutes", ge=0)
    max_video_minutes: float = Field(alias="maxVideoMinutes", gt=0)
    resets_at: datetime = Field(alias="resetsAt")
