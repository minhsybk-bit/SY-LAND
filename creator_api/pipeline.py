from __future__ import annotations

import asyncio
import json
import math
import re
import shutil
import subprocess
import textwrap
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable

from .config import Settings
from .sources import validate_source_url
from .translation_validation import validated_translations


Progress = Callable[[str, int], None]

VOICE_PRESETS = {
    # Edge TTS hiện có hai giọng vi-VN chính. Bắc/Nam là biến thể tốc độ và cao độ,
    # không được quảng cáo như bản sao giọng vùng miền thuần.
    "female-north": {"voice": "vi-VN-HoaiMyNeural", "rate": "+2%", "pitch": "+2Hz"},
    "male-north": {"voice": "vi-VN-NamMinhNeural", "rate": "+0%", "pitch": "+1Hz"},
    "female-south": {"voice": "vi-VN-HoaiMyNeural", "rate": "-3%", "pitch": "-2Hz"},
    "male-south": {"voice": "vi-VN-NamMinhNeural", "rate": "-4%", "pitch": "-2Hz"},
}
SUBTITLE_SIZES = {"small": 46, "medium": 58, "large": 70}


def run_command(args: list[str], cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        args,
        cwd=str(cwd) if cwd else None,
        capture_output=True,
        text=True,
        timeout=3600,
    )
    if result.returncode:
        message = (result.stderr or result.stdout or "Lệnh xử lý media thất bại.").strip()
        raise RuntimeError(message[-3000:])
    return result


def require_media_tools() -> None:
    if not shutil.which("ffmpeg") or not shutil.which("ffprobe"):
        raise RuntimeError("Máy chủ chưa cài FFmpeg/FFprobe.")


def probe_duration(path: Path) -> float:
    result = run_command(
        [
            "ffprobe",
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(path),
        ]
    )
    duration = float(result.stdout.strip())
    if not math.isfinite(duration) or duration <= 0:
        raise RuntimeError("Không đọc được thời lượng video.")
    return duration


def download_video(url: str, workdir: Path, settings: Settings) -> Path:
    import yt_dlp

    validate_source_url(url)
    template = str(workdir / "source.%(ext)s")
    options: dict[str, Any] = {
        "outtmpl": template,
        "format": "bv*+ba/b",
        "merge_output_format": "mp4",
        "noplaylist": True,
        "quiet": True,
        "no_warnings": True,
        "retries": 2,
        "socket_timeout": settings.yt_dlp_socket_timeout,
        "cookiefile": None,
        "username": None,
        "password": None,
        "geo_bypass": False,
    }
    with yt_dlp.YoutubeDL(options) as downloader:
        downloader.download([url])
    candidates = [
        item
        for item in workdir.glob("source.*")
        if item.is_file() and item.suffix.lower() not in {".part", ".ytdl", ".json"}
    ]
    if not candidates:
        raise RuntimeError("Không tải được video công khai.")
    return max(candidates, key=lambda item: item.stat().st_size)


def extract_audio(video: Path, output: Path) -> None:
    run_command(
        [
            "ffmpeg",
            "-y",
            "-i",
            str(video),
            "-vn",
            "-ac",
            "1",
            "-ar",
            "16000",
            "-c:a",
            "pcm_s16le",
            str(output),
        ]
    )


@lru_cache(maxsize=3)
def _whisper_model(name: str):
    import torch
    import whisper

    device = "cuda" if torch.cuda.is_available() else "cpu"
    return whisper.load_model(name, device=device), device


def transcribe(audio: Path, model_name: str) -> tuple[str, list[dict[str, Any]]]:
    model, device = _whisper_model(model_name)
    result = model.transcribe(
        str(audio),
        task="transcribe",
        fp16=device == "cuda",
        verbose=False,
        condition_on_previous_text=True,
    )
    segments = []
    for index, item in enumerate(result.get("segments") or []):
        source = str(item.get("text") or "").strip()
        start = max(0.0, float(item.get("start") or 0))
        end = max(start + 0.25, float(item.get("end") or start + 0.25))
        if source:
            segments.append({"id": index, "start": round(start, 3), "end": round(end, 3), "source": source})
    if not segments:
        raise RuntimeError("Whisper không nhận diện được lời thoại trong video.")
    return str(result.get("language") or "unknown"), segments


def _prompt(batch: list[dict[str, Any]], language: str) -> str:
    payload = [{"id": item["id"], "text": item["source"]} for item in batch]
    return f"""
Ngôn ngữ nguồn: {language}.
Dịch lời thoại sang tiếng Việt tự nhiên, đời thường và phù hợp video ngắn.
Giữ đúng ý, tên riêng, số liệu và sắc thái. Không bịa thêm, không hashtag,
không lời chào, không giải thích. Ưu tiên độ dài gần câu gốc để lồng tiếng.
Chỉ trả về mảng JSON hợp lệ dạng [{{"id":0,"vi":"..."}}], giữ đủ và đúng id.

Dữ liệu:
{json.dumps(payload, ensure_ascii=False)}
""".strip()


def _json_array(raw: str) -> list[dict[str, Any]]:
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip(), flags=re.I)
    start, end = cleaned.find("["), cleaned.rfind("]")
    if start < 0 or end <= start:
        raise RuntimeError("AI không trả về định dạng bản dịch hợp lệ.")
    value = json.loads(cleaned[start : end + 1])
    if not isinstance(value, list):
        raise RuntimeError("Bản dịch không phải danh sách JSON.")
    return value


def _translate_openai(prompt: str, settings: Settings) -> list[dict[str, Any]]:
    if not settings.openai_api_key:
        raise RuntimeError("Máy chủ chưa cấu hình OPENAI_API_KEY.")
    from openai import OpenAI

    client = OpenAI(api_key=settings.openai_api_key)
    response = client.responses.create(
        model=settings.openai_translation_model,
        instructions="Bạn là biên dịch viên video ngắn. Tuân thủ chính xác định dạng JSON được yêu cầu.",
        input=prompt,
        store=False,
    )
    return _json_array(response.output_text)


def _translate_gemini(prompt: str, settings: Settings) -> list[dict[str, Any]]:
    if not settings.gemini_api_key:
        raise RuntimeError("Máy chủ chưa cấu hình GEMINI_API_KEY.")
    from google import genai

    client = genai.Client(api_key=settings.gemini_api_key)
    response = client.models.generate_content(model=settings.gemini_translation_model, contents=prompt)
    if not response.text:
        raise RuntimeError("Gemini không trả về bản dịch.")
    return _json_array(response.text)


def translate(
    segments: list[dict[str, Any]],
    language: str,
    provider: str,
    progress: Progress,
) -> list[dict[str, Any]]:
    settings = Settings()
    translated: dict[int, str] = {}
    batches = [segments[index : index + 25] for index in range(0, len(segments), 25)]
    for index, batch in enumerate(batches):
        prompt = _prompt(batch, language)
        rows = _translate_openai(prompt, settings) if provider == "openai" else _translate_gemini(prompt, settings)
        expected_ids = {int(item["id"]) for item in batch}
        translated.update(validated_translations(rows, expected_ids))
        progress("translating", 45 + round((index + 1) / len(batches) * 15))
    missing = [item["id"] for item in segments if item["id"] not in translated]
    if missing:
        raise RuntimeError(f"AI dịch thiếu {len(missing)} đoạn.")
    return [{**item, "vi": translated[item["id"]]} for item in segments]


async def _save_tts(text: str, path: Path, preset: dict[str, str]) -> None:
    import edge_tts

    await edge_tts.Communicate(
        text=text,
        voice=preset["voice"],
        rate=preset["rate"],
        pitch=preset["pitch"],
    ).save(str(path))


def _atempo(speed: float) -> str:
    speed = max(0.25, min(speed, 8.0))
    factors = []
    while speed > 2:
        factors.append(2.0)
        speed /= 2
    while speed < 0.5:
        factors.append(0.5)
        speed /= 0.5
    factors.append(speed)
    return ",".join(f"atempo={factor:.5f}" for factor in factors)


def build_dub(
    segments: list[dict[str, Any]],
    voice: str,
    duration: float,
    workdir: Path,
    progress: Progress,
) -> Path:
    from pydub import AudioSegment

    preset = VOICE_PRESETS[voice]
    master = AudioSegment.silent(duration=math.ceil(duration * 1000) + 500, frame_rate=44100).set_channels(2)
    for index, segment in enumerate(segments):
        raw = workdir / f"tts_{index:04d}.mp3"
        fitted = workdir / f"tts_{index:04d}.wav"
        asyncio.run(_save_tts(segment["vi"], raw, preset))
        slot = max(0.35, float(segment["end"]) - float(segment["start"]))
        speed = max(0.85, probe_duration(raw) / slot)
        run_command(["ffmpeg", "-y", "-i", str(raw), "-filter:a", _atempo(speed), "-ar", "44100", "-ac", "2", str(fitted)])
        clip = AudioSegment.from_file(fitted)[: int((slot + 0.15) * 1000)].apply_gain(-1)
        master = master.overlay(clip, position=int(float(segment["start"]) * 1000))
        progress("dubbing", 62 + round((index + 1) / len(segments) * 18))
    output = workdir / "dub.wav"
    master[: int(duration * 1000)].export(output, format="wav", parameters=["-ac", "2", "-ar", "44100"])
    return output


def _ass_time(seconds: float) -> str:
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    return f"{hours}:{minutes:02d}:{seconds % 60:05.2f}"


def _ass_text(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", text).strip().replace("\\", r"\\").replace("{", r"\{").replace("}", r"\}")
    return r"\N".join(textwrap.wrap(cleaned, width=34, break_long_words=False, break_on_hyphens=False)[:3])


def write_subtitles(segments: list[dict[str, Any]], path: Path, size: str) -> None:
    header = f"""[Script Info]
ScriptType: v4.00+
PlayResX: 1080
PlayResY: 1920
ScaledBorderAndShadow: yes
WrapStyle: 2

[V4+ Styles]
Format: Name,Fontname,Fontsize,PrimaryColour,SecondaryColour,OutlineColour,BackColour,Bold,Italic,Underline,StrikeOut,ScaleX,ScaleY,Spacing,Angle,BorderStyle,Outline,Shadow,Alignment,MarginL,MarginR,MarginV,Encoding
Style: Default,DejaVu Sans,{SUBTITLE_SIZES[size]},&H00FFFFFF,&H000000FF,&H00000000,&H96000000,-1,0,0,0,100,100,0,0,1,4,1,2,60,60,110,1

[Events]
Format: Layer,Start,End,Style,Name,MarginL,MarginR,MarginV,Effect,Text
"""
    events = [
        f"Dialogue: 0,{_ass_time(float(item['start']))},{_ass_time(float(item['end']))},Default,,0,0,0,,{_ass_text(item['vi'])}\n"
        for item in segments
    ]
    path.write_text(header + "".join(events), encoding="utf-8-sig")


def render(video: Path, dub: Path, subtitles: Path, background: int, output: Path) -> None:
    source = subtitles.parent / f"input{video.suffix.lower()}"
    if source != video:
        shutil.copy2(video, source)
    background_track = subtitles.parent / "background.wav"
    try:
        run_command(["ffmpeg", "-y", "-i", source.name, "-vn", "-ar", "44100", "-ac", "2", background_track.name], subtitles.parent)
    except RuntimeError:
        duration = probe_duration(source)
        run_command(["ffmpeg", "-y", "-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo", "-t", str(duration), background_track.name], subtitles.parent)
    volume = max(0, min(background, 30)) / 100
    run_command(
        [
            "ffmpeg", "-y", "-i", source.name, "-i", background_track.name, "-i", dub.name,
            "-filter_complex",
            f"[1:a]volume={volume:.3f}[bg];[2:a]volume=1[dub];[bg][dub]amix=inputs=2:duration=first:dropout_transition=0[a];[0:v]ass={subtitles.name}[v]",
            "-map", "[v]", "-map", "[a]", "-c:v", "libx264", "-preset", "veryfast", "-crf", "20",
            "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", "-shortest", output.name,
        ],
        subtitles.parent,
    )


def process_media(
    source: Path,
    output: Path,
    job_settings: dict[str, Any],
    settings: Settings,
    progress: Progress,
) -> float:
    require_media_tools()
    duration = probe_duration(source)
    if duration > settings.max_video_seconds:
        raise RuntimeError(f"Video vượt giới hạn {settings.max_video_seconds // 60} phút.")
    audio = source.parent / "source.wav"
    progress("transcribing", 25)
    extract_audio(source, audio)
    language, segments = transcribe(audio, job_settings["whisper_model"])
    progress("translating", 42)
    segments = translate(segments, language, job_settings["provider"], progress)
    (source.parent / "transcript.json").write_text(
        json.dumps({"source_language": language, "segments": segments}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    progress("dubbing", 62)
    dub = build_dub(segments, job_settings["voice"], duration, source.parent, progress)
    progress("rendering", 85)
    subtitles = source.parent / "subtitles.ass"
    write_subtitles(segments, subtitles, job_settings["subtitle_size"])
    render(source, dub, subtitles, job_settings["background_volume"], output)
    return duration
