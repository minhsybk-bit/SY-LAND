from __future__ import annotations

from pathlib import Path


def valid_video_signature(path: Path, extension: str) -> bool:
    with path.open("rb") as handle:
        header = handle.read(16)
    if extension in {".mp4", ".mov"}:
        return len(header) >= 8 and header[4:8] == b"ftyp"
    if extension in {".webm", ".mkv"}:
        return header.startswith(b"\x1aE\xdf\xa3")
    return False
