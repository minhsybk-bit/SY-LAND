from __future__ import annotations

from typing import Any


def validated_translations(
    rows: list[dict[str, Any]],
    expected_ids: set[int],
) -> dict[int, str]:
    translated: dict[int, str] = {}
    for row in rows:
        if not isinstance(row, dict) or "id" not in row or "vi" not in row:
            raise RuntimeError("AI trả về một đoạn dịch thiếu id hoặc nội dung.")
        try:
            row_id = int(row["id"])
        except (TypeError, ValueError) as error:
            raise RuntimeError("AI trả về id bản dịch không hợp lệ.") from error
        text = str(row["vi"]).strip()
        if row_id not in expected_ids:
            raise RuntimeError("AI trả về id không thuộc đoạn lời thoại đã gửi.")
        if row_id in translated:
            raise RuntimeError("AI trả về id bản dịch bị trùng.")
        if not text:
            raise RuntimeError("AI trả về một đoạn dịch trống.")
        if len(text) > 2000:
            raise RuntimeError("AI trả về một đoạn dịch dài bất thường.")
        translated[row_id] = text
    return translated
