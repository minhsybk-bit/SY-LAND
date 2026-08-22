#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from creator_api.preflight import check_env_file


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Kiểm tra cấu hình SỸ LAND Creator mà không hiển thị giá trị bí mật."
    )
    parser.add_argument("env_file", nargs="?", default=".env.creator")
    args = parser.parse_args()

    path = Path(args.env_file)
    errors = check_env_file(path)
    if errors:
        print(f"Cấu hình Creator chưa sẵn sàng ({len(errors)} lỗi):")
        for error in errors:
            print(f"- {error}")
        return 1
    print("Cấu hình Creator hợp lệ; không phát hiện giá trị mẫu hoặc biến bắt buộc bị thiếu.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
