#!/usr/bin/env python3
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    notice_path = root / "OSS_LICENSES.md"

    if shutil.which("pip-licenses") is None:
        print(
            "pip-licenses が見つかりません。"
            "この環境に pip-licenses をインストールしてから再実行してください。"
        )
        return 1

    with tempfile.NamedTemporaryFile("w+", encoding="utf-8", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with tmp_path.open("w", encoding="utf-8") as f:
            result = subprocess.run(
                ["pip-licenses", "--format=markdown"],
                cwd=root,
                stdout=f,
                stderr=subprocess.PIPE,
                text=True,
            )
        if result.returncode != 0:
            print(result.stderr.strip())
            return result.returncode

        new_text = tmp_path.read_text(encoding="utf-8")
        old_text = notice_path.read_text(encoding="utf-8") if notice_path.exists() else ""

        if new_text != old_text:
            notice_path.write_text(new_text, encoding="utf-8")
            print("OSS_LICENSES.md を更新しました。")
        else:
            print("OSS_LICENSES.md は最新です。")
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
