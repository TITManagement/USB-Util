"""pip install report から internal-PyPI 由来依存を検証する。

目的:
- AbnormalSoundDetection が internal-PyPI 配布前提としている依存を、CI で自動検証する。
- 対象パッケージが official PyPI から混入した場合に fail-fast する。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_REQUIRED_INTERNAL = [
    "usb-util",
    "transport-core",
    "opcua-drivers",
    "lads-drivers",
    "autokobopy",
]

DEFAULT_INTERNAL_HINTS = [
    ".ci/internal-PyPI",
    "/internal-PyPI/",
    "/internal-pypi/",
]


def _normalize_name(name: str) -> str:
    return name.strip().lower().replace("_", "-")


def _load_report(path: Path) -> dict:
    if not path.exists():
        raise SystemExit(f"pip report not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--report",
        required=True,
        help="Path to pip --report JSON file.",
    )
    parser.add_argument(
        "--require-internal",
        nargs="*",
        default=DEFAULT_REQUIRED_INTERNAL,
        help="Package names that must be resolved from internal-PyPI.",
    )
    parser.add_argument(
        "--internal-index-url",
        default="",
        help="internal-PyPI index URL (e.g. file:///.../simple or http://host:18080/simple).",
    )
    parser.add_argument(
        "--internal-hints",
        nargs="*",
        default=DEFAULT_INTERNAL_HINTS,
        help="URL substrings that identify internal-PyPI sources.",
    )
    return parser.parse_args()


def _derive_internal_prefixes(index_url: str) -> list[str]:
    if not index_url:
        return []
    parsed = urlparse(index_url)
    if parsed.scheme not in {"file", "http", "https"}:
        return []

    path = parsed.path or ""
    normalized_path = path.rstrip("/")
    if normalized_path.endswith("/simple"):
        normalized_path = normalized_path[: -len("/simple")]
    base_path = normalized_path.rstrip("/") + "/"

    if parsed.scheme == "file":
        root = f"file://{parsed.netloc}{base_path}"
    else:
        root = f"{parsed.scheme}://{parsed.netloc}{base_path}"
    return [root, f"{root}dist/"]


def main() -> int:
    args = _parse_args()
    report = _load_report(Path(args.report))
    install_rows = report.get("install", [])

    by_name: dict[str, str] = {}
    for row in install_rows:
        metadata = row.get("metadata") or {}
        name = metadata.get("name")
        if not name:
            continue
        url = ((row.get("download_info") or {}).get("url")) or ""
        by_name[_normalize_name(str(name))] = str(url)

    required = [_normalize_name(name) for name in args.require_internal]
    hints = [hint.lower() for hint in args.internal_hints]
    prefixes = [prefix.lower() for prefix in _derive_internal_prefixes(args.internal_index_url)]

    missing: list[str] = []
    violations: list[tuple[str, str]] = []
    for name in required:
        url = by_name.get(name)
        if not url:
            missing.append(name)
            continue
        lower_url = url.lower()
        from_hint = any(hint in lower_url for hint in hints)
        from_prefix = any(lower_url.startswith(prefix) for prefix in prefixes)
        if not (from_hint or from_prefix):
            violations.append((name, url))

    if missing:
        print("ERROR: required packages are missing in pip report:")
        for name in missing:
            print(f"  - {name}")
        print("Hint: run pip install with --report in a clean environment.")
        return 1

    if violations:
        print("ERROR: required internal packages were not resolved from internal-PyPI:")
        for name, url in violations:
            print(f"  - {name}: {url}")
        return 1

    print("OK: internal-PyPI source verification passed.")
    for name in required:
        print(f"  - {name}: {by_name.get(name, '')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
