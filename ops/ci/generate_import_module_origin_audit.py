"""Generate a markdown audit table of package origins from pip --report.

The script focuses on packages declared in pyproject dependencies
(`project.dependencies` and optional dependency groups).
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path
from urllib.parse import urlparse

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:  # pragma: no cover - Python 3.10 fallback
    import tomli as tomllib  # type: ignore[no-redef]


DEFAULT_INTERNAL_HINTS = [
    ".ci/internal-PyPI",
    "/internal-PyPI/",
    "/internal-pypi/",
]


def _normalize_name(name: str) -> str:
    return name.strip().lower().replace("_", "-")


def _extract_name(requirement: str) -> str:
    raw = requirement.split(";", 1)[0].strip()
    m = re.match(r"^([A-Za-z0-9_.-]+)", raw)
    return _normalize_name(m.group(1)) if m else ""


def _derive_internal_prefixes(index_url: str) -> list[str]:
    if not index_url:
        return []
    parsed = urlparse(index_url)
    if parsed.scheme not in {"file", "http", "https"}:
        return []

    path = (parsed.path or "").rstrip("/")
    if path.endswith("/simple"):
        path = path[: -len("/simple")]
    path = path.rstrip("/") + "/"

    if parsed.scheme == "file":
        root = f"file://{parsed.netloc}{path}"
    else:
        root = f"{parsed.scheme}://{parsed.netloc}{path}"
    return [root.lower(), f"{root}dist/".lower()]


def _classify(url: str, internal_hints: list[str], internal_prefixes: list[str]) -> str:
    lower = url.lower()
    if not lower:
        return "missing-url"
    if any(lower.startswith(prefix) for prefix in internal_prefixes):
        return "internal-PyPI"
    if any(hint in lower for hint in internal_hints):
        return "internal-PyPI"
    if lower.startswith("file://"):
        return "local-file/editable"
    if "pypi.org" in lower or "pythonhosted.org" in lower:
        return "official-PyPI"
    return "other"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, help="Path to pip --report JSON file.")
    parser.add_argument("--pyproject", required=True, help="Path to pyproject.toml.")
    parser.add_argument("--output", required=True, help="Output markdown path.")
    parser.add_argument(
        "--internal-index-url",
        default="",
        help="internal index URL (file:///.../simple or http://host:18080/simple).",
    )
    parser.add_argument(
        "--internal-hints",
        nargs="*",
        default=DEFAULT_INTERNAL_HINTS,
        help="URL substrings treated as internal-PyPI hints.",
    )
    parser.add_argument(
        "--include-optional-groups",
        nargs="*",
        default=[],
        help="Optional dependency groups to include (default: none).",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    report_path = Path(args.report)
    pyproject_path = Path(args.pyproject)
    output_path = Path(args.output)

    if not report_path.exists():
        print(f"ERROR: report not found: {report_path}")
        return 1
    if not pyproject_path.exists():
        print(f"ERROR: pyproject not found: {pyproject_path}")
        return 1

    report = json.loads(report_path.read_text(encoding="utf-8"))
    with pyproject_path.open("rb") as f:
        pyproject = tomllib.load(f)

    install_rows = report.get("install", [])
    by_name: dict[str, dict[str, str]] = {}
    for row in install_rows:
        metadata = row.get("metadata") or {}
        name = metadata.get("name")
        if not name:
            continue
        normalized = _normalize_name(str(name))
        by_name[normalized] = {
            "name": str(name),
            "version": str(metadata.get("version") or ""),
            "url": str(((row.get("download_info") or {}).get("url")) or ""),
        }

    project = pyproject.get("project") or {}
    deps: list[str] = list(project.get("dependencies") or [])
    optional = project.get("optional-dependencies") or {}
    include_groups = set(args.include_optional_groups)
    for group_name in sorted(optional):
        if group_name not in include_groups:
            continue
        group_items = optional.get(group_name) or []
        deps.extend(group_items)

    required_specs: dict[str, str] = {}
    for req in deps:
        name = _extract_name(req)
        if not name:
            continue
        required_specs[name] = req

    internal_hints = [hint.lower() for hint in args.internal_hints]
    internal_prefixes = _derive_internal_prefixes(args.internal_index_url)

    lines: list[str] = []
    lines.append("# Import Module Origin Audit (Auto Generated)")
    lines.append("")
    lines.append(f"- generated_at_utc: `{dt.datetime.utcnow().isoformat()}Z`")
    lines.append(f"- report: `{report_path}`")
    lines.append(f"- pyproject: `{pyproject_path}`")
    lines.append("")
    lines.append("| package | requirement | resolved_version | source | download_info.url |")
    lines.append("| --- | --- | --- | --- | --- |")

    for name in sorted(required_specs):
        spec = required_specs[name]
        resolved = by_name.get(name)
        if resolved is None:
            lines.append(f"| {name} | `{spec}` | (missing) | missing | (not found in report) |")
            continue
        url = resolved["url"]
        source = _classify(url=url, internal_hints=internal_hints, internal_prefixes=internal_prefixes)
        lines.append(
            f"| {resolved['name']} | `{spec}` | `{resolved['version']}` | {source} | `{url}` |"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Generated: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
