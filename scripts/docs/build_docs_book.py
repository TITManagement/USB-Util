#!/usr/bin/env python3
"""Build a single printable book from docs/ based on README.md TOC order.

This script performs:
1) chapter extraction in README order,
2) validation (missing/duplicate chapters, broken links, basic markdown checks),
3) merged markdown generation,
4) optional PDF/DOCX build via pandoc,
5) JSON and log output for CI usage.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import re
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote


LINK_PATTERN = re.compile(r"!?\[[^\]]*\]\(([^)]+)\)")
HEADING_PATTERN = re.compile(r"^(#{1,6})\s+.+")
FENCE_PATTERN = re.compile(r"^\s*```")
BOOK_TOC_SECTION_HEADING = "## 文書一冊化目次"
PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TOC_FILE = PROJECT_ROOT / "README.md"


@dataclass
class BuildReport:
    generated_at: str
    docs_root: str
    toc_file: str
    chapter_count: int = 0
    chapters: list[str] = field(default_factory=list)
    duplicate_chapters: list[str] = field(default_factory=list)
    missing_chapters: list[str] = field(default_factory=list)
    orphan_markdown_files: list[str] = field(default_factory=list)
    broken_links: list[dict[str, str]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    outputs: dict[str, str] = field(default_factory=dict)
    success: bool = False


def configure_logger(log_file: Path) -> logging.Logger:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("docs-book-builder")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a single printable docs book")
    parser.add_argument("--docs-root", default="docs", help="Docs root directory")
    parser.add_argument(
        "--toc-file",
        default="",
        help="Markdown file whose dedicated TOC section defines chapter source order",
    )
    parser.add_argument("--out-dir", default="docs/_book", help="Output directory")
    parser.add_argument(
        "--book-title",
        default="",
        help="Title used in the merged markdown metadata",
    )
    parser.add_argument(
        "--format",
        choices=["none", "pdf", "docx", "both"],
        default="pdf",
        help="Output format; none means validation+merge only",
    )
    parser.add_argument(
        "--strict-warnings",
        action="store_true",
        help="Treat warnings as build errors (recommended for CI)",
    )
    parser.add_argument(
        "--keep-going",
        action="store_true",
        help="Try pandoc build even if warnings exist (errors still stop)",
    )
    parser.add_argument(
        "--collection-profile",
        choices=["all", "spec-design"],
        default="all",
        help="Chapter selection profile: all (current behavior) or spec-design (requirements/architecture only)",
    )
    return parser.parse_args()


def _is_spec_design_doc(path: Path, docs_root: Path) -> bool:
    try:
        rel = path.resolve().relative_to(docs_root.resolve()).as_posix()
    except ValueError:
        return False
    return rel.startswith("requirements/") or rel.startswith("architecture/")


def filter_chapters(chapters: list[Path], *, docs_root: Path, profile: str) -> list[Path]:
    if profile == "all":
        return chapters
    if profile == "spec-design":
        return [chapter for chapter in chapters if _is_spec_design_doc(chapter, docs_root)]
    return chapters


def _read_toc_lines(toc_file: Path) -> list[str]:
    lines = toc_file.read_text(encoding="utf-8").splitlines()
    start_index: int | None = None

    for index, line in enumerate(lines):
        if line.strip() == BOOK_TOC_SECTION_HEADING:
            start_index = index + 1
            break

    if start_index is None:
        return lines

    scoped_lines: list[str] = []
    for line in lines[start_index:]:
        if line.startswith("## "):
            break
        scoped_lines.append(line)
    return scoped_lines


def _display_path(path: Path, docs_root: Path) -> str:
    path = path.resolve()
    docs_root = docs_root.resolve()
    try:
        return path.relative_to(docs_root).as_posix()
    except ValueError:
        pass

    project_root = docs_root.parent
    try:
        return path.relative_to(project_root).as_posix()
    except ValueError:
        return path.as_posix()


def extract_links_in_order(
    toc_file: Path,
    docs_root: Path,
    seen_toc_files: set[Path] | None = None,
) -> tuple[list[Path], list[str], list[str]]:
    toc_file = toc_file.resolve()
    if seen_toc_files is None:
        seen_toc_files = set()
    if toc_file in seen_toc_files:
        return [], [], []
    seen_toc_files.add(toc_file)

    lines = _read_toc_lines(toc_file)
    chapters: list[Path] = []
    missing: list[str] = []
    duplicates: list[str] = []
    seen: set[Path] = set()
    root_docs_readme = (docs_root / "README.md").resolve()

    for line in lines:
        for match in LINK_PATTERN.finditer(line):
            raw_target = match.group(1).strip()
            target = raw_target.split()[0]
            target = target.split("#", maxsplit=1)[0]
            if not target:
                continue
            if re.match(r"^(https?|mailto|tel):", target):
                continue

            candidate = (toc_file.parent / target).resolve()
            if candidate.suffix.lower() != ".md":
                continue

            try:
                candidate.relative_to(docs_root.resolve())
            except ValueError:
                # Ignore links outside docs/ for book assembly.
                continue

            if toc_file.resolve() != root_docs_readme and candidate == root_docs_readme:
                rel = candidate.relative_to(docs_root.resolve())
                if not candidate.exists():
                    missing.append(rel.as_posix())
                    continue

                if candidate in seen:
                    duplicates.append(rel.as_posix())
                else:
                    seen.add(candidate)
                    chapters.append(candidate)

                nested_chapters, nested_duplicates, nested_missing = extract_links_in_order(
                    toc_file=candidate,
                    docs_root=docs_root,
                    seen_toc_files=seen_toc_files,
                )
                chapters.extend(nested_chapters)
                duplicates.extend(nested_duplicates)
                missing.extend(nested_missing)
                continue

            rel = candidate.relative_to(docs_root.resolve())
            if not candidate.exists():
                missing.append(rel.as_posix())
                continue

            if candidate in seen:
                duplicates.append(rel.as_posix())
                continue

            seen.add(candidate)
            chapters.append(candidate)

    return chapters, duplicates, missing


def collect_markdown_files(docs_root: Path, ignore_dirs: Iterable[str]) -> list[Path]:
    ignored = set(ignore_dirs)
    all_files: list[Path] = []
    for path in docs_root.rglob("*.md"):
        if any(part in ignored for part in path.parts):
            continue
        all_files.append(path.resolve())
    return sorted(all_files)


def scan_markdown_warnings(path: Path, docs_root: Path) -> list[str]:
    warnings: list[str] = []
    lines = path.read_text(encoding="utf-8").splitlines()

    display = _display_path(path, docs_root)
    non_empty_lines = [line for line in lines if line.strip()]
    if len(non_empty_lines) < 5:
        warnings.append(f"short-file: {display} has fewer than 5 non-empty lines")

    last_level = 0
    for index, line in enumerate(lines, start=1):
        heading = HEADING_PATTERN.match(line)
        if heading:
            level = len(heading.group(1))
            if last_level and level - last_level > 1:
                warnings.append(
                    f"heading-jump: {display}:{index} jumps H{last_level} -> H{level}"
                )
            last_level = level

    fence_count = sum(1 for line in lines if FENCE_PATTERN.match(line))
    if fence_count % 2 != 0:
        warnings.append(f"unclosed-code-fence: {display}")

    return warnings


def scan_broken_links(path: Path, docs_root: Path) -> list[dict[str, str]]:
    broken: list[dict[str, str]] = []
    lines = path.read_text(encoding="utf-8").splitlines()
    display = _display_path(path, docs_root)
    for index, line in enumerate(lines, start=1):
        for match in LINK_PATTERN.finditer(line):
            token = match.group(1).strip().split()[0]
            if not token:
                continue
            if token.startswith("#"):
                continue
            if re.match(r"^(https?|mailto|tel):", token):
                continue
            if token.startswith("<") and token.endswith(">"):
                # Placeholder/path-template references are treated as external docs hints.
                continue
            if token.startswith("/abs/path"):
                # Historical record files may contain platform-specific absolute placeholders.
                continue

            is_image = line[max(0, match.start() - 1) : match.start()] == "!"
            target_file = unquote(token.split("#", maxsplit=1)[0])
            resolved = (path.parent / target_file).resolve()
            if not resolved.exists():
                # Fallback for project docs that use docs-root-relative links.
                normalized_target = target_file.lstrip("./")
                resolved = (docs_root / normalized_target).resolve()

            if not resolved.exists():
                broken.append(
                    {
                        "file": display,
                        "line": str(index),
                        "type": "image" if is_image else "link",
                        "target": token,
                    }
                )
    return broken


def chapter_anchor_id(path: Path, docs_root: Path) -> str:
    project_root = docs_root.resolve().parent
    try:
        rel = path.resolve().relative_to(project_root).as_posix()
    except ValueError:
        rel = _display_path(path, docs_root)
    slug = re.sub(r"[^a-z0-9]+", "-", rel.lower()).strip("-")
    return f"chapter-{slug or 'root'}"


def rewrite_links_for_combined_book(
    markdown_text: str,
    *,
    current_file: Path,
    docs_root: Path,
    chapter_anchors: dict[Path, str],
) -> str:
    book_link_pattern = re.compile(r"(!?)\[([^\]]*)\]\(([^)]+)\)")

    def _replace(match: re.Match[str]) -> str:
        whole = match.group(0)
        is_image = bool(match.group(1))
        label = match.group(2)
        token = match.group(3).strip()

        # Keep image links and external references unchanged.
        if is_image:
            return whole
        if not token or token.startswith("#"):
            return whole
        if re.match(r"^(https?|mailto|tel):", token):
            return whole

        target_file = unquote(token.split("#", maxsplit=1)[0])
        if not target_file:
            return whole

        resolved = (current_file.parent / target_file).resolve()
        if not resolved.exists():
            normalized_target = target_file.lstrip("./")
            resolved = (docs_root / normalized_target).resolve()

        anchor = chapter_anchors.get(resolved)
        if not anchor:
            if Path(target_file).suffix.lower() == ".md":
                fallback_label = label or target_file
                return f"{fallback_label}（未収録: {target_file}）"
            return whole

        return f"[{label}](#{anchor})"

    return book_link_pattern.sub(_replace, markdown_text)


def write_combined_markdown(
    chapters: list[Path],
    docs_root: Path,
    output_md: Path,
    title: str,
    chapter_anchors: dict[Path, str],
) -> None:
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    output_md.parent.mkdir(parents=True, exist_ok=True)

    with output_md.open("w", encoding="utf-8") as f:
        f.write("---\n")
        f.write(f'title: "{title}"\n')
        f.write(f'date: "{now}"\n')
        f.write("toc: true\n")
        f.write("numbersections: true\n")
        f.write("---\n\n")
        f.write("# Book Build Metadata\n\n")
        f.write(f"Generated at: {now}\\n\n")

        for chapter in chapters:
            rel = _display_path(chapter, docs_root)
            anchor = chapter_anchors.get(chapter.resolve(), chapter_anchor_id(chapter, docs_root))
            f.write("\n\n---\n\n")
            f.write(f"<!-- source: {rel} -->\n\n")
            f.write(f"<a id=\"{anchor}\"></a>\n\n")
            chapter_text = chapter.read_text(encoding="utf-8").rstrip()
            chapter_text = rewrite_links_for_combined_book(
                chapter_text,
                current_file=chapter,
                docs_root=docs_root,
                chapter_anchors=chapter_anchors,
            )
            f.write(chapter_text)
            f.write("\n")


def run_pandoc(input_md: Path, output_file: Path, logger: logging.Logger) -> None:
    command = [
        "pandoc",
        str(input_md),
        "--from",
        "gfm",
        "--standalone",
        "--toc",
        "--number-sections",
        "-o",
        str(output_file),
    ]
    logger.info("Run pandoc: %s", " ".join(command))
    subprocess.run(command, check=True)


def main() -> int:
    args = parse_args()
    docs_root = Path(args.docs_root).resolve()
    toc_file = Path(args.toc_file).resolve() if args.toc_file else DEFAULT_TOC_FILE.resolve()
    out_dir = Path(args.out_dir).resolve()
    log_file = out_dir / "build_docs_book.log"

    logger = configure_logger(log_file)
    logger.info("Start docs book build")
    logger.info("docs_root=%s toc_file=%s out_dir=%s", docs_root, toc_file, out_dir)

    if not docs_root.exists() or not docs_root.is_dir():
        logger.error("docs_root does not exist or is not a directory: %s", docs_root)
        return 2
    if not toc_file.exists():
        logger.error("toc_file does not exist: %s", toc_file)
        return 2

    repo_name = PROJECT_ROOT.name
    repo_slug = re.sub(r"[^A-Za-z0-9._-]+", "_", repo_name)
    book_title = args.book_title or f"{repo_name} Documentation Book"

    report = BuildReport(
        generated_at=dt.datetime.now().isoformat(),
        docs_root=str(docs_root),
        toc_file=str(toc_file),
    )

    chapters, duplicates, missing = extract_links_in_order(toc_file=toc_file, docs_root=docs_root)

    chapters = filter_chapters(chapters, docs_root=docs_root, profile=args.collection_profile)

    # Include top-level README as preface chapter when the TOC source is README.md.
    if (
        args.collection_profile == "all"
        and toc_file.suffix.lower() == ".md"
        and toc_file.name.lower() == "readme.md"
    ):
        chapters = [toc_file, *[chapter for chapter in chapters if chapter.resolve() != toc_file.resolve()]]

    # Fallback: if no chapter links were found in TOC, include all markdown files under docs/.
    if not chapters:
        chapters = collect_markdown_files(docs_root=docs_root, ignore_dirs={"_book", ".git"})

    report.chapters = [_display_path(chapter, docs_root) for chapter in chapters]
    report.chapter_count = len(chapters)
    chapter_anchors = {chapter.resolve(): chapter_anchor_id(chapter, docs_root) for chapter in chapters}
    report.duplicate_chapters = sorted(set(duplicates))
    report.missing_chapters = sorted(set(missing))

    all_md = collect_markdown_files(docs_root=docs_root, ignore_dirs={"_book", ".git"})
    chapter_set = {path.resolve() for path in chapters}
    orphans = [path.relative_to(docs_root).as_posix() for path in all_md if path.resolve() not in chapter_set]
    report.orphan_markdown_files = orphans

    for chapter in chapters:
        report.warnings.extend(scan_markdown_warnings(chapter, docs_root))
        report.broken_links.extend(scan_broken_links(chapter, docs_root))

    combined_md = out_dir / f"{repo_slug}_docs_book.md"
    report.outputs["combined_markdown"] = str(combined_md)

    has_errors = bool(report.missing_chapters or report.broken_links)
    has_warnings = bool(report.warnings)

    if report.chapter_count == 0:
        logger.error("No chapters resolved from TOC file")
        has_errors = True

    if report.missing_chapters:
        logger.error("Missing chapters: %s", report.missing_chapters)
    if report.duplicate_chapters:
        logger.info("Duplicate chapters in TOC (deduplicated by first appearance): %s", report.duplicate_chapters)
    if report.broken_links:
        logger.error("Broken links/images: %d", len(report.broken_links))
    if report.warnings:
        logger.warning("Warnings: %d", len(report.warnings))

    out_dir.mkdir(parents=True, exist_ok=True)
    write_combined_markdown(
        chapters=chapters,
        docs_root=docs_root,
        output_md=combined_md,
        title=book_title,
        chapter_anchors=chapter_anchors,
    )

    if args.format != "none":
        if shutil.which("pandoc") is None:
            logger.error("pandoc command not found")
            has_errors = True
        elif has_errors:
            logger.error("Skip pandoc due to errors")
        elif has_warnings and args.strict_warnings and not args.keep_going:
            logger.error("Skip pandoc due to strict warning policy")
        else:
            try:
                if args.format in {"pdf", "both"}:
                    pdf_path = out_dir / f"{repo_slug}_docs_book.pdf"
                    run_pandoc(combined_md, pdf_path, logger)
                    report.outputs["pdf"] = str(pdf_path)
                if args.format in {"docx", "both"}:
                    docx_path = out_dir / f"{repo_slug}_docs_book.docx"
                    run_pandoc(combined_md, docx_path, logger)
                    report.outputs["docx"] = str(docx_path)
            except subprocess.CalledProcessError as exc:
                logger.error("pandoc failed: %s", exc)
                has_errors = True

    report.success = not has_errors and not (has_warnings and args.strict_warnings)

    report_path = out_dir / f"{repo_slug}_docs_report.json"
    report_path.write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2), encoding="utf-8")
    report.outputs["report_json"] = str(report_path)
    logger.info("Wrote report: %s", report_path)

    # Re-write report so outputs include report_json field as well.
    report_path.write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2), encoding="utf-8")

    if report.success:
        logger.info("Build succeeded")
        return 0

    logger.error("Build failed")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
