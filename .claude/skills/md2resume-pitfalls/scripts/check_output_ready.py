#!/usr/bin/env python3
"""Fail if a DOCX target is locked or a deliverable directory has non-DOCX artifacts."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ALLOWED = {".docx"}
FORBIDDEN = {".pdf", ".png", ".html", ".htm"}


def lock_siblings(path: Path) -> list[Path]:
    parent = path.parent
    name = path.name
    candidates = [
        parent / f"~${name}",
        parent / f".~{name}",
        parent / f"~{name}",
    ]
    return [p for p in candidates if p.exists()]


def extra_artifacts(directory: Path) -> list[Path]:
    if not directory.is_dir():
        return []
    found: list[Path] = []
    for child in directory.iterdir():
        if child.is_file() and child.suffix.lower() in FORBIDDEN:
            found.append(child)
    return sorted(found)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("target", type=Path, help="Intended DOCX output path")
    parser.add_argument(
        "--clean-extras",
        action="store_true",
        help="Delete forbidden artifacts next to the target",
    )
    args = parser.parse_args()
    target = args.target.expanduser().resolve()
    if target.suffix.lower() not in ALLOWED:
        print(f"error: target must be .docx, got {target.suffix}", file=sys.stderr)
        return 2

    locks = lock_siblings(target)
    if locks:
        print("error: editor lock files present:", file=sys.stderr)
        for lock in locks:
            print(f"  {lock}", file=sys.stderr)
        print("Close WPS/Word or choose a new output name.", file=sys.stderr)
        return 1

    extras = extra_artifacts(target.parent)
    if extras:
        if args.clean_extras:
            for extra in extras:
                extra.unlink()
                print(f"removed {extra}")
        else:
            print("error: non-DOCX artifacts in output directory:", file=sys.stderr)
            for extra in extras:
                print(f"  {extra}", file=sys.stderr)
            return 1

    print(f"ready: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
