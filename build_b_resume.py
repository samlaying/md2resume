"""Headless builder: Markdown -> DOCX.

Usage:
    python3 build_b_resume.py                          # default sample
    python3 build_b_resume.py examples/sample_zh.md
    python3 build_b_resume.py --scale 0.90 --dense
Output files are written next to the markdown, named after its stem.
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from core.parser import parse_markdown
from core.docx_renderer import render_docx

ROOT = Path(__file__).resolve().parent
MD = ROOT / "examples" / "sample_zh.md"
PHOTO = None  # keep samples photo-free; pass a local file only at build time


def main():
    scale = 0.90
    dense = True
    md_path = MD
    photo = PHOTO
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == "--scale" and i + 1 < len(args):
            scale = float(args[i + 1]); i += 2
        elif args[i] == "--dense":
            dense = True; i += 1
        elif args[i] == "--no-dense":
            dense = False; i += 1
        elif args[i] == "--photo" and i + 1 < len(args):
            photo = args[i + 1]; i += 2
        elif args[i] == "--md" and i + 1 < len(args):
            md_path = Path(args[i + 1]); i += 2
        elif not args[i].startswith("--"):
            md_path = Path(args[i]); i += 1
        else:
            i += 1

    out_dir = md_path.parent
    stem = md_path.stem
    md_text = md_path.read_text(encoding="utf-8")
    data = parse_markdown(md_text)

    print("=== parsed ===")
    print("name:", data.name)
    print("contact:", data.contact)
    for sec, entries in data.sections.items():
        print(f"## {sec}  ({len(entries)} entries)")
        for e in entries:
            print(f"   - {e.title!r} sub={e.subtitle!r} date={e.date!r} details={len(e.details)}")

    docx_bytes = render_docx(data, photo_path=photo, scale=scale, dense=dense)
    docx_path = out_dir / f"{stem}.docx"
    docx_path.write_bytes(docx_bytes)
    print(f"\n[docx] {docx_path}  ({len(docx_bytes)} bytes)  scale={scale} dense={dense}")


if __name__ == "__main__":
    main()
