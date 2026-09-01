#!/usr/bin/env python3
"""Strip author metadata and preview thumbnails from a DOCX package."""
from __future__ import annotations

import argparse
import re
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

THUMBNAIL = "docProps/thumbnail.emf"
CORE = "docProps/core.xml"
RELS = "_rels/.rels"


def _scrub_core(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    text = re.sub(r"(<dc:creator>)[^<]*(</dc:creator>)", r"\1md2resume\2", text)
    text = re.sub(
        r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)",
        r"\1md2resume\2",
        text,
    )
    return text.encode("utf-8")


def _scrub_rels(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    text = re.sub(
        r'<Relationship[^>]*Type="[^"]*thumbnail"[^>]*/>',
        "",
        text,
    )
    return text.encode("utf-8")


def sanitize(path: Path, image: Path | None = None) -> None:
    buf = BytesIO()
    with ZipFile(path) as src, ZipFile(buf, "w", compression=ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename == THUMBNAIL:
                continue
            data = src.read(item.filename)
            if item.filename == CORE:
                data = _scrub_core(data)
            elif item.filename == RELS:
                data = _scrub_rels(data)
            elif image is not None and item.filename == "word/media/image1.jpeg":
                data = image.read_bytes()
            dst.writestr(item, data)
    path.write_bytes(buf.getvalue())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--image", type=Path, help="Replacement JPEG for word/media/image1.jpeg")
    args = parser.parse_args()
    sanitize(args.docx, args.image)
    print(f"sanitized={args.docx}")


if __name__ == "__main__":
    main()
