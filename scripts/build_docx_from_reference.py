from __future__ import annotations

import argparse
import re
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph


ACCENT = RGBColor(0x2E, 0x75, 0xB6)
DARK_BLUE = RGBColor(0x1F, 0x4E, 0x79)
GRAY = RGBColor(0x59, 0x59, 0x59)
BODY = RGBColor(0x40, 0x40, 0x40)
LATIN_FONT = "Arial"
CJK_FONT = "微软雅黑"

# Fonts whose embedding rights trip PDF "受限字体无法嵌入" warnings. The reference
# template ships bare "Courier" in two unused macro styles + fontTable.xml; remap
# to an embeddable substitute so downstream PDF export stays clean. Match the
# quoted form only, so "Courier New" (embeddable) is left untouched.
RESTRICTED_FONT_REMAP = {"Courier": "Arial"}


def _sanitize_fonts(data: bytes) -> bytes:
    if not data:
        return data
    text = data.decode("utf-8", errors="replace")
    for bad, good in RESTRICTED_FONT_REMAP.items():
        text = text.replace(f'"{bad}"', f'"{good}"').replace(f"'{bad}'", f"'{good}'")
    return text.encode("utf-8")


@dataclass
class Entry:
    company: str
    subtitle: str
    date: str
    project: str = ""
    bullets: list[str] = field(default_factory=list)


@dataclass
class Resume:
    name: str = ""
    role: str = ""
    contacts: str = ""
    target: str = ""
    positioning: str = ""
    education: Entry | None = None
    education_bullets: list[str] = field(default_factory=list)
    experience: list[Entry] = field(default_factory=list)
    capabilities: list[str] = field(default_factory=list)


def strip_outer_bold(text: str) -> str:
    text = text.strip()
    if text.startswith("**") and text.endswith("**"):
        return text[2:-2].strip()
    return text


def parse_entry_header(text: str) -> Entry:
    parts = [part.strip() for part in text.split("｜") if part.strip()]
    if len(parts) < 3:
        raise ValueError(f"Cannot parse entry header: {text}")
    return Entry(parts[0], " · ".join(parts[1:-1]), parts[-1])


def parse_resume(path: Path) -> Resume:
    resume = Resume()
    section = "header"
    current: Entry | None = None

    def finish_entry() -> None:
        nonlocal current
        if current is not None:
            resume.experience.append(current)
            current = None

    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line:
            continue

        if line.startswith("# ") and not line.startswith("## "):
            name_role = [part.strip() for part in line[2:].strip().split("｜", 1)]
            resume.name = name_role[0]
            resume.role = name_role[1] if len(name_role) == 2 else ""
            continue

        if section == "header" and line.startswith("电话："):
            resume.contacts = line.replace("  ", " ")
            continue
        if section == "header" and line.startswith("求职方向："):
            resume.target = line
            continue
        if section == "header" and line.startswith("个人定位："):
            resume.positioning = line
            continue

        if line.startswith("## "):
            heading = line[3:].strip()
            if heading in {"教育背景", "实习经历", "核心能力"}:
                if heading == "核心能力":
                    finish_entry()
                section = heading
                continue
            if section == "实习经历" and "实习生" in heading:
                finish_entry()
                current = parse_entry_header(heading)
                continue

        if line.startswith("### "):
            heading = line[4:].strip()
            if not heading:
                continue
            if section == "实习经历":
                finish_entry()
                current = parse_entry_header(heading)
                continue

        if section == "教育背景":
            if line.startswith("**") and "｜" in line:
                resume.education = parse_entry_header(strip_outer_bold(line))
            elif not line.startswith("#"):
                # 任意非标题行 → 教育要点（核心课程/排名等）
                resume.education_bullets.append(line)
            continue

        if section == "实习经历" and current is not None:
            if line.startswith("**") and not line.startswith("- "):
                current.project = strip_outer_bold(line)
            elif line.startswith("- "):
                detail = line[2:].strip()
                if detail.startswith(">"):
                    detail = detail[1:].strip()
                current.bullets.append(detail)
            continue

        if section == "核心能力" and line.startswith("- "):
            resume.capabilities.append(line[2:].strip())

    finish_entry()
    if not resume.education or len(resume.experience) < 1:
        raise ValueError(
            f"Unexpected source structure: education={bool(resume.education)}, "
            f"experience={len(resume.experience)} (need >=1)"
        )
    return resume


def _clear_runs(paragraph: Paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def _copy_rpr(run, template_rpr) -> None:
    if template_rpr is None:
        return
    current = run._r.find(qn("w:rPr"))
    if current is not None:
        run._r.remove(current)
    run._r.insert(0, deepcopy(template_rpr))


def _add_run(paragraph: Paragraph, text: str, template_rpr=None):
    run = paragraph.add_run(text)
    _copy_rpr(run, template_rpr)
    return run


def _font(run, *, size: float, bold: bool, color: RGBColor) -> None:
    run.font.name = LATIN_FONT
    rpr = run._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), CJK_FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def _inline_parts(text: str) -> list[tuple[str, bool]]:
    parts: list[tuple[str, bool]] = []
    for index, segment in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if segment:
            parts.append((segment, index % 2 == 1))
    return parts


def _clone_paragraph(pattern: Paragraph) -> Paragraph:
    element = deepcopy(pattern._p)
    return Paragraph(element, pattern._parent)


def _fill_section(pattern: Paragraph, title: str) -> Paragraph:
    paragraph = _clone_paragraph(pattern)
    rprs = [deepcopy(run._r.rPr) for run in pattern.runs[:2]]
    _clear_runs(paragraph)
    _add_run(paragraph, "● ", rprs[0])
    _add_run(paragraph, title, rprs[1])
    return paragraph


def _fill_entry_head(pattern: Paragraph, entry: Entry) -> Paragraph:
    paragraph = _clone_paragraph(pattern)
    rprs = [deepcopy(run._r.rPr) for run in pattern.runs]
    _clear_runs(paragraph)
    _add_run(paragraph, entry.company, rprs[0])
    _add_run(paragraph, "  " + entry.subtitle, rprs[1])
    _add_run(paragraph, "\t", rprs[2])
    _add_run(paragraph, entry.date, rprs[-1])
    return paragraph


def _fill_project(pattern: Paragraph, text: str) -> Paragraph:
    paragraph = _clone_paragraph(pattern)
    rpr = deepcopy(pattern.runs[0]._r.rPr)
    _clear_runs(paragraph)
    _add_run(paragraph, text, rpr)
    return paragraph


def _fill_bullet(pattern: Paragraph, text: str) -> Paragraph:
    paragraph = _clone_paragraph(pattern)
    marker_rpr = deepcopy(pattern.runs[0]._r.rPr)
    tab_rpr = deepcopy(pattern.runs[1]._r.rPr)
    bold_rpr = deepcopy(pattern.runs[2]._r.rPr)
    normal_template = next((run for run in pattern.runs[2:] if run.bold is False), pattern.runs[-1])
    normal_rpr = deepcopy(normal_template._r.rPr)
    _clear_runs(paragraph)
    _add_run(paragraph, "•", marker_rpr)
    _add_run(paragraph, "\t", tab_rpr)
    for segment, bold in _inline_parts(text):
        _add_run(paragraph, segment, bold_rpr if bold else normal_rpr)
    return paragraph


def _fill_result(pattern: Paragraph, text: str) -> Paragraph:
    paragraph = _clone_paragraph(pattern)
    normal_template = next((run for run in pattern.runs if run.bold is not True), pattern.runs[0])
    bold_template = next((run for run in pattern.runs if run.bold is True), pattern.runs[0])
    normal_rpr = deepcopy(normal_template._r.rPr)
    bold_rpr = deepcopy(bold_template._r.rPr)
    _clear_runs(paragraph)
    for segment, bold in _inline_parts(text):
        _add_run(paragraph, segment, bold_rpr if bold else normal_rpr)
    return paragraph


def _replace_header(resume: Resume, doc: Document) -> None:
    left = doc.tables[0].cell(0, 0)
    name_p = left.paragraphs[0]
    contact_p = left.paragraphs[1]

    name_rpr = deepcopy(name_p.runs[0]._r.rPr)
    _clear_runs(name_p)
    _add_run(name_p, resume.name, name_rpr)
    role_run = _add_run(name_p, "  " + resume.role, name_rpr)
    _font(role_run, size=11, bold=True, color=DARK_BLUE)

    contact_rprs = [deepcopy(run._r.rPr) for run in contact_p.runs]
    normal_rpr = contact_rprs[0]
    separator_rpr = contact_rprs[1] if len(contact_rprs) > 1 else normal_rpr
    _clear_runs(contact_p)
    contact_parts = [part.strip() for part in resume.contacts.split("｜")]
    for index, part in enumerate(contact_parts):
        _add_run(contact_p, part.replace("：", " ", 1), normal_rpr)
        if index != len(contact_parts) - 1:
            _add_run(contact_p, "  |  ", separator_rpr)

    previous = contact_p._p
    for text in (resume.target, resume.positioning):
        element = deepcopy(contact_p._p)
        previous.addnext(element)
        paragraph = Paragraph(element, left)
        _clear_runs(paragraph)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0.4)
        paragraph.paragraph_format.line_spacing = 1.0
        for segment, bold in _inline_parts(text):
            run = _add_run(paragraph, segment, normal_rpr)
            _font(run, size=8.5, bold=bold, color=DARK_BLUE if bold else BODY)
        previous = element


def _is_result(entry: Entry, index: int, text: str) -> bool:
    # Each experience's last bullet is its headline result → highlighted blue bar.
    # Company-agnostic: highlight the last achievement of each experience block.
    return index == len(entry.bullets) - 1


def _iter_all_paragraphs(doc: Document):
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def _scale_typography_and_rhythm(doc: Document, scale: float) -> None:
    """Apply the minimal scale needed for the explicit one-page constraint."""
    for paragraph in _iter_all_paragraphs(doc):
        paragraph_format = paragraph.paragraph_format
        if paragraph_format.space_before is not None:
            paragraph_format.space_before = Pt(paragraph_format.space_before.pt * scale)
        if paragraph_format.space_after is not None:
            paragraph_format.space_after = Pt(paragraph_format.space_after.pt * scale)
        for run in paragraph.runs:
            if run.font.size is not None:
                run.font.size = Pt(run.font.size.pt * scale)

    # Word/WPS keep a 1.27cm footer gap by default, which leaves a blank band
    # at the bottom. Pull footer distance down to the real bottom margin.
    for section in doc.sections:
        if section.bottom_margin is not None:
            section.footer_distance = section.bottom_margin


def build(source: Path, reference: Path, output: Path, scale: float = 0.95) -> None:
    resume = parse_resume(source)
    doc = Document(reference)
    patterns = list(doc.paragraphs)
    if len(patterns) < 27 or not doc.tables:
        raise ValueError("Reference structure no longer matches the distilled template contract")

    section_pattern = patterns[0]
    first_head_pattern = patterns[1]
    bullet_pattern = patterns[2]
    project_pattern = patterns[6]
    result_pattern = patterns[10]
    divided_head_pattern = patterns[11]

    _replace_header(resume, doc)

    for paragraph in list(doc.paragraphs):
        paragraph._p.getparent().remove(paragraph._p)

    built: list[Paragraph] = []
    built.append(_fill_section(section_pattern, "教育背景"))
    assert resume.education is not None
    built.append(_fill_entry_head(first_head_pattern, resume.education))
    for detail in resume.education_bullets:
        built.append(_fill_bullet(bullet_pattern, detail))

    built.append(_fill_section(section_pattern, "实习经历"))
    for entry_index, entry in enumerate(resume.experience):
        head_pattern = first_head_pattern if entry_index == 0 else divided_head_pattern
        built.append(_fill_entry_head(head_pattern, entry))
        built.append(_fill_project(project_pattern, entry.project))
        for bullet_index, detail in enumerate(entry.bullets):
            if _is_result(entry, bullet_index, detail):
                built.append(_fill_result(result_pattern, detail))
            else:
                built.append(_fill_bullet(bullet_pattern, detail))

    built.append(_fill_section(section_pattern, "核心能力"))
    for detail in resume.capabilities:
        built.append(_fill_bullet(bullet_pattern, detail))

    body = doc._element.body
    sect_pr = body.sectPr
    for paragraph in built:
        body.insert(body.index(sect_pr), paragraph._p)

    _scale_typography_and_rhythm(doc, scale)

    output.parent.mkdir(parents=True, exist_ok=True)

    # Save the edited document XML once, then rebuild the final package from
    # the untouched reference. This keeps every preserve-only theme, style,
    # numbering, relationship, custom XML, and media part byte-for-byte intact.
    with NamedTemporaryFile(suffix=".docx") as generated:
        doc.save(generated.name)
        with ZipFile(reference) as source_zip, ZipFile(generated.name) as generated_zip:
            new_document_xml = generated_zip.read("word/document.xml")
            with ZipFile(output, "w", compression=ZIP_DEFLATED) as output_zip:
                for item in source_zip.infolist():
                    data = (
                        new_document_xml
                        if item.filename == "word/document.xml"
                        else source_zip.read(item.filename)
                    )
                    if item.filename.endswith((".xml", ".rels")):
                        data = _sanitize_fonts(data)
                    output_zip.writestr(item, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("--reference", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=0.99)
    args = parser.parse_args()
    build(args.source, args.reference, args.output, args.scale)
    print(f"created={args.output}")


if __name__ == "__main__":
    main()
