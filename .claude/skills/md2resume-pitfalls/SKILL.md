---
name: md2resume-pitfalls
description: >
  WHAT: Convert a Markdown resume into one deliverable one-page DOCX while
  avoiding template font pollution, pagination overflow, privacy leaks, lost
  tmp scripts, editor lock files, and leftover PDF/HTML/PNG artifacts.
  WHEN: Use whenever the task involves md2resume, a reference DOCX template,
  one-page Chinese resumes, --scale calibration, Courier font warnings, WPS/Word
  lock files, desensitized templates, or an explicit DOCX-only deliverable.
  KEYWORDS: md2resume, Markdown to DOCX, one-page resume, reference template,
  Courier, font warning, scale, WPS lock, PII, DOCX-only.
---

# md2resume 踩坑防护

## Overview

Treat Markdown as the only content source and a reference DOCX as the layout skeleton. Deliver exactly one inspected DOCX. Do not treat PDF, PNG, or HTML as official artifacts; if visual QA needs them, generate them in a temp directory and delete them before delivery.

Load this skill together with `.claude/skills/md2resume/SKILL.md`. This file is the process and failure-mode layer; that file is the Markdown dialect and renderer map.

## Process

1. Inspect source Markdown, the reference template, and the output path. Confirm the committed template has no real name, phone, WeChat, email, or photo.
2. Prefer the reference-template clone path:

   ```bash
   python3 scripts/build_docx_from_reference.py input.md \
     --reference assets/reference-template.docx \
     --output output/resume.docx --scale 0.99
   ```

3. Pass real photos only through an external `--reference` file. Keep the in-repo template as a fake-name silhouette sample.
4. Before writing, isolate or rename any existing target file. If WPS or Word has the DOCX open, close it or use a new output name. Run `scripts/check_output_ready.py` from this skill when the path may be locked.
5. Verify content, font XML, page count, and visual layout. One-page claims require Word or LibreOffice render checks; do not infer page count from Markdown length.
6. Before delivery, confirm the output directory contains only `.docx`. Delete generated PDF, PNG, HTML, and scratch drafts. Do not delete source Markdown, scripts, or the desensitized template.

## One-page calibration

- Re-measure after content grows. This project overflowed at `scale=0.95` and stayed on one page at `scale=0.90`.
- Compress redundant copy and paragraph spacing first; then lower `--scale` in small steps. Do not shrink body text past readability.
- Check that the last bullet is not pushed to page 2, that result callout bars are intact, and that the footer is neither clipped nor overly empty.

## Template and fonts

- Keep the template's name hierarchy, blue rules, light-blue result bars, left blue accent, photo slot, and margins.
- Bold 1–2 keywords per bullet. Do not bold entire sentences.
- Unused macro styles or `fontTable.xml` may still contain a bare `Courier`. Replace exactly `"Courier"` with `Arial`; leave `"Courier New"` alone.
- Highlight only the last achievement of each experience block. Do not turn every description into a color block.

## Examples

**Input:** Markdown resume plus `assets/reference-template.docx`, request is "只要一页 DOCX".

**Output:** `output/resume.docx` at `scale=0.90`, page count verified, no PDF/PNG/HTML beside it.

**Anti-example:** Running the old multi-format renderer, leaving Chromium-missing Playwright errors, and committing a real photo template.

## Guidelines

- Promote stable builders into `scripts/`. `tmp/` is disposable.
- Do not copy personal paths, tokens, names, or image URLs from Claude Code logs into skills or docs.
- Default one-page mode in `core/renderer.py` already uses `scale=0.90`.
- For symptoms, causes, and actions, read [pitfalls.md](references/pitfalls.md).
