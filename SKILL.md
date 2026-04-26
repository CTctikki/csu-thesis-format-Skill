---
name: csu-thesis-format
description: Assist with formatting, reviewing, and repairing Central South University (中南大学, CSU) undergraduate graduation design/thesis Word documents. Use when working on 中南大学本科毕业设计(论文) DOCX files, extracting or applying CSU thesis template rules, checking page setup, cover, abstract, table of contents, body headings, figures, tables, equations, page numbering, citations, or references against the official Word template.
---

# CSU Thesis Format

## Quick Start

Use this skill for 中南大学本科毕业设计(论文) Word formatting work. First read `references/csu-thesis-format-rules.md`; it is the extracted source of truth from the bundled template.

For a candidate thesis DOCX, run:

```bash
python scripts/check_csu_thesis_docx.py path/to/thesis.docx
```

For a style inventory before editing, run:

```bash
python scripts/extract_docx_format.py path/to/thesis.docx
```

Use `assets/csu-thesis-template.docx` as the canonical template asset when creating a new document, copying styles, or comparing page setup. It is converted from the provided official `.doc` template so agents do not need legacy Word conversion support.

## Workflow

1. Confirm the input document type.
   - If the user provides `.docx`, inspect it directly.
   - If the user provides legacy `.doc`, convert it to `.docx` with LibreOffice before analysis:

```bash
soffice --headless --convert-to docx --outdir output_dir input.doc
```

2. Load the rules reference before making formatting decisions.
   - Use the reference for typography, spacing, page setup, numbering, and captions.
   - Treat the bundled template as authoritative when a rule is ambiguous.

3. Analyze before editing.
   - Run `extract_docx_format.py` to see actual sections, styles, and paragraph samples.
   - Run `check_csu_thesis_docx.py` to get a checklist of likely mismatches.

4. Edit conservatively.
   - Prefer existing thesis content and structure.
   - Preserve figures, tables, equations, captions, cross-references, and field codes.
   - Use `python-docx` for paragraph styles and page setup; use direct OOXML only when `python-docx` cannot express the required setting.
   - Do not delete user content while fixing format. If content must move across sections, keep a clear audit trail.

5. Validate visually when layout matters.
   - Export the edited DOCX to PDF with LibreOffice.
   - Inspect pages around cover, abstracts, TOC, chapter starts, figure/table-heavy pages, and references.
   - Re-run the checker after each meaningful formatting pass.

## Editing Priorities

Fix in this order:

1. Page setup and section breaks: A4, margins, header/footer distance, first page without header/footer, Roman pre-body pages, Arabic body pages.
2. Required front matter: cover, Chinese abstract, English abstract, table of contents.
3. Body hierarchy: chapter starts, title levels, body paragraphs, citations.
4. Objects: equations, tables, figures, captions.
5. Back matter: conclusion or acknowledgments, references.

## Script Notes

The scripts require Python plus `python-docx` and `lxml`. They produce Markdown reports and are intentionally heuristic: Word documents often mix styles, direct formatting, text boxes, and field codes. Use script output to find likely issues, then verify visually against the template.

## Bundled Resources

- `references/csu-thesis-format-rules.md`: extracted formatting rules and Word style notes.
- `assets/csu-thesis-template.docx`: reusable CSU thesis template asset.
- `scripts/extract_docx_format.py`: inspect a DOCX structure, styles, sections, tables, and paragraph samples.
- `scripts/check_csu_thesis_docx.py`: quick CSU-format checklist for a DOCX thesis.
