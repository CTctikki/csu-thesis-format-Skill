#!/usr/bin/env python3
"""Heuristic checker for CSU undergraduate thesis DOCX formatting."""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag: str) -> str:
    return f"{{{W}}}{tag}"


def w_attr(el, name: str = "val"):
    return el.get(qn(name)) if el is not None else None


def approx(actual: float, expected: float, tolerance: float = 0.08) -> bool:
    return abs(actual - expected) <= tolerance


@dataclass
class Finding:
    level: str
    location: str
    message: str


class Report:
    def __init__(self) -> None:
        self.findings: list[Finding] = []

    def add(self, level: str, location: str, message: str) -> None:
        self.findings.append(Finding(level, location, message))

    def error(self, location: str, message: str) -> None:
        self.add("ERROR", location, message)

    def warn(self, location: str, message: str) -> None:
        self.add("WARN", location, message)

    def info(self, location: str, message: str) -> None:
        self.add("INFO", location, message)

    @property
    def error_count(self) -> int:
        return sum(1 for finding in self.findings if finding.level == "ERROR")

    @property
    def warn_count(self) -> int:
        return sum(1 for finding in self.findings if finding.level == "WARN")


def text_of(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs).strip()


def p_pr(paragraph):
    return paragraph._p.pPr


def jc(paragraph) -> str | None:
    props = p_pr(paragraph)
    return w_attr(props.find(qn("jc"))) if props is not None and props.find(qn("jc")) is not None else None


def spacing_multiple(paragraph) -> float | None:
    props = p_pr(paragraph)
    if props is None:
        return None
    spacing = props.find(qn("spacing"))
    if spacing is None:
        return None
    line = spacing.get(qn("line"))
    rule = spacing.get(qn("lineRule"))
    if not line or rule in {"exact", "atLeast"}:
        return None
    return round(int(line) / 240, 3)


def first_line_chars(paragraph) -> int | None:
    props = p_pr(paragraph)
    if props is None:
        return None
    indent = props.find(qn("ind"))
    if indent is None:
        return None
    value = indent.get(qn("firstLineChars"))
    if value is not None:
        return int(value)
    value = indent.get(qn("firstLine"))
    if value is not None:
        return round(int(value) / 240)
    return None


def outline_level(paragraph) -> int | None:
    props = p_pr(paragraph)
    if props is None:
        return None
    outline = props.find(qn("outlineLvl"))
    value = w_attr(outline)
    return int(value) if value is not None else None


def first_run_rpr(paragraph):
    for run in paragraph.runs:
        if run.text.strip():
            return run._r.rPr
    return None


def first_run_size(paragraph) -> float | None:
    r_pr = first_run_rpr(paragraph)
    if r_pr is None:
        return None
    size = r_pr.find(qn("sz"))
    value = w_attr(size)
    return round(int(value) / 2, 2) if value else None


def first_run_fonts(paragraph) -> set[str]:
    r_pr = first_run_rpr(paragraph)
    if r_pr is None:
        return set()
    fonts = r_pr.find(qn("rFonts"))
    if fonts is None:
        return set()
    result = set()
    for key in ["ascii", "hAnsi", "eastAsia", "cs"]:
        value = fonts.get(qn(key))
        if value:
            result.add(value)
    return result


def has_font(paragraph, names: list[str]) -> bool:
    fonts = first_run_fonts(paragraph)
    return any(any(name.lower() in font.lower() for font in fonts) for name in names)


def is_bold(paragraph) -> bool | None:
    r_pr = first_run_rpr(paragraph)
    if r_pr is None:
        return None
    bold = r_pr.find(qn("b"))
    if bold is None:
        return None
    return w_attr(bold) not in {"0", "false", "False"}


def is_centered(paragraph) -> bool:
    return jc(paragraph) == "center"


def paragraph_location(index: int, text: str) -> str:
    return f"paragraph {index}: {text[:32]}"


def find_body_start(document: Document) -> int:
    """Return the 1-based paragraph index where the main body likely starts."""
    candidates = []
    for index, paragraph in enumerate(document.paragraphs, 1):
        text = text_of(paragraph)
        if not re.match(r"^第[0-9一二三四五六七八九十百]+章\s+\S", text):
            continue
        if "\t" in text or re.search(r"\.{2,}\s*\d+$|…+\s*\d+$|\s+\d+$", text):
            continue
        candidates.append(index)
    return candidates[0] if candidates else 1


def find_references_start(document: Document) -> int | None:
    for index, paragraph in enumerate(document.paragraphs, 1):
        if text_of(paragraph) == "参考文献":
            return index
    return None


def check_sections(document: Document, report: Report) -> None:
    expected = {
        "page_width": 21.0,
        "page_height": 29.7,
        "top": 2.5,
        "bottom": 2.5,
        "left": 3.0,
        "right": 2.0,
        "footer": 1.75,
    }
    for index, section in enumerate(document.sections, 1):
        location = f"section {index}"
        values = {
            "page_width": section.page_width.cm,
            "page_height": section.page_height.cm,
            "top": section.top_margin.cm,
            "bottom": section.bottom_margin.cm,
            "left": section.left_margin.cm,
            "right": section.right_margin.cm,
            "footer": section.footer_distance.cm,
        }
        for key, expected_value in expected.items():
            if not approx(values[key], expected_value):
                report.error(location, f"{key} is {values[key]:.2f} cm; expected {expected_value:.2f} cm")
        header = section.header_distance.cm
        if not (approx(header, 1.25, 0.12) or approx(header, 1.5, 0.12)):
            report.warn(location, f"header distance is {header:.2f} cm; template uses about 1.25 cm or 1.50 cm")
    if document.sections and not document.sections[0].different_first_page_header_footer:
        report.warn("section 1", "cover/front section should usually use different first page header/footer")


def check_title_paragraph(index: int, paragraph, report: Report) -> None:
    text = text_of(paragraph)
    location = paragraph_location(index, text)
    size = first_run_size(paragraph)

    if text in {"摘要", "目录", "参考文献"} or text in {"结束语", "致谢", "结束语(或致谢)"}:
        if not is_centered(paragraph):
            report.warn(location, "major front/back heading should be centered")
        if size and not approx(size, 16, 0.8):
            report.warn(location, f"heading size is {size} pt; expected third-size about 16 pt")
        if not has_font(paragraph, ["黑体", "SimHei"]):
            report.warn(location, "heading should use Heiti/SimHei")

    if text == "ABSTRACT":
        if not is_centered(paragraph):
            report.warn(location, "ABSTRACT should be centered")
        if size and not approx(size, 16, 0.8):
            report.warn(location, f"ABSTRACT size is {size} pt; expected about 16 pt")
        if has_font(paragraph, ["Times New Roman"]) and is_bold(paragraph) is False:
            report.warn(location, "ABSTRACT should be bold Times New Roman")


def check_body_heading(index: int, paragraph, report: Report, body_start: int) -> None:
    if index < body_start:
        return
    text = text_of(paragraph)
    location = paragraph_location(index, text)
    size = first_run_size(paragraph)

    if re.match(r"^第[0-9一二三四五六七八九十百]+章\s+\S", text):
        if not is_centered(paragraph):
            report.warn(location, "chapter heading should be centered")
        if size and not approx(size, 16, 0.8):
            report.warn(location, f"chapter heading size is {size} pt; expected about 16 pt")
        if not has_font(paragraph, ["黑体", "SimHei"]):
            report.warn(location, "chapter heading should use Heiti/SimHei")
        return

    if re.match(r"^\d+\.\d+\.\d+\s+\S", text):
        if size and not approx(size, 12, 0.8):
            report.warn(location, f"third-level heading size is {size} pt; expected 12 pt")
        if not has_font(paragraph, ["楷", "Kai"]):
            report.warn(location, "third-level heading should use KaiTi/楷体GB2312")
        chars = first_line_chars(paragraph)
        if chars not in {None, 200, 2}:
            report.warn(location, "third-level heading should indent about two Chinese characters")
        return

    if re.match(r"^\d+\.\d+\s+\S", text):
        if size and not approx(size, 12, 0.8):
            report.warn(location, f"second-level heading size is {size} pt; expected 12 pt")
        if not has_font(paragraph, ["黑体", "SimHei"]):
            report.warn(location, "second-level heading should use Heiti/SimHei")
        chars = first_line_chars(paragraph)
        if chars not in {None, 200, 2}:
            report.warn(location, "second-level heading should indent about two Chinese characters")


def check_caption_or_equation(index: int, paragraph, report: Report, body_start: int) -> None:
    if index < body_start:
        return
    text = text_of(paragraph)
    location = paragraph_location(index, text)
    size = first_run_size(paragraph)

    if re.match(r"^表\d+[-－]\d+\s+\S", text):
        if not is_centered(paragraph):
            report.warn(location, "table title should be centered above the table")
        if size and not approx(size, 10.5, 0.8):
            report.warn(location, f"table title size is {size} pt; expected fifth-size about 10.5 pt")
        if not has_font(paragraph, ["黑体", "SimHei"]):
            report.warn(location, "table title should use Heiti/SimHei")

    if re.match(r"^图\d+[-－]\d+\s+\S", text):
        if not is_centered(paragraph):
            report.warn(location, "figure title should be centered below the figure")
        if size and not approx(size, 10.5, 0.8):
            report.warn(location, f"figure title size is {size} pt; expected fifth-size about 10.5 pt")
        if not has_font(paragraph, ["黑体", "SimHei"]):
            report.warn(location, "figure title should use Heiti/SimHei")

    if re.fullmatch(r"[（(]\d+[-－]\d+[）)]", text):
        if jc(paragraph) not in {"right", "end"}:
            report.warn(location, "standalone equation number should be right aligned")


def looks_like_body(text: str) -> bool:
    if len(text) < 25:
        return False
    if re.match(r"^(第[0-9一二三四五六七八九十百]+章|\d+\.\d+|图\d+|表\d+|摘要|ABSTRACT|目录|参考文献|结束语|致谢)", text):
        return False
    if re.fullmatch(r"[（(]\d+[-－]\d+[）)]", text):
        return False
    return True


def check_body_paragraphs(document: Document, report: Report, body_start: int, references_start: int | None, limit: int = 25) -> None:
    seen = 0
    for index, paragraph in enumerate(document.paragraphs, 1):
        if index < body_start:
            continue
        if references_start is not None and index >= references_start:
            continue
        text = text_of(paragraph)
        if not looks_like_body(text):
            continue
        seen += 1
        if seen > limit:
            return
        location = paragraph_location(index, text)
        line = spacing_multiple(paragraph)
        if line is not None and not approx(line, 1.5, 0.05):
            report.warn(location, f"body paragraph line spacing is {line}; expected 1.5")
        chars = first_line_chars(paragraph)
        if chars not in {None, 200, 2}:
            report.warn(location, "body paragraph should have first-line indent of about two Chinese characters")
        size = first_run_size(paragraph)
        if size and not approx(size, 12, 1.0):
            report.warn(location, f"body text size is {size} pt; expected small fourth-size about 12 pt")


def check_keywords(document: Document, report: Report) -> None:
    for index, paragraph in enumerate(document.paragraphs, 1):
        text = text_of(paragraph)
        if text.startswith("关键词"):
            terms = re.split(r"\s{2,}|　|；|;|，|,", text.split("：", 1)[-1].strip())
            terms = [term for term in terms if term]
            if not (3 <= len(terms) <= 8):
                report.warn(paragraph_location(index, text), f"Chinese keywords count is {len(terms)}; expected 3-8")
        if text.lower().startswith(("key words", "keywords")):
            terms = re.split(r"\s{2,}|;|,|；|，", re.split(r"[:：]", text, 1)[-1].strip())
            terms = [term for term in terms if term]
            if terms and not (3 <= len(terms) <= 8):
                report.warn(paragraph_location(index, text), f"English keywords count is {len(terms)}; expected 3-8")


def check_references(document: Document, report: Report) -> None:
    in_refs = False
    numbers: list[int] = []
    for index, paragraph in enumerate(document.paragraphs, 1):
        text = text_of(paragraph)
        if text == "参考文献":
            in_refs = True
            continue
        if not in_refs:
            continue
        match = re.match(r"^\[(\d+)\]\s+", text)
        if match:
            numbers.append(int(match.group(1)))
            size = first_run_size(paragraph)
            if size and not approx(size, 10.5, 1.0):
                report.warn(paragraph_location(index, text), f"reference entry size is {size} pt; expected fifth-size about 10.5 pt")
            if not has_font(paragraph, ["楷", "Kai"]):
                report.warn(paragraph_location(index, text), "reference entry should use KaiTi/楷体GB2312")
    if numbers:
        expected = list(range(1, len(numbers) + 1))
        if numbers != expected:
            report.warn("references", f"reference numbers are {numbers}; expected continuous order {expected}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--strict", action="store_true", help="exit non-zero when errors are found")
    args = parser.parse_args()

    if not args.docx.exists():
        raise SystemExit(f"File not found: {args.docx}")

    document = Document(str(args.docx))
    report = Report()
    body_start = find_body_start(document)
    references_start = find_references_start(document)

    check_sections(document, report)
    for index, paragraph in enumerate(document.paragraphs, 1):
        if not text_of(paragraph):
            continue
        check_title_paragraph(index, paragraph, report)
        check_body_heading(index, paragraph, report, body_start)
        check_caption_or_equation(index, paragraph, report, body_start)
    check_body_paragraphs(document, report, body_start, references_start)
    check_keywords(document, report)
    check_references(document, report)

    print(f"# CSU Thesis DOCX Check: {args.docx.name}\n")
    print("## Summary")
    print(f"- Errors: {report.error_count}")
    print(f"- Warnings: {report.warn_count}")
    print(f"- Paragraphs checked: {len(document.paragraphs)}")
    print(f"- Sections checked: {len(document.sections)}")
    print(f"- Tables found: {len(document.tables)}")
    print(f"- Main body starts near paragraph: {body_start}")

    print("\n## Findings")
    if not report.findings:
        print("- No obvious CSU-format issues found by the heuristic checker.")
    else:
        for finding in report.findings:
            print(f"- [{finding.level}] {finding.location}: {finding.message}")

    print("\n## Reminder")
    print("- This checker is heuristic. Render the DOCX to PDF and inspect cover, abstracts, TOC, chapter starts, captions, and references visually.")

    return 1 if args.strict and report.error_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
