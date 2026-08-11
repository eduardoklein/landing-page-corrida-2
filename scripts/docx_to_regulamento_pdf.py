#!/usr/bin/env python3
"""Convert regulation DOCX files to PDF for landing pages."""

from __future__ import annotations

import re
import sys
import xml.sax.saxutils as saxutils
from pathlib import Path

from docx import Document
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "assets" / "docs"

SOURCES = {
    "regulamento-aracaju-2026.pdf": Path(
        "/Users/eduardoklein/Downloads/REGULAMENTO-ARACAJU-2026-v6.docx"
    ),
    "regulamento-mateus-etapa-balsas-2026.pdf": Path(
        "/Users/eduardoklein/Downloads/REGULAMENTO-BALSAS- MATEUS.docx"
    ),
    "regulamento-sao-luis-2026.pdf": Path(
        "/Users/eduardoklein/Downloads/REGULAMENTO-SAO-LUIS-2026-v5.docx"
    ),
    "regulamento-mateus-etapa-teresina-2026.pdf": Path(
        "/Users/eduardoklein/Downloads/REGULAMENTO-TERESINA-MATEUS.docx"
    ),
}


def escape(text: str) -> str:
    return saxutils.escape(text)


def is_heading(text: str) -> bool:
    if not text:
        return False
    if text.isupper() and len(text) < 120:
        return True
    return bool(re.match(r"^\d+(\.\d+)*\.\s", text)) and len(text) < 160


def build_pdf(docx_path: Path, pdf_path: Path) -> None:
    document = Document(docx_path)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "RegTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        alignment=TA_CENTER,
        spaceAfter=14,
    )
    heading_style = ParagraphStyle(
        "RegHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=13,
        spaceBefore=8,
        spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "RegBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        alignment=TA_JUSTIFY,
        spaceAfter=6,
    )

    story = []
    for index, para in enumerate(document.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        safe = escape(text)
        if index == 0 or (text.isupper() and len(text) < 120):
            story.append(Paragraph(safe, title_style if index == 0 else heading_style))
        elif is_heading(text):
            story.append(Paragraph(safe, heading_style))
        else:
            story.append(Paragraph(safe, body_style))

    if not story:
        raise ValueError(f"No content found in {docx_path}")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    pdf = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title=pdf_path.stem,
    )
    pdf.build(story)


def main() -> int:
    for filename, source in SOURCES.items():
        if not source.exists():
            print(f"Missing source: {source}", file=sys.stderr)
            return 1
        target = OUT_DIR / filename
        build_pdf(source, target)
        print(f"Wrote {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
