from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches


def make_docx(path: Path) -> None:
    doc = Document()
    doc.add_heading("Services Agreement", level=1)
    doc.add_paragraph("This agreement is between Acme Co. and Northwind.")
    doc.add_paragraph("Payment is due within 30 days of the invoice date.")
    doc.add_paragraph("The parties may revise the scope in writing.")
    doc.save(path)


def make_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget"
    ws["A1"] = "Growth"
    ws["B1"] = 0.08
    ws["A2"] = "Base"
    ws["B2"] = 10000
    ws["A3"] = "Projected"
    ws["B3"] = "=B2*(1+B1)"
    ws["A4"] = "Q1"
    ws["B4"] = 1200
    ws["A5"] = "Q2"
    ws["B5"] = 1300
    ws["A6"] = "Q3"
    ws["B6"] = 1400
    ws["A7"] = "Q4"
    ws["B7"] = 1500
    ws["A8"] = "Total"
    ws["B8"] = "=SUM(B4:B7)"
    wb.save(path)


def _textbox(slide, left, top, width, height, text, name):
    shape = slide.shapes.add_textbox(left, top, width, height)
    shape.text = text
    shape.name = name
    return shape


def make_pptx(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Quarterly Review"
    slide.placeholders[1].text = "Acme Co."
    _textbox(slide, Inches(1), Inches(2.5), Inches(6), Inches(1), "Northwind partnership", "hero-box")

    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    _textbox(slide2, Inches(1), Inches(1), Inches(7), Inches(1), "Revenue grew 12%", "metric-box")
    _textbox(slide2, Inches(1), Inches(2), Inches(4), Inches(1), "Pipeline is healthy", "status-box")

    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    _textbox(slide3, Inches(1), Inches(1), Inches(6), Inches(1), "Risks and next steps", "risk-box")

    prs.save(path)


def make_pptx_v2(path: Path, source: Path) -> None:
    prs = Presentation(source)
    prs.slides[0].shapes.title.text = "Quarterly Review v2"
    prs.slides.move(prs.slides[2], 1)
    slide = prs.slides[1]
    _textbox(slide, Inches(1), Inches(3), Inches(4), Inches(1), "New action item", "action-box")
    prs.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output_dir")
    args = parser.parse_args()
    out = Path(args.output_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    make_docx(out / "Services Agreement.docx")
    make_xlsx(out / "Budget FY26.xlsx")
    make_pptx(out / "Quarterly Review.pptx")
    make_pptx_v2(out / "Quarterly Review v2.pptx", out / "Quarterly Review.pptx")


if __name__ == "__main__":
    main()
