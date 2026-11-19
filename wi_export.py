"""Draft Work Instruction DOCX export; no invented engineering measurements."""
from io import BytesIO
from datetime import date
from docx import Document
from docx.shared import Inches, Pt

def build_wi_docx(project, part, method, tools="", prepared_by="", revision="DRAFT"):
    doc = Document()
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "WORK INSTRUCTION | ENGINEERING DRAFT"
    doc.add_heading("WORK INSTRUCTION", 0)
    meta = doc.add_table(rows=0, cols=2)
    meta.style = "Table Grid"
    for label, value in [
        ("Project", project), ("Part", part), ("Revision", revision),
        ("Date", date.today().isoformat()), ("Prepared by", prepared_by or "Not specified"),
        ("Approval status", "DRAFT - NOT APPROVED"),
    ]:
        cells = meta.add_row().cells
        cells[0].text, cells[1].text = label, str(value)
    doc.add_heading("1. Purpose and scope", level=1)
    doc.add_paragraph("To be completed and verified by the responsible engineer.")
    doc.add_heading("2. Work method / control", level=1)
    doc.add_paragraph(method.strip() or "Not specified - engineer input required.")
    doc.add_heading("3. Tools and equipment", level=1)
    doc.add_paragraph(tools.strip() or "Not specified - engineer input required.")
    doc.add_heading("4. Acceptance criteria", level=1)
    doc.add_paragraph("Not specified - define and verify against approved engineering requirements.")
    doc.add_heading("5. Review and approval", level=1)
    doc.add_paragraph("This document is an unapproved draft. Verify all steps, values and safety controls before use.")
    doc.styles["Normal"].font.size = Pt(10)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
