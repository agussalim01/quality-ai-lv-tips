"""Create an unapproved WI draft while preserving the source template."""
from io import BytesIO
from datetime import date
from pathlib import Path
from docx import Document

TEMPLATE_PATH = Path(__file__).resolve().parent / "TMP-Common-General-0002-Template WI Landscape.docx"

def build_wi_docx(project, part, method, tools="", prepared_by="", revision="DRAFT", orientation="Landscape"):
    if not TEMPLATE_PATH.is_file():
        raise FileNotFoundError(
            "Template WI landscape tidak ditemukan di repository. "
            "Periksa TMP-Common-General-0002-Template WI Landscape.docx."
        )
    if not method or not method.strip():
        raise ValueError("Work method / Control wajib diisi.")
    if orientation not in ("Landscape", "Portrait"):
        raise ValueError("Orientation harus Landscape atau Portrait.")
    if orientation == "Portrait":
        from docx.enum.section import WD_ORIENT
        from docx.shared import Cm
        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        doc.add_heading("WORK INSTRUCTION - DRAFT", level=0)
        doc.add_paragraph("Portrait layout - separate draft, not the approved master template.")
        meta = doc.add_table(rows=0, cols=2)
        for label, value in [
            ("Project", project), ("Part", part), ("Revision", revision),
            ("Date", date.today().isoformat()), ("Prepared by", prepared_by or "Not specified"),
            ("Approval", "DRAFT - NOT APPROVED"), ("Work method / Control", method.strip()),
            ("Tools / Equipment", tools.strip() or "Not specified"),
            ("Acceptance criteria", "Not specified - engineer review required"),
        ]:
            cells = meta.add_row().cells
            cells[0].text, cells[1].text = label, str(value)
        doc.add_paragraph("Engineering verification and formal approval required before production use.")
        output = BytesIO()
        doc.save(output)
        return output.getvalue()
    # Open the actual master template. Keep its existing paragraphs, tables,
    # images, headers, footers, page dimensions and styles unchanged.
    doc = Document(str(TEMPLATE_PATH))
    doc.add_page_break()
    heading = doc.add_paragraph()
    heading.add_run("ENGINEERING INPUT - UNAPPROVED DRAFT").bold = True
    doc.add_paragraph(
        "This page is an automatically generated draft appendix. "
        "It does not populate or approve the controlled fields in the master WI template."
    )
    table = doc.add_table(rows=0, cols=2)
    # Preserve the source template table defaults; do not require Table Grid.
    fields = [
        ("Project", project),
        ("Part", part),
        ("Revision (draft)", revision),
        ("Date", date.today().isoformat()),
        ("Prepared by", prepared_by or "Not specified"),
        ("Approval", "DRAFT - NOT APPROVED"),
        ("Work method / Control", method.strip()),
        ("Tools / Equipment", tools.strip() or "Not specified"),
        ("Acceptance criteria", "Not specified - engineering review required"),
    ]
    for label, value in fields:
        cells = table.add_row().cells
        cells[0].text = str(label)
        cells[1].text = str(value)
    doc.add_paragraph(
        "Engineering review required: confirm instructions, equipment, "
        "inspection criteria, safety requirements and document approval before use."
    )
    output = BytesIO()
    doc.save(output)
    return output.getvalue()
