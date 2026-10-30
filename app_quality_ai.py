import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
from datetime import datetime

st.set_page_config(page_title="WI Generator - WIN NT8", layout="wide")

def create_wi_exact(doc_no, product, temp_no, rev, process_desc, part_req, tools_req, time_df, steps_text):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(9)

    # === HEADER EXACT SEPERTI PDF ===
    t = doc.add_table(rows=5, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Set lebar
    t.autofit = False
    for row in t.rows:
        row.height = Cm(0.9)
    
    # Row 0
    t.cell(0,0).text = "WORK INSTRUCTION"
    t.cell(0,1).text = f"Revision No.\n{rev}"
    t.cell(0,2).text = f"Document No.\n{doc_no}"
    t.cell(0,3).text = f"Product\n{product}"
    # Row 1
    t.cell(1,0).text = ""
    t.cell(1,1).text = ""
    t.cell(1,2).text = f"Temp. Number : {temp_no}"
    t.cell(1,3).text = f"Release Date\n{datetime.now().strftime('%d/%m/%Y')}"
    # Row 2 - Process Desc merge
    c = t.cell(2,0).merge(t.cell(2,3))
    c.text = f"Process Desc.\n{process_desc}"
    # Row 3 - Work Instruction
    c2 = t.cell(3,0).merge(t.cell(3,3))
    c2.text = "Work Instruction\nQC 100% Visual Inspection LV Tips"
    # Row 4 - Judul besar tengah
    c3 = t.cell(4,0).merge(t.cell(4,3))
    p = c3.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"\n{process_desc}\n{product}\n\n")
    run.bold = True
    run.font.size = Pt(12)

    # Bold semua header
    for r in t.rows[:3]:
        for cell in r.cells:
            for par in cell.paragraphs:
                for run in par.runs:
                    run.bold = True
                    run.font.size = Pt(8)

    doc.add_paragraph("") # spasi

    # === TABEL LANGKAH KERJA + TIME (Page 1 contoh) ===
    p = doc.add_paragraph()
    run = p.add_run("Langkah Kerja")
    run.bold = True
    run.font.size = Pt(10)

    # Buat tabel ringkasan
    summary = doc.add_table(rows=1, cols=3)
    summary.alignment = WD_TABLE_ALIGNMENT.CENTER
    summary.style = 'Table Grid'
    hdr = summary.rows[0].cells
    hdr[0].text = "Langkah Kerja"
    hdr[1].text = "Time\n(min)"
    hdr[2].text = "Total"
    for c in hdr: 
        for par in c.paragraphs:
            for run in par.runs: run.bold = True

    total = 0
    for _, row in time_df.iterrows():
        r = summary.add_row().cells
        r[0].text = str(row['step'])
        r[1].text = str(row['time'])
        r[2].text = ""
        try: total += int(row['time'])
        except: pass
    
    r = summary.add_row().cells
    r[0].text = "Total"
    r[1].text = str(total)
    r[2].text = str(total)

    # Bold total
    for c in r:
        for par in c.paragraphs:
            for run in par.runs: run.bold = True

    doc.add_page_break()

    # === HALAMAN 2: PART & TOOLS ===
    # Ulang header kecil
    p = doc.add_paragraph(f"Part required :")
    p.runs[0].bold = True
    for item in part_req.split('\n'):
        if item.strip():
            doc.add_paragraph(f"●  {item.strip()}", style='List Bullet')
    
    p = doc.add_paragraph(f"Tools required :")
    p.runs[0].bold = True
    for item in tools_req.split('\n'):
        if item.strip():
            doc.add_paragraph(f"●  {item.strip()}", style='List Bullet')

    doc.add_paragraph("")

    # === DETAIL STEP (seperti halaman 2-3 PDF kamu) ===
    # steps_text format: pisah per Step
    import re
    steps = re.split(r'(Step\s*\d+\s*:.*)', steps_text, flags=re.I)
    # Gabungkan judul + isi
    final_steps = []
    for i in range(1, len(steps), 2):
        title = steps[i].strip()
        content = steps[i+1] if i+1 < len(steps) else ""
        bullets = [b.strip() for b in content.split('\n') if b.strip()]
        final_steps.append((title, bullets))

    if not final_steps: # fallback jika user tidak ketik Step
        final_steps = [("Steps", [s for s in steps_text.split('\n') if s.strip()])]

    for title, bullets in final_steps:
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(10)
        for b in bullets:
            if b:
                clean = re.sub(r'^[•\-\*]\s*', '', b)
                doc.add_paragraph(f"●  {clean}", style='List Bullet')
        doc.add_paragraph("")

    # Footer
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_para.text = f"Page 1 of 6 | {doc_no} | Generated {datetime.now().strftime('%d/%m/%Y')}"

    return doc

# UI
st.title("📄 WIN Generator - Format Plek Asli NT8")

with st.container():
    c1, c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Document No", "WIN-13-005-0015")
        product = st.text_input("Product", "NT8")
        temp_no = st.text_input("Temp Number", "TMP-11-001-0002 rev.1")
        rev = st.text_input("Revision No", "0")
        process_desc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips")
    with c2:
        part_req = st.text_area("Part Required", "1 batch LV Tips after cutting", height=80)
        tools_req = st.text_area("Tools Required", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120)

    st.markdown("**Tabel Waktu (Halaman 1)**")
    time_df = st.data_editor(pd.DataFrame([
        {"step": "Step 1 :  Susun Tips", "time": 10},
        {"step": "Step 2 :  Check visual defect", "time": 80},
        {"step": "Step 3 :  Data Tips OK dan Tips Reject", "time": 5},
        {"step": "Step 4 :  Packing LV Tips", "time": 5},
        {"step": "Step 5 : Manual Filling Process", "time": 30},
        {"step": "Step 6 : 2nd Visual Check", "time": 5},
        {"step": "Step 7 : Ionizer", "time": 5},
    ]), num_rows="dynamic", use_container_width=True)

    raw = st.text_area("Detail Langkah (Copy persis dari contoh kamu, pakai format Step X : )", 
"""Step 1 :  Susun Tips
Susun LV tips di atas wadah (tatakan) LV tips secara horizontal dengan menggunakan pinset.
Step 2 :  Check visual defect
Cek tips secara kasat mata dengan pencahayaan lampu LED dengan jarak 30cm
Cek visual bagian ujung sampai pangkal tips
Pastikan tidak terdapat bubble, sinkmark, penyok, short, flashing, kontaminasi
Step 3 :  Data Tips OK dan Tips Reject
Catat jumlah tips OK dan tips yang Reject hasil QC 100% di form QC 100%
Step 4 :  Packing LV Tips
Masukkan kembali tips yang OK ke dalam plastik ziplock sebelumnya.""", height=250)

    if st.button("✨ Generate WIN Plek Asli", type="primary", use_container_width=True):
        doc = create_wi_exact(doc_no, product, temp_no, rev, process_desc, part_req, tools_req, time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success("WIN berhasil dibuat plek contoh! Ada Page Break, Footer, dan Table Grid seperti PDF asli.")
        st.download_button("📥 Download WIN NT8 Final", data=bio, file_name=f"{doc_no}_{product}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
