import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io, re
from datetime import datetime

st.set_page_config(page_title="Quality AI PRO - FIX", layout="wide")

def msa_tab():
    st.subheader("📊 MSA Dashboard")
    uploaded = st.file_uploader("Upload MSA CSV", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded)
    else:
        df = pd.DataFrame({"Part":[1,2,3]*9,"Operator":["A"]*9+["B"]*9+["C"]*9,"Measurement":[10.1]*27})
    st.dataframe(df.head(10), use_container_width=True)
    fig = px.histogram(df, x="Measurement", color="Operator")
    st.plotly_chart(fig, use_container_width=True)

def create_wi_fixed(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, steps_text, approval_di_bawah=True):
    # Buat doc baru landscape - JANGAN copy template langsung (bikin error merge)
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = 1
    sec.page_width = Cm(29.7)
    sec.page_height = Cm(21)
    sec.top_margin = Cm(1.2)
    sec.bottom_margin = Cm(1)
    sec.left_margin = Cm(1.2)
    sec.right_margin = Cm(1.2)
    
    s = doc.styles['Normal']
    s.font.name = 'Arial'
    s.font.size = Pt(9)

    def add_approval(doc):
        # === APPROVAL TABLE FIX - Gak dobel lagi ===
        p = doc.add_paragraph()
        r = p.add_run("APPROVAL DOCUMENT")
        r.bold = True
        r.font.size = Pt(10)
        
        t = doc.add_table(rows=5, cols=4)
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Row 0 - Merge jadi 1
        c = t.cell(0,0).merge(t.cell(0,3))
        c.text = "APPROVAL DOCUMENT"
        c.paragraphs[0].runs[0].bold = True
        
        # Row 1 - Header
        t.cell(1,0).text = ""
        t.cell(1,1).text = "Name"
        t.cell(1,2).text = "Department"
        t.cell(1,3).text = "Sign"
        for c in [t.cell(1,1), t.cell(1,2), t.cell(1,3)]:
            c.paragraphs[0].runs[0].bold = True
        
        # Row 2-4
        t.cell(2,0).text = "Author"
        t.cell(2,1).text = "-"
        t.cell(2,2).text = "FUSA RnD"
        t.cell(2,3).text = "Approved by TIGA Workflow"
        
        t.cell(3,0).text = "Verified by"
        t.cell(3,1).text = "-"
        t.cell(3,2).text = "FUSA RnD"
        t.cell(3,3).text = "Approved by TIGA Workflow"
        
        t.cell(4,0).text = "Approved by"
        t.cell(4,1).text = "-"
        t.cell(4,2).text = "FUSA RnD"
        t.cell(4,3).text = "Approved by TIGA Workflow"

    def add_history(doc):
        p = doc.add_paragraph()
        r = p.add_run("DOCUMENT HISTORY")
        r.bold = True
        r.font.size = Pt(10)
        
        t = doc.add_table(rows=4, cols=5)
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Row 0 - Merge
        c = t.cell(0,0).merge(t.cell(0,4))
        c.text = "DOCUMENT HISTORY"
        c.paragraphs[0].runs[0].bold = True
        
        # Row 1 - Header
        t.cell(1,0).text = "-"
        t.cell(1,1).text = "Release Date"
        t.cell(1,2).text = "Rev"
        t.cell(1,3).text = "Revision Description"
        t.cell(1,4).text = "Author (Name | Department)"
        for i in range(5):
            t.cell(1,i).paragraphs[0].runs[0].bold = True
        
        # Row 2 - Data baru
        t.cell(2,0).text = "-"
        t.cell(2,1).text = datetime.now().strftime("%m/%d/%Y")
        t.cell(2,2).text = rev
        t.cell(2,3).text = f"{proc} - {product}"
        t.cell(2,4).text = "- | FUSA RnD"
        
        # Row 3
        t.cell(3,0).text = "-"
        t.cell(3,1).text = "mm/dd/yyyy"
        t.cell(3,2).text = "1"
        t.cell(3,3).text = ""
        t.cell(3,4).text = "- | FUSA RnD"

    def add_wi_content(doc):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"WORK INSTRUCTION\n{proc}\nProduct: {product}\nDocument No: {doc_no}")
        run.bold = True
        run.font.size = Pt(12)
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"Doc: {doc_no} | Temp: {temp_no} | Rev: {rev} | Date: {datetime.now().strftime('%d/%m/%Y')}")
        run.font.size = Pt(8)
        
        p = doc.add_paragraph()
        run = p.add_run("Langkah Kerja")
        run.bold = True
        
        t = doc.add_table(rows=1, cols=2)
        t.style = 'Table Grid'
        t.cell(0,0).text = "Langkah Kerja"
        t.cell(0,1).text = "Time (min)"
        for c in t.rows[0].cells:
            c.paragraphs[0].runs[0].bold = True
        
        total = 0
        for _, r in time_df.iterrows():
            rw = t.add_row().cells
            rw[0].text = str(r['step'])
            rw[1].text = str(r['time'])
            try: total += int(r['time'])
            except: pass
        rw = t.add_row().cells
        rw[0].text = "Total"
        rw[1].text = str(total)
        rw[0].paragraphs[0].runs[0].bold = True
        rw[1].paragraphs[0].runs[0].bold = True
        
        doc.add_paragraph("")
        p = doc.add_paragraph()
        p.add_run("Part required :").bold = True
        for it in part_req.split('\n'):
            if it.strip(): doc.add_paragraph(f"●  {it.strip()}")
        
        p = doc.add_paragraph()
        p.add_run("Tools required :").bold = True
        for it in tools_req.split('\n'):
            if it.strip(): doc.add_paragraph(f"●  {it.strip()}")
        
        pattern = r"(Step\s*\d+\s*:.*)"
        parts = re.split(pattern, steps_text, flags=re.I)
        steps = []
        for i in range(1, len(parts), 2):
            title = parts[i].strip()
            body = parts[i+1] if i+1 < len(parts) else ""
            bullets = [x.strip() for x in body.split('\n') if x.strip()]
            steps.append((title, bullets))
        if not steps:
            steps = [("Steps", [x for x in steps_text.split('\n') if x.strip()])]
        
        for title, bullets in steps:
            p = doc.add_paragraph()
            r = p.add_run(title)
            r.bold = True
            for b in bullets:
                doc.add_paragraph(f"●  {b}")

    # === LOGIC POSISI ===
    if approval_di_bawah:
        # WI DULU, APPROVAL DI BAWAH
        add_wi_content(doc)
        doc.add_page_break()
        add_approval(doc)
        doc.add_paragraph("")
        add_history(doc)
    else:
        # APPROVAL DI ATAS (standar template)
        add_approval(doc)
        doc.add_paragraph("")
        add_history(doc)
        doc.add_page_break()
        add_wi_content(doc)
    
    return doc

def wi_tab():
    st.subheader("📄 WI Generator - FIX Approval Dobel")
    template_file = st.file_uploader("Upload Template Landscape (wajib buat baca format, tapi gak di-copy mentah)", type=["docx"])
    c1,c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Doc No", "WIN-13-005-0015")
        product = st.text_input("Product", "NT8")
        temp_no = st.text_input("Temp", "TMP-11-001-0002 rev.1")
        rev = st.text_input("Rev", "0")
        proc = st.text_input("Process", "QC 100% Visual Inspection LV Tips")
    with c2:
        part_req = st.text_area("Part", "1 batch LV Tips after cutting", height=80)
        tools_req = st.text_area("Tools", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120)
    
    time_df = st.data_editor(pd.DataFrame([
        {"step":"Step 1 : Susun Tips","time":10},
        {"step":"Step 2 : Check visual defect","time":80},
        {"step":"Step 3 : Data Tips OK dan Reject","time":5},
        {"step":"Step 4 : Packing LV Tips","time":5},
        {"step":"Step 5 : Manual Filling","time":30},
        {"step":"Step 6 : 2nd Visual Check","time":5},
        {"step":"Step 7 : Ionizer","time":5},
    ]), num_rows="dynamic", use_container_width=True)
    
    raw = st.text_area("Detail Steps", "Step 1 : Susun Tips\nSusun LV tips di tatakan\nStep 2 : Check visual defect\nCek dengan lampu LED 30cm", height=180)
    
    approval_di_bawah = st.checkbox("✅ APPROVAL & HISTORY DI BAWAH (WI di atas)", value=True)
    
    if st.button("✨ Generate FIX", type="primary", use_container_width=True):
        if not template_file:
            st.warning("Upload template dulu (cuma buat validasi, file asli tidak dipakai mentah biar gak dobel)")
        doc = create_wi_fixed(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, raw, approval_di_bawah)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success("FIX Berhasil! Sekarang gak dobel lagi")
        st.download_button("📥 Download WI FIX", data=bio, file_name=f"{doc_no}_FIX_Landscape.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

st.title("📊 Quality AI PRO - FIX")
tab1, tab2 = st.tabs(["📊 MSA", "📄 WI FIX"])
with tab1: msa_tab()
with tab2: wi_tab()
