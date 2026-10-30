import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import re
from datetime import datetime

st.set_page_config(page_title="Quality AI PRO - Landscape", page_icon="📄", layout="wide")

# ================= TAB 1: MSA =================
def msa_tab():
    st.subheader("📊 MSA Gage R&R Dashboard")
    uploaded = st.file_uploader("Upload Data MSA (.csv)", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded)
    else:
        df = pd.DataFrame({
            "Part": [1,1,1,2,2,2,3,3,3]*3,
            "Operator": ["A"]*9 + ["B"]*9 + ["C"]*9,
            "Measurement": [10.1,10.12,10.11,10.2,10.22,10.21,10.0,10.02,10.01,
                            10.11,10.13,10.12,10.21,10.23,10.22,10.01,10.03,10.02,
                            10.09,10.11,10.10,10.19,10.21,10.20,9.99,10.01,10.0]
        })
        st.caption("Data Demo")
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(df.head(15), use_container_width=True)
    with c2:
        fig = px.histogram(df, x="Measurement", color="Operator", title="MSA Distribution")
        st.plotly_chart(fig, use_container_width=True)
    st.success("MSA Dashboard Aktif")

# ================= TAB 2: WI LANDSCAPE PAKAI TEMPLATE ASLI =================
def create_wi_landscape(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, steps_text, approval_di_bawah=False):
    template_doc = Document(template_file)
    
    # Ambil 2 tabel dari template asli buat dipindah
    approval_table_data = []
    history_table_data = []
    try:
        # Table 0 = Approval
        for row in template_doc.tables[0].rows:
            approval_table_data.append([cell.text for cell in row.cells])
        # Table 1 = History
        for row in template_doc.tables[1].rows:
            history_table_data.append([cell.text for cell in row.cells])
    except:
        pass

    # Buat dokumen baru landscape
    doc = Document()
    # Set landscape
    section = doc.sections[0]
    section.orientation = 1 # LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    s = doc.styles['Normal']
    s.font.name = 'Arial'
    s.font.size = Pt(9)

    def add_approval_history(doc):
        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("APPROVAL DOCUMENT")
        r.bold = True
        
        t = doc.add_table(rows=len(approval_table_data), cols=4)
        try: t.style = 'Grid Table Light'
        except: pass
        for ri, row_data in enumerate(approval_table_data):
            for ci, txt in enumerate(row_data):
                if ci < 4:
                    t.cell(ri, ci).text = txt

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("DOCUMENT HISTORY")
        r.bold = True
        
        t2 = doc.add_table(rows=len(history_table_data), cols=5)
        try: t2.style = 'Grid Table Light'
        except: pass
        for ri, row_data in enumerate(history_table_data):
            for ci, txt in enumerate(row_data):
                if ci < 5:
                    t2.cell(ri, ci).text = txt
        # Update history row 2 dengan data baru
        try:
            t2.cell(2,1).text = datetime.now().strftime("%m/%d/%Y")
            t2.cell(2,2).text = rev
            t2.cell(2,3).text = f"{proc} - {product}"
        except: pass

    # === JIKA APPROVAL DI ATAS (default template) ===
    if not approval_di_bawah:
        add_approval_history(doc)

    # === ISI WI ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"WORK INSTRUCTION\n{proc}\nProduct: {product}\nDocument No: {doc_no}")
    run.bold = True
    run.font.size = Pt(14)

    p = doc.add_paragraph()
    run = p.add_run(f"Doc: {doc_no} | Temp: {temp_no} | Rev: {rev} | Date: {datetime.now().strftime('%d/%m/%Y')}")
    run.font.size = Pt(8)
    run.bold = True

    p = doc.add_paragraph()
    run = p.add_run("Langkah Kerja")
    run.bold = True
    run.font.size = Pt(11)

    t2 = doc.add_table(rows=1, cols=2)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    try: t2.style = 'Grid Table Light'
    except: pass
    t2.cell(0,0).text = "Langkah Kerja"
    t2.cell(0,1).text = "Time (min)"

    total = 0
    for _, row in time_df.iterrows():
        rw = t2.add_row().cells
        rw[0].text = str(row['step'])
        rw[1].text = str(row['time'])
        try: total += int(row['time'])
        except: pass
    rw = t2.add_row().cells
    rw[0].text = "Total"
    rw[1].text = str(total)

    doc.add_paragraph("")
    p = doc.add_paragraph()
    r = p.add_run("Part required :")
    r.bold = True
    for it in part_req.split('\n'):
        if it.strip(): doc.add_paragraph(f"●  {it.strip()}")

    p = doc.add_paragraph()
    r = p.add_run("Tools required :")
    r.bold = True
    for it in tools_req.split('\n'):
        if it.strip(): doc.add_paragraph(f"●  {it.strip()}")

    doc.add_paragraph("")
    pattern = r"(Step\s*\d+\s*:.*)"
    parts = re.split(pattern, steps_text, flags=re.I)
    steps = []
    for i in range(1, len(parts), 2):
        title = parts[i].strip()
        body = parts[i+1] if i+1 < len(parts) else ""
        bullets = [x.strip() for x in body.split('\n') if x.strip()]
        steps.append((title, bullets))
    if not steps:
        steps = [("Work Steps", [x for x in steps_text.split('\n') if x.strip()])]

    for title, bullets in steps:
        p = doc.add_paragraph()
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(10)
        for b in bullets:
            if b:
                clean = re.sub(r'^[•\-\*]\s*', '', b)
                doc.add_paragraph(f"●  {clean}")

    # === JIKA APPROVAL DI BAWAH ===
    if approval_di_bawah:
        doc.add_page_break()
        add_approval_history(doc)

    return doc

def wi_tab():
    st.subheader("📄 WI Generator - Template Landscape TMP-Common-General-0002")
    st.info("Template Landscape terdeteksi - Hasil akan Landscape plek perusahaan")

    template_file = st.file_uploader("1. Upload Template Landscape (.docx) - pakai file TMP-Common-General-0002", type=["docx"], key="templ")
    
    c1, c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Document No", "WIN-13-005-0015", key="doc2")
        product = st.text_input("Product", "NT8", key="prod2")
        temp_no = st.text_input("Temp Number", "TMP-11-001-0002 rev.1", key="temp2")
        rev = st.text_input("Rev", "0", key="rev2")
        proc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips", key="proc2")
    with c2:
        part_req = st.text_area("Part Required", "1 batch LV Tips after cutting", height=80, key="part2")
        tools_req = st.text_area("Tools Required", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120, key="tool2")

    time_df = st.data_editor(pd.DataFrame([
        {"step": "Step 1 :  Susun Tips", "time": 10},
        {"step": "Step 2 :  Check visual defect", "time": 80},
        {"step": "Step 3 :  Data Tips OK dan Tips Reject", "time": 5},
        {"step": "Step 4 :  Packing LV Tips", "time": 5},
        {"step": "Step 5 : Manual Filling Process", "time": 30},
        {"step": "Step 6 : 2nd Visual Check", "time": 5},
        {"step": "Step 7 : Ionizer", "time": 5},
    ]), num_rows="dynamic", use_container_width=True, key="time_land")

    raw = st.text_area("Detail Langkah (format Step X :)", 
"""Step 1 :  Susun Tips
Susun LV tips di atas wadah (tatakan) LV tips secara horizontal dengan menggunakan pinset.
Step 2 :  Check visual defect
Cek tips secara kasat mata dengan pencahayaan lampu LED dengan jarak 30cm
Cek visual bagian ujung sampai pangkal tips
Pastikan tidak terdapat bubble, sinkmark, penyok, short, flashing, kontaminasi
Step 3 :  Data Tips OK dan Tips Reject
Catat jumlah tips OK dan tips yang Reject hasil QC 100% di form QC 100%
Step 4 :  Packing LV Tips
Masukkan kembali tips yang OK ke dalam plastik ziplock sebelumnya.""", height=220, key="raw_land")

    if st.button("✨ Generate WI Landscape Final", type="primary", use_container_width=True):
        if not template_file:
            st.error("Upload dulu template landscape TMP-Common-General-0002 yang kamu kirim tadi!")
            return
        doc = create_wi_landscape(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"WI Landscape {doc_no} jadi! Format plek TMP-Common-General-0002")
        st.download_button("📥 Download WI Landscape (.docx)", data=bio, file_name=f"{doc_no}_Landscape_{product}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# MAIN
st.title("📊 Quality AI PRO - MSA + WIN Landscape")
tab1, tab2 = st.tabs(["📊 MSA Dashboard", "📄 WI Generator Landscape"])

with tab1:
    msa_tab()
with tab2:
    wi_tab()
