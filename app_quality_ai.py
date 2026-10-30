import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import re
from datetime import datetime
import os

st.set_page_config(page_title="Quality AI PRO - Auto Template", page_icon="📄", layout="wide")

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

def create_wi_final(template_path_or_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, steps_text):
    doc = Document(template_path_or_file)
    saved_tables = list(doc.tables)
    for tbl in saved_tables:
        try:
            tbl._tbl.getparent().remove(tbl._tbl)
        except:
            pass
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"WORK INSTRUCTION\n{proc}\nProduct: {product}\nDocument No: {doc_no}")
    run.bold = True
    run.font.size = Pt(14)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Doc: {doc_no} | Temp: {temp_no} | Rev: {rev} | Date: {datetime.now().strftime('%d/%m/%Y')}")
    run.font.size = Pt(8)
    run.bold = True
    p = doc.add_paragraph()
    run = p.add_run("Langkah Kerja")
    run.bold = True
    run.font.size = Pt(11)
    t2 = doc.add_table(rows=1, cols=2)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.style = 'Table Grid'
    t2.cell(0,0).text = "Langkah Kerja"
    t2.cell(0,1).text = "Time (min)"
    t2.cell(0,0).paragraphs[0].runs[0].bold = True
    t2.cell(0,1).paragraphs[0].runs[0].bold = True
    total = 0
    for _, row in time_df.iterrows():
        rw = t2.add_row().cells
        rw[0].text = str(row['step'])
        rw[1].text = str(row['time'])
        try:
            total += int(row['time'])
        except:
            pass
    rw = t2.add_row().cells
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
        p.add_run(title).bold = True
        for b in bullets:
            if b:
                clean = re.sub(r'^[•\-\*]\s*', '', b)
                doc.add_paragraph(f"●  {clean}")
    doc.add_page_break()
    for tbl in saved_tables:
        doc._body._element.append(tbl._tbl)
        doc.add_paragraph("")
    try:
        hist_table = doc.tables[-1]
        hist_table.cell(2,1).text = datetime.now().strftime("%m/%d/%Y")
        hist_table.cell(2,2).text = rev
        hist_table.cell(2,3).text = f"{proc} - {product}"
    except:
        pass
    return doc

def wi_tab():
    st.subheader("📄 WI Generator - Auto Template")
    auto_template_path = "TMP-Common-General-0002-Template WI Landscape.docx"
    # Cek semua kemungkinan nama file template
    possible_names = [
        "TMP-Common-General-0002-Template WI Landscape.docx",
        "TMP-Common-General-0002-Template_WI_Landscape.docx",
        "TMP-Common-General-0002-Template WI Lan.docx"
    ]
    template_file = None
    found_name = None
    for name in possible_names:
        if os.path.exists(name):
            found_name = name
            template_file = name
            break
    # Juga cek file yang mengandung TMP
    if not template_file:
        for f in os.listdir("."):
            if "TMP" in f and f.endswith(".docx"):
                found_name = f
                template_file = f
                break
    
    if template_file:
        st.success(f"✅ Template otomatis ditemukan: {found_name} - Logo tetap, Approval & History di BAWAH!")
    else:
        st.warning("Template belum terdeteksi. Upload manual:")
        template_file = st.file_uploader("Upload Template (.docx)", type=["docx"])

    c1, c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Document No", "WIN-13-005-0015")
        product = st.text_input("Product", "NT8")
        temp_no = st.text_input("Temp Number", "TMP-11-001-0002 rev.1")
        rev = st.text_input("Rev", "0")
        proc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips")
    with c2:
        part_req = st.text_area("Part Required", "1 batch LV Tips after cutting", height=80)
        tools_req = st.text_area("Tools Required", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120)
    time_df = st.data_editor(pd.DataFrame([
        {"step": "Step 1 :  Susun Tips", "time": 10},
        {"step": "Step 2 :  Check visual defect", "time": 80},
        {"step": "Step 3 :  Data Tips OK dan Tips Reject", "time": 5},
        {"step": "Step 4 :  Packing LV Tips", "time": 5},
        {"step": "Step 5 : Manual Filling Process", "time": 30},
        {"step": "Step 6 : 2nd Visual Check", "time": 5},
        {"step": "Step 7 : Ionizer", "time": 5},
    ]), num_rows="dynamic", use_container_width=True)
    raw = st.text_area("Detail Langkah", "Step 1 :  Susun Tips\nSusun LV tips di atas tatakan\nStep 2 :  Check visual defect\nCek dengan lampu LED 30cm\nStep 3 :  Data Tips OK dan Tips Reject\nCatat jumlah tips OK dan Reject\nStep 4 :  Packing LV Tips\nMasukkan kembali tips OK ke plastik", height=220)

    if st.button("✨ Generate WI FINAL - Approval di Bawah", type="primary", use_container_width=True):
        if not template_file:
            st.error("Template tidak ditemukan!")
            return
        doc = create_wi_final(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"WI {doc_no} jadi! Logo tetap, WI di atas, Approval & History di BAWAH!")
        st.download_button("📥 Download WI FINAL", data=bio, file_name=f"{doc_no}_FINAL_ApprovalBawah_{product}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

st.title("📊 Quality AI PRO - MSA + WI Auto Template")
tab1, tab2 = st.tabs(["📊 MSA Dashboard", "📄 WI Generator"])
with tab1:
    msa_tab()
with tab2:
    wi_tab()
