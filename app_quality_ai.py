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
import difflib

st.set_page_config(page_title="Quality AI PRO + Comparator", page_icon="🔍", layout="wide")

# ================= MSA TAB =================
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

# ================= WI GENERATOR (LAMA) =================
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
    try:
        t2.style = 'Table Grid'
    except:
        pass
    t2.cell(0,0).text = "Langkah Kerja"
    t2.cell(0,1).text = "Time (min)"
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
    template_file = None
    found_name = None
    for f in os.listdir("."):
        if "TMP" in f and f.endswith(".docx"):
            found_name = f
            template_file = f
            break
    if template_file:
        st.success(f"✅ Template otomatis: {found_name} | Approval di BAWAH!")
    else:
        st.warning("Template belum ada, upload manual:")
        template_file = st.file_uploader("Upload Template (.docx)", type=["docx"])
    c1, c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Document No", "WIN-13-005-0015", key="docno_gen")
        product = st.text_input("Product", "NT8", key="prod_gen")
        temp_no = st.text_input("Temp Number", "TMP-11-001-0002 rev.1", key="temp_gen")
        rev = st.text_input("Rev", "0", key="rev_gen")
        proc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips", key="proc_gen")
    with c2:
        part_req = st.text_area("Part Required", "1 batch LV Tips after cutting", height=80, key="part_gen")
        tools_req = st.text_area("Tools Required", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120, key="tool_gen")
    time_df = st.data_editor(pd.DataFrame([
        {"step": "Step 1 :  Susun Tips", "time": 10},
        {"step": "Step 2 :  Check visual defect", "time": 80},
        {"step": "Step 3 :  Data Tips OK dan Tips Reject", "time": 5},
        {"step": "Step 4 :  Packing LV Tips", "time": 5},
        {"step": "Step 5 : Manual Filling Process", "time": 30},
        {"step": "Step 6 : 2nd Visual Check", "time": 5},
        {"step": "Step 7 : Ionizer", "time": 5},
    ]), num_rows="dynamic", use_container_width=True, key="time_gen")
    raw = st.text_area("Detail Langkah", "Step 1 :  Susun Tips\nSusun LV tips di atas tatakan\nStep 2 :  Check visual defect\nCek dengan lampu LED 30cm", height=180, key="raw_gen")
    if st.button("✨ Generate WI FINAL - Approval di Bawah", type="primary", use_container_width=True, key="btn_gen"):
        if not template_file:
            st.error("Template tidak ditemukan!")
            return
        doc = create_wi_final(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"WI {doc_no} jadi!")
        st.download_button("📥 Download WI FINAL", data=bio, file_name=f"{doc_no}_FINAL.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_gen")

# ================= NEW: WI COMPARATOR TAB =================
def extract_all_text(docx_file):
    doc = Document(docx_file)
    texts = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            texts.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    texts.append(f"[TABLE] {t}")
    return texts

def comparator_tab():
    st.subheader("🔍 WI Comparator - Cari Perubahan")
    st.caption("Upload WI Lama (Rev 0) vs WI Baru (Rev 1) - sistem akan highlight bedanya")
    
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA (Rev Lama)", type=["docx"], key="old_wi")
    with c2:
        new_file = st.file_uploader("📂 WI BARU (Rev Baru)", type=["docx"], key="new_wi")
    
    if old_file and new_file:
        old_texts = extract_all_text(old_file)
        new_texts = extract_all_text(new_file)
        
        # Hitung similarity
        matcher = difflib.SequenceMatcher(None, old_texts, new_texts)
        similarity = matcher.ratio() * 100
        
        st.divider()
        colA, colB, colC = st.columns(3)
        colA.metric("Kemiripan", f"{similarity:.1f}%")
        colB.metric("Baris Lama", len(old_texts))
        colC.metric("Baris Baru", len(new_texts))
        
        # Buat diff
        diff = list(difflib.ndiff(old_texts, new_texts))
        
        added = [d[2:] for d in diff if d.startswith("+ ")]
        removed = [d[2:] for d in diff if d.startswith("- ")]
        same = [d[2:] for d in diff if d.startswith("  ")]
        
        # Tampilkan summary
        st.subheader("📌 Ringkasan Perubahan")
        m1, m2, m3 = st.columns(3)
        m1.error(f"❌ Dihapus: {len(removed)} baris")
        m2.success(f"✅ Ditambah: {len(added)} baris")
        m3.info(f"➖ Tetap: {len(same)} baris")
        
        # Detail perubahan
        st.subheader("🔎 Detail Point yang Berubah")
        
        changes_data = []
        for line in diff:
            if line.startswith("- "):
                changes_data.append({"Status": "❌ DIHAPUS", "Text": line[2:], "Type": "Removed"})
            elif line.startswith("+ "):
                changes_data.append({"Status": "✅ DITAMBAH", "Text": line[2:], "Type": "Added"})
        
        if changes_data:
            df_changes = pd.DataFrame(changes_data)
            # Filter hanya yang penting (Step, Part, Tools, Time)
            st.dataframe(df_changes, use_container_width=True, height=400)
            
            # Highlight Step yang berubah
            st.divider()
            st.subheader("🚩 Step / Point Penting yang Berubah")
            important_changes = [c for c in changes_data if any(k in c["Text"].lower() for k in ["step", "susun", "check", "visual", "packing", "ionizer", "lampu", "pinset", "part required", "tools required", "time"])]
            if important_changes:
                for ch in important_changes[:20]: # limit 20
                    if "DIHAPUS" in ch["Status"]:
                        st.markdown(f"<div style='background:#ffcccc;padding:8px;border-radius:5px;margin:4px'>❌ {ch['Text']}</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div style='background:#ccffcc;padding:8px;border-radius:5px;margin:4px'>✅ {ch['Text']}</div>", unsafe_allow_html=True)
            else:
                st.info("Tidak ada perubahan di keyword penting, tapi ada perubahan teks lain di atas.")
            
            # Download laporan perubahan
            bio = io.BytesIO()
            report_doc = Document()
            report_doc.add_heading(f"Laporan Perbandingan WI - {datetime.now().strftime('%d/%m/%Y')}", 1)
            report_doc.add_paragraph(f"Kemiripan: {similarity:.1f}% | Dihapus: {len(removed)} | Ditambah: {len(added)}")
            report_doc.add_heading("Dihapus:", 2)
            for r in removed:
                report_doc.add_paragraph(f"- {r}", style='List Bullet')
            report_doc.add_heading("Ditambah:", 2)
            for a in added:
                report_doc.add_paragraph(f"+ {a}", style='List Bullet')
            report_doc.save(bio)
            bio.seek(0)
            st.download_button("📥 Download Laporan Perubahan (.docx)", data=bio, file_name=f"WI_Comparator_Report_{datetime.now().strftime('%Y%m%d')}.docx", use_container_width=True)
        else:
            st.success("🎉 Tidak ada perbedaan! Dokumen identik 100%")
    else:
        st.info("👆 Upload kedua file WI dulu untuk mulai bandingkan")

# ================= MAIN =================
st.title("📊 Quality AI PRO - MSA + WI + Comparator")
tab1, tab2, tab3 = st.tabs(["📊 MSA Dashboard", "📄 WI Generator", "🔍 WI Comparator"])
with tab1:
    msa_tab()
with tab2:
    wi_tab()
with tab3:
    comparator_tab()
