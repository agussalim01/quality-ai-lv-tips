import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io, re, os, difflib
from datetime import datetime

st.set_page_config(page_title="Quality AI PRO + Comparator", page_icon="🔍", layout="wide")

def clean_text(t):
    # Bersihkan bullet dan spasi biar ● 1 batch == 1 batch
    t = re.sub(r'^[●•\-\*\d\.\s]+', '', t)
    t = t.strip()
    return t

def extract_all(docx_file):
    doc = Document(docx_file)
    items = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t and len(t) >= 2:
            items.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and len(t) >= 2:
                    # Skip header yang sama persis
                    if t.lower() in ["work instruction", "revision no.", "document no.", "product", "temp. number", "release date", "process desc."]:
                        continue
                    if "formulatrix" in t.lower() and len(t) < 30:
                        continue
                    items.append(t)
    return items

def create_side_by_side_html(old_list, new_list):
    # Normalisasi untuk compare
    old_clean = [clean_text(x) for x in old_list]
    new_clean = [clean_text(x) for x in new_list]
    
    # Cari yang bener-bener beda
    sm = difflib.SequenceMatcher(None, old_clean, new_clean)
    html = ""
    changes = 0
    
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            continue
        changes += 1
        if tag == 'delete':
            for k in range(i1, i2):
                if old_list[k].strip() and len(clean_text(old_list[k])) > 2:
                    html += f"<tr style='background:#ffebee'><td style='padding:8px;color:red'>❌ DIHAPUS</td><td style='padding:8px'><s>{old_list[k]}</s></td><td style='padding:8px;color:#ccc'>-</td></tr>"
        elif tag == 'insert':
            for k in range(j1, j2):
                if new_list[k].strip() and len(clean_text(new_list[k])) > 2:
                    is_ddff = "ddff" in new_list[k].lower()
                    bg = "#ffcdd2" if is_ddff else "#e8f5e9"
                    color = "red" if is_ddff else "green"
                    icon = "🔴" if is_ddff else "✅"
                    html += f"<tr style='background:{bg}'><td style='padding:8px;color:{color}'>{icon} BARU</td><td style='padding:8px;color:#ccc'>-</td><td style='padding:8px;color:{color}'><b>{new_list[k]}</b></td></tr>"
        elif tag == 'replace':
            # Old vs New
            old_chunk = "<br>".join(old_list[i1:i2])
            new_chunk = "<br>".join(new_list[j1:j2])
            # Kalau setelah dibersihkan sama, skip
            if "".join(old_clean[i1:i2]).lower() == "".join(new_clean[j1:j2]).lower():
                continue
            html += f"<tr style='background:#fff9c4'><td style='padding:8px;color:orange'>⚠️ DIUBAH</td><td style='padding:8px;background:#ffebee'><s style='color:red'>{old_chunk}</s></td><td style='padding:8px;background:#e8f5e9;color:green'><b>{new_chunk}</b></td></tr>"
    
    return html, changes

def comparator_tab():
    st.subheader("🔍 WI Comparator - Side by Side (Fix Berantakan)")
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA", type=["docx"], key="old_wi")
    with c2:
        new_file = st.file_uploader("📂 WI BARU (ada ddff)", type=["docx"], key="new_wi")
    
    if old_file and new_file:
        old_list = extract_all(old_file)
        new_list = extract_all(new_file)
        
        sim = difflib.SequenceMatcher(None, [clean_text(x) for x in old_list], [clean_text(x) for x in new_list]).ratio()*100
        st.metric("Kemiripan (Sudah Bersih, Tanpa Bullet)", f"{sim:.1f}%")
        
        html, changes = create_side_by_side_html(old_list, new_list)
        
        if changes == 0:
            st.balloons()
            st.success("🎉 Tidak ada perbedaan isi kerja!")
        else:
            st.warning(f"🔍 Ditemukan {changes} group perubahan")
        
        st.divider()
        st.subheader("📄 Preview Perbandingan - Kiri LAMA | Kanan BARU")
        st.markdown(f"""
        <div style="overflow-x:auto;background:white;border:1px solid #ddd;border-radius:8px">
        <table style="width:100%;border-collapse:collapse;font-size:14px">
            <tr style="background:#263238;color:white">
                <th style="padding:10px;width:15%">Status</th>
                <th style="padding:10px;width:42%">WI LAMA</th>
                <th style="padding:10px;width:42%">WI BARU</th>
            </tr>
            {html if html else "<tr><td colspan=3 style='padding:20px;text-align:center;color:green'>✅ 100% Sama - Tidak ada ddff atau perubahan</td></tr>"}
        </table>
        </div>
        """, unsafe_allow_html=True)
        
        # Highlight khusus ddff
        if any("ddff" in x.lower() for x in new_list):
            st.divider()
            st.error("🔴 PERHATIAN: Teks 'ddff' warna merah terdeteksi di WI BARU!")
            for x in new_list:
                if "ddff" in x.lower():
                    st.markdown(f"<div style='background:red;color:white;padding:15px;border-radius:8px;font-size:20px;text-align:center'><b>ddff</b> ditemukan!</div>", unsafe_allow_html=True)
        
        # Download
        st.divider()
        doc = Document()
        doc.add_heading("Laporan Perbandingan WI", 1)
        doc.add_paragraph(f"Kemiripan: {sim:.1f}% | Perubahan: {changes} group")
        for line in html.split("<tr"):
            doc.add_paragraph(line[:200])
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.download_button("📥 Download Laporan .DOCX", data=bio, file_name="WI_Comparison_SideBySide.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
    else:
        st.info("👆 Upload 2 WI dulu - preview akan langsung side-by-side kayak Word")

def wi_tab():
    st.subheader("📄 WI Generator")
    st.info("Ready")

def msa_tab():
    st.subheader("📊 MSA")
    st.info("Ready")

st.title("📊 Quality AI PRO")
tab1, tab2, tab3 = st.tabs(["📊 MSA", "📄 WI Generator", "🔍 Comparator Side-by-Side"])
with tab1: msa_tab()
with tab2: wi_tab()
with tab3: comparator_tab()
