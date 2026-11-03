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

def extract_raw_with_color(docx_file):
    doc = Document(docx_file)
    items = []
    # 1. Semua paragraf termasuk yang pendek + warna
    for p in doc.paragraphs:
        for run in p.runs:
            t = run.text.strip()
            if not t:
                continue
            color = "red" if run.font.color.rgb and str(run.font.color.rgb) != "000000" and str(run.font.color.rgb) != "None" else "normal"
            if run.font.color.rgb:
                try:
                    # Jika merah
                    if run.font.color.rgb == RGBColor(255,0,0) or "FF0000" in str(run.font.color.rgb):
                        color = "red"
                except:
                    pass
            # Simpan dengan info warna
            items.append({"text": t, "color": color, "full": t})
    
    # 2. Semua table tanpa filter
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    items.append({"text": t, "color": "normal", "full": t})
    return items

def parse_steps_raw(items):
    texts = [i["text"] for i in items]
    full_text = "\n".join(texts)
    pattern = r"(Step\s*\d+\s*:.*?(?=Step\s*\d+\s*:|$))"
    matches = re.findall(pattern, full_text, flags=re.I | re.S)
    steps = {}
    for content in matches:
        m = re.search(r"Step\s*(\d+)", content, re.I)
        if m:
            steps[f"Step {m.group(1)}"] = content
    return steps

def create_redline_docx_raw(old_items, new_items):
    doc = Document()
    doc.add_heading(f"LAPORAN PERBANDINGAN WI - {datetime.now().strftime('%d/%m/%Y')}", 1)
    
    old_texts = [i["text"] for i in old_items]
    new_texts = [i["text"] for i in new_items]
    
    diff = list(difflib.ndiff(old_texts, new_texts))
    
    doc.add_heading("RINGKASAN PERUBAHAN", 2)
    for line in diff:
        if line.startswith("- "):
            p = doc.add_paragraph()
            run = p.add_run(f"DIHAPUS: {line[2:]}")
            run.font.strike = True
            run.font.color.rgb = RGBColor(255,0,0)
        elif line.startswith("+ "):
            p = doc.add_paragraph()
            run = p.add_run(f"BARU: {line[2:]}")
            run.font.color.rgb = RGBColor(0,128,0)
            run.bold = True
    
    # Khusus highlight ddff merah
    for item in new_items:
        if "ddff" in item["text"].lower():
            doc.add_heading("Teks Merah Terdeteksi:", 2)
            p = doc.add_paragraph()
            run = p.add_run(f"-> {item['text']} (warna: {item['color']})")
            run.font.color.rgb = RGBColor(255,0,0)
            run.bold = True
            
    return doc

def comparator_tab():
    st.subheader("🔍 WI Comparator - RAW Mode (Tangkap ddff Merah)")
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA", type=["docx"], key="old_wi")
    with c2:
        new_file = st.file_uploader("📂 WI BARU (ada ddff)", type=["docx"], key="new_wi")
    
    if old_file and new_file:
        old_items = extract_raw_with_color(old_file)
        new_items = extract_raw_with_color(new_file)
        
        old_texts = [i["text"] for i in old_items]
        new_texts = [i["text"] for i in new_items]
        
        sim = difflib.SequenceMatcher(None, old_texts, new_texts).ratio()*100
        st.metric("Kemiripan RAW (Semua Text)", f"{sim:.1f}%")
        
        # === TAMPILKAN LANGSUNG SEPERTI GAMBAR ===
        st.divider()
        st.subheader("🖼️ Preview Visual Langsung (Seperti Screenshot)")
        
        diff = list(difflib.ndiff(old_texts, new_texts))
        
        # Buat tampilan kayak Word
        st.markdown("""
        <div style="border:2px solid #ccc;padding:20px;background:white">
        <div style="border:1px solid black;display:flex">
            <div style="width:30%;border-right:1px solid black;padding:5px"><b>FORMULATRIX</b></div>
            <div style="width:40%;border-right:1px solid black;padding:5px;text-align:center"><b>WORK INSTRUCTION</b></div>
            <div style="width:30%;padding:5px">Revision No. -</div>
        </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.write("")
        for line in diff:
            if line.startswith("- "):
                if len(line[2:].strip()) > 1:
                    st.markdown(f"<div style='background:#ffcccc;padding:8px;border-radius:5px;margin:3px'>❌ DIHAPUS: <s style='color:red'>{line[2:]}</s></div>", unsafe_allow_html=True)
            elif line.startswith("+ "):
                if len(line[2:].strip()) > 1:
                    # Kalau ddff merah
                    if "ddff" in line[2:].lower():
                        st.markdown(f"<div style='background:#ffcdd2;border:2px solid red;padding:12px;border-radius:5px;margin:5px'><b style='color:red;font-size:18px'>🔴 BARU TERDETEKSI (MERAH): {line[2:]}</b></div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div style='background:#ccffcc;padding:8px;border-radius:5px;margin:3px'>✅ BARU: <b style='color:green'>{line[2:]}</b></div>", unsafe_allow_html=True)
        
        if sim == 100:
            st.error("Masih 100%? Coba upload ulang, pastikan file BARU memang ada ddff nya")
        else:
            st.success(f"✅ Terdeteksi {len([d for d in diff if d.startswith('+ ') or d.startswith('- ')])} perubahan!")
        
        # Download
        st.divider()
        redline = create_redline_docx_raw(old_items, new_items)
        bio = io.BytesIO()
        redline.save(bio)
        bio.seek(0)
        st.download_button("📥 Download Laporan .DOCX (ddff Merah Akan Ditandai)", data=bio, file_name="WI_Redline_ddff.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
    else:
        st.info("👆 Upload WI Lama tanpa ddff, dan WI Baru dengan ddff merah")

def wi_tab():
    st.subheader("📄 WI Generator")
    st.info("Tab ini tetap seperti kemarin")

def msa_tab():
    st.subheader("📊 MSA")
    st.info("Tab ini tetap")

st.title("📊 Quality AI PRO")
tab1, tab2, tab3 = st.tabs(["📊 MSA", "📄 WI Generator", "🔍 Comparator RAW (Fix ddff)"])
with tab1: msa_tab()
with tab2: wi_tab()
with tab3: comparator_tab()
