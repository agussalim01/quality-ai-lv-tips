import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, re, difflib, os
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title="Quality AI PRO", layout="wide")

def extract_all(docx_file):
    doc = Document(docx_file)
    items = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            items.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and len(t) > 1:
                    items.append(t)
    return items

def create_side_screenshot(old_list, new_list, output_path="/tmp/side_by_side.png"):
    W, H = 1600, 1200
    img = Image.new("RGB", (W, H), "#f0f0f0")
    draw = ImageDraw.Draw(img)
    
    try:
        font_title = ImageFont.truetype("arial.ttf", 22)
        font_body = ImageFont.truetype("arial.ttf", 16)
    except:
        font_title = ImageFont.load_default()
        font_body = ImageFont.load_default()

    # Background 2 kertas
    # Kiri
    draw.rectangle([10, 10, W//2-10, H-10], fill="white", outline="#999", width=2)
    draw.rectangle([W//2+10, 10, W-10, H-10], fill="white", outline="#999", width=2)
    
    # Judul
    draw.rectangle([10, 10, W//2-10, 50], fill="#263238")
    draw.text((20, 15), "WI LAMA", fill="white", font=font_title)
    draw.rectangle([W//2+10, 10, W-10, 50], fill="#2e7d32")
    draw.text((W//2+20, 15), "WI BARU (ada ddff)", fill="white", font=font_title)
    
    y_left = 70
    y_right = 70
    
    # Bandingkan
    matcher = difflib.SequenceMatcher(None, old_list, new_list)
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if y_left > H-50 or y_right > H-50:
            break
            
        if tag == 'equal':
            for k in range(i1, min(i2, i1+8)): # tampil max 8 baris yang sama
                txt = old_list[k][:60]
                draw.text((20, y_left), f"{txt}", fill="black", font=font_body)
                draw.text((W//2+20, y_right), f"{txt}", fill="black", font=font_body)
                y_left += 22
                y_right += 22
        elif tag == 'replace' or tag == 'delete' or tag == 'insert':
            # Old yang dihapus - merah
            if tag in ('replace', 'delete'):
                for k in range(i1, i2):
                    if y_left > H-60: break
                    txt = old_list[k][:55]
                    draw.rectangle([15, y_left-2, W//2-15, y_left+20], fill="#ffcdd2")
                    draw.text((20, y_left), f"- {txt}", fill="#b71c1c", font=font_body)
                    y_left += 26
            # New yang baru - hijau/merah kalau ddff
            if tag in ('replace', 'insert'):
                for k in range(j1, j2):
                    if y_right > H-60: break
                    txt = new_list[k][:55]
                    is_ddff = "ddff" in txt.lower()
                    bg = "#ff0000" if is_ddff else "#c8e6c9"
                    fg = "white" if is_ddff else "#1b5e20"
                    draw.rectangle([W//2+15, y_right-2, W-15, y_right+20], fill=bg)
                    draw.text((W//2+20, y_right), f"+ {txt}", fill=fg, font=font_body)
                    y_right += 26
            y_left += 10
            y_right += 10
            # Samakan tinggi biar sejajar
            max_y = max(y_left, y_right)
            y_left = max_y
            y_right = max_y
    
    img.save(output_path)
    return output_path

def comparator_tab():
    st.subheader("🔍 Comparator - Tampilan Screenshot Kiri Kanan")
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA", type=["docx"], key="old_wi")
    with c2:
        new_file = st.file_uploader("📂 WI BARU", type=["docx"], key="new_wi")
    
    if old_file and new_file:
        old_list = extract_all(old_file)
        new_list = extract_all(new_file)
        
        sim = difflib.SequenceMatcher(None, old_list, new_list).ratio()*100
        st.metric("Kemiripan", f"{sim:.1f}%")
        
        # Generate screenshot
        img_path = create_side_screenshot(old_list, new_list)
        
        st.divider()
        st.subheader("🖼️ Preview Screenshot - Kiri LAMA vs Kanan BARU")
        st.caption("Ini preview gambar persis kayak buka 2 Word berdampingan")
        st.image(img_path, use_container_width=True)
        
        # Download gambar
        with open(img_path, "rb") as f:
            st.download_button("📥 Download Screenshot Kiri-Kanan (PNG)", data=f, file_name="WI_Kiri_Kanan.png", mime="image/png", type="primary", use_container_width=True)
        
        # Juga tampilkan detail text
        st.divider()
        st.subheader("📋 Detail Perubahan")
        diff = list(difflib.ndiff(old_list, new_list))
        for line in diff:
            if line.startswith("+ ") and "ddff" in line.lower():
                st.error(f"🔴 {line}")
            elif line.startswith("+ ") or line.startswith("- "):
                if len(line[2:].strip()) > 2:
                    st.write(line)
    else:
        st.info("👆 Upload 2 file WI dulu")

def wi_tab():
    st.subheader("📄 WI Generator - Auto Template")
    st.info("Ready - pakai template TMP-11-001-0002")

def msa_tab():
    st.subheader("📊 MSA Dashboard")
    uploaded = st.file_uploader("Upload CSV", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded)
        st.dataframe(df.head())

st.title("📊 Quality AI PRO")
tab1, tab2, tab3 = st.tabs(["📊 MSA", "📄 WI Generator", "🖼️ Screenshot Kiri Kanan"])
with tab1: msa_tab()
with tab2: wi_tab()
with tab3: comparator_tab()
