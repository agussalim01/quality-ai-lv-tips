import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io, re, os, difflib
from datetime import datetime

st.set_page_config(page_title="Quality AI PRO + Comparator", page_icon="🔍", layout="wide")

def extract_all_text(docx_file):
    doc = Document(docx_file)
    texts = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t and len(t) > 2:
            texts.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and len(t) > 1:
                    texts.append(t)
    return texts

def parse_steps(text_list):
    full_text = "\n".join(text_list)
    # Ambil Step 1 : xxx sampai Step berikutnya
    pattern = r"(Step\s*\d+\s*:.*?(?=Step\s*\d+\s*:|$))"
    matches = re.findall(pattern, full_text, flags=re.I | re.S)
    steps = {}
    for content in matches:
        content = content.strip()
        m = re.search(r"Step\s*(\d+)", content, re.I)
        if m:
            key = f"Step {m.group(1)}"
            # ambil bullets
            bullets = [b.strip() for b in re.split(r'\n|●|•', content) if b.strip()]
            steps[key] = {"full": content, "bullets": bullets}
    return steps

def comparator_tab():
    st.subheader("🔍 WI Comparator - Hanya Tampilkan Yang Berubah")
    
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA", type=["docx"], key="old_wi")
    with c2:
        new_file = st.file_uploader("📂 WI BARU", type=["docx"], key="new_wi")

    if old_file and new_file:
        old_texts = extract_all_text(old_file)
        new_texts = extract_all_text(new_file)
        old_steps = parse_steps(old_texts)
        new_steps = parse_steps(new_texts)

        matcher = difflib.SequenceMatcher(None, old_texts, new_texts)
        similarity = matcher.ratio()*100

        st.metric("Kemiripan Total Dokumen", f"{similarity:.1f}%")
        
        # ===== LANGSUNG TAMPILKAN POINT BERUBAH =====
        st.divider()
        st.subheader("🎯 POINT YANG BERUBAH")

        all_keys = sorted(set(list(old_steps.keys()) + list(new_steps.keys())), 
                          key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999)
        
        changed_list = []
        
        for k in all_keys:
            old = old_steps.get(k)
            new = new_steps.get(k)
            
            if not old and new:
                changed_list.append(k)
                st.success(f"✅ {k} - BARU DITAMBAHKAN")
                for b in new["bullets"][1:]: # skip title
                    st.markdown(f"<div style='background:#c8e6c9;padding:6px;margin:3px;border-radius:5px'>+ {b}</div>", unsafe_allow_html=True)
                st.write("")
                
            elif old and not new:
                changed_list.append(k)
                st.error(f"❌ {k} - DIHAPUS TOTAL")
                st.markdown(f"<div style='background:#ffcdd2;padding:6px;border-radius:5px'>- {old['full'][:400]}</div>", unsafe_allow_html=True)
                st.write("")
                
            elif old and new and old["full"].strip() != new["full"].strip():
                # Cari bullet mana yang berubah
                old_bullets = set(old["bullets"])
                new_bullets = set(new["bullets"])
                removed_b = old_bullets - new_bullets
                added_b = new_bullets - old_bullets
                
                if added_b or removed_b:
                    changed_list.append(k)
                    st.warning(f"⚠️ {k} - ADA POINT BERUBAH")
                    colL, colR = st.columns(2)
                    with colL:
                        if removed_b:
                            st.caption("❌ Point Lama (dihapus/diubah):")
                            for b in removed_b:
                                if len(b) > 5:
                                    st.markdown(f"<div style='background:#ffcccc;padding:5px;margin:2px;border-radius:4px;font-size:13px'>- {b}</div>", unsafe_allow_html=True)
                    with colR:
                        if added_b:
                            st.caption("✅ Point Baru:")
                            for b in added_b:
                                if len(b) > 5:
                                    st.markdown(f"<div style='background:#ccffcc;padding:5px;margin:2px;border-radius:4px;font-size:13px'>+ {b}</div>", unsafe_allow_html=True)
                    st.divider()

        if not changed_list:
            st.balloons()
            st.success("🎉 Tidak ada perubahan sama sekali! WI 100% identik.")
        else:
            st.info(f"Total {len(changed_list)} Step yang berubah: {', '.join(changed_list)}")
            
            # Export laporan hanya yang berubah
            bio = io.BytesIO()
            rep = Document()
            rep.add_heading(f"Laporan Perubahan WI - {datetime.now().strftime('%d/%m/%Y')}", 1)
            rep.add_paragraph(f"Step Berubah: {', '.join(changed_list)}")
            for k in changed_list:
                rep.add_heading(k, 2)
                if k in old_steps and k in new_steps:
                    rep.add_paragraph(f"LAMA: {old_steps[k]['full']}")
                    rep.add_paragraph(f"BARU: {new_steps[k]['full']}")
                elif k in new_steps:
                    rep.add_paragraph(f"BARU DITAMBAH: {new_steps[k]['full']}")
                else:
                    rep.add_paragraph(f"DIHAPUS: {old_steps[k]['full']}")
            rep.save(bio)
            bio.seek(0)
            st.download_button("📥 Download Laporan Hanya Yang Berubah", data=bio, file_name="WI_Perubahan_Only.docx", use_container_width=True)

    else:
        st.info("👆 Upload WI Lama & Baru dulu, nanti langsung keluar point yang berubah aja")

# ==== MSA + WI TAB (singkat) ====
def wi_tab():
    st.subheader("📄 WI Generator")
    st.caption("Fitur lama tetap jalan")
    # ... (pakai fungsi create_wi_final kamu yang kemarin ya, biar gak kepanjangan) ...
    st.success("Tab WI Generator ready - pakai template auto seperti kemarin")

def msa_tab():
    st.subheader("📊 MSA")
    st.success("MSA Dashboard ready")

st.title("📊 Quality AI PRO")
tab1, tab2, tab3 = st.tabs(["📊 MSA", "📄 WI Generator", "🔍 WI Comparator - Point Berubah"])
with tab1: msa_tab()
with tab2: wi_tab()
with tab3: comparator_tab()
