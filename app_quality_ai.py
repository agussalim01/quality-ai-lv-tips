import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import re
import os
import difflib
from datetime import datetime

st.set_page_config(page_title="Quality AI PRO + Comparator", page_icon="🔍", layout="wide")

# ================= HELPER FILTERED =================
def extract_all_text_filtered(docx_file):
    doc = Document(docx_file)
    texts = []
    ignore_keywords = ["department", "verified by", "approval document", "approved by", "fusa rnd", "tiga workflow", "history", "revision", "date", "description"]
    for p in doc.paragraphs:
        t = p.text.strip()
        if t and len(t) > 3:
            if any(k in t.lower() for k in ignore_keywords):
                continue
            if len(t) < 15 and ("WIN-" in t or "TMP-" in t):
                continue
            texts.append(t)
    for tbl in doc.tables:
        if not tbl.rows:
            continue
        first_cell = tbl.cell(0,0).text.lower() if tbl.rows else ""
        # Hanya ambil table Langkah Kerja, bukan table approval
        if "langkah kerja" in first_cell or "step" in first_cell.lower():
            for row in tbl.rows[1:]:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t and not any(k in t.lower() for k in ignore_keywords):
                        texts.append(t)
    return texts

def parse_steps_filtered(text_list):
    full_text = "\n".join(text_list)
    pattern = r"(Step\s*\d+\s*:.*?(?=Step\s*\d+\s*:|$))"
    matches = re.findall(pattern, full_text, flags=re.I | re.S)
    steps = {}
    for content in matches:
        content = content.strip()
        if len(content) < 10:
            continue
        m = re.search(r"Step\s*(\d+)", content, re.I)
        if m:
            key = f"Step {m.group(1)}"
            bullets = [b.strip() for b in re.split(r'\n|●|•', content) if b.strip() and len(b.strip())>3]
            bullets = [b for b in bullets if not any(k in b.lower() for k in ["department","verified","approval","approved by","fusa","tiga workflow"])]
            steps[key] = {"full": content, "bullets": bullets}
    return steps

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

def create_redline_docx(old_steps, new_steps):
    doc = Document()
    title = doc.add_heading(f"LAPORAN PERBANDINGAN WI - {datetime.now().strftime('%d/%m/%Y')}", 1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    all_keys = sorted(set(list(old_steps.keys()) + list(new_steps.keys())), key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999)
    changed = 0
    for k in all_keys:
        old = old_steps.get(k)
        new = new_steps.get(k)
        if not old and new:
            changed += 1
            doc.add_heading(f"{k} - BARU DITAMBAHKAN", 2)
            for b in new["bullets"]:
                p = doc.add_paragraph()
                run = p.add_run(f"+ {b}")
                run.font.color.rgb = RGBColor(0, 128, 0)
                run.bold = True
        elif old and not new:
            changed += 1
            doc.add_heading(f"{k} - DIHAPUS", 2)
            p = doc.add_paragraph()
            run = p.add_run(old["full"])
            run.font.strike = True
            run.font.color.rgb = RGBColor(255, 0, 0)
        elif old and new and old["full"].strip() != new["full"].strip():
            old_b = set(old["bullets"])
            new_b = set(new["bullets"])
            removed_b = old_b - new_b
            added_b = new_b - old_b
            if added_b or removed_b:
                changed += 1
                doc.add_heading(f"{k} - BERUBAH", 2)
                if removed_b:
                    p = doc.add_paragraph()
                    p.add_run("Dihapus/Diubah:").bold = True
                    for b in removed_b:
                        if len(b) > 4:
                            p = doc.add_paragraph()
                            run = p.add_run(f"- {b}")
                            run.font.strike = True
                            run.font.color.rgb = RGBColor(255, 0, 0)
                if added_b:
                    p = doc.add_paragraph()
                    p.add_run("Baru:").bold = True
                    for b in added_b:
                        if len(b) > 4:
                            p = doc.add_paragraph()
                            run = p.add_run(f"+ {b}")
                            run.font.color.rgb = RGBColor(0, 128, 0)
                            run.bold = True
    if changed == 0:
        doc.add_paragraph("Tidak ada perubahan - 100% identik")
    return doc, changed

def create_diff_image(old_steps, new_steps):
    from PIL import Image, ImageDraw
    W, H = 1000, 1600
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)
    y = 20
    draw.text((20, y), "PERBANDINGAN WI - Point Berubah", fill="black")
    y += 50
    all_keys = sorted(set(list(old_steps.keys()) + list(new_steps.keys())), key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999)
    for k in all_keys:
        old = old_steps.get(k)
        new = new_steps.get(k)
        if old and new and old["full"].strip() != new["full"].strip():
            old_b = set(old["bullets"]) - set(new["bullets"])
            new_b = set(new["bullets"]) - set(old["bullets"])
            if old_b or new_b:
                draw.rectangle([10, y, W-10, y+35], fill="#FFF9C4")
                draw.text((20, y+8), f"{k} - BERUBAH", fill="black")
                y += 45
                for b in old_b:
                    if len(b)>4:
                        draw.rectangle([20, y, W-20, y+30], fill="#FFCDD2")
                        draw.text((30, y+6), f"- {b[:85]}", fill="#B71C1C")
                        y += 35
                for b in new_b:
                    if len(b)>4:
                        draw.rectangle([20, y, W-20, y+30], fill="#C8E6C9")
                        draw.text((30, y+6), f"+ {b[:85]}", fill="#1B5E20")
                        y += 35
                y += 15
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio

# ================= TABS =================
def comparator_tab():
    st.subheader("🔍 WI Comparator - Tampil Langsung Gambar + Docs")
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA (Rev 0)", type=["docx"], key="old_wi")
    with c2:
        new_file = st.file_uploader("📂 WI BARU (Rev 1)", type=["docx"], key="new_wi")
    if old_file and new_file:
        old_texts = extract_all_text_filtered(old_file)
        new_texts = extract_all_text_filtered(new_file)
        old_steps = parse_steps_filtered(old_texts)
        new_steps = parse_steps_filtered(new_texts)
        sim = difflib.SequenceMatcher(None, old_texts, new_texts).ratio()*100
        st.metric("Kemiripan (Isi Kerja Saja, Approval Diabaikan)", f"{sim:.1f}%")
        
        # Gambar preview
        img_bio = create_diff_image(old_steps, new_steps)
        st.divider()
        st.subheader("🖼️ Preview Gambar (Screenshot Otomatis)")
        st.image(img_bio, caption="Merah = Dihapus, Hijau = Baru", use_container_width=True)
        
        # Preview HTML
        st.subheader("📄 Preview Dokumen - Point Berubah")
        has_change = False
        for k in sorted(set(list(old_steps.keys()) + list(new_steps.keys())), key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999):
            old = old_steps.get(k)
            new = new_steps.get(k)
            if not old and new:
                has_change = True
                st.success(f"✅ {k} BARU DITAMBAHKAN")
                st.markdown(f"<div style='background:#e8f5e9;border-left:5px solid green;padding:10px'>{new['full']}</div>", unsafe_allow_html=True)
            elif old and not new:
                has_change = True
                st.error(f"❌ {k} DIHAPUS")
            elif old and new and old["full"].strip() != new["full"].strip():
                rb = set(old["bullets"]) - set(new["bullets"])
                ab = set(new["bullets"]) - set(old["bullets"])
                if rb or ab:
                    has_change = True
                    with st.container(border=True):
                        st.markdown(f"**⚠️ {k} - BERUBAH**")
                        colL, colR = st.columns(2)
                        with colL:
                            st.caption("LAMA")
                            for b in rb:
                                if len(b)>4:
                                    st.markdown(f"<div style='background:#ffebee;padding:6px;border-radius:5px;margin:2px'>❌ {b}</div>", unsafe_allow_html=True)
                        with colR:
                            st.caption("BARU")
                            for b in ab:
                                if len(b)>4:
                                    st.markdown(f"<div style='background:#e8f5e9;padding:6px;border-radius:5px;margin:2px'>✅ {b}</div>", unsafe_allow_html=True)
        if not has_change:
            st.balloons()
            st.success("🎉 Tidak ada perubahan di Step Kerja! 100% identik")
        
        st.divider()
        redline_doc, changed = create_redline_docx(old_steps, new_steps)
        bio = io.BytesIO()
        redline_doc.save(bio)
        bio.seek(0)
        colA, colB = st.columns(2)
        with colA:
            st.download_button("📥 Download .DOCX Redline", data=bio, file_name=f"WI_Redline_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
        with colB:
            st.download_button("🖼️ Download Gambar PNG", data=img_bio, file_name=f"WI_Perbedaan_{datetime.now().strftime('%Y%m%d')}.png", mime="image/png", use_container_width=True)
    else:
        st.info("👆 Upload WI Lama & Baru dulu")

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
        doc_no = st.text_input("Document No", "WIN-13-005-0015")
        product = st.text_input("Product", "NT8")
        temp_no = st.text_input("Temp Number", "TMP-11-001-0002 rev.1")
        rev = st.text_input("Rev", "0")
        proc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips")
    with c2:
        part_req = st.text_area("Part Required", "1 batch LV Tips after cutting", height=80)
        tools_req = st.text_area("Tools Required", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120)
    time_df = st.data_editor(pd.DataFrame([{"step": "Step 1 :  Susun Tips", "time": 10},{"step": "Step 2 :  Check visual defect", "time": 80},{"step": "Step 3 :  Data Tips OK dan Tips Reject", "time": 5},{"step": "Step 4 :  Packing LV Tips", "time": 5},{"step": "Step 5 : Manual Filling Process", "time": 30},{"step": "Step 6 : 2nd Visual Check", "time": 5},{"step": "Step 7 : Ionizer", "time": 5}]), num_rows="dynamic", use_container_width=True)
    raw = st.text_area("Detail Langkah", "Step 1 :  Susun Tips\nSusun LV tips di atas tatakan\nStep 2 :  Check visual defect\nCek dengan lampu LED 30cm", height=180)
    if st.button("✨ Generate WI FINAL", type="primary", use_container_width=True):
        if not template_file:
            st.error("Template tidak ditemukan!")
            return
        doc = create_wi_final(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"WI {doc_no} jadi!")
        st.download_button("📥 Download WI FINAL", data=bio, file_name=f"{doc_no}_FINAL.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

def msa_tab():
    st.subheader("📊 MSA Dashboard")
    uploaded = st.file_uploader("Upload Data MSA (.csv)", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded)
    else:
        df = pd.DataFrame({"Part": [1,1,1,2,2,2,3,3,3]*3, "Operator": ["A"]*9 + ["B"]*9 + ["C"]*9, "Measurement": [10.1,10.12,10.11,10.2,10.22,10.21,10.0,10.02,10.01,10.11,10.13,10.12,10.21,10.23,10.22,10.01,10.03,10.02,10.09,10.11,10.10,10.19,10.21,10.20,9.99,10.01,10.0]})
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(df.head(15), use_container_width=True)
    with c2:
        fig = px.histogram(df, x="Measurement", color="Operator", title="MSA Distribution")
        st.plotly_chart(fig, use_container_width=True)

st.title("📊 Quality AI PRO")
tab1, tab2, tab3 = st.tabs(["📊 MSA", "📄 WI Generator", "🔍 WI Comparator - Gambar & Docs"])
with tab1:
    msa_tab()
with tab2:
    wi_tab()
with tab3:
    comparator_tab()
