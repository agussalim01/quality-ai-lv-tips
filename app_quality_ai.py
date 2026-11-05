# app_quality_ai.py - V6 FINAL FIXED V3 - FULL 7 TABS WITH DYNAMIC CATALOG & COMPARE
import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import io
from datetime import datetime
import difflib
import requests

# Safe Import AI Helper
AI_AVAILABLE = False
try:
    import ai_helper
    from ai_helper import analyze_wi_text, analyze_trial_screenshot
    AI_AVAILABLE = True
except Exception:
    pass

st.set_page_config(page_title="Quality AI V6 FINAL V3 HEADER REAL", page_icon="📊", layout="wide")

# =================== CACHING & UTILS ===================
@st.cache_data(show_spinner=False)
def extract_file_content(file_bytes, file_name):
    try:
        if file_name.lower().endswith('.docx'):
            doc = Document(io.BytesIO(file_bytes))
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ''])
        else:
            import fitz
            pdf = fitz.open(stream=file_bytes, filetype="pdf")
            txt = ""
            for pg in pdf:
                txt += pg.get_text() + "\n"
            return txt
    except Exception as e:
        return str(e)

def generate_side_by_side_html(old_text, new_text):
    """Menghasilkan HTML bersisian (Kiri = OLD, Kanan = NEW) ala Google Docs"""
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    
    old_html = []
    new_html = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            word_str = " ".join(old_words[i1:i2])
            old_html.append(word_str)
            new_html.append(word_str)
        elif tag == 'delete':
            deleted = " ".join(old_words[i1:i2])
            old_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
        elif tag == 'insert':
            inserted = " ".join(new_words[j1:j2])
            new_html.append(f'<span style="color: red; font-weight: bold; background-color: #dafbe1; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
        elif tag == 'replace':
            deleted = " ".join(old_words[i1:i2])
            inserted = " ".join(new_words[j1:j2])
            old_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
            new_html.append(f'<span style="color: red; font-weight: bold; background-color: #dafbe1; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
            
    return " ".join(old_html), " ".join(new_html)

def safe_join(row):
    vals = []
    for v in row.values:
        try:
            if pd.isna(v): continue
            s = str(v).strip()
            if s.lower() not in ['nan', '']: vals.append(s)
        except: continue
    return ' '.join(vals)

def universal_parser(file_or_link):
    result = {"cav_df": pd.DataFrame(), "pareto": pd.DataFrame(), "define": "", "measure": "", "analyze": "", "improve": "", "control": "", "info": ""}
    try:
        xls = None
        if isinstance(file_or_link, str) and 'docs.google.com' in file_or_link:
            fid = file_or_link.split('/d/')[1].split('/')[0]
            url = f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx"
            xls = pd.ExcelFile(url)
            result["info"] = "Link OK; "
        elif hasattr(file_or_link, 'name'):
            xls = pd.ExcelFile(file_or_link)
        else: return result
        
        for sh in xls.sheet_names:
            df0 = xls.parse(sh, header=None)
            txt = df0.apply(safe_join, axis=1).str.cat(sep=' ').lower()
            if 'define' in txt:
                for _, r in df0.iterrows():
                    rt = safe_join(r)
                    if len(rt) < 12: continue
                    if 'flashing' in rt.lower(): result["define"] += rt + "\n"
                    if 'actual condition' in rt.lower(): result["measure"] += rt + "\n"
                    if 'mechanism' in rt.lower(): result["analyze"] += rt + "\n"
                    if 'main solution' in rt.lower(): result["improve"] += rt + "\n"
                    if 'routine monitoring' in rt.lower(): result["control"] += rt + "\n"
        for sh in xls.sheet_names:
            df0 = xls.parse(sh, header=None)
            idx = -1
            for i, r in df0.iterrows():
                if 'Cav 1' in safe_join(r): idx = i; break
            if idx >= 0:
                df_cav = xls.parse(sh, header=idx)
                df_cav = df_cav.loc[:, ~df_cav.columns.astype(str).str.contains('Unnamed')]
                result["cav_df"] = df_cav
                cav_cols = [c for c in df_cav.columns if 'Cav' in str(c)]
                if cav_cols:
                    m = df_cav.melt(value_vars=cav_cols, var_name='Cavity', value_name='Result')
                    m['Result'] = m['Result'].astype(str).str.upper()
                    ng = m[m['Result'].str.contains('NG|FLASH')]
                    result["pareto"] = ng.groupby('Cavity').size().reset_index(name='Qty') if not ng.empty else pd.DataFrame([{"Cavity": "All OK", "Qty": 0}])
                break
        return result
    except Exception as e:
        result["info"] = f"{e}"
        return result

# =================== MASTER DATA ===================
MASTER_PROJECTS = {
    "FAST": {"parts": ["LV Pippete tip", "HV PIppete Tip", "Plunger LV", "LV Caddy", "Cover LV", "LV Plunger A - 001434C", "LV Plunger B - 001434C", "LV Plunger C - 001434C", "LV Plunger D - 001434C", "HV Caddy", "Plunger HV"], "mesin": "SUMITOMO 50T/100T/130T, JSW 80T/100T/220T"},
    "Flo PD (FLo PD)": {"parts": ["Caddy Cover", "Tip HV", "Tip LV", "Base Caddy", "Top Caddy", "Plunger LV", "Plunger HV", "Plunger Pos.Ring", "Sleeve"], "mesin": "JSW 100T A/B"},
    "Mantis": {"parts": ["Tube support", "Bottle Cap Top", "Bottle Cap lower", "Cap IMHVS_PI_AVA", "Air Layer IMHVS_PI_AVA", "Thin Layer IMHVS_PI_AVA", "Thick Layer IMHVS_PI_AVA", "Cap Mantis IMLVS_PI_AVA", "Air Layer IMLVS_PI_AVA", "Thin Layer Mantis IMLVS_PI_AVA", "Thick Layer Mantis IMLVS_PI_AVA", "Thin Layer IMCFPFE_PI", "Air Layer IMCFPFE_PI", "Thick Layer IMCFPFE_PI", "Thin Liquid Layer IMHV3PFE Dome", "Thick Liquid Layer IMHV3PFE Dome"], "mesin": "ARBURG 70T LSR"},
    "Mantis LVDS": {"parts": ["Air layer V1", "Air layer V2", "Thin Liquid Layer V1", "Thin Liquid Layer V2", "Thick Liquid Layer V1", "Thick Liquid Layer V2"], "mesin": "Mantis LVDS"},
    "Mantis HVDS": {"parts": ["Air layer", "Thin Liquid Layer", "Thick Liquid Layer"], "mesin": "Mantis HVDS"},
    "Tempest": {"parts": ["Tempest", "Tempest LV", "Tempest HV"], "mesin": "Silicon Molding"},
    "NT8": {"parts": ["LV Tips", "NT8 Chamber"], "mesin": "JSW 100T"},
    "HT96 Microplate": {"parts": ["HT96 Microplate"], "mesin": "SUMITOMO 130T"},
    "Topas Formulator": {"parts": ["Topas Blank", "Formulator"], "mesin": "General"},
    "RX": {"parts": ["RX", "RI360"], "mesin": "General"},
    "Qiagen": {"parts": ["Coding Frame"], "mesin": "Forindo 140T"},
    "Pulse": {"parts": ["Acc", "RI360", "Stempede"], "mesin": "Silicon Molding"},
}

if "master_data" not in st.session_state:
    st.session_state.master_data = {}
    for proj, pdata in MASTER_PROJECTS.items():
        for part in pdata["parts"]:
            st.session_state.master_data[f"{proj}||{part}"] = {
                "issue": "Flashing on tail & stem areas" if "Plunger" in part or "Tip" in part else "General Issue",
                "define": f"Flashing on {part} ({proj}) - Insert no axial support",
                "measure": "Day1 Minimal, Day2 visible, Day3 increases. Baseline 1.3mm jig",
                "analyze": "Mechanism: Insert no axial support, melt pressure push back",
                "improve": "Main Solution: Add support pillars, grind mold base, trial Sample 1-10 x Cav 1-8 OK",
                "control": "Routine monitoring: Daily visual per Cav 1-8 + 1.3mm jig, weekly audit",
                "status": "Open"
            }

if "sel_proj" not in st.session_state: st.session_state.sel_proj = "FAST"
if "sel_part" not in st.session_state: st.session_state.sel_part = "Plunger LV"

st.title("📊 Quality AI PRO V6 FINAL V3 - HEADER REAL")
with st.container(border=True):
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        proj_list = list(MASTER_PROJECTS.keys())
        sel_proj = st.selectbox("1️⃣ PROJECT", proj_list, index=proj_list.index(st.session_state.sel_proj) if st.session_state.sel_proj in proj_list else 0, key="v6f_proj")
        st.session_state.sel_proj = sel_proj
    with c2:
        part_list = MASTER_PROJECTS[sel_proj]["parts"]
        sel_part = st.selectbox(f"2️⃣ PART di {sel_proj}", part_list, index=part_list.index(st.session_state.sel_part) if st.session_state.sel_part in part_list else 0, key="v6f_part")
        st.session_state.sel_part = sel_part
    with c3:
        key = f"{sel_proj}||{sel_part}"
        data = st.session_state.master_data.get(key, {})
        st.metric("3️⃣ Status", data.get("status", "Open"))

current_key = f"{sel_proj}||{sel_part}"
current_data = st.session_state.master_data.get(current_key, {})
st.info(f"🎯 Aktif: **{sel_proj} > {sel_part}** → 7 Tab - HEADER REAL from template_WI_landscape.docx")

# =================== TAB DEFINITIONS ===================
def tab_dmaic():
    st.subheader(f"🌀 DMAIC - {sel_proj} > {sel_part}")
    up = st.file_uploader("Upload DMAIC (xlsx)", type=["xlsx", "xls"], key=f"dmaic_up_{current_key}")
    if up:
        parsed = universal_parser(up)
        st.success(parsed["info"])
        if not parsed["cav_df"].empty: st.dataframe(parsed["cav_df"], use_container_width=True)
        current_data.update({"define": parsed["define"] or current_data.get("define", ""), "measure": parsed["measure"] or current_data.get("measure", ""), "analyze": parsed["analyze"] or current_data.get("analyze", ""), "improve": parsed["improve"] or current_data.get("improve", ""), "control": parsed["control"] or current_data.get("control", "")})
        st.session_state.master_data[current_key] = current_data
    issue = st.text_input("Issue", current_data.get("issue", "Flashing"), key=f"issue_{current_key}")
    define = st.text_area("1. DEFINE", current_data.get("define", ""), height=80, key=f"def_{current_key}")
    measure = st.text_area("2. MEASURE", current_data.get("measure", ""), height=60, key=f"meas_{current_key}")
    analyze = st.text_area("3. ANALYZE", current_data.get("analyze", ""), height=60, key=f"ana_{current_key}")
    improve = st.text_area("4. IMPROVE", current_data.get("improve", ""), height=60, key=f"imp_{current_key}")
    control = st.text_area("5. CONTROL", current_data.get("control", ""), height=60, key=f"ctrl_{current_key}")
    st.session_state.master_data[current_key] = {"issue": issue, "define": define, "measure": measure, "analyze": analyze, "improve": improve, "control": control, "status": current_data.get("status", "Open")}

def tab_msa(): 
    st.subheader(f"📏 MSA 44 Rows - {sel_proj} > {sel_part}")
    msa_list = []
    for i in range(44):
        msa_list.append({
            "Sample": f"S{i//8+1}",
            "Cavity": f"Cav {i%8+1}",
            "Part": sel_part,
            "Result": "OK",
            "Jig 1.3mm": "PASS"
        })
    df = pd.DataFrame(msa_list)
    st.dataframe(df, use_container_width=True)

def tab_wi():
    st.subheader(f"📄 WI Generator - {sel_proj} > {sel_part}")
    d = current_data
    method = st.text_area("Control", d.get("control", "Daily visual per Cav 1-8 + 1.3mm jig"), height=80, key=f"wi_{current_key}")
    wi_txt = f"WORK INSTRUCTION - {sel_proj} - {sel_part}\nMethod: {method}\nTool: 1.3mm jig"
    st.code(wi_txt)

def tab_screenshot():
    st.subheader(f"📸 Screenshot - {sel_proj} > {sel_part}")
    f_before = st.file_uploader(f"Before {sel_part}", type=["jpg", "png"], key=f"ssb_{current_key}")
    f_after = st.file_uploader(f"After {sel_part}", type=["jpg", "png"], key=f"ssa_{current_key}")
    
    if f_before and f_after and AI_AVAILABLE:
        if st.button("✨ Auto-Analisis Before/After dengan Gemini AI", key=f"btn_ai_ss_{current_key}"):
            from PIL import Image
            img_b = Image.open(f_before)
            img_a = Image.open(f_after)
            with st.spinner("AI sedang menganalisis perbedaan foto..."):
                res_ai = analyze_trial_screenshot(img_b, img_a, sel_proj, sel_part)
                st.markdown("### 🤖 Hasil Analisis Visual AI:")
                st.info(res_ai)

def tab_trial():
    st.subheader(f"🧪 Trial Report - {sel_proj} > {sel_part}")
    st.text_area(f"Trial for {sel_part}", f"Add support pillars for {sel_part}", height=80, key=f"trial_{current_key}")

def tab_catalog():
    st.subheader(f"📚 Catalog Issue & Action Plan - {sel_proj} > {sel_part}")
    st.caption("Katalog masalah spesifik beserta rekomendasi tindakan iteratif untuk mesin dan part terdeteksi.")

    # Master Katalog Rekomendasi berdasarkan Isu/Proyek
    catalog_database = [
        {
            "Project": "FAST",
            "Part": "Plunger LV",
            "Mesin": "SUMITOMO 50T/100T / JSW",
            "Defect / Issue": "Flashing on tail & stem areas",
            "Penyebab Utama (Root Cause)": "Insert no axial support, melt pressure push back",
            "Iterasi 1 (Action Immediate)": "Turunkan hold pressure & periksa V/P transfer point",
            "Iterasi 2 (Action Mold/Tooling)": "Tambahkan support pillar pada mold base & ratakan permukaan insert",
            "Iterasi 3 (Standardisasi & Control)": "Gunakan Jig 1.3mm per Cav 1-8 harian & audit mingguan",
            "Status": "Open"
        },
        {
            "Project": "Tempest",
            "Part": "Tempest HV",
            "Mesin": "Silicon Molding",
            "Defect / Issue": "Bowing / Leaking Issue",
            "Penyebab Utama (Root Cause)": "Perbedaan suhu MTC Core & Cavity, Ejector hold time kurang",
            "Iterasi 1 (Action Immediate)": "Setel MTC Cavity 60°C dan Core 60°C (sebelumnya Cavity 70°C)",
            "Iterasi 2 (Action Mold/Tooling)": "Naikkan Ejector holding time dari 0 detik ke 3 detik",
            "Iterasi 3 (Standardisasi & Control)": "Ukur gap dengan feeler gauge (target < 0.35mm)",
            "Status": "Closed"
        },
        {
            "Project": "NT8",
            "Part": "LV Tips",
            "Mesin": "JSW 100T",
            "Defect / Issue": "Penyok Ujung & Cut/Runner NG",
            "Penyebab Utama (Root Cause)": "Misalignment arm & cutting die, flashing di pangkal tip",
            "Iterasi 1 (Action Immediate)": "Ganti cutting dies set baru & bersihkan runner residue",
            "Iterasi 2 (Action Mold/Tooling)": "Ubah metode ukur Keyence dari X-axis ke Z-axis (flat surface)",
            "Iterasi 3 (Standardisasi & Control)": "Gunakan jig V-groove dengan anodized finish",
            "Status": "In Progress"
        },
        {
            "Project": "RX",
            "Part": "RX / RI360",
            "Mesin": "General / JSW 100T",
            "Defect / Issue": "Silver Marks & Bubble",
            "Penyebab Utama (Root Cause)": "Material lembap, V/P Transfer (1st) & Hold Pressure terlalu tinggi",
            "Iterasi 1 (Action Immediate)": "Jalankan baking material PC & pertahankan Hold Pressure < 3000 kgf/cm²",
            "Iterasi 2 (Action Mold/Tooling)": "Optimasi V/P Transfer (1st & 2nd stage) via OFAT/DOE",
            "Iterasi 3 (Standardisasi & Control)": "Monitoring gate freeze curve (optimal hold time 5-6s)",
            "Status": "Open"
        },
        {
            "Project": "Mantis",
            "Part": "Cap Mantis / Tube support",
            "Mesin": "ARBURG 70T LSR",
            "Defect / Issue": "Uneven Surface & Depth Issue",
            "Penyebab Utama (Root Cause)": "Pengukuran area dome LSR bergelombang / fleksibel",
            "Iterasi 1 (Action Immediate)": "Gunakan jig acrylic / 3D print khusus untuk penahan part",
            "Iterasi 2 (Action Mold/Tooling)": "Gunakan Keyence Microscope untuk area flat & dome",
            "Iterasi 3 (Standardisasi & Control)": "Jalankan GRnR jig depth micrometer (A8/I95)",
            "Status": "Open"
        }
    ]

    df_cat = pd.DataFrame(catalog_database)

    # Filter tampilan sesuai Project yang aktif dipilih di header
    filtered_df = df_cat[df_cat["Project"] == sel_proj]
    
    if filtered_df.empty:
        st.warning(f"Belum ada data spesifik untuk Project **{sel_proj}**. Menampilkan seluruh katalog isu:")
        st.dataframe(df_cat, use_container_width=True)
    else:
        st.success(f"Menampilkan katalog isu terdaftar untuk **{sel_proj}** ({MASTER_PROJECTS[sel_proj]['mesin']}):")
        st.dataframe(filtered_df, use_container_width=True)

    st.divider()

    # Section Visual / Upload Gambar Contoh Defek & Solusi
    st.markdown("### 🖼️ Contoh Visual & Dokumen Perbaikan")
    c_img1, c_img2 = st.columns(2)
    with c_img1:
        st.markdown("**1️⃣ Foto Defek (Before Action)**")
        img_b = st.file_uploader(f"Upload Foto NG ({sel_part})", type=["png", "jpg", "jpeg"], key=f"cat_img_b_{current_key}")
        if img_b:
            st.image(img_b, caption=f"Defek Sebelum Perbaikan - {sel_part}", use_column_width=True)
    with c_img2:
        st.markdown("**2️⃣ Foto Hasil Trial OK (After Action)**")
        img_a = st.file_uploader(f"Upload Foto OK ({sel_part})", type=["png", "jpg", "jpeg"], key=f"cat_img_a_{current_key}")
        if img_a:
            st.image(img_a, caption=f"Kondisi OK Setelah Trial - {sel_part}", use_column_width=True)

def tab_wi_comparison():
    st.subheader(f"🔍 WI Comparison - OLD vs NEW - {sel_proj} > {sel_part}")
    st.caption("Membandingkan dari Upload File (.docx/.pdf) atau langsung via Link Google Docs & Google Sheets.")

    input_type = st.radio("Pilih Metode Input:", ["📄 Upload File", "🔗 Link Dokumen / Sheets"], horizontal=True, key=f"radio_in_{current_key}")
    old_text, new_text = "", ""

    if input_type == "📄 Upload File":
        c_old, c_new = st.columns(2)
        with c_old:
            old_file = st.file_uploader("📤 Upload OLD WI", type=["docx", "pdf"], key=f"wi_old_{current_key}")
        with c_new:
            new_file = st.file_uploader("📤 Upload NEW WI", type=["docx", "pdf"], key=f"wi_new_{current_key}")

        if old_file and new_file:
            old_text = extract_file_content(old_file.getvalue(), old_file.name)
            new_text = extract_file_content(new_file.getvalue(), new_file.name)
    else:
        c_link_old, c_link_new = st.columns(2)
        with c_link_old:
            url_old = st.text_input("🔗 Link OLD Document/Sheets:", placeholder="https://docs.google.com/...", key=f"url_old_{current_key}")
        with c_link_new:
            url_new = st.text_input("🔗 Link NEW Document/Sheets:", placeholder="https://docs.google.com/...", key=f"url_new_{current_key}")

        if url_old.strip() and url_new.strip():
            with st.spinner("Mengambil konten dari link..."):
                def fetch_text_from_link(url):
                    try:
                        clean_url = url.strip()
                        if "/d/" in clean_url:
                            doc_id = clean_url.split("/d/")[1].split("/")[0]
                        else:
                            return ""

                        if "docs.google.com/document" in clean_url:
                            export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
                            res = requests.get(export_url)
                            if res.status_code == 200:
                                return res.text
                            else:
                                st.error("⚠️ Gagal membaca Google Docs. Pastikan akses di-set ke 'Anyone with the link' (Public).")
                                return ""
                        elif "docs.google.com/spreadsheets" in clean_url:
                            export_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=csv"
                            res = requests.get(export_url)
                            if res.status_code == 200:
                                return res.text
                            else:
                                return universal_parser(clean_url).get("define", "")
                        return ""
                    except Exception as e:
                        st.error(f"Error parse URL: {e}")
                        return ""

                old_text = fetch_text_from_link(url_old)
                new_text = fetch_text_from_link(url_new)

    if old_text and new_text:
        old_html_res, new_html_res = generate_side_by_side_html(old_text, new_text)
        st.subheader("📑 Perbandingan Dokumen (Side-by-Side Google Docs Style)")
        
        col_doc_left, col_doc_right = st.columns(2)
        with col_doc_left:
            st.markdown("### 📄 OLD Document")
            st.markdown(
                f'<div style="line-height: 1.6; font-family: sans-serif; background: #ffffff; padding: 20px; border-radius: 5px; border: 1px solid #ccc; height: 650px; overflow-y: auto;">{old_html_res}</div>', 
                unsafe_allow_html=True
            )
            
        with col_doc_right:
            st.markdown("### 📝 NEW Document")
            st.markdown(
                f'<div style="line-height: 1.6; font-family: sans-serif; background: #ffffff; padding: 20px; border-radius: 5px; border: 1px solid #ccc; height: 650px; overflow-y: auto;">{new_html_res}</div>', 
                unsafe_allow_html=True
            )

        if AI_AVAILABLE:
            st.divider()
            if st.button("🤖 Review Hasil Perubahan WI Ini Menggunakan Gemini AI", type="primary", key=f"btn_ai_wicomp_{current_key}"):
                with st.spinner("AI sedang mengevaluasi..."):
                    review_res = analyze_wi_text(old_text, new_text, sel_proj, sel_part)
                    st.markdown("### 📋 Laporan Review AI:")
                    st.markdown(review_res)
    else:
        st.info("👆 Masukkan kedua file atau link (Google Docs/Sheets) di atas untuk melihat hasil perbandingan bersisian.")

# =================== RENDER TAB ===================
t1, t2, t3, t4, t5, t6, t7 = st.tabs(["🌀 DMAIC", "📏 MSA 44 Rows", "📄 WI Generator", "📸 Screenshot", "🧪 Trial Report", "📚 Catalog All Issue", "🔍 WI Comparison V3 HEADER REAL"])
with t1: tab_dmaic()
with t2: tab_msa()
with t3: tab_wi()
with t4: tab_screenshot()
with t5: tab_trial()
with t6: tab_catalog()
with t7: tab_wi_comparison()

st.success(f"V6 FINAL V3 - {sel_proj} > {sel_part} - Ready!")
