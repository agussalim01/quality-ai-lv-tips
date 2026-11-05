# app_quality_ai.py - V6 FINAL FIXED V3 - FULL 7 TABS SAFE RUN
import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import io
from datetime import datetime
import difflib

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

def generate_doc_diff_html(old_text, new_text):
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    diff_html = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            diff_html.append(" ".join(old_words[i1:i2]))
        elif tag == 'delete':
            deleted = " ".join(old_words[i1:i2])
            diff_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
        elif tag == 'insert':
            inserted = " ".join(new_words[j1:j2])
            diff_html.append(f'<span style="background-color: #dafbe1; color: #1a7f37; font-weight: bold; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
        elif tag == 'replace':
            deleted = " ".join(old_words[i1:i2])
            inserted = " ".join(new_words[j1:j2])
            diff_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
            diff_html.append(f'<span style="background-color: #dafbe1; color: #1a7f37; font-weight: bold; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
            
    return " ".join(diff_html)

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
    df = pd.DataFrame([{"Sample": f"S{i//8+1}", "Cavity": f"Cav {i%8+1}", "Part": sel_part, "Result": "OK", "Jig
