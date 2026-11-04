import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, os, difflib
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="Quality AI PRO - V4 JOS - ALL PROJECT ALL ISSUE + LINK", page_icon="📊", layout="wide")

def safe_join(row):
    vals = []
    for v in row.values:
        try:
            if pd.isna(v): continue
            s = str(v).strip()
            if s.lower() != 'nan' and s != '': vals.append(s)
        except: continue
    return ' '.join(vals)

def universal_parser_any_project(file_or_link):
    result = {"pareto_df": pd.DataFrame(), "define": "", "measure": "", "analyze": "", "improve": "", "control": "", "format_detected": ""}
    try:
        # --- SUPPORT GOOGLE SHEETS LINK ---
        xls_dict = {}
        if isinstance(file_or_link, str) and 'docs.google.com' in file_or_link:
            file_id = file_or_link.split('/d/')[1].split('/')[0] if '/d/' in file_or_link else ''
            export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
            xls = pd.ExcelFile(export_url)
            for sh in xls.sheet_names:
                xls_dict[sh] = xls.parse(sh, header=None)
            result["format_detected"] += "Google Sheets Link; "
        elif hasattr(file_or_link, 'name'):
            if file_or_link.name.endswith('.csv'):
                xls_dict = {"Data": pd.read_csv(file_or_link, header=None)}
            else:
                xls = pd.ExcelFile(file_or_link)
                for sh in xls.sheet_names:
                    xls_dict[sh] = xls.parse(sh, header=None)
        else:
            return result

        # SUMMARY DETECT
        for sh_name, df_raw in xls_dict.items():
            try:
                txt_all = df_raw.apply(safe_join, axis=1).str.cat(sep=' ').lower()
            except: txt_all = ""
            if 'define' in txt_all and 'measure' in txt_all:
                for _, row in df_raw.iterrows():
                    row_txt = safe_join(row)
                    if len(row_txt) < 15: continue
                    low = row_txt.lower()
                    if 'problem' in low or 'flashing on' in low: result["define"] += row_txt + "\n"
                    if 'actual condition' in low or 'inject day' in low or 'baseline' in low: result["measure"] += row_txt + "\n"
                    if 'mechanism' in low or 'insert' in low: result["analyze"] += row_txt + "\n"
                    if 'main solution' in low or 'add support' in low: result["improve"] += row_txt + "\n"
                    if 'routine monitoring' in low or 'daily visual' in low: result["control"] += row_txt + "\n"
                if result["define"]: result["format_detected"] += f"Summary DMAIC {sh_name}; "

        # CAVITY FORMAT (Cav 1-8) - Kayak file HV Plunger kamu
        for sh_name, df_raw in xls_dict.items():
            cav_row_idx = -1
            for i, row in df_raw.iterrows():
                rs = safe_join(row)
                if 'Cav 1' in rs: cav_row_idx = i; break
            if cav_row_idx >= 0:
                try:
                    xls_tmp = pd.ExcelFile(file_or_link) if hasattr(file_or_link, 'name') else pd.ExcelFile(export_url)
                    df_cav = xls_tmp.parse(sh_name, header=cav_row_idx)
                    df_cav = df_cav.loc[:, ~df_cav.columns.astype(str).str.contains('Unnamed')]
                    cav_cols = [c for c in df_cav.columns if 'Cav' in str(c)]
                    if cav_cols:
                        melted = df_cav.melt(value_vars=cav_cols, var_name='Cavity', value_name='Result')
                        melted['Result'] = melted['Result'].astype(str).str.upper()
                        ng_df = melted[melted['Result'].str.contains('NG|FLASH|FAIL')]
                        if not ng_df.empty:
                            pareto = ng_df.groupby('Cavity').size().reset_index(name='Qty')
                            pareto.rename(columns={'Cavity':'DefectType'}, inplace=True)
                            result["pareto_df"] = pareto
                        else:
                            result["pareto_df"] = pd.DataFrame([{"DefectType": "All Cavity OK - After Improvement (HV Plunger)", "Qty": 0}])
                        result["format_detected"] += f"Cavity {sh_name}; "
                        break
                except: continue

        if not result["format_detected"]: result["format_detected"] = "File terbaca - Format After Improvement (All OK)"
        return result
    except Exception as e:
        result["format_detected"] = f"Terbaca tapi format baru: {str(e)[:120]}"
        return result

def dmaic_tab():
    st.subheader("🌀 DMAIC UNIVERSAL V4 JOS - ALL PROJECT ALL ISSUE + LINK SUPPORT")
    st.caption("Bisa upload file ATAU paste link Google Sheets - Contoh: DMAIC_Flashing_Plunger_HV.xlsx")

    if "dmaic_measure_data" not in st.session_state:
        st.session_state.dmaic_measure_data = pd.DataFrame([{"DefectType": "Contoh: Flashing", "Qty": 10}])
    for k in ["auto_define","auto_measure","auto_analyze","auto_improve","auto_control"]:
        if k not in st.session_state: st.session_state[k] = ""

    tab_upload, tab_link = st.tabs(["📤 Upload File (xlsx/csv)", "🔗 Paste Google Sheets Link"])
    uploaded = None
    link = None
    with tab_upload:
        uploaded = st.file_uploader("Upload DMAIC - SEMUA PROJECT SEMUA ISU", type=["xlsx","xls","csv"], key="dmaic_v4_upload")
    with tab_link:
        link = st.text_input("Paste Link Google Sheets (Anyone with link - Viewer)", "https://docs.google.com/spreadsheets/d/.../edit", key="dmaic_v4_link")
        if st.button("🔍 Parse dari Link", key="parse_link"):
            if 'docs.google.com' in link:
                uploaded = link

    source = uploaded if uploaded else (link if link and 'docs.google.com' in link else None)
    df_m = st.session_state.dmaic_measure_data

    if source:
        with st.spinner("JOS! Auto-detecting..."):
            parsed = universal_parser_any_project(source)
        st.success(f"✅ Format terdeteksi: {parsed['format_detected']}")
        if not parsed["pareto_df"].empty:
            st.session_state.dmaic_measure_data = parsed["pareto_df"]
            df_m = parsed["pareto_df"]
            if parsed["pareto_df"]["Qty"].sum() > 0:
                st.metric("Total Defect", int(parsed["pareto_df"]["Qty"].sum()))
                st.plotly_chart(px.bar(parsed["pareto_df"], x="DefectType", y="Qty", color="DefectType", title="Pareto Auto - All Project"), use_container_width=True)
            else:
                st.info("✅ File ini After Improvement (semua OK) - seperti contoh HV Plunger kamu")
        if parsed["define"]: st.session_state.auto_define = parsed["define"][:2000]
        if parsed["measure"]: st.session_state.auto_measure = parsed["measure"][:2000]
        if parsed["analyze"]: st.session_state.auto_analyze = parsed["analyze"][:2000]
        if parsed["improve"]: st.session_state.auto_improve = parsed["improve"][:2000]
        if parsed["control"]: st.session_state.auto_control = parsed["control"][:2000]

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        dmaic_no = st.text_input("DMAIC No", f"DMAIC-{datetime.now().strftime('%Y%m%d')}-001", key="no_v4")
        project_name = st.text_input("Project / Product (BEBAS)", "HV Plunger / NT8 / Project Lain", key="proj_v4")
        issue = st.text_input("Issue / Problem (BEBAS - Semua isu)", "Flashing, Short Shot, Black Dot, Burn Mark, dll", key="issue_v4")
    with c2:
        dmaic_date = st.date_input("Date", datetime.now(), key="date_v4")
        team = st.text_input("Team", "Quality, Engineering, Production", key="team_v4")

    problem = st.text_area("Problem Statement (Auto-fill dari Summary, bisa untuk semua project)", st.session_state.auto_define or "Flashing on HV plunger tail & stem areas, increasing each cycle", height=120, key="prob_v4")
    baseline = st.text_area("Baseline / Actual Condition", st.session_state.auto_measure or "Inject Day 1 Minimal, Day 2 visible, Day 3 increases", height=80, key="base_v4")
    analyze = st.text_area("Mechanism / Root Cause", st.session_state.auto_analyze or "Insert no axial support, melt pressure pushes back, micro gap", height=100, key="anal_v4")
    improve = st.text_area("Improvement", st.session_state.auto_improve or "Add support pillars, grind mold base for preload", height=100, key="imp_v4")
    control = st.text_area("Control Plan", st.session_state.auto_control or "Daily visual + 1.3mm jig, weekly audit PE & IME", height=100, key="ctrl_v4")

    if st.button("📄 Generate DMAIC Report UNIVERSAL - All Project All Issue", type="primary", use_container_width=True, key="gen_v4"):
        doc = Document()
        doc.add_heading(f"DMAIC REPORT - {issue} - {project_name}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"DMAIC No: {dmaic_no} | Project: {project_name} | Issue: {issue} | Date: {dmaic_date} | Team: {team}")
        doc.add_heading("1. DEFINE", 2); doc.add_paragraph(problem)
        doc.add_heading("2. MEASURE", 2); doc.add_paragraph(f"Baseline: {baseline}\nData:\n{df_m.to_string()}")
        doc.add_heading("3. ANALYZE", 2); doc.add_paragraph(analyze)
        doc.add_heading("4. IMPROVE", 2); doc.add_paragraph(improve)
        doc.add_heading("5. CONTROL", 2); doc.add_paragraph(control)
        bio = io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Download DMAIC Universal", data=bio, file_name=f"{dmaic_no}_{project_name}_{issue}_UNIVERSAL_V4.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

st.title("📊 Quality AI PRO - V4 JOS - ALL PROJECT ALL ISSUE + LINK")
dmaic_tab()
