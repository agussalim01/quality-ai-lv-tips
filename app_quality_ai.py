import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io, os, difflib
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="Quality AI PRO - UNIVERSAL ALL PROJECT ALL ISSUE V3", page_icon="📊", layout="wide")

def universal_parser_any_project(uploaded_file):
    result = {"pareto_df": pd.DataFrame(), "define": "", "measure": "", "analyze": "", "improve": "", "control": "", "format_detected": ""}
    try:
        if uploaded_file.name.endswith('.csv'):
            xls_dict = {"Data": pd.read_csv(uploaded_file)}
        else:
            xls = pd.ExcelFile(uploaded_file)
            xls_dict = {}
            for sh in xls.sheet_names:
                try:
                    xls_dict[sh] = xls.parse(sh, header=None)
                except:
                    pass
        
        for sh_name, df_raw in xls_dict.items():
            txt_all = df_raw.astype(str).apply(lambda x: ' '.join(x), axis=1).str.cat(sep=' ').lower()
            if 'define' in txt_all and 'measure' in txt_all:
                for _, row in df_raw.iterrows():
                    row_txt = ' '.join([str(v) for v in row.values if str(v)!='nan'])
                    if len(row_txt) < 20: continue
                    low = row_txt.lower()
                    if 'problem' in low or 'flashing on' in low or 'issue' in low:
                        result["define"] += row_txt + "\n"
                    if 'actual condition' in low or 'inject day' in low or 'baseline' in low:
                        result["measure"] += row_txt + "\n"
                    if 'mechanism' in low or 'insert' in low or 'root cause' in low:
                        result["analyze"] += row_txt + "\n"
                    if 'main solution' in low or 'add support' in low or 'improvement' in low:
                        result["improve"] += row_txt + "\n"
                    if 'routine monitoring' in low or 'daily visual' in low or 'control' in low:
                        result["control"] += row_txt + "\n"
                if result["define"]:
                    result["format_detected"] += f"Summary DMAIC in {sh_name}; "

        for sh_name, df_raw in xls_dict.items():
            cav_row_idx = -1
            for i, row in df_raw.iterrows():
                row_str = ' '.join([str(v) for v in row.values])
                if 'Cav 1' in row_str or 'CAV' in row_str and '1' in row_str:
                    cav_row_idx = i
                    break
            if cav_row_idx >= 0:
                try:
                    xls2 = pd.ExcelFile(uploaded_file) if not uploaded_file.name.endswith('.csv') else None
                    if xls2:
                        df_cav = xls2.parse(sh_name, header=cav_row_idx)
                    else:
                        df_cav = pd.DataFrame()
                    df_cav = df_cav.loc[:, ~df_cav.columns.astype(str).str.contains('Unnamed')]
                    cav_cols = [c for c in df_cav.columns if 'Cav' in str(c)]
                    if cav_cols:
                        melted = df_cav.melt(value_vars=cav_cols, var_name='Cavity', value_name='Result')
                        melted['Result'] = melted['Result'].astype(str).str.upper()
                        ng_df = melted[melted['Result'].str.contains('NG|FLASH|FAIL|REJECT')]
                        if not ng_df.empty:
                            pareto = ng_df.groupby('Cavity').size().reset_index(name='Qty')
                            pareto.rename(columns={'Cavity':'DefectType'}, inplace=True)
                            result["pareto_df"] = pareto
                            result["format_detected"] += f"Cavity format (Cav 1-8) in {sh_name} - {len(ng_df)} NG; "
                        else:
                            result["pareto_df"] = pd.DataFrame([{"DefectType": "All Cavity OK - After Improvement", "Qty": 0}])
                            result["format_detected"] += f"Cavity all OK (After) in {sh_name}; "
                        break
                except:
                    pass

        if result["pareto_df"].empty:
            for sh_name, df_raw in xls_dict.items():
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = xls_dict[sh_name]
                    else:
                        xls2 = pd.ExcelFile(uploaded_file)
                        df = xls2.parse(sh_name)
                    df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed')]
                    df = df.dropna(how='all')
                    if df.shape[0] < 2: continue
                    cols_lower = {str(c).lower(): c for c in df.columns}
                    defect_col = next((cols_lower[k] for k in cols_lower if 'defect' in k or 'jenis' in k or 'kategori' in k), None)
                    qty_col = next((cols_lower[k] for k in cols_lower if 'qty' in k or 'jumlah' in k), None)
                    status_col = next((cols_lower[k] for k in cols_lower if 'status' in k or 'result' in k or 'ok/ng' in k), None)
                    
                    if qty_col and defect_col:
                        df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
                        pareto = df.groupby(defect_col)[qty_col].sum().reset_index()
                        pareto.columns = ['DefectType','Qty']
                        pareto = pareto[pareto['Qty']>0]
                        if not pareto.empty:
                            result["pareto_df"] = pareto
                            result["format_detected"] += f"Defect+Qty in {sh_name}; "
                            break
                    elif status_col:
                        tmp = df.copy()
                        tmp[status_col] = tmp[status_col].astype(str).str.upper()
                        ng = tmp[tmp[status_col].str.contains('NG|REJECT|FAIL')]
                        if not ng.empty:
                            if defect_col:
                                pareto = ng.groupby(defect_col).size().reset_index(name='Qty')
                                pareto.columns = ['DefectType','Qty']
                            else:
                                pareto = pd.DataFrame([{"DefectType": "TOTAL NG", "Qty": len(ng)}])
                            result["pareto_df"] = pareto
                            result["format_detected"] += f"Status OK/NG in {sh_name}; "
                            break
                    elif defect_col:
                        pareto = df.groupby(defect_col).size().reset_index(name='Qty')
                        pareto.columns = ['DefectType','Qty']
                        pareto = pareto[~pareto['DefectType'].astype(str).str.upper().isin(['OK','GOOD','PASS'])]
                        if not pareto.empty:
                            result["pareto_df"] = pareto
                            result["format_detected"] += f"Count per Defect in {sh_name}; "
                            break
                except:
                    pass
        return result
    except Exception as e:
        result["format_detected"] = f"Error: {e}"
        return result

def dmaic_tab():
    st.subheader("🌀 DMAIC Generator - UNIVERSAL ALL PROJECT ALL ISSUE V3")
    st.caption("Upload contoh DMAIC apapun (HV Plunger, NT8, project lain) - Sistem auto-learn & bisa dipakai untuk semua project & semua isu")

    if "dmaic_measure_data" not in st.session_state:
        st.session_state.dmaic_measure_data = pd.DataFrame([{"DefectType": "Contoh: Flashing", "Qty": 10}])
    if "auto_define" not in st.session_state: st.session_state.auto_define = ""
    if "auto_measure" not in st.session_state: st.session_state.auto_measure = ""
    if "auto_analyze" not in st.session_state: st.session_state.auto_analyze = ""
    if "auto_improve" not in st.session_state: st.session_state.auto_improve = ""
    if "auto_control" not in st.session_state: st.session_state.auto_control = ""

    uploaded = st.file_uploader("Upload Contoh DMAIC - SEMUA PROJECT SEMUA ISU (xlsx/csv)", type=["xlsx","xls","csv"], key="dmaic_universal_v3")
    df_m = st.session_state.dmaic_measure_data

    if uploaded:
        with st.spinner("🔍 Auto-detecting format untuk semua project..."):
            parsed = universal_parser_any_project(uploaded)
        st.success(f"✅ Format terdeteksi: {parsed['format_detected']}")
        if not parsed["pareto_df"].empty:
            st.session_state.dmaic_measure_data = parsed["pareto_df"]
            df_m = parsed["pareto_df"]
            if parsed["pareto_df"]["Qty"].sum() > 0:
                st.metric("Total Defect Terdeteksi (Bukan 0 lagi)", int(parsed["pareto_df"]["Qty"].sum()))
                fig = px.bar(parsed["pareto_df"], x="DefectType", y="Qty", color="DefectType", title=f"Pareto Auto - {parsed['format_detected']}")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("File ini kondisi After Improvement (semua OK) - seperti contoh HV Plunger kamu")
        if parsed["define"]: st.session_state.auto_define = parsed["define"][:2000]
        if parsed["measure"]: st.session_state.auto_measure = parsed["measure"][:2000]
        if parsed["analyze"]: st.session_state.auto_analyze = parsed["analyze"][:2000]
        if parsed["improve"]: st.session_state.auto_improve = parsed["improve"][:2000]
        if parsed["control"]: st.session_state.auto_control = parsed["control"][:2000]

    st.divider()
    st.markdown("### 1️⃣ D - DEFINE - Untuk Semua Project Semua Isu")
    c1, c2 = st.columns(2)
    with c1:
        dmaic_no = st.text_input("DMAIC No", f"DMAIC-{datetime.now().strftime('%Y%m%d')}-001", key="dmaic_no_v3")
        project_name = st.text_input("Project / Product", "HV Plunger / NT8 / Project Lain - Bebas", key="proj_v3")
        issue = st.text_input("Issue / Problem (BEBAS - Semua isu)", "Flashing, Short Shot, Black Dot, dll", key="issue_v3")
    with c2:
        dmaic_date = st.date_input("Date", datetime.now(), key="date_v3")
        team = st.text_input("Team", "Quality, Engineering, Production", key="team_v3")

    problem = st.text_area("Problem Statement (Auto-fill jika upload Summary)", st.session_state.auto_define if st.session_state.auto_define else "Isi problem untuk project apapun", height=120, key="prob_v3")

    st.markdown("### 2️⃣ M - MEASURE - Universal Pareto")
    if not df_m.empty:
        st.dataframe(df_m, use_container_width=True)
    
    baseline = st.text_area("Baseline / Actual Condition", st.session_state.auto_measure if st.session_state.auto_measure else "Day 1 Minimal, Day 2 visible", height=100, key="base_v3")
    st.markdown("### 3️⃣ A - ANALYZE")
    analyze = st.text_area("Mechanism / Root Cause", st.session_state.auto_analyze if st.session_state.auto_analyze else "Root cause untuk project apapun", height=120, key="anal_v3")
    st.markdown("### 4️⃣ I - IMPROVE")
    improve = st.text_area("Improvement Action", st.session_state.auto_improve if st.session_state.auto_improve else "Improvement untuk semua project", height=120, key="imp_v3")
    st.markdown("### 5️⃣ C - CONTROL")
    control = st.text_area("Control Plan", st.session_state.auto_control if st.session_state.auto_control else "Control untuk semua project", height=120, key="ctrl_v3")

    st.divider()
    if st.button("📄 Generate DMAIC Report UNIVERSAL - All Project All Issue", type="primary", use_container_width=True, key="gen_dmaic_v3"):
        doc = Document()
        doc.add_heading(f"DMAIC REPORT - {issue} - {project_name}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"DMAIC No: {dmaic_no} | Project: {project_name} | Issue: {issue} | Date: {dmaic_date} | Team: {team}")
        doc.add_heading("1. DEFINE", 2); doc.add_paragraph(problem)
        doc.add_heading("2. MEASURE", 2); doc.add_paragraph(f"Baseline: {baseline}\nData:\n{df_m.to_string() if not df_m.empty else 'No data'}")
        doc.add_heading("3. ANALYZE", 2); doc.add_paragraph(analyze)
        doc.add_heading("4. IMPROVE", 2); doc.add_paragraph(improve)
        doc.add_heading("5. CONTROL", 2); doc.add_paragraph(control)
        bio = io.BytesIO(); doc.save(bio); bio.seek(0)
        st.success(f"DMAIC Universal untuk {project_name} - {issue} jadi!")
        st.download_button("📥 Download DMAIC Universal Report", data=bio, file_name=f"{dmaic_no}_{project_name}_{issue}_UNIVERSAL.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

st.title("📊 Quality AI PRO - UNIVERSAL ALL PROJECT ALL ISSUE V3")
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📊 MSA", "📄 WI Generator", "🖼️ Screenshot Compare", "🧪 Trial Report", "📚 Catalog Reject", "🌀 DMAIC UNIVERSAL V3"])
def msa_tab():
    st.subheader("📊 MSA")
    st.dataframe(pd.DataFrame({"Part":[1,2,3]}))
def wi_tab():
    st.subheader("📄 WI")
def comparator_tab():
    st.subheader("🖼️ Compare")
def trial_report_tab():
    st.subheader("🧪 Trial")
def catalog_reject_tab():
    st.subheader("📚 Catalog All Issue")
with tab1:
    msa_tab()
with tab2:
    wi_tab()
with tab3:
    comparator_tab()
with tab4:
    trial_report_tab()
with tab5:
    catalog_reject_tab()
with tab6:
    dmaic_tab()
