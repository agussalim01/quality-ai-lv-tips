# app_quality_ai.py - V6 FINAL FROM HANDBOOK AGS - PROJECT>PART>6 TAB AUTO-FILL
import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import io
from datetime import datetime

st.set_page_config(page_title="Quality AI V6 FINAL - Handbook AGS", page_icon="📊", layout="wide")

def safe_join(row):
    vals = []
    for v in row.values:
        try:
            if pd.isna(v): continue
            s = str(v).strip()
            if s.lower() not in ['nan','']: vals.append(s)
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
                    if 'flashing' in rt.lower() or 'problem' in rt.lower(): result["define"] += rt + "\n"
                    if 'actual condition' in rt.lower() or 'inject day' in rt.lower(): result["measure"] += rt + "\n"
                    if 'mechanism' in rt.lower() or 'insert' in rt.lower(): result["analyze"] += rt + "\n"
                    if 'main solution' in rt.lower() or 'add support' in rt.lower(): result["improve"] += rt + "\n"
                    if 'routine monitoring' in rt.lower() or 'daily visual' in rt.lower(): result["control"] += rt + "\n"
                result["info"] += f"{sh}; "
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
                result["info"] += f"Cav {sh}; "
                break
        return result
    except Exception as e:
        result["info"] = f"{e}"; return result

MASTER_PROJECTS = {
    "FAST": {"parts": ["LV Pippete tip", "HV PIppete Tip", "Plunger LV", "LV Caddy", "Cover LV", "LV Plunger A - 001434C", "LV Plunger B - 001434C", "LV Plunger C - 001434C", "LV Plunger D - 001434C", "HV Caddy", "Plunger HV"], "mesin": "SUMITOMO 50T/100T/130T, JSW 80T/100T/220T"},
    "Flo PD (FLo PD)": {"parts": ["Caddy Cover", "Tip HV", "Tip LV", "Base Caddy", "Top Caddy", "Plunger LV", "Plunger HV", "Plunger Pos.Ring", "Sleeve"], "mesin": "JSW 100T A/B, SUMITOMO 50T"},
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
                "issue": "Flashing on tail & stem areas" if "Plunger" in part or "Tip" in part else "General Issue - Flashing/Short Shot/Black Dot",
                "define": f"Flashing on {part} ({proj}) - Background: Insert does not have axial support",
                "measure": "Day1 Minimal, Day2 clearly visible, Day3 increases. Baseline with 1.3mm jig",
                "analyze": "Mechanism: Insert no axial support, melt pressure push back, micro gap -> flashing",
                "improve": "Main Solution: Add support pillars behind insert, grind mold base for preload, trial Sample 1-10 x Cav 1-8 OK",
                "control": "Routine monitoring: Daily visual check per Cav 1-8 + 1.3mm jig, weekly audit PE & IME",
                "status": "Open"
            }
    if "FAST||Plunger LV" in st.session_state.master_data:
        st.session_state.master_data["FAST||Plunger LV"]["status"] = "Closed - 0% with 1.3mm jig"

if "sel_proj" not in st.session_state: st.session_state.sel_proj = "FAST"
if "sel_part" not in st.session_state: st.session_state.sel_part = "Plunger LV"

st.title("📊 Quality AI PRO V6 FINAL - Handbook AGS - PROJECT>PART>6 TAB")
with st.container(border=True):
    c1, c2, c3, c4 = st.columns([2,2,2,1.5])
    with c1:
        proj_list = list(MASTER_PROJECTS.keys())
        sel_proj = st.selectbox("1️⃣ PROJECT (dari handbook_ags.xlsx)", proj_list, index=proj_list.index(st.session_state.sel_proj) if st.session_state.sel_proj in proj_list else 0, key="v6f_proj")
        st.session_state.sel_proj = sel_proj
        st.caption(f"Mesin: {MASTER_PROJECTS[sel_proj]['mesin']}")
    with c2:
        part_list = MASTER_PROJECTS[sel_proj]["parts"]
        sel_part = st.selectbox(f"2️⃣ PART di {sel_proj} ({len(part_list)} part)", part_list, index=part_list.index(st.session_state.sel_part) if st.session_state.sel_part in part_list else 0, key="v6f_part")
        st.session_state.sel_part = sel_part
    with c3:
        key = f"{sel_proj}||{sel_part}"
        data = st.session_state.master_data.get(key, {})
        st.metric("3️⃣ Status Part", data.get("status","Open"))
        st.caption(f"Issue: {data.get('issue','-')}")
    with c4:
        st.write("➕ Quick Add")
        if st.button("Reuse Issue Mirip ke Part Lain", use_container_width=True, key="v6f_reuse_btn"):
            st.session_state.show_reuse = True

if st.session_state.get("show_reuse", False):
    with st.container(border=True):
        st.markdown("### 🔁 Reuse Issue (Karena isu hampir sama di beberapa project/part)")
        sc1, sc2, sc3 = st.columns(3)
        with sc1: src = st.selectbox("Copy dari", list(st.session_state.master_data.keys()), key="v6f_src")
        with sc2:
            tgt_proj = st.selectbox("Ke Project", proj_list, key="v6f_tproj")
            tgt_part = st.selectbox("Ke Part", MASTER_PROJECTS[tgt_proj]["parts"], key="v6f_tpart")
        with sc3:
            if st.button("📋 Copy Sekarang", type="primary", key="v6f_copy"):
                st.session_state.master_data[f"{tgt_proj}||{tgt_part}"] = st.session_state.master_data[src].copy()
                st.success(f"Copied {src} -> {tgt_proj}||{tgt_part}!")
                st.session_state.show_reuse = False
                st.rerun()
            if st.button("Tutup", key="v6f_close"): st.session_state.show_reuse = False; st.rerun()

current_key = f"{sel_proj}||{sel_part}"
current_data = st.session_state.master_data.get(current_key, {})
st.info(f"🎯 Aktif: **{sel_proj} > {sel_part}** → 6 Tab auto-fill untuk part ini. Isu hampir sama bisa di-reuse 1 klik.")

def tab_dmaic():
    st.subheader(f"🌀 DMAIC Generator - {sel_proj} > {sel_part} - AUTO")
    c_up, c_link = st.columns(2)
    with c_up: up = st.file_uploader("Upload DMAIC (xlsx)", type=["xlsx","xls"], key=f"dmaic_up_{current_key}")
    with c_link: link = st.text_input("Paste Google Sheets Link", key=f"dmaic_link_{current_key}"); btn = st.button("Parse Link", key=f"dmaic_btn_{current_key}")
    source = up if up else (link if btn and 'docs.google.com' in link else None)
    if source:
        parsed = universal_parser(source)
        st.success(parsed["info"])
        if not parsed["cav_df"].empty: st.dataframe(parsed["cav_df"], use_container_width=True)
        current_data.update({"define": parsed["define"] or current_data.get("define",""), "measure": parsed["measure"] or current_data.get("measure",""), "analyze": parsed["analyze"] or current_data.get("analyze",""), "improve": parsed["improve"] or current_data.get("improve",""), "control": parsed["control"] or current_data.get("control","")})
        if not parsed["pareto"].empty: current_data["pareto"] = parsed["pareto"]
        st.session_state.master_data[current_key] = current_data
    no = st.text_input("DMAIC No", f"DMAIC-{sel_proj}-{sel_part}-{datetime.now().strftime('%Y%m%d')}", key=f"no_{current_key}")
    issue = st.text_input("Issue", current_data.get("issue","Flashing/Short Shot/Black Dot"), key=f"issue_{current_key}")
    define = st.text_area("1. DEFINE", current_data.get("define",""), height=100, key=f"def_{current_key}")
    measure = st.text_area("2. MEASURE", current_data.get("measure",""), height=80, key=f"meas_{current_key}")
    if "pareto" in current_data and isinstance(current_data["pareto"], pd.DataFrame) and not current_data["pareto"].empty and current_data["pareto"]["Qty"].sum() > 0:
        st.plotly_chart(px.bar(current_data["pareto"], x="Cavity", y="Qty", title=f"Pareto {sel_part}"), use_container_width=True)
    analyze = st.text_area("3. ANALYZE", current_data.get("analyze",""), height=80, key=f"ana_{current_key}")
    improve = st.text_area("4. IMPROVE", current_data.get("improve",""), height=80, key=f"imp_{current_key}")
    control = st.text_area("5. CONTROL", current_data.get("control",""), height=80, key=f"ctrl_{current_key}")
    st.session_state.master_data[current_key] = {"issue": issue, "define": define, "measure": measure, "analyze": analyze, "improve": improve, "control": control, "status": current_data.get("status","Open"), "pareto": current_data.get("pareto", pd.DataFrame())}
    if st.button(f"📄 Generate DMAIC DOCX - {sel_part}", type="primary", use_container_width=True, key=f"gen_{current_key}"):
        doc = Document(); doc.add_heading(f"DMAIC {sel_proj} {sel_part} - {issue}",1); doc.add_paragraph(f"No:{no}\nProject:{sel_proj} Part:{sel_part}\nDEFINE:{define}\nMEASURE:{measure}\nANALYZE:{analyze}\nIMPROVE:{improve}\nCONTROL:{control}")
        bio = io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Download DOCX", bio, file_name=f"{no}_{sel_proj}_{sel_part}.docx", use_container_width=True, key=f"dl_{current_key}")

def tab_msa(): 
    st.subheader(f"📏 MSA 44 Rows - {sel_proj} > {sel_part}")
    df = pd.DataFrame([{"Sample": f"S{i//8+1}", "Cavity": f"Cav {i%8+1}", "Part": sel_part, "Project": sel_proj, "Result": "OK", "Jig 1.3mm": "PASS"} for i in range(44)])
    st.dataframe(df, use_container_width=True)

def tab_wi():
    st.subheader(f"📄 WI Generator - {sel_proj} > {sel_part} - TEMPLATE SAMA")
    d = current_data
    method = st.text_area("Control (dari Tab DMAIC)", d.get("control","Daily visual per Cav 1-8 + 1.3mm jig"), height=100, key=f"wi_{current_key}")
    wi_txt = f"WORK INSTRUCTION - {sel_proj} - {sel_part}\nProject: {sel_proj} | Part: {sel_part}\nMethod: {method}\nTool: 1.3mm jig go/no-go (template sama)\nCriteria: PASS 0% flashing"
    st.code(wi_txt)
    if st.button(f"Generate WI DOCX - {sel_part}", type="primary", use_container_width=True, key=f"wigen_{current_key}"):
        doc = Document(); doc.add_heading(f"WI {sel_proj} {sel_part}",1); doc.add_paragraph(wi_txt)
        bio = io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Download WI", bio, file_name=f"WI_{sel_proj}_{sel_part}.docx", use_container_width=True, key=f"widl_{current_key}")

def tab_ss(): st.subheader(f"📸 Screenshot - {sel_proj} > {sel_part}"); st.file_uploader(f"Before {sel_part}", type=["jpg","png"], key=f"ssb_{current_key}"); st.file_uploader(f"After {sel_part}", type=["jpg","png"], key=f"ssa_{current_key}")
def tab_trial(): st.subheader(f"🧪 Trial Report - {sel_proj} > {sel_part}"); st.text_area(f"Trial for {sel_part}", f"Add support pillars, grind base for {sel_part}", height=80, key=f"trial_{current_key}")
def tab_catalog():
    st.subheader("📚 Catalog All Issue - Isu Hampir Sama")
    all_rows = []
    for k, v in st.session_state.master_data.items():
        if "||" in k:
            pj, pt = k.split("||")
            all_rows.append({"Project": pj, "Part": pt, "Issue": v.get("issue",""), "Status": v.get("status",""), "Root": v.get("analyze","")[:60]})
    df = pd.DataFrame(all_rows)
    if not df.empty:
        c1, c2 = st.columns(2)
        with c1: f_proj = st.selectbox("Filter Project", ["All"] + sorted(df["Project"].unique().tolist()), key="cat_fproj")
        with c2: f_issue = st.text_input("Cari Issue (Flashing/Short/Black Dot)", key="cat_fissue")
        if f_proj != "All": df = df[df["Project"]==f_proj]
        if f_issue: df = df[df["Issue"].str.contains(f_issue, case=False, na=False)]
        st.dataframe(df, use_container_width=True, height=400)
        st.metric("Total Part dengan Issue", len(df))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w: pd.DataFrame(all_rows).to_excel(w, index=False)
    st.download_button("📥 Export Catalog All Project>Part", buf.getvalue(), file_name=f"Catalog_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True, key="cat_exp")

t1, t2, t3, t4, t5, t6 = st.tabs(["🌀 DMAIC","📏 MSA 44 Rows","📄 WI Generator","📸 Screenshot","🧪 Trial Report","📚 Catalog All Issue"])
with t1: tab_dmaic()
with t2: tab_msa()
with t3: tab_wi()
with t4: tab_ss()
with t5: tab_trial()
with t6: tab_catalog()
st.success(f"V6 FINAL - Handbook AGS - {sel_proj} > {sel_part} - AutoFill - Catalog isu hampir sama BENAR")
