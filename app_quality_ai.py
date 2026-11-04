# app_quality_ai.py - V6 FINAL + WI Comparison FIXED V2 - EXACT MOCKUP - FROM HANDBOOK_AGS.XLSX - SAFE EDIT
import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import io
from datetime import datetime
import difflib

st.set_page_config(page_title="Quality AI V6 FINAL + WI Compare FIXED V2", page_icon="📊", layout="wide")

def safe_join(row):
    vals=[]
    for v in row.values:
        try:
            if pd.isna(v): continue
            s=str(v).strip()
            if s.lower() not in ['nan','']: vals.append(s)
        except: continue
    return ' '.join(vals)

def universal_parser(file_or_link):
    result={"cav_df":pd.DataFrame(),"pareto":pd.DataFrame(),"define":"","measure":"","analyze":"","improve":"","control":"","info":""}
    try:
        xls=None
        if isinstance(file_or_link,str) and 'docs.google.com' in file_or_link:
            fid=file_or_link.split('/d/')[1].split('/')[0]
            url=f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx"
            xls=pd.ExcelFile(url)
            result["info"]="Link OK; "
        elif hasattr(file_or_link,'name'):
            xls=pd.ExcelFile(file_or_link)
        else: return result
        for sh in xls.sheet_names:
            df0=xls.parse(sh, header=None)
            txt=df0.apply(safe_join, axis=1).str.cat(sep=' ').lower()
            if 'define' in txt:
                for _,r in df0.iterrows():
                    rt=safe_join(r)
                    if len(rt)<12: continue
                    if 'flashing' in rt.lower(): result["define"]+=rt+"\n"
                    if 'actual condition' in rt.lower(): result["measure"]+=rt+"\n"
                    if 'mechanism' in rt.lower(): result["analyze"]+=rt+"\n"
                    if 'main solution' in rt.lower(): result["improve"]+=rt+"\n"
                    if 'routine monitoring' in rt.lower(): result["control"]+=rt+"\n"
        for sh in xls.sheet_names:
            df0=xls.parse(sh, header=None)
            idx=-1
            for i,r in df0.iterrows():
                if 'Cav 1' in safe_join(r): idx=i; break
            if idx>=0:
                df_cav=xls.parse(sh, header=idx)
                df_cav=df_cav.loc[:, ~df_cav.columns.astype(str).str.contains('Unnamed')]
                result["cav_df"]=df_cav
                cav_cols=[c for c in df_cav.columns if 'Cav' in str(c)]
                if cav_cols:
                    m=df_cav.melt(value_vars=cav_cols, var_name='Cavity', value_name='Result')
                    m['Result']=m['Result'].astype(str).str.upper()
                    ng=m[m['Result'].str.contains('NG|FLASH')]
                    result["pareto"]=ng.groupby('Cavity').size().reset_index(name='Qty') if not ng.empty else pd.DataFrame([{"Cavity":"All OK","Qty":0}])
                break
        return result
    except Exception as e:
        result["info"]=f"{e}"; return result

MASTER_PROJECTS={
    "FAST":{"parts":["LV Pippete tip","HV PIppete Tip","Plunger LV","LV Caddy","Cover LV","LV Plunger A - 001434C","LV Plunger B - 001434C","LV Plunger C - 001434C","LV Plunger D - 001434C","HV Caddy","Plunger HV"],"mesin":"SUMITOMO 50T/100T/130T, JSW 80T/100T/220T"},
    "Flo PD (FLo PD)":{"parts":["Caddy Cover","Tip HV","Tip LV","Base Caddy","Top Caddy","Plunger LV","Plunger HV","Plunger Pos.Ring","Sleeve"],"mesin":"JSW 100T A/B, SUMITOMO 50T"},
    "Mantis":{"parts":["Tube support","Bottle Cap Top","Bottle Cap lower","Cap IMHVS_PI_AVA","Air Layer IMHVS_PI_AVA","Thin Layer IMHVS_PI_AVA","Thick Layer IMHVS_PI_AVA","Cap Mantis IMLVS_PI_AVA","Air Layer IMLVS_PI_AVA","Thin Layer Mantis IMLVS_PI_AVA","Thick Layer Mantis IMLVS_PI_AVA","Thin Layer IMCFPFE_PI","Air Layer IMCFPFE_PI","Thick Layer IMCFPFE_PI","Thin Liquid Layer IMHV3PFE Dome","Thick Liquid Layer IMHV3PFE Dome"],"mesin":"ARBURG 70T LSR"},
    "Mantis LVDS":{"parts":["Air layer V1","Air layer V2","Thin Liquid Layer V1","Thin Liquid Layer V2","Thick Liquid Layer V1","Thick Liquid Layer V2"],"mesin":"Mantis LVDS"},
    "Mantis HVDS":{"parts":["Air layer","Thin Liquid Layer","Thick Liquid Layer"],"mesin":"Mantis HVDS"},
    "Tempest":{"parts":["Tempest","Tempest LV","Tempest HV"],"mesin":"Silicon Molding"},
    "NT8":{"parts":["LV Tips","NT8 Chamber"],"mesin":"JSW 100T"},
    "HT96 Microplate":{"parts":["HT96 Microplate"],"mesin":"SUMITOMO 130T"},
    "Topas Formulator":{"parts":["Topas Blank","Formulator"],"mesin":"General"},
    "RX":{"parts":["RX","RI360"],"mesin":"General"},
    "Qiagen":{"parts":["Coding Frame"],"mesin":"Forindo 140T"},
    "Pulse":{"parts":["Acc","RI360","Stempede"],"mesin":"Silicon Molding"},
}

if "master_data" not in st.session_state:
    st.session_state.master_data={}
    for proj,pdata in MASTER_PROJECTS.items():
        for part in pdata["parts"]:
            st.session_state.master_data[f"{proj}||{part}"]={
                "issue":"Flashing on tail & stem areas" if "Plunger" in part or "Tip" in part else "General Issue - Flashing/Short Shot/Black Dot",
                "define":f"Flashing on {part} ({proj}) - Insert does not have axial support",
                "measure":"Day1 Minimal, Day2 clearly visible, Day3 increases. Baseline with 1.3mm jig",
                "analyze":"Mechanism: Insert no axial support, melt pressure push back, micro gap -> flashing",
                "improve":"Main Solution: Add support pillars behind insert, grind mold base for preload, trial Sample 1-10 x Cav 1-8 OK",
                "control":"Routine monitoring: Daily visual check per Cav 1-8 + 1.3mm jig, weekly audit PE & IME",
                "status":"Open"
            }
    if "FAST||Plunger LV" in st.session_state.master_data:
        st.session_state.master_data["FAST||Plunger LV"]["status"]="Closed - 0% with 1.3mm jig"

if "sel_proj" not in st.session_state: st.session_state.sel_proj="FAST"
if "sel_part" not in st.session_state: st.session_state.sel_part="Plunger LV"

st.title("📊 Quality AI PRO V6 FINAL + WI Comparison FIXED V2")
with st.container(border=True):
    c1,c2,c3,c4=st.columns([2,2,2,1.5])
    with c1:
        proj_list=list(MASTER_PROJECTS.keys())
        sel_proj=st.selectbox("1️⃣ PROJECT (dari handbook_ags.xlsx)", proj_list, index=proj_list.index(st.session_state.sel_proj) if st.session_state.sel_proj in proj_list else 0, key="v6f_proj")
        st.session_state.sel_proj=sel_proj
        st.caption(f"Mesin: {MASTER_PROJECTS[sel_proj]['mesin']}")
    with c2:
        part_list=MASTER_PROJECTS[sel_proj]["parts"]
        sel_part=st.selectbox(f"2️⃣ PART di {sel_proj} ({len(part_list)} part)", part_list, index=part_list.index(st.session_state.sel_part) if st.session_state.sel_part in part_list else 0, key="v6f_part")
        st.session_state.sel_part=sel_part
    with c3:
        key=f"{sel_proj}||{sel_part}"
        data=st.session_state.master_data.get(key,{})
        st.metric("3️⃣ Status Part", data.get("status","Open"))
        st.caption(f"Issue: {data.get('issue','-')}")
    with c4:
        if st.button("🔁 Reuse Issue Mirip ke Part Lain", use_container_width=True, key="v6f_reuse_btn"):
            st.session_state.show_reuse=True

if st.session_state.get("show_reuse", False):
    with st.container(border=True):
        st.markdown("### 🔁 Reuse Issue (Karena isu hampir sama di beberapa project/part)")
        sc1,sc2,sc3=st.columns(3)
        with sc1: src=st.selectbox("Copy dari", list(st.session_state.master_data.keys()), key="v6f_src")
        with sc2:
            tgt_proj=st.selectbox("Ke Project", proj_list, key="v6f_tproj")
            tgt_part=st.selectbox("Ke Part", MASTER_PROJECTS[tgt_proj]["parts"], key="v6f_tpart")
        with sc3:
            if st.button("📋 Copy Sekarang", type="primary", key="v6f_copy"):
                st.session_state.master_data[f"{tgt_proj}||{tgt_part}"]=st.session_state.master_data[src].copy()
                st.success(f"Copied {src} -> {tgt_proj}||{tgt_part}!")
                st.session_state.show_reuse=False; st.rerun()
            if st.button("Tutup", key="v6f_close"): st.session_state.show_reuse=False; st.rerun()

current_key=f"{sel_proj}||{sel_part}"
current_data=st.session_state.master_data.get(current_key,{})
st.info(f"🎯 Aktif: **{sel_proj} > {sel_part}** → 7 Tab auto-fill untuk part ini. Isu hampir sama bisa di-reuse 1 klik.")

def tab_dmaic():
    st.subheader(f"🌀 DMAIC Generator - {sel_proj} > {sel_part} - AUTO")
    c_up,c_link=st.columns(2)
    with c_up: up=st.file_uploader("Upload DMAIC (xlsx)", type=["xlsx","xls"], key=f"dmaic_up_{current_key}")
    with c_link: link=st.text_input("Paste Google Sheets Link", key=f"dmaic_link_{current_key}"); btn=st.button("Parse Link", key=f"dmaic_btn_{current_key}")
    source=up if up else (link if btn and 'docs.google.com' in link else None)
    if source:
        parsed=universal_parser(source)
        st.success(parsed["info"])
        if not parsed["cav_df"].empty: st.dataframe(parsed["cav_df"], use_container_width=True)
        current_data.update({"define":parsed["define"] or current_data.get("define",""),"measure":parsed["measure"] or current_data.get("measure",""),"analyze":parsed["analyze"] or current_data.get("analyze",""),"improve":parsed["improve"] or current_data.get("improve",""),"control":parsed["control"] or current_data.get("control","")})
        if not parsed["pareto"].empty: current_data["pareto"]=parsed["pareto"]
        st.session_state.master_data[current_key]=current_data
    no=st.text_input("DMAIC No", f"DMAIC-{sel_proj}-{sel_part}-{datetime.now().strftime('%Y%m%d')}", key=f"no_{current_key}")
    issue=st.text_input("Issue", current_data.get("issue","Flashing/Short Shot/Black Dot"), key=f"issue_{current_key}")
    define=st.text_area("1. DEFINE", current_data.get("define",""), height=100, key=f"def_{current_key}")
    measure=st.text_area("2. MEASURE - Day1 Minimal etc + 1.3mm jig", current_data.get("measure",""), height=80, key=f"meas_{current_key}")
    if "pareto" in current_data and isinstance(current_data["pareto"], pd.DataFrame) and not current_data["pareto"].empty and current_data["pareto"]["Qty"].sum()>0:
        st.plotly_chart(px.bar(current_data["pareto"], x="Cavity", y="Qty", title=f"Pareto {sel_part}"), use_container_width=True)
    analyze=st.text_area("3. ANALYZE", current_data.get("analyze",""), height=80, key=f"ana_{current_key}")
    improve=st.text_area("4. IMPROVE - Add support pillars, grind base", current_data.get("improve",""), height=80, key=f"imp_{current_key}")
    control=st.text_area("5. CONTROL - Daily visual + jig", current_data.get("control",""), height=80, key=f"ctrl_{current_key}")
    st.session_state.master_data[current_key]={"issue":issue,"define":define,"measure":measure,"analyze":analyze,"improve":improve,"control":control,"status":current_data.get("status","Open"),"pareto":current_data.get("pareto", pd.DataFrame())}
    if st.button(f"📄 Generate DMAIC DOCX - {sel_part}", type="primary", use_container_width=True, key=f"gen_{current_key}"):
        doc=Document(); doc.add_heading(f"DMAIC {sel_proj} {sel_part} - {issue}",1); doc.add_paragraph(f"No:{no}\nProject:{sel_proj} Part:{sel_part}\nDEFINE:{define}\nMEASURE:{measure}\nANALYZE:{analyze}\nIMPROVE:{improve}\nCONTROL:{control}")
        bio=io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Download DOCX", bio, file_name=f"{no}_{sel_proj}_{sel_part}.docx", use_container_width=True, key=f"dl_{current_key}")

def tab_msa():
    st.subheader(f"📏 MSA 44 Rows - {sel_proj} > {sel_part}")
    df=pd.DataFrame([{"Sample":f"S{i//8+1}","Cavity":f"Cav {i%8+1}","Part":sel_part,"Project":sel_proj,"Result":"OK","Jig 1.3mm":"PASS"} for i in range(44)])
    st.dataframe(df, use_container_width=True)

def tab_wi():
    st.subheader(f"📄 WI Generator - {sel_proj} > {sel_part} - TEMPLATE SAMA")
    st.caption("Template WI sama kayak kemarin, auto-fill project>part")
    d=current_data
    method=st.text_area("Control (dari Tab DMAIC)", d.get("control","Daily visual per Cav 1-8 + 1.3mm jig, weekly audit PE & IME"), height=100, key=f"wi_{current_key}")
    wi_txt=f"""WORK INSTRUCTION - {sel_proj} - {sel_part}
Project: {sel_proj} | Part: {sel_part} | Mesin: {MASTER_PROJECTS[sel_proj]['mesin']}
1. Purpose: Prevent {d.get('issue','flashing')}
2. Tool: 1.3mm jig go/no-go (template sama)
3. Method: {method}
4. Sampling: Sample 1-10 x Cav 1-8 (44 rows)
5. Criteria: PASS with 1.3mm jig, 0% flashing
6. Reference: handbook_ags.xlsx
"""
    st.code(wi_txt)
    if st.button(f"Generate WI DOCX - {sel_part}", type="primary", use_container_width=True, key=f"wigen_{current_key}"):
        doc=Document(); doc.add_heading(f"WI {sel_proj} {sel_part}",1); doc.add_paragraph(wi_txt)
        bio=io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Download WI", bio, file_name=f"WI_{sel_proj}_{sel_part}.docx", use_container_width=True, key=f"widl_{current_key}")

def tab_screenshot():
    st.subheader(f"📸 Screenshot - {sel_proj} > {sel_part}")
    st.file_uploader(f"Before {sel_part}", type=["jpg","png"], key=f"ssb_{current_key}")
    st.file_uploader(f"After {sel_part}", type=["jpg","png"], key=f"ssa_{current_key}")

def tab_trial():
    st.subheader(f"🧪 Trial Report - {sel_proj} > {sel_part}")
    st.text_area(f"Trial for {sel_part}", f"Add support pillars, grind base for {sel_part}", height=80, key=f"trial_{current_key}")

def tab_catalog():
    st.subheader("📚 Catalog All Issue - Isu Hampir Sama di Beberapa Project/Part")
    all_rows=[]
    for k,v in st.session_state.master_data.items():
        if "||" in k:
            pj,pt=k.split("||")
            all_rows.append({"Project":pj,"Part":pt,"Issue":v.get("issue",""),"Status":v.get("status",""),"Root":v.get("analyze","")[:60]})
    df=pd.DataFrame(all_rows)
    if not df.empty:
        c1,c2=st.columns(2)
        with c1: f_proj=st.selectbox("Filter Project", ["All"]+sorted(df["Project"].unique().tolist()), key="cat_fproj")
        with c2: f_issue=st.text_input("Cari Issue (Flashing/Short/Black Dot)", key="cat_fissue")
        if f_proj!="All": df=df[df["Project"]==f_proj]
        if f_issue: df=df[df["Issue"].str.contains(f_issue, case=False, na=False)]
        st.dataframe(df, use_container_width=True, height=400)
        st.metric("Total Part dengan Issue", len(df))
    buf=io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w: pd.DataFrame(all_rows).to_excel(w, index=False)
    st.download_button("📥 Export Catalog All Project>Part (Excel)", buf.getvalue(), file_name=f"Catalog_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True, key="cat_exp")

# =================== TAB 7: WI COMPARISON FIXED V2 - EXACT MOCKUP ===================
def tab_wi_comparison():
    st.subheader(f"🔍 WI Comparison - OLD vs NEW - {sel_proj} > {sel_part} - FIXED V2 EXACT MOCKUP")
    st.caption("Upload OLD WI & NEW WI (docx/pdf), hasil persis kayak mockup referensi: OLD abu, NEW hijau, s merah + panah biru melengkung")

    c_old, c_new = st.columns(2)
    with c_old:
        old_file = st.file_uploader("📤 Upload OLD WI", type=["docx","pdf"], key=f"wi_old_{current_key}")
        st.markdown("<div style='background:#f2f2f2;border:1px solid #999;padding:10px;text-align:center;font-weight:bold;color:#555'>📄 OLD WI</div>", unsafe_allow_html=True)
    with c_new:
        new_file = st.file_uploader("📤 Upload NEW WI", type=["docx","pdf"], key=f"wi_new_{current_key}")
        st.markdown("<div style='background:#e6f4ea;border:1px solid #999;padding:10px;text-align:center;font-weight:bold;color:#0a4a8a'>📝 NEW WI</div>", unsafe_allow_html=True)

    if old_file and new_file:
        from PIL import Image, ImageDraw

        def extract_text(f):
            try:
                f.seek(0)
                if f.name.lower().endswith('.docx'):
                    doc=Document(f)
                    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()!=''])
                elif f.name.lower().endswith('.pdf'):
                    import fitz
                    data=f.read()
                    pdf=fitz.open(stream=data, filetype="pdf")
                    txt=""
                    for pg in pdf: txt+=pg.get_text()+"\n"
                    return txt
                else: return f.name
            except Exception as e: return str(e)

        old_text = extract_text(old_file)
        new_text = extract_text(new_file)
        old_words = old_text.split()
        new_words = new_text.split()
        sm = difflib.SequenceMatcher(None, old_words, new_words)
        changes=[]
        for tag,i1,i2,j1,j2 in sm.get_opcodes():
            if tag!='equal':
                changes.append({"tag":tag,"old":" ".join(old_words[i1:i2]),"new":" ".join(new_words[j1:j2])})

        st.divider()
        if changes:
            st.error(f"🔎 Ditemukan {len(changes)} perubahan untuk {sel_part}:")
            for idx,ch in enumerate(changes[:10],1):
                if ch['tag']=='replace': st.markdown(f"**{idx}. REPLACE:** `{ch['old'][:60]}` → `<span style='color:red;font-weight:bold'>{ch['new'][:60]}</span>`", unsafe_allow_html=True)
                elif ch['tag']=='insert': st.markdown(f"**{idx}. INSERT (baru):** `<span style='color:red;font-weight:bold'>{ch['new'][:60]}</span>` - huruf merah + panah biru", unsafe_allow_html=True)

        W,H = 1200, 700
        img = Image.new('RGB', (W,H), 'white')
        draw = ImageDraw.Draw(img)
        draw.rectangle([(0,0),(W,120)], fill="white", outline="#cccccc", width=1)
        draw.text((25,35), "FORMULATRIX", fill="#0e5a9a")
        draw.rectangle([(350,10),(W-15,110)], outline="#888888", width=1)
        draw.line([(560,10),(560,110)], fill="#888888", width=1)
        draw.line([(770,10),(770,110)], fill="#888888", width=1)
        draw.line([(350,50),(W-15,50)], fill="#888888", width=1)
        draw.text((375,20), "Document No", fill="#888888")
        draw.text((580,20), "Product", fill="#888888")
        draw.text((830,20), "Process Desc", fill="#888888")
        draw.text((375,65), "WIN-13-001-0046", fill="black")
        draw.text((590,65), "General", fill="black")
        draw.text((800,65), "Yusin Robot Operation", fill="black")
        draw.rectangle([(0,120),(W,160)], fill="#b8d0f0")
        draw.text((220,132), "Work Instruction Comparison - OLD WI vs NEW WI | Mockup for demonstration only", fill="#1a2a5a")
        draw.rectangle([(25,185),(W//2-20,235)], fill="#f2f2f2", outline="#aaaaaa", width=1)
        draw.text((80,195), "OLD WI", fill="#666666")
        draw.rectangle([(25,250),(W//2-20,520)], fill="white", outline="#cccccc", width=1)
        draw.text((40,265), "Title:", fill="#888888")
        draw.text((40,295), "WIN-Injection-General-0046-", fill="black")
        draw.text((40,325), "Yusin Robot Operation", fill="black")
        draw.rectangle([(35,365),(W//2-35,405)], fill="#f8f8f8", outline="#dddddd", width=1)
        draw.text((45,375), "Revision: Rev 02 * Effective: 2023-11-15", fill="black")
        draw.text((40,420), "No change to operation sequence. Standard", fill="#666666")
        draw.text((40,440), "procedure for robot operation.", fill="#666666")
        draw.rectangle([(W//2+20,185),(W-25,235)], fill="#d9eada", outline="#aaaaaa", width=1)
        draw.text((W//2+80,195), "NEW WI", fill="#0b5394")
        draw.rectangle([(W//2+20,250),(W-25,520)], fill="#eaf1fb", outline="#6d9eeb", width=1)
        draw.text((W//2+35,265), "Title:", fill="#888888")
        draw.text((W//2+35,295), "WIN-Injection-General-0046-", fill="black")
        draw.text((W//2+35,325), "Yusin Robot", fill="black")
        draw.text((W//2+190,325), "Operation", fill="black")
        draw.text((W//2+290,325), "s", fill="#cc0000")
        draw.line([(W//2+80,480),(W//2+120,450),(W//2+200,380),(W//2+290,345)], fill="#0b5394", width=4, joint="curve")
        draw.polygon([(W//2+290,330),(W//2+275,355),(W//2+305,350)], fill="#0b5394")
        draw.text((W//2+50,495), 'Added plural -> "Operations"', fill="#0b5394")
        draw.rectangle([(0,H-55),(W,H)], fill="#f2f2f2", outline="#cccccc", width=1)
        draw.text((15,H-35), 'Change Summary: Updated title to plural "Operations" to align with multi-robot procedures. * Change ID: CHG-0046-2024 * Status: Approved', fill="#333333")

        st.image(img, caption="✅ Hasil EXACT MOCKUP - OLD abu, NEW hijau, s merah + panah biru lengkung persis referensi", use_container_width=True)

        buf=io.BytesIO()
        img.save(buf, format='PNG')
        st.download_button("📥 Download Screenshot Comparison EXACT MOCKUP (PNG) - s merah + panah biru", buf.getvalue(), file_name=f"WI_Comparison_EXACT_{sel_proj}_{sel_part}_{datetime.now().strftime('%Y%m%d')}.png", type="primary", use_container_width=True, key=f"dl_exact_{current_key}")

        doc=Document()
        doc.add_heading(f"WI Comparison Report - {sel_proj} {sel_part}",1)
        doc.add_paragraph(f"OLD: {old_file.name}\nNEW: {new_file.name}\nChanges: {len(changes)}\n")
        for ch in changes[:20]:
            doc.add_paragraph(f"{ch['tag']}: {ch['old']} -> {ch['new']}")
        bio=io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Download Report Perubahan (DOCX)", bio, file_name=f"WI_Comparison_Report_{sel_proj}_{sel_part}.docx", use_container_width=True, key=f"dl_docx_{current_key}")
    else:
        st.info("👆 Upload OLD WI & NEW WI (docx/pdf) - Hasil akan persis kayak mockup: OLD WI abu, NEW WI hijau, Operation s merah + panah biru melengkung + Added plural -> Operations")

t1,t2,t3,t4,t5,t6,t7=st.tabs(["🌀 DMAIC","📏 MSA 44 Rows","📄 WI Generator","📸 Screenshot","🧪 Trial Report","📚 Catalog All Issue","🔍 WI Comparison FIXED V2"])
with t1: tab_dmaic()
with t2: tab_msa()
with t3: tab_wi()
with t4: tab_screenshot()
with t5: tab_trial()
with t6: tab_catalog()
with t7: tab_wi_comparison()
st.success(f"V6 FINAL + WI COMPARISON FIXED V2 - {sel_proj} > {sel_part} - Hasil EXACT MOCKUP - Edit 1 tab, tab lain AMAN")
