import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io, re, os, difflib
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="Quality AI PRO - 6 Modules Auto Chart", page_icon="📊", layout="wide")

# ================= HELPER WI GENERATOR =================
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
    W, H = 1600, 1300
    img = Image.new("RGB", (W, H), "#f0f0f0")
    draw = ImageDraw.Draw(img)
    try:
        font_title = ImageFont.truetype("arial.ttf", 22)
        font_body = ImageFont.truetype("arial.ttf", 16)
    except:
        font_title = ImageFont.load_default()
        font_body = ImageFont.load_default()
    draw.rectangle([10, 10, W//2-10, H-10], fill="white", outline="#999", width=2)
    draw.rectangle([W//2+10, 10, W-10, H-10], fill="white", outline="#999", width=2)
    draw.rectangle([10, 10, W//2-10, 50], fill="#263238")
    draw.text((20, 15), "WI LAMA", fill="white", font=font_title)
    draw.rectangle([W//2+10, 10, W-10, 50], fill="#2e7d32")
    draw.text((W//2+20, 15), "WI BARU", fill="white", font=font_title)
    y_left, y_right = 70, 70
    matcher = difflib.SequenceMatcher(None, old_list, new_list)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if y_left > H-50 or y_right > H-50:
            break
        if tag == 'equal':
            for k in range(i1, min(i2, i1+12)):
                txt = old_list[k][:58]
                draw.text((20, y_left), txt, fill="black", font=font_body)
                draw.text((W//2+20, y_right), txt, fill="black", font=font_body)
                y_left += 22
                y_right += 22
        else:
            if tag in ('replace', 'delete'):
                for k in range(i1, i2):
                    if y_left > H-60: break
                    txt = old_list[k][:55]
                    draw.rectangle([15, y_left-2, W//2-15, y_left+20], fill="#ffcdd2")
                    draw.text((20, y_left), f"- {txt}", fill="#b71c1c", font=font_body)
                    y_left += 26
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
            max_y = max(y_left, y_right)
            y_left = y_right = max_y + 10
    img.save(output_path)
    return output_path

# ================= TABS =================
def comparator_tab():
    st.subheader("🖼️ Screenshot Kiri Kanan")
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
        img_path = create_side_screenshot(old_list, new_list)
        st.divider()
        st.image(img_path, caption="Kiri = Lama | Kanan = Baru", use_container_width=True)
        with open(img_path, "rb") as f:
            st.download_button("📥 Download Screenshot PNG", data=f, file_name="WI_Kiri_Kanan.png", mime="image/png", type="primary", use_container_width=True)
    else:
        st.info("👆 Upload 2 WI dulu")

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
        template_file = st.file_uploader("Upload Template (.docx)", type=["docx"], key="wi_template")
    c1, c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Document No", "WIN-13-005-0015", key="wi_docno")
        product = st.text_input("Product", "NT8", key="wi_prod")
        temp_no = st.text_input("Temp Number", "TMP-11-001-0002 rev.1", key="wi_temp")
        rev = st.text_input("Rev", "0", key="wi_rev")
        proc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips", key="wi_proc")
    with c2:
        part_req = st.text_area("Part Required", "1 batch LV Tips after cutting", height=80, key="wi_part")
        tools_req = st.text_area("Tools Required", "Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer", height=120, key="wi_tools")
    st.write("**Time Table**")
    time_df = st.data_editor(pd.DataFrame([
        {"step": "Step 1 :  Susun Tips", "time": 10},
        {"step": "Step 2 :  Check visual defect", "time": 80},
        {"step": "Step 3 :  Data Tips OK dan Tips Reject", "time": 5},
        {"step": "Step 4 :  Packing LV Tips", "time": 5},
        {"step": "Step 5 : Manual Filling Process", "time": 30},
        {"step": "Step 6 : 2nd Visual Check", "time": 5},
        {"step": "Step 7 : Ionizer", "time": 5}
    ]), num_rows="dynamic", use_container_width=True, key="time_editor")
    raw = st.text_area("Detail Langkah Kerja", "Step 1 :  Susun Tips\nSusun LV tips di atas tatakan\n\nStep 2 :  Check visual defect\nCek dengan lampu LED 30cm", height=200, key="wi_raw")
    if st.button("✨ Generate WI FINAL", type="primary", use_container_width=True):
        if not template_file:
            st.error("Template tidak ditemukan! Upload TMP-11-001-0002")
            return
        doc = create_wi_final(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"WI {doc_no} jadi! Approval sudah di bawah.")
        st.download_button("📥 Download WI FINAL", data=bio, file_name=f"{doc_no}_FINAL.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

def msa_tab():
    st.subheader("📊 MSA Dashboard")
    uploaded = st.file_uploader("Upload Data MSA (.csv)", type=["csv"], key="msa_upload")
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

def trial_report_tab():
    st.subheader("🧪 Trial Report - Standard + Analisa Otomatis")
    c1, c2 = st.columns(2)
    with c1:
        trial_no = st.text_input("Trial No", "TR-2024-001", key="tr_no")
        product = st.text_input("Product", "NT8", key="tr_prod")
        trial_date = st.date_input("Trial Date", datetime.now(), key="tr_date")
        pic = st.text_input("PIC / Engineer", "", key="tr_pic")
        objective = st.text_area("Objective / Tujuan", "Validasi proses baru untuk mengurangi defect visual", height=100, key="tr_obj")
    with c2:
        parameters = st.text_area("Parameter Trial", "1. Suhu: 25C\n2. Speed: 100rpm\n3. Material: LV Tips batch A\n4. Lampu: LED 6W", height=100, key="tr_param")
        method = st.text_area("Method", "1. Siapkan 100 pcs sample\n2. Running QC visual\n3. Catat OK/NG\n4. Analisa defect", height=100, key="tr_method")
    st.divider()
    st.write("**📊 Hasil Trial**")
    tab_upload, tab_manual = st.tabs(["📤 Upload Excel/CSV", "✏️ Input Manual"])
    df_result = None
    with tab_upload:
        data_file = st.file_uploader("Upload hasil trial", type=["csv", "xlsx"], key="trial_data")
        if data_file:
            if data_file.name.endswith(".csv"):
                df_result = pd.read_csv(data_file)
            else:
                df_result = pd.read_excel(data_file)
            st.dataframe(df_result, use_container_width=True)
    with tab_manual:
        manual_df = st.data_editor(pd.DataFrame([
            {"Sample": "S001", "Parameter": "Visual", "Result": "OK", "Notes": ""},
            {"Sample": "S002", "Parameter": "Visual", "Result": "NG - Flash", "Notes": "Flash di tip"},
            {"Sample": "S003", "Parameter": "Visual", "Result": "OK", "Notes": ""},
        ]), num_rows="dynamic", use_container_width=True, key="trial_manual")
        if df_result is None:
            df_result = manual_df
    analysis = st.text_area("🧠 Analisa", "Dari hasil trial, defect utama adalah Flash (X%).", height=130, key="tr_analysis")
    conclusion = st.text_area("✅ Conclusion", "Trial PASS dengan yield X%.", height=100, key="tr_conc")
    if st.button("📄 Generate Trial Report FINAL", type="primary", use_container_width=True, key="gen_trial"):
        doc = Document()
        doc.add_heading(f"TRIAL REPORT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Trial No: {trial_no} | Product: {product} | Date: {trial_date} | PIC: {pic}")
        doc.add_heading("1. Objective", 2)
        doc.add_paragraph(objective)
        doc.add_heading("2. Parameters", 2)
        doc.add_paragraph(parameters)
        doc.add_heading("3. Method", 2)
        doc.add_paragraph(method)
        doc.add_heading("4. Results", 2)
        if df_result is not None and not df_result.empty:
            t = doc.add_table(rows=1, cols=len(df_result.columns))
            t.style = 'Table Grid'
            for i, col in enumerate(df_result.columns):
                t.cell(0,i).text = str(col)
            for _, row in df_result.iterrows():
                rw = t.add_row().cells
                for i, val in enumerate(row):
                    rw[i].text = str(val)
        doc.add_heading("5. Analysis", 2)
        doc.add_paragraph(analysis)
        doc.add_heading("6. Conclusion", 2)
        doc.add_paragraph(conclusion)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.download_button("📥 Download Trial Report", data=bio, file_name=f"{trial_no}_Trial_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

def catalog_reject_tab():
    st.subheader("📚 Catalog Reject Injection")
    os.makedirs("catalog_images", exist_ok=True)
    catalog = [
        {"name": "SHORT SHOT", "alias": "Isi Tidak Penuh", "level": "CRITICAL", "desc": "Part tidak terisi penuh, ujung tip kosong/patah.", "cause": ["Material kurang", "Suhu barrel rendah", "Injection pressure rendah", "Venting buntu"], "action": ["Cek hopper", "Naikkan temp +10-15C", "Naikkan pressure 5-10%", "Bersihkan venting"], "criteria": "REJECT 100%"},
        {"name": "FLASH", "alias": "Selaput di parting line", "level": "MAJOR", "desc": "Kelebihan material tipis di sekeliling part.", "cause": ["Clamping kurang", "Mold aus", "Pressure tinggi"], "action": ["Naikkan clamping", "Repair mold", "Turunkan pressure"], "criteria": "REJECT jika >0.1mm"},
        {"name": "BLACK DOT", "alias": "Titik hitam", "level": "CRITICAL", "desc": "Titik hitam kecil di part clear - defect LV Tips paling sering.", "cause": ["Material kontaminasi", "Barrel kotor", "Purging kurang"], "action": ["Purging 10 shots", "Bersihkan hopper", "Ganti batch"], "criteria": "REJECT 100% LED 6W"},
        {"name": "BURN MARK", "alias": "Gosong", "level": "CRITICAL", "desc": "Warna coklat/hitam terbakar karena gas terjebak.", "cause": ["Venting buntu", "Speed tinggi", "Suhu tinggi"], "action": ["STOP - Panggil teknisi", "Bersihkan venting", "Turunkan speed 20%"], "criteria": "REJECT 100%"},
        {"name": "SINK MARK", "alias": "Cekungan", "level": "MINOR-MAJOR", "desc": "Permukaan cekung seperti kawah di area tebal.", "cause": ["Holding kurang", "Cooling kurang"], "action": ["Naikkan holding", "Tambah cooling 2-3 detik"], "criteria": "REJECT jika di sealing"},
    ]
    search = st.text_input("🔍 Cari defect", "", key="cat_search")
    for item in catalog:
        if search and search.lower() not in item["name"].lower():
            continue
        with st.expander(f"🔴 {item['name']} - {item['alias']} [{item['level']}]"):
            c1, c2 = st.columns([1, 1.5])
            with c1:
                st.info(f"Contoh: {item['name']}\nUpload foto asli ke catalog_images/{item['name']}.jpg")
                st.error(f"Kriteria: {item['criteria']}")
            with c2:
                st.write(f"**Desc:** {item['desc']}")
                st.write("**Cause:** " + ", ".join(item['cause']))
                st.write("**Action:**")
                for a in item["action"]:
                    st.write(f"● {a}")

def dmaic_tab():
    st.subheader("🌀 DMAIC Generator - Injection Issue + Auto Chart")
    st.caption("Auto Pareto 80/20, Before-After Chart, Trend - langsung jadi")
    
    st.markdown("### 1️⃣ D - DEFINE")
    c1, c2 = st.columns(2)
    with c1:
        dmaic_no = st.text_input("DMAIC No", f"DMAIC-{datetime.now().strftime('%Y%m%d')}-001", key="dmaic_no")
        product = st.text_input("Product / Part", "LV Tips NT8", key="dmaic_prod")
        issue_type = st.selectbox("Jenis Defect (dari Catalog)", ["SHORT SHOT", "FLASH / BURR", "BLACK DOT / CONTAMINATION", "BURN MARK", "SINK MARK", "WELD LINE", "SPLAY", "WARPAGE", "Other"], key="dmaic_issue")
        other_issue = st.text_input("Jika Other, tulis issue", "", key="dmaic_other")
        problem = st.text_area("Problem Statement", f"Terjadi peningkatan defect {issue_type} di line injection LV Tips. Defect rate X% di atas target Y%.", height=100, key="dmaic_prob")
    with c2:
        dmaic_date = st.date_input("Date", datetime.now(), key="dmaic_date")
        team = st.text_input("Team / PIC", "Quality, Engineering, Production, Mold Maintenance", key="dmaic_team")
        goal = st.text_area("Goal / Target", "Menurunkan defect rate dari X% menjadi <0.5% dalam 2 minggu.\nMeningkatkan yield dari 92% ke 98%", height=80, key="dmaic_goal")
        scope = st.text_area("Scope", "Line: Injection Machine #3\nMold: 32 cavity\nMaterial: PP Clear\nPeriode: 1-15 Nov 2024", height=100, key="dmaic_scope")
    
    final_issue = other_issue if other_issue else issue_type
    
    st.divider()
    st.markdown("### 2️⃣ M - MEASURE + Auto Pareto 80/20")
    m_col1, m_col2 = st.columns([1,1])
    with m_col1:
        measure_file = st.file_uploader("Upload data defect (.csv / .xlsx) - kolom: DefectType, Qty", type=["csv","xlsx"], key="dmaic_measure_file")
        if measure_file:
            if measure_file.name.endswith(".csv"):
                df_m = pd.read_csv(measure_file)
            else:
                df_m = pd.read_excel(measure_file)
        else:
            df_m = st.data_editor(pd.DataFrame([
                {"DefectType": "BLACK DOT", "Qty": 45, "Date": "2024-11-01"},
                {"DefectType": "FLASH", "Qty": 30, "Date": "2024-11-01"},
                {"DefectType": "SHORT SHOT", "Qty": 15, "Date": "2024-11-02"},
                {"DefectType": "SINK MARK", "Qty": 10, "Date": "2024-11-02"},
                {"DefectType": "BLACK DOT", "Qty": 50, "Date": "2024-11-02"},
                {"DefectType": "BURN MARK", "Qty": 5, "Date": "2024-11-03"},
            ]), num_rows="dynamic", use_container_width=True, key="dmaic_measure_edit")
        
        total_defect = df_m["Qty"].sum() if "Qty" in df_m.columns else 0
        c_m1, c_m2 = st.columns(2)
        c_m1.metric("Total Defect", total_defect)
        if not df_m.empty:
            top_defect = df_m.groupby("DefectType")["Qty"].sum().sort_values(ascending=False).head(1)
            if not top_defect.empty:
                c_m2.metric("Top Defect", f"{top_defect.index[0]} ({top_defect.values[0]})")
    
    with m_col2:
        if not df_m.empty and "DefectType" in df_m.columns and "Qty" in df_m.columns:
            df_pareto = df_m.groupby("DefectType")["Qty"].sum().reset_index().sort_values("Qty", ascending=False)
            df_pareto["Cumulative"] = df_pareto["Qty"].cumsum()
            df_pareto["Cumulative%"] = df_pareto["Cumulative"] / df_pareto["Qty"].sum() * 100
            
            fig_combined = make_subplots(specs=[[{"secondary_y": True}]])
            fig_combined.add_trace(go.Bar(x=df_pareto["DefectType"], y=df_pareto["Qty"], name="Qty", marker_color="#ef5350"), secondary_y=False)
            fig_combined.add_trace(go.Scatter(x=df_pareto["DefectType"], y=df_pareto["Cumulative%"], name="Cumulative %", mode="lines+markers", line=dict(color="#1e88e5", width=3)), secondary_y=True)
            fig_combined.add_hline(y=80, line_dash="dash", line_color="green", secondary_y=True, annotation_text="80%")
            fig_combined.update_layout(title="Pareto 80/20 - Auto", yaxis=dict(title="Qty"), yaxis2=dict(title="Cumulative %", range=[0,100]))
            st.plotly_chart(fig_combined, use_container_width=True)
            
            top_80 = df_pareto[df_pareto["Cumulative%"] <= 80]
            if not top_80.empty:
                st.success(f"🎯 Fokus: {', '.join(top_80['DefectType'].tolist())} = 80% masalah (Pareto Principle)")
        
        baseline = st.text_input("Baseline Yield / Defect Rate", "Defect 5.2%, Yield 94.8%", key="dmaic_baseline")
        measurement_method = st.text_area("How we measure", "Visual 100% under LED 6W 580LM, jarak 30cm", height=60, key="dmaic_meas_method")
    
    if "Date" in df_m.columns:
        st.write("**📈 Trend Defect Over Time - Auto**")
        try:
            df_trend = df_m.groupby(["Date", "DefectType"])["Qty"].sum().reset_index()
            fig_trend = px.line(df_trend, x="Date", y="Qty", color="DefectType", markers=True, title="Trend Defect Harian")
            st.plotly_chart(fig_trend, use_container_width=True)
        except:
            pass
    
    st.divider()
    st.markdown("### 3️⃣ A - ANALYZE (Fishbone + 5 Why)")
    a1, a2 = st.columns(2)
    with a1:
        man = st.text_area("Man", "Operator baru, training kurang", height=60, key="fish_man")
        machine = st.text_area("Machine", "Nozzle buntu, barrel temp tidak stabil", height=60, key="fish_machine")
        material = st.text_area("Material", "Material kontaminasi, moisture tinggi", height=60, key="fish_mat")
        method = st.text_area("Method", "Purging kurang, parameter tidak standar", height=60, key="fish_method")
        mold = st.text_area("Mold / Measurement", "Venting buntu, mold aus, gate kecil", height=60, key="fish_mold")
    with a2:
        why1 = st.text_input("Why 1", f"Kenapa terjadi {final_issue}?", key="why1")
        ans1 = st.text_input("Answer 1", "Karena material tidak mengisi penuh", key="ans1")
        why2 = st.text_input("Why 2", "Kenapa bisa begitu?", key="why2")
        ans2 = st.text_input("Answer 2", "Karena nozzle buntu", key="ans2")
        why3 = st.text_input("Why 3", "Kenapa bisa begitu?", key="why3")
        ans3 = st.text_input("Answer 3", "Karena purging kurang", key="ans3")
        why4 = st.text_input("Why 4", "Kenapa bisa begitu?", key="why4")
        ans4 = st.text_input("Answer 4", "Tidak ada checklist purging", key="ans4")
        why5 = st.text_input("Why 5 - ROOT CAUSE", "Root Cause", key="why5")
        ans5 = st.text_area("Root Cause", "Tidak ada standard purging dan PM venting mold tidak terjadwal", height=60, key="ans5")
    
    st.divider()
    st.markdown("### 4️⃣ I - IMPROVE + Auto Before-After Chart")
    i1, i2 = st.columns(2)
    with i1:
        improve_action = st.data_editor(pd.DataFrame([
            {"No": 1, "Action": "Purging standard 10 shots", "PIC": "Production", "Due": "Immediate", "Status": "Done"},
            {"No": 2, "Action": "Bersihkan venting mold", "PIC": "Mold Maint", "Due": "2 hari", "Status": "On Progress"},
            {"No": 3, "Action": "Checklist drying <0.02%", "PIC": "Quality", "Due": "3 hari", "Status": "Plan"},
            {"No": 4, "Action": "Training parameter standard", "PIC": "Engineering", "Due": "1 minggu", "Status": "Plan"},
        ]), num_rows="dynamic", use_container_width=True, key="improve_table")
        
        st.write("**📊 Input Before vs After untuk Auto Chart**")
        before_after = st.data_editor(pd.DataFrame([
            {"Phase": "Before", "Defect": "BLACK DOT", "Qty": 45, "Yield": 94.8},
            {"Phase": "After", "Defect": "BLACK DOT", "Qty": 5, "Yield": 98.2},
            {"Phase": "Before", "Defect": "FLASH", "Qty": 30, "Yield": 94.8},
            {"Phase": "After", "Defect": "FLASH", "Qty": 8, "Yield": 98.2},
        ]), num_rows="dynamic", use_container_width=True, key="before_after")
    
    with i2:
        if not before_after.empty:
            fig_ba = px.bar(before_after, x="Defect", y="Qty", color="Phase", barmode="group", title="Before vs After - Defect Qty (Auto)", color_discrete_map={"Before": "#ef5350", "After": "#66bb6a"})
            st.plotly_chart(fig_ba, use_container_width=True)
            
            fig_yield = px.bar(before_after.drop_duplicates("Phase"), x="Phase", y="Yield", title="Yield Before vs After", color="Phase", color_discrete_map={"Before": "#ef5350", "After": "#66bb6a"})
            fig_yield.add_hline(y=98, line_dash="dash", line_color="green", annotation_text="Target 98%")
            st.plotly_chart(fig_yield, use_container_width=True)
            
            if len(before_after) >= 2:
                before_total = before_after[before_after["Phase"]=="Before"]["Qty"].sum()
                after_total = before_after[before_after["Phase"]=="After"]["Qty"].sum()
                if before_total > 0:
                    imp_pct = (before_total - after_total) / before_total * 100
                    st.success(f"🚀 Improvement: {imp_pct:.1f}% penurunan defect! Dari {before_total} ke {after_total}")
        
        improve_result = st.text_area("Hasil Improvement", "Setelah implementasi:\n- Defect BLACK DOT turun 89%\n- Yield naik dari 94.8% ke 98.2%", height=100, key="improve_result")
        cost = st.text_input("Cost / Benefit", "Cost: Purging 2kg, Benefit: Saving Rp 5jt/minggu", key="improve_cost")
    
    st.divider()
    st.markdown("### 5️⃣ C - CONTROL")
    c_c1, c_c2 = st.columns(2)
    with c_c1:
        control_plan = st.text_area("Control Plan", "1. Checklist purging wajib tiap ganti material\n2. PM venting tiap 2 minggu\n3. Monitoring defect hourly board\n4. Alert jika defect >1%", height=120, key="control_plan")
        monitoring = st.text_area("Monitoring", "Dashboard harian defect injection. Review mingguan Quality Meeting.", height=80, key="control_monitor")
    with c_c2:
        standardization = st.text_area("Standardization", "Update WI: WI-INJ-003 Rev 2\nUpdate Checklist Form\nUpdate PM Schedule", height=120, key="control_std")
        lesson = st.text_area("Lesson Learned", "Pentingnya purging standard & PM venting berkala", height=80, key="control_lesson")
    
    st.divider()
    if st.button("📄 Generate DMAIC Report FINAL + Charts", type="primary", use_container_width=True):
        doc = Document()
        h = doc.add_heading(f"DMAIC REPORT - {final_issue}", 1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"DMAIC No: {dmaic_no} | Product: {product} | Date: {dmaic_date} | Team: {team}")
        
        doc.add_heading("1. DEFINE", 2)
        doc.add_paragraph(f"Problem: {problem}\nGoal: {goal}\nScope: {scope}\nBaseline: {baseline}")
        
        doc.add_heading("2. MEASURE - Pareto", 2)
        doc.add_paragraph(f"Method: {measurement_method}")
        if not df_m.empty:
            t = doc.add_table(rows=1, cols=len(df_m.columns))
            t.style = 'Table Grid'
            for i, col in enumerate(df_m.columns):
                t.cell(0,i).text = str(col)
            for _, row in df_m.iterrows():
                rw = t.add_row().cells
                for i, val in enumerate(row):
                    rw[i].text = str(val)
        
        doc.add_heading("3. ANALYZE - Fishbone", 2)
        doc.add_paragraph(f"Man: {man}\nMachine: {machine}\nMaterial: {material}\nMethod: {method}\nMold: {mold}")
        doc.add_heading("5 Why - Root Cause", 2)
        doc.add_paragraph(f"{why1}: {ans1}\n{why2}: {ans2}\n{why3}: {ans3}\n{why4}: {ans4}\nROOT: {why5}: {ans5}")
        
        doc.add_heading("4. IMPROVE", 2)
        if not improve_action.empty:
            t = doc.add_table(rows=1, cols=len(improve_action.columns))
            t.style = 'Table Grid'
            for i, col in enumerate(improve_action.columns):
                t.cell(0,i).text = str(col)
            for _, row in improve_action.iterrows():
                rw = t.add_row().cells
                for i, val in enumerate(row):
                    rw[i].text = str(val)
        doc.add_paragraph(f"Result: {improve_result}\nCost: {cost}")
        
        doc.add_heading("5. CONTROL", 2)
        doc.add_paragraph(f"Control: {control_plan}\nMonitoring: {monitoring}\nStandardization: {standardization}\nLesson: {lesson}")
        
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"DMAIC {dmaic_no} jadi! Charts ada di dashboard, report .docx siap print.")
        st.download_button("📥 Download DMAIC Report", data=bio, file_name=f"{dmaic_no}_{final_issue}_DMAIC.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# ================= MAIN =================
st.title("📊 Quality AI PRO - 6 Modules Auto Chart")
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📊 MSA", "📄 WI Generator", "🖼️ Screenshot Compare", "🧪 Trial Report", "📚 Catalog Reject", "🌀 DMAIC Auto Chart"])
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
