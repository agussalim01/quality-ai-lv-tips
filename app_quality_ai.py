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

st.set_page_config(page_title="Quality AI PRO - 5 Modules", page_icon="📊", layout="wide")

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
            if "Result" in df_result.columns or "Status" in df_result.columns:
                col_name = "Result" if "Result" in df_result.columns else "Status"
                ok_count = len(df_result[df_result[col_name].astype(str).str.contains("OK|PASS|Good", case=False, na=False)])
                ng_count = len(df_result) - ok_count
                yield_pct = ok_count / len(df_result) * 100 if len(df_result) > 0 else 0
                cA, cB, cC = st.columns(3)
                cA.metric("Total", len(df_result))
                cB.metric("OK", ok_count)
                cC.metric("Yield", f"{yield_pct:.1f}%")
                fig = px.pie(values=[ok_count, ng_count], names=["OK", "NG"], title="Hasil Trial")
                st.plotly_chart(fig, use_container_width=True)
    with tab_manual:
        manual_df = st.data_editor(pd.DataFrame([
            {"Sample": "S001", "Parameter": "Visual", "Result": "OK", "Notes": ""},
            {"Sample": "S002", "Parameter": "Visual", "Result": "NG - Flash", "Notes": "Flash di tip"},
            {"Sample": "S003", "Parameter": "Visual", "Result": "OK", "Notes": ""},
        ]), num_rows="dynamic", use_container_width=True, key="trial_manual")
        if df_result is None:
            df_result = manual_df
    analysis = st.text_area("🧠 Analisa", "Dari hasil trial, defect utama adalah Flash (X%). Penyebab kemungkinan...", height=130, key="tr_analysis")
    conclusion = st.text_area("✅ Conclusion & Recommendation", "Trial PASS dengan yield X%. Rekomendasi: Lanjutkan ke mass production / Perlu perbaikan...", height=100, key="tr_conc")
    if st.button("📄 Generate Trial Report FINAL", type="primary", use_container_width=True, key="gen_trial"):
        doc = Document()
        title = doc.add_heading(f"TRIAL REPORT", 1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Trial No: {trial_no} | Product: {product} | Date: {trial_date} | PIC: {pic}").alignment = WD_ALIGN_PARAGRAPH.CENTER
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
        doc.add_heading("6. Conclusion & Recommendation", 2)
        doc.add_paragraph(conclusion)
        doc.add_heading("7. Approval", 2)
        appr = doc.add_table(rows=2, cols=3)
        appr.style = 'Table Grid'
        appr.cell(0,0).text = "Prepared by"
        appr.cell(0,1).text = "Checked by"
        appr.cell(0,2).text = "Approved by"
        appr.cell(1,0).text = pic
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success(f"Trial Report {trial_no} jadi!")
        st.download_button("📥 Download Trial Report", data=bio, file_name=f"{trial_no}_Trial_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

def catalog_reject_tab():
    st.subheader("📚 Catalog Reject Injection")
    st.info("Katalog ini berisi defect injection umum untuk LV Tips & general injection. Klik untuk detail + action. Upload foto asli kamu untuk menggantikan placeholder.")
    
    # Buat folder catalog jika belum ada
    os.makedirs("catalog_images", exist_ok=True)
    
    catalog = [
        {"name": "SHORT SHOT", "alias": "Isi Tidak Penuh / Tip Kosong", "level": "CRITICAL", "desc": "Part tidak terisi penuh, ujung tip kosong/patah. Sering terjadi di LV Tips clear.", "cause": ["Material kurang", "Suhu barrel rendah", "Injection pressure rendah", "Mold venting buntu", "Nozzle buntu"], "action": ["Cek hopper & material feed", "Naikkan barrel temp +10-15C", "Naikkan injection pressure 5-10%", "Bersihkan venting mold", "Purging nozzle"], "criteria": "REJECT 100% - Sort lot"},
        {"name": "FLASH / BURR", "alias": "Selaput tipis di parting line", "level": "MAJOR", "desc": "Kelebihan material tipis di sekeliling part, bahaya masuk ke produk.", "cause": ["Clamping force kurang", "Mold aus/rusak", "Injection pressure terlalu tinggi", "Mold kotor"], "action": ["Naikkan clamping force", "Cek kondisi mold - repair jika aus", "Turunkan injection pressure/speed", "Bersihkan parting line", "Trial 5 shots"], "criteria": "REJECT jika flash >0.1mm"},
        {"name": "SINK MARK", "alias": "Cekungan / Lekukan permukaan", "level": "MINOR-MAJOR", "desc": "Permukaan ada cekungan seperti kawah di area tebal.", "cause": ["Holding pressure kurang", "Cooling time kurang", "Shrinkage tinggi", "Gate kecil"], "action": ["Naikkan holding pressure & time", "Tambah cooling time 2-3 detik", "Turunkan barrel temp", "Cek gate design"], "criteria": "REJECT jika di sealing area"},
        {"name": "BURN MARK", "alias": "Gosong / Hitam kecoklatan", "level": "CRITICAL", "desc": "Warna coklat/hitam terbakar karena gas terjebak (diesel effect).", "cause": ["Venting buntu - gas trap", "Injection speed terlalu tinggi", "Suhu terlalu tinggi"], "action": ["STOP - Panggil teknisi", "Bersihkan/tambah venting", "Turunkan injection speed 20%", "Turunkan suhu barrel"], "criteria": "REJECT 100% Kontaminasi"},
        {"name": "WELD LINE", "alias": "Garis sambungan aliran", "level": "MAJOR", "desc": "Garis terlihat di part, kekuatan lemah.", "cause": ["Suhu rendah", "Speed rendah", "Mold dingin"], "action": ["Naikkan barrel & mold temp", "Naikkan injection speed", "Cek gate position"], "criteria": "REJECT jika di critical area"},
        {"name": "BLACK DOT / CONTAMINATION", "alias": "Titik hitam", "level": "CRITICAL untuk LV Tips", "desc": "Titik hitam kecil di part clear - defect paling sering di LV Tips.", "cause": ["Material kontaminasi", "Barrel kotor/degradasi", "Purging kurang", "Hopper kotor"], "action": ["Purging 10 shots", "Bersihkan hopper & dryer", "Ganti batch material", "Cek screw & barrel"], "criteria": "REJECT 100% visual under LED 6W"},
        {"name": "SPLAY / SILVER STREAK", "alias": "Garis perak / Moisture", "level": "MAJOR", "desc": "Garis perak di permukaan karena moisture.", "cause": ["Material basah", "Drying kurang", "Suhu terlalu tinggi"], "action": ["Cek dryer - pastikan <0.02% moisture", "Tambah drying time 2 jam", "Turunkan nozzle temp", "Cek material storage"], "criteria": "REJECT jika terlihat jelas"},
        {"name": "WARPAGE", "alias": "Bengkok / Tidak rata", "level": "MAJOR", "desc": "Part bengkok / tidak rata setelah eject.", "cause": ["Cooling tidak merata", "Ejector tidak balance", "Shrinkage tidak merata"], "action": ["Cek cooling channel", "Balance cooling time", "Cek ejector pin", "Turunkan mold temp"], "criteria": "REJECT jika melebihi tolerance"},
    ]
    
    search = st.text_input("🔍 Cari defect", "", key="cat_search")
    level_filter = st.selectbox("Filter Level", ["ALL", "CRITICAL", "MAJOR", "MINOR-MAJOR"], key="cat_filter")
    
    for item in catalog:
        if search and search.lower() not in item["name"].lower() and search.lower() not in item["alias"].lower():
            continue
        if level_filter != "ALL" and level_filter not in item["level"]:
            continue
            
        with st.expander(f"🔴 {item['name']} - {item['alias']} [{item['level']}]", expanded=False):
            c1, c2 = st.columns([1, 1.5])
            with c1:
                # Cek apakah ada foto asli di catalog_images
                custom_path = f"catalog_images/{item['name'].replace(' ', '_').replace('/', '_')}.jpg"
                if os.path.exists(custom_path):
                    st.image(custom_path, caption=f"Foto Asli: {item['name']}", use_container_width=True)
                else:
                    st.info(f"📸 Contoh: {item['name']}\n\nUpload foto asli kamu dengan nama:\n{custom_path}\n\nUntuk mengganti placeholder ini.")
                    # Tampilkan upload di dalam expander
                    up = st.file_uploader(f"Upload foto {item['name']}", type=["jpg","png","jpeg"], key=f"up_{item['name']}")
                    if up:
                        with open(custom_path, "wb") as f:
                            f.write(up.getbuffer())
                        st.success("Foto tersimpan! Refresh halaman")
                        st.image(custom_path, use_container_width=True)
                st.error(f"Kriteria: {item['criteria']}")
            with c2:
                st.write(f"**Deskripsi:** {item['desc']}")
                st.write("**⚠️ Penyebab Umum:**")
                for cause in item["cause"]:
                    st.write(f"- {cause}")
                st.write("**✅ Action / Rekomendasi:**")
                for act in item["action"]:
                    st.write(f"● {act}")
                st.warning("Action QC: Sort 100% lot yang sama, simpan sample NG untuk review Engineering & Mold")
    
    st.divider()
    st.subheader("📤 Upload Foto Reject Asli Massal")
    st.caption("Upload banyak foto sekaligus, nanti masuk ke catalog_images")
    multi = st.file_uploader("Upload foto-foto reject (bisa banyak)", type=["jpg","png","jpeg"], accept_multiple_files=True, key="multi_upload")
    if multi:
        for f in multi:
            path = f"catalog_images/{f.name}"
            with open(path, "wb") as out:
                out.write(f.getbuffer())
        st.success(f"{len(multi)} foto tersimpan di catalog_images/")

# ================= MAIN =================
st.title("📊 Quality AI PRO - 5 Modules")
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 MSA", "📄 WI Generator", "🖼️ Screenshot Compare", "🧪 Trial Report", "📚 Catalog Reject"])
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
