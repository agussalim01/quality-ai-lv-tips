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

st.set_page_config(page_title="Quality AI PRO - 6 Modules Auto Chart FINAL UNIVERSAL", page_icon="📊", layout="wide")

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
                    draw.rectangle([W//2+15, y_right-2, W-15, y_right+20], fill="#c8e6c9")
                    draw.text((W//2+20, y_right), f"+ {txt}", fill="#1b5e20", font=font_body)
                    y_right += 26
            max_y = max(y_left, y_right)
            y_left = y_right = max_y + 10
    img.save(output_path)
    return output_path

def comparator_tab():
    st.subheader("🖼️ Screenshot Kiri Kanan")
    c1, c2 = st.columns(2)
    with c1:
        old_file = st.file_uploader("📂 WI LAMA", type=["docx"], key="old_wi_univ")
    with c2:
        new_file = st.file_uploader("📂 WI BARU", type=["docx"], key="new_wi_univ")
    if old_file and new_file:
        old_list = extract_all(old_file)
        new_list = extract_all(new_file)
        sim = difflib.SequenceMatcher(None, old_list, new_list).ratio()*100
        st.metric("Kemiripan", f"{sim:.1f}%")
        img_path = create_side_screenshot(old_list, new_list)
        st.divider()
        st.image(img_path, caption="Kiri = Lama | Kanan = Baru", use_container_width=True)

def wi_tab():
    st.subheader("📄 WI Generator - Auto Template")
    template_file = None
    for f in os.listdir("."):
        if "TMP" in f and f.endswith(".docx"):
            template_file = f
            break
    if not template_file:
        template_file = st.file_uploader("Upload Template (.docx)", type=["docx"], key="wi_template_univ")
    if template_file:
        st.success(f"Template: {template_file}")
    c1, c2 = st.columns(2)
    with c1:
        doc_no = st.text_input("Document No", "WIN-13-005-0015", key="wi_docno_univ")
        product = st.text_input("Product", "NT8", key="wi_prod_univ")
    with c2:
        proc = st.text_input("Process Desc", "QC 100% Visual Inspection LV Tips", key="wi_proc_univ")
    time_df = st.data_editor(pd.DataFrame([{"step": "Step 1 :  Susun Tips", "time": 10},{"step": "Step 2 :  Check visual defect", "time": 80}]), num_rows="dynamic", use_container_width=True, key="time_editor_univ")
    raw = st.text_area("Detail Langkah Kerja", "Step 1 :  Susun Tips\nSusun LV tips", height=200, key="wi_raw_univ")
    if st.button("✨ Generate WI FINAL", type="primary", use_container_width=True, key="gen_wi_univ"):
        doc = create_wi_final(template_file, doc_no, product, "TMP-11-001-0002", "0", proc, "1 batch", "Lampu", time_df, raw)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.download_button("📥 Download WI FINAL", data=bio, file_name=f"{doc_no}_FINAL.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

def msa_tab():
    st.subheader("📊 MSA Dashboard - FIX 44 BARIS + ISOLATED")
    if "msa_df_fix" not in st.session_state:
        st.session_state.msa_df_fix = pd.DataFrame({"Part": [1,1,1,2,2,2,3,3,3]*3, "Operator": ["A"]*9 + ["B"]*9 + ["C"]*9, "Measurement": [10.1,10.12,10.11,10.2,10.22,10.21,10.0,10.02,10.01,10.11,10.13,10.12,10.21,10.23,10.22,10.01,10.03,10.02,10.09,10.11,10.10,10.19,10.21,10.20,9.99,10.01,10.0]})
    uploaded = st.file_uploader("Upload Data MSA (.csv/.xlsx) - AUTO deteksi header berantakan", type=["csv","xlsx","xls"], key="msa_upload_univ")
    if uploaded:
        try:
            if uploaded.name.endswith('.csv'):
                temp_df = pd.read_csv(uploaded, header=None, nrows=20)
                header_row = 0
                for i, row in temp_df.iterrows():
                    row_str = ' '.join([str(x).lower() for x in row.values])
                    if 'part' in row_str and 'operator' in row_str:
                        header_row = i
                        break
                uploaded.seek(0)
                df = pd.read_csv(uploaded, header=header_row)
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            else:
                temp_df = pd.read_excel(uploaded, header=None, nrows=20)
                header_row = 0
                for i, row in temp_df.iterrows():
                    row_str = ' '.join([str(x).lower() for x in row.values])
                    if 'part' in row_str or 'measurement' in row_str:
                        header_row = i
                        break
                df = pd.read_excel(uploaded, header=header_row)
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            col_map = {}
            for c in df.columns:
                cl = str(c).lower()
                if 'part' in cl: col_map[c] = 'Part'
                elif 'operator' in cl or 'appraiser' in cl: col_map[c] = 'Operator'
                elif 'measur' in cl or 'value' in cl or 'result' in cl: col_map[c] = 'Measurement'
            df.rename(columns=col_map, inplace=True)
            if 'Part' in df.columns:
                st.session_state.msa_df_fix = df.dropna(subset=['Part']).head(100)
                st.success(f"✅ Terbaca {len(st.session_state.msa_df_fix)} baris FIX 44 BARIS")
            else:
                st.error(f"Kolom Part tidak ketemu. Kolom: {list(df.columns)}")
        except Exception as e:
            st.error(f"Error MSA: {e}")
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(st.session_state.msa_df_fix.head(15), use_container_width=True)
    with c2:
        if not st.session_state.msa_df_fix.empty and 'Measurement' in st.session_state.msa_df_fix.columns:
            fig = px.histogram(st.session_state.msa_df_fix, x="Measurement", color="Operator" if 'Operator' in st.session_state.msa_df_fix.columns else None, title="MSA Distribution")
            st.plotly_chart(fig, use_container_width=True)

def trial_report_tab():
    st.subheader("🧪 Trial Report - ISOLATED FIX")
    if "trial_manual_data" not in st.session_state:
        st.session_state.trial_manual_data = pd.DataFrame([{"Sample": "S001", "Parameter": "Visual", "Result": "OK", "Notes": ""},{"Sample": "S002", "Parameter": "Visual", "Result": "NG - Flash", "Notes": "Flash di tip"}])
    edited_trial = st.data_editor(st.session_state.trial_manual_data, num_rows="dynamic", use_container_width=True, key="editor_trial_manual_univ_final")
    st.session_state.trial_manual_data = edited_trial
    if st.button("📄 Generate Trial Report FINAL", type="primary", use_container_width=True, key="gen_trial_univ"):
        doc = Document()
        doc.add_heading("TRIAL REPORT", 1)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.download_button("📥 Download Trial Report", data=bio, file_name="Trial_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

def catalog_reject_tab():
    st.subheader("📚 Catalog Reject Injection - FULL - ISOLATED")
    catalog = [
        {"name": "SHORT SHOT", "alias": "Isi Tidak Penuh", "level": "CRITICAL", "desc": "Part tidak terisi penuh.", "criteria": "REJECT 100%"},
        {"name": "FLASH", "alias": "Selaput di parting line", "level": "MAJOR", "desc": "Kelebihan material tipis.", "criteria": "REJECT jika >0.1mm"},
        {"name": "BLACK DOT", "alias": "Titik hitam", "level": "CRITICAL", "desc": "Titik hitam kecil di part clear.", "criteria": "REJECT 100% LED 6W"},
        {"name": "BURN MARK", "alias": "Gosong", "level": "CRITICAL", "desc": "Warna coklat/hitam terbakar.", "criteria": "REJECT 100%"},
        {"name": "SINK MARK", "alias": "Cekungan", "level": "MINOR-MAJOR", "desc": "Permukaan cekung.", "criteria": "REJECT jika di sealing"},
    ]
    search = st.text_input("🔍 Cari defect", "", key="cat_search_univ")
    for item in catalog:
        if search and search.lower() not in item["name"].lower(): continue
        with st.expander(f"🔴 {item['name']} - {item['alias']} [{item['level']}]"):
            st.error(f"Kriteria: {item['criteria']}")
            st.write(item['desc'])

def dmaic_tab():
    st.subheader("🌀 DMAIC Generator - UNIVERSAL ALL PROJECT - FINAL")
    st.caption("Bisa untuk NT8, project lain, Daily, Monthly - tinggal mapping kolom - FIX 0 defect")
    if "dmaic_measure_data" not in st.session_state:
        st.session_state.dmaic_measure_data = pd.DataFrame([
            {"DefectType": "BLACK DOT", "Qty": 45, "Date": "2024-11-01"},
            {"DefectType": "FLASH", "Qty": 30, "Date": "2024-11-01"},
            {"DefectType": "SHORT SHOT", "Qty": 15, "Date": "2024-11-02"},
        ])
    if "dmaic_ba_data" not in st.session_state:
        st.session_state.dmaic_ba_data = pd.DataFrame([
            {"Phase": "Before", "Defect": "BLACK DOT", "Qty": 45, "Yield": 94.8},
            {"Phase": "After", "Defect": "BLACK DOT", "Qty": 5, "Yield": 98.2},
        ])
    st.markdown("### 2️⃣ M - MEASURE - UNIVERSAL PARSER (ALL PROJECT)")
    uploaded = st.file_uploader("Upload data defect - SEMUA PROJECT (csv/xlsx 4MB+ ok) - Daily NT8 / Project lain", type=["csv","xlsx","xls"], key="universal_dmaic_upload_final")
    df_m = st.session_state.dmaic_measure_data
    if uploaded:
        try:
            if uploaded.name.endswith('.csv'):
                xls_dict = {"CSV": pd.read_csv(uploaded)}
            else:
                xls = pd.ExcelFile(uploaded)
                xls_dict = {name: xls.parse(name) for name in xls.sheet_names}
            sheet_names = list(xls_dict.keys())
            selected_sheet = st.selectbox("Pilih Sheet", sheet_names, key="sheet_selector_universal_final")
            raw_df = xls_dict[selected_sheet].copy()
            raw_df = raw_df.loc[:, ~raw_df.columns.astype(str).str.contains('^Unnamed')]
            raw_df = raw_df.dropna(how='all')
            st.write(f"**Preview Sheet: {selected_sheet} - {raw_df.shape[0]} baris x {raw_df.shape[1]} kolom**")
            st.dataframe(raw_df.head(15), use_container_width=True)
            st.divider()
            st.markdown("#### 🔧 Mapping Kolom - Pilih sesuai file project kamu (BIAR GAK 0 LAGI)")
            cols = ["-- Tidak Ada / Auto Count --"] + list(raw_df.columns.astype(str))
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                defect_col = st.selectbox("Kolom Jenis Defect / NG", cols, key="map_defect_final")
            with c2:
                qty_col = st.selectbox("Kolom Qty / Jumlah (kalau ada)", cols, key="map_qty_final")
            with c3:
                status_col = st.selectbox("Kolom Status OK/NG (kalau ada)", cols, key="map_status_final")
            with c4:
                date_col = st.selectbox("Kolom Tanggal (opsional)", cols, key="map_date_final")
            mode = st.radio("Mode Hitung", ["Count NG (hitung baris NG) - UNIVERSAL", "Sum Qty (jumlahkan kolom Qty)", "Auto - Deteksi OK/NG"], horizontal=True, key="mode_universal_final")
            if st.button("✅ Proses Data Jadi Pareto", type="primary", use_container_width=True, key="proses_universal_final"):
                processed = pd.DataFrame()
                try:
                    if mode == "Sum Qty (jumlahkan kolom Qty)" and qty_col != "-- Tidak Ada / Auto Count --" and defect_col != "-- Tidak Ada / Auto Count --":
                        raw_df[qty_col] = pd.to_numeric(raw_df[qty_col], errors='coerce').fillna(0)
                        processed = raw_df.groupby(defect_col)[qty_col].sum().reset_index()
                        processed.columns = ["DefectType", "Qty"]
                    elif status_col != "-- Tidak Ada / Auto Count --":
                        tmp = raw_df.copy()
                        tmp[status_col] = tmp[status_col].astype(str).str.upper()
                        ng_df = tmp[tmp[status_col].str.contains("NG|REJECT|N.G|NOT OK|FAIL")]
                        if defect_col != "-- Tidak Ada / Auto Count --" and not ng_df.empty:
                            processed = ng_df.groupby(defect_col).size().reset_index(name="Qty")
                            processed.columns = ["DefectType", "Qty"]
                        else:
                            if not ng_df.empty:
                                processed = pd.DataFrame([{"DefectType": "TOTAL NG", "Qty": len(ng_df)}])
                    elif defect_col != "-- Tidak Ada / Auto Count --":
                        filtered = raw_df[raw_df[defect_col].notna()]
                        filtered = filtered[~filtered[defect_col].astype(str).str.strip().isin(["", "nan", "NaN", "OK", "GOOD"])]
                        filtered = filtered[~filtered[defect_col].astype(str).str.upper().isin(["OK", "GOOD", "PASS"])]
                        processed = filtered.groupby(defect_col).size().reset_index(name="Qty")
                        processed.columns = ["DefectType", "Qty"]
                        processed = processed[processed["Qty"] > 0]
                    if not processed.empty:
                        processed["Date"] = selected_sheet
                        processed = processed[processed["DefectType"].astype(str).str.strip() != ""]
                        processed = processed[processed["Qty"] > 0]
                        st.session_state.dmaic_measure_data = processed
                        df_m = processed
                        st.success(f"✅ Berhasil! Total Defect = {int(processed['Qty'].sum())} - Bukan 0 lagi!")
                        st.dataframe(processed, use_container_width=True)
                    else:
                        st.error("Hasil masih 0. Coba ganti kolom Defect atau mode Count NG")
                except Exception as e:
                    st.error(f"Error proses: {e}")
        except Exception as e:
            st.error(f"Error baca file: {e}")
    else:
        edited_m = st.data_editor(st.session_state.dmaic_measure_data, num_rows="dynamic", use_container_width=True, key="editor_dmaic_measure_universal_final2")
        st.session_state.dmaic_measure_data = edited_m
        df_m = edited_m
    st.divider()
    if not df_m.empty and "DefectType" in df_m.columns and "Qty" in df_m.columns:
        total = int(df_m["Qty"].sum())
        c1, c2 = st.columns(2)
        c1.metric("Total Defect", total)
        if total == 0:
            st.error("Total masih 0 - cek mapping kolom di atas!")
        else:
            top = df_m.groupby("DefectType")["Qty"].sum().sort_values(ascending=False).head(1)
            if not top.empty:
                c2.metric("Top Defect", f"{top.index[0]} ({top.values[0]})")
            df_pareto = df_m.groupby("DefectType")["Qty"].sum().reset_index().sort_values("Qty", ascending=False)
            df_pareto["Cumulative"] = df_pareto["Qty"].cumsum()
            df_pareto["Cumulative%"] = df_pareto["Cumulative"] / df_pareto["Qty"].sum() * 100
            fig_combined = make_subplots(specs=[[{"secondary_y": True}]])
            fig_combined.add_trace(go.Bar(x=df_pareto["DefectType"], y=df_pareto["Qty"], name="Qty", marker_color="#ef5350"), secondary_y=False)
            fig_combined.add_trace(go.Scatter(x=df_pareto["DefectType"], y=df_pareto["Cumulative%"], name="Cumulative %", mode="lines+markers", line=dict(color="#1e88e5", width=3)), secondary_y=True)
            fig_combined.add_hline(y=80, line_dash="dash", line_color="green", secondary_y=True)
            fig_combined.update_layout(title=f"Pareto 80/20 - Total {total} (All Project)")
            st.plotly_chart(fig_combined, use_container_width=True)

st.title("📊 Quality AI PRO - 6 Modules Auto Chart - FINAL UNIVERSAL")
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
