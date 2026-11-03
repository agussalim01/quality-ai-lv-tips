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

st.set_page_config(page_title="Quality AI PRO - 7 Modules AUTO DETECT", page_icon="📊", layout="wide")

def create_wi_final(template_path_or_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, steps_text):
    doc = Document(template_path_or_file)
    saved_tables = list(doc.tables)
    for tbl in saved_tables:
        try: tbl._tbl.getparent().remove(tbl._tbl)
        except: pass
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"WORK INSTRUCTION\n{proc}\nProduct: {product}\nDocument No: {doc_no}"); run.bold=True; run.font.size=Pt(14)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Doc: {doc_no} | Temp: {temp_no} | Rev: {rev} | Date: {datetime.now().strftime('%d/%m/%Y')}"); run.font.size=Pt(8); run.bold=True
    p = doc.add_paragraph(); run = p.add_run("Langkah Kerja"); run.bold=True; run.font.size=Pt(11)
    t2 = doc.add_table(rows=1, cols=2); t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    try: t2.style='Table Grid'
    except: pass
    t2.cell(0,0).text="Langkah Kerja"; t2.cell(0,1).text="Time (min)"
    total=0
    for _, row in time_df.iterrows():
        rw = t2.add_row().cells; rw[0].text=str(row['step']); rw[1].text=str(row['time'])
        try: total+=int(row['time'])
        except: pass
    rw = t2.add_row().cells; rw[0].text="Total"; rw[1].text=str(total)
    doc.add_paragraph("")
    p = doc.add_paragraph(); p.add_run("Part required :").bold=True
    for it in part_req.split('\n'):
        if it.strip(): doc.add_paragraph(f"●  {it.strip()}")
    p = doc.add_paragraph(); p.add_run("Tools required :").bold=True
    for it in tools_req.split('\n'):
        if it.strip(): doc.add_paragraph(f"●  {it.strip()}")
    doc.add_paragraph("")
    pattern=r"(Step\s*\d+\s*:.*)"; parts=re.split(pattern, steps_text, flags=re.I); steps=[]
    for i in range(1, len(parts), 2):
        title=parts[i].strip(); body=parts[i+1] if i+1 < len(parts) else ""; bullets=[x.strip() for x in body.split('\n') if x.strip()]; steps.append((title,bullets))
    if not steps: steps=[("Work Steps",[x for x in steps_text.split('\n') if x.strip()])]
    for title, bullets in steps:
        p=doc.add_paragraph(); p.add_run(title).bold=True
        for b in bullets:
            if b: clean=re.sub(r'^[•\-\*]\s*','',b); doc.add_paragraph(f"●  {clean}")
    doc.add_page_break()
    for tbl in saved_tables: doc._body._element.append(tbl._tbl); doc.add_paragraph("")
    try:
        hist_table=doc.tables[-1]; hist_table.cell(2,1).text=datetime.now().strftime("%m/%d/%Y"); hist_table.cell(2,2).text=rev; hist_table.cell(2,3).text=f"{proc} - {product}"
    except: pass
    return doc

def extract_all(docx_file):
    doc=Document(docx_file); items=[]
    for p in doc.paragraphs:
        t=p.text.strip()
        if t: items.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t=cell.text.strip()
                if t and len(t)>1: items.append(t)
    return items

def create_side_screenshot(old_list, new_list, output_path="/tmp/side_by_side.png"):
    W,H=1600,1300; img=Image.new("RGB",(W,H),"#f0f0f0"); draw=ImageDraw.Draw(img)
    try: font_title=ImageFont.truetype("arial.ttf",22); font_body=ImageFont.truetype("arial.ttf",16)
    except: font_title=ImageFont.load_default(); font_body=ImageFont.load_default()
    draw.rectangle([10,10,W//2-10,H-10],fill="white",outline="#999",width=2); draw.rectangle([W//2+10,10,W-10,H-10],fill="white",outline="#999",width=2)
    draw.rectangle([10,10,W//2-10,50],fill="#263238"); draw.text((20,15),"WI LAMA",fill="white",font=font_title)
    draw.rectangle([W//2+10,10,W-10,50],fill="#2e7d32"); draw.text((W//2+20,15),"WI BARU",fill="white",font=font_title)
    y_left,y_right=70,70; matcher=difflib.SequenceMatcher(None,old_list,new_list)
    for tag,i1,i2,j1,j2 in matcher.get_opcodes():
        if y_left>H-50 or y_right>H-50: break
        if tag=='equal':
            for k in range(i1,min(i2,i1+12)): txt=old_list[k][:58]; draw.text((20,y_left),txt,fill="black",font=font_body); draw.text((W//2+20,y_right),txt,fill="black",font=font_body); y_left+=22; y_right+=22
        else:
            if tag in ('replace','delete'):
                for k in range(i1,i2):
                    if y_left>H-60: break; txt=old_list[k][:55]; draw.rectangle([15,y_left-2,W//2-15,y_left+20],fill="#ffcdd2"); draw.text((20,y_left),f"- {txt}",fill="#b71c1c",font=font_body); y_left+=26
            if tag in ('replace','insert'):
                for k in range(j1,j2):
                    if y_right>H-60: break; txt=new_list[k][:55]; is_ddff="ddff" in txt.lower(); bg="#ff0000" if is_ddff else "#c8e6c9"; fg="white" if is_ddff else "#1b5e20"; draw.rectangle([W//2+15,y_right-2,W-15,y_right+20],fill=bg); draw.text((W//2+20,y_right),f"+ {txt}",fill=fg,font=font_body); y_right+=26
            max_y=max(y_left,y_right); y_left=y_right=max_y+10
    img.save(output_path); return output_path

def msa_tab():
    st.subheader("📊 MSA - AUTO DETECT Variabel & Attribute [FIX UNTUK FILE KAMU]")
    st.caption("File kamu GRNR...SA FL.csv = Attribute Gage R&R (NG/G) - Akan auto convert")
    uploaded=st.file_uploader("Upload Data MSA (.csv)",type=["csv"],key="msa_upload_final")
    try:
        if not uploaded:
            df=pd.DataFrame({"Part":[1,1,1,2,2,2,3,3,3]*3,"Operator":["A"]*9+["B"]*9+["C"]*9,"Measurement":[10.1,10.12,10.11,10.2,10.22,10.21,10.0,10.02,10.01,10.11,10.13,10.12,10.21,10.23,10.22,10.01,10.03,10.02,10.09,10.11,10.10,10.19,10.21,10.20,9.99,10.01,10.0]})
            st.dataframe(df.head(10),use_container_width=True)
            st.info("Upload file GRNR Ji...SA FL.csv kamu di sini")
            return
        
        # Cek apakah file Attribute
        raw_preview = pd.read_csv(uploaded, header=None, encoding='utf-8-sig', nrows=15).astype(str)
        raw_text = " ".join(raw_preview.values.flatten()).upper()
        is_attribute = "ATTRIBUTE GAGE" in raw_text or "KNOWN POPULATION" in raw_text or "TRIAL #1" in raw_text or "SAMPLE #" in raw_text
        
        uploaded.seek(0)
        
        if is_attribute:
            st.info("🔍 Terdeteksi format Attribute Gage R&R - Convert...")
            df_raw = pd.read_csv(uploaded, header=None, encoding='utf-8-sig')
            # Cari baris header Sample #
            header_row_idx = 0
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                if "SAMPLE" in row_str and ("ATTRIBUTE" in row_str or "KNOWN" in row_str or "TRIAL" in row_str):
                    header_row_idx = i
                    break
            
            # Ambil nama operator dari baris sebelumnya
            operator_names = []
            if header_row_idx>0:
                op_row = df_raw.iloc[header_row_idx-1].tolist()
                # Baris Known Population biasanya: None, Known Population, None, PE, None, Intern...
                # Kita ambil yang bukan None
                for x in op_row:
                    xs = str(x).strip()
                    if xs.lower() not in ['none','nan','','known population'] and len(xs)>1:
                        operator_names.append(xs)
            if not operator_names:
                # Cari baris lagi
                for i in range(max(0, header_row_idx-2), header_row_idx):
                    row = df_raw.iloc[i].tolist()
                    for x in row:
                        xs = str(x).strip()
                        if xs.upper() in ['PE','INTERN','OPERATOR A','OPERATOR B','QA','QC','APPRAISER']:
                            if xs not in operator_names: operator_names.append(xs)
            if not operator_names:
                operator_names = ["PE","Intern"]
            
            st.write(f"Operator terdeteksi: {operator_names}")
            
            # Header dan data
            header = df_raw.iloc[header_row_idx].tolist()
            data_rows = df_raw.iloc[header_row_idx+1:].reset_index(drop=True)
            
            long_rows=[]
            for _, row in data_rows.iterrows():
                try:
                    # Row format: None, Sample#, Attribute, Trial#1, Trial#2, Trial#1, Trial#2...
                    # Bersihkan
                    row_vals = [str(x).strip() if pd.notna(x) else "" for x in row.tolist()]
                    # Sample number biasanya di index 1
                    sample = ""
                    attr = ""
                    for idx, v in enumerate(row_vals):
                        if v.isdigit() or (v.replace('.','',1).isdigit()):
                            if not sample: sample = v
                        if v.upper() in ['NG','G','OK','A','R']:
                            if not attr: attr = v
                            break
                    # Lebih robust: ambil kolom 1 dan 2
                    if len(row_vals)>=3:
                        if row_vals[1] and row_vals[1].lower() not in ['none','nan','']:
                            sample = row_vals[1]
                        if len(row_vals)>=3 and row_vals[2] and row_vals[2].lower() not in ['none','nan','']:
                            attr = row_vals[2]
                    
                    if not sample or not attr or attr.lower() in ['none','nan','']: continue
                    if sample.lower() in ['none','nan','']: continue
                    
                    std_map = {"G":"OK", "NG":"NG", "OK":"OK", "A":"OK", "R":"NG"}
                    std_std = std_map.get(attr.upper(), attr.upper())
                    if std_std not in ["OK","NG"]: continue
                    
                    # Trials mulai dari kolom 3
                    col_idx = 3
                    op_idx = 0
                    while col_idx < len(row_vals):
                        val = row_vals[col_idx].strip().upper()
                        if val in ['G','NG','OK','A','R']:
                            op_name = operator_names[op_idx//2] if op_idx//2 < len(operator_names) else f"Op{op_idx//2+1}"
                            trial_no = (op_idx%2)+1
                            res = std_map.get(val, val)
                            if res in ["OK","NG"]:
                                long_rows.append({"Part": f"P-{sample}", "Operator": op_name, "Trial": trial_no, "Result": res, "Standard": std_std, "Defect": ""})
                            op_idx+=1
                        elif val not in ['', 'NONE','NAN']:
                            # Tetap hitung sebagai trial walau kosong
                            pass
                        col_idx+=1
                        # Safety break
                        if op_idx > 20: break
                except Exception as ex:
                    continue
            
            df_converted = pd.DataFrame(long_rows)
            if df_converted.empty:
                st.error("Gagal convert! Cek format file. Coba upload di tab AAA Visual dengan format Part,Operator,Result,Standard")
                st.dataframe(df_raw.head(10), use_container_width=True)
                return
            
            st.success(f"✅ Convert OK: {len(df_converted)} inspeksi, {df_converted['Part'].nunique()} part, {df_converted['Operator'].nunique()} operator")
            st.dataframe(df_converted.head(20), use_container_width=True)
            
            total=len(df_converted)
            df_converted["Result"]=df_converted["Result"].astype(str).str.upper()
            df_converted["Standard"]=df_converted["Standard"].astype(str).str.upper()
            accuracy=(df_converted["Result"]==df_converted["Standard"]).sum()/total*100 if total>0 else 0
            
            c1,c2,c3=st.columns(3)
            c1.metric("Total Inspeksi", total)
            c2.metric("Accuracy vs Known", f"{accuracy:.1f}%")
            c3.metric("Mismatch", f"{100-accuracy:.1f}%")
            
            if accuracy>=90: st.success("✅ LULUS Attribute Gage R&R")
            elif accuracy>=80: st.warning("⚠️ MARGINAL")
            else: st.error("❌ GAGAL")
            
            # Pareto mismatch
            df_miss=df_converted[df_converted["Result"]!=df_converted["Standard"]]
            if not df_miss.empty:
                st.write("**Detail Miss:**")
                st.dataframe(df_miss, use_container_width=True)
                fig=px.bar(df_miss.groupby(["Part","Standard","Operator"]).size().reset_index(name="Miss Count"), x="Part", y="Miss Count", color="Operator", title="Miss per Part")
                st.plotly_chart(fig, use_container_width=True)
            
            st.session_state['aaa_from_msa'] = df_converted
            return
        else:
            # Variabel
            uploaded.seek(0)
            df=pd.read_csv(uploaded, encoding='utf-8-sig')
            df.columns=[c.strip().replace('\ufeff','') for c in df.columns]
            col_map={c.lower():c for c in df.columns}
            if 'measurement' in col_map and 'Measurement' not in df.columns: df.rename(columns={col_map['measurement']:'Measurement'},inplace=True)
            if 'operator' in col_map and 'Operator' not in df.columns: df.rename(columns={col_map['operator']:'Operator'},inplace=True)
            df=df.dropna(how='all')
            if df.empty: st.error("File kosong"); return
            st.dataframe(df.head(20),use_container_width=True)
            if "Measurement" in df.columns and "Operator" in df.columns:
                df["Measurement"]=pd.to_numeric(df["Measurement"],errors='coerce'); df_clean=df.dropna(subset=["Measurement"])
                if not df_clean.empty:
                    fig=px.histogram(df_clean,x="Measurement",color="Operator",title="MSA Distribution"); st.plotly_chart(fig,use_container_width=True)
            else: st.warning(f"Butuh Measurement & Operator. Kolom: {list(df.columns)}")
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback; st.code(traceback.format_exc())

def aaa_visual_tab():
    st.subheader("👁️ AAA Visual - FINAL FIX Anti 'Part' Error")
    st.caption("Format: Part,Operator,Trial,Result,Standard,Defect")
    # Jika ada data dari MSA tab
    if 'aaa_from_msa' in st.session_state and not st.session_state['aaa_from_msa'].empty:
        st.info("📥 Data dari tab MSA Attribute terdeteksi, pakai itu?")
        if st.checkbox("Pakai data dari MSA Tab", value=True):
            df_aaa = st.session_state['aaa_from_msa']
        else:
            df_aaa = pd.DataFrame()
    else:
        df_aaa = pd.DataFrame()
    
    aaa_file=st.file_uploader("Upload CSV AAA (opsional kalau sudah ada dari MSA)",type=["csv","xlsx"],key="aaa_file_final")
    try:
        if aaa_file:
            if aaa_file.name.endswith(".csv"): df_aaa=pd.read_csv(aaa_file, encoding='utf-8-sig')
            else:
                try: df_aaa=pd.read_excel(aaa_file)
                except ImportError: st.error("Butuh openpyxl, pakai CSV dulu"); df_aaa=pd.DataFrame()
            if not df_aaa.empty:
                df_aaa.columns=[str(c).strip().replace('\ufeff','') for c in df_aaa.columns]
                lower_map={str(c).lower():c for c in df_aaa.columns}
                rename_dict={}
                for k in ['part','operator','result','standard','trial','defect']:
                    if k in lower_map: rename_dict[lower_map[k]]=k.capitalize()
                df_aaa.rename(columns=rename_dict,inplace=True)
        if df_aaa.empty:
            df_aaa=pd.DataFrame([{"Part":"P01","Operator":"A","Trial":1,"Result":"NG","Standard":"NG","Defect":"BLACK DOT"},{"Part":"P01","Operator":"B","Trial":1,"Result":"OK","Standard":"NG","Defect":"BLACK DOT"}])
        
        st.success(f"Kolom: {list(df_aaa.columns)} - {len(df_aaa)} rows")
        df_aaa=st.data_editor(df_aaa, num_rows="dynamic", use_container_width=True, key="aaa_edit_main", height=350)
        
        if df_aaa.empty: return
        for col in ["Part","Operator","Result","Standard"]:
            if col not in df_aaa.columns:
                st.error(f"Kolom '{col}' tidak ada!"); return
        
        df_aaa["Result"]=df_aaa["Result"].astype(str).str.upper().str.strip()
        df_aaa["Standard"]=df_aaa["Standard"].astype(str).str.upper().str.strip()
        df_aaa["Part"]=df_aaa["Part"].astype(str).str.strip()
        df_aaa["Operator"]=df_aaa["Operator"].astype(str).str.strip()
        
        total=len(df_aaa)
        repeat_scores=[]
        for (part, op), group in df_aaa.groupby(["Part","Operator"]):
            if len(group)>1: repeat_scores.append(1 if group["Result"].nunique()==1 else 0)
        repeat_pct=(sum(repeat_scores)/len(repeat_scores)*100) if repeat_scores else 100.0
        accuracy=(df_aaa["Result"]==df_aaa["Standard"]).sum()/total*100 if total>0 else 0
        inter_agree=[]
        for part, group in df_aaa.groupby("Part"):
            try:
                op_results=group.groupby("Operator")["Result"].agg(lambda x: x.mode().iloc[0] if not x.mode().empty else x.iloc[0])
                inter_agree.append(1 if op_results.nunique()==1 else 0)
            except: inter_agree.append(0)
        inter_pct=(sum(inter_agree)/len(inter_agree)*100) if inter_agree else 0
        
        c1,c2,c3=st.columns(3); c1.metric("Repeatability",f"{repeat_pct:.1f}%"); c2.metric("Inter-Operator",f"{inter_pct:.1f}%"); c3.metric("Accuracy",f"{accuracy:.1f}%")
        if accuracy>=90 and repeat_pct>=90 and inter_pct>=90: st.success("✅ LULUS")
        elif accuracy>=80: st.warning("⚠️ MARGINAL")
        else: st.error("❌ GAGAL")
    except Exception as e:
        st.error(f"Error AAA: {e}"); import traceback; st.code(traceback.format_exc())

def comparator_tab():
    st.subheader("🖼️ Screenshot Kiri Kanan")
    c1,c2=st.columns(2)
    with c1: old_file=st.file_uploader("📂 WI LAMA",type=["docx"],key="old_wi")
    with c2: new_file=st.file_uploader("📂 WI BARU",type=["docx"],key="new_wi")
    if old_file and new_file:
        old_list=extract_all(old_file); new_list=extract_all(new_file)
        sim=difflib.SequenceMatcher(None,old_list,new_list).ratio()*100; st.metric("Kemiripan",f"{sim:.1f}%")
        img_path=create_side_screenshot(old_list,new_list); st.image(img_path,caption="Kiri = Lama | Kanan = Baru",use_container_width=True)
        with open(img_path,"rb") as f: st.download_button("📥 Download PNG",data=f,file_name="WI_Kiri_Kanan.png",mime="image/png",type="primary",use_container_width=True)
    else: st.info("👆 Upload 2 WI dulu")

def wi_tab():
    st.subheader("📄 WI Generator")
    template_file=None; found_name=None
    for f in os.listdir("."):
        if "TMP" in f and f.endswith(".docx"): found_name=f; template_file=f; break
    if template_file: st.success(f"✅ Template: {found_name}")
    else: template_file=st.file_uploader("Upload Template (.docx)",type=["docx"],key="wi_template")
    c1,c2=st.columns(2)
    with c1:
        doc_no=st.text_input("Document No","WIN-13-005-0015",key="wi_docno"); product=st.text_input("Product","NT8",key="wi_prod")
        temp_no=st.text_input("Temp Number","TMP-11-001-0002 rev.1",key="wi_temp"); rev=st.text_input("Rev","0",key="wi_rev"); proc=st.text_input("Process Desc","QC 100% Visual Inspection LV Tips",key="wi_proc")
    with c2:
        part_req=st.text_area("Part Required","1 batch LV Tips after cutting",height=80,key="wi_part"); tools_req=st.text_area("Tools Required","Lampu LED 6 watt 580 LM\nTatakan LV tips\nPinset\nIonizer",height=120,key="wi_tools")
    time_df=st.data_editor(pd.DataFrame([{"step":"Step 1 :  Susun Tips","time":10},{"step":"Step 2 :  Check visual defect","time":80},{"step":"Step 3 :  Data Tips OK dan Tips Reject","time":5},{"step":"Step 4 :  Packing LV Tips","time":5},{"step":"Step 5 : Manual Filling Process","time":30},{"step":"Step 6 : 2nd Visual Check","time":5},{"step":"Step 7 : Ionizer","time":5}]),num_rows="dynamic",use_container_width=True,key="time_editor")
    raw=st.text_area("Detail Langkah Kerja","Step 1 :  Susun Tips\nSusun LV tips di atas tatakan\n\nStep 2 :  Check visual defect\nCek dengan lampu LED 30cm",height=200,key="wi_raw")
    if st.button("✨ Generate WI FINAL",type="primary",use_container_width=True):
        if not template_file: st.error("Template tidak ada"); return
        doc=create_wi_final(template_file,doc_no,product,temp_no,rev,proc,part_req,tools_req,time_df,raw); bio=io.BytesIO(); doc.save(bio); bio.seek(0)
        st.success(f"WI {doc_no} jadi!"); st.download_button("📥 Download WI FINAL",data=bio,file_name=f"{doc_no}_FINAL.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)

def trial_report_tab():
    st.subheader("🧪 Trial Report")
    c1,c2=st.columns(2)
    with c1: trial_no=st.text_input("Trial No","TR-2024-001",key="tr_no"); product=st.text_input("Product","NT8",key="tr_prod"); trial_date=st.date_input("Trial Date",datetime.now(),key="tr_date"); pic=st.text_input("PIC","",key="tr_pic"); objective=st.text_area("Objective","Validasi proses baru",height=100,key="tr_obj")
    with c2: parameters=st.text_area("Parameter Trial","1. Suhu: 25C\n2. Speed: 100rpm",height=100,key="tr_param"); method=st.text_area("Method","1. Siapkan 100 pcs\n2. Running QC",height=100,key="tr_method")
    data_file=st.file_uploader("Upload hasil trial",type=["csv","xlsx"],key="trial_data")
    if data_file:
        if data_file.name.endswith(".csv"): df_result=pd.read_csv(data_file, encoding='utf-8-sig')
        else:
            try: df_result=pd.read_excel(data_file)
            except: df_result=pd.read_csv(data_file, encoding='utf-8-sig')
        st.dataframe(df_result,use_container_width=True)
    else: df_result=st.data_editor(pd.DataFrame([{"Sample":"S001","Parameter":"Visual","Result":"OK","Notes":""},{"Sample":"S002","Parameter":"Visual","Result":"NG - Flash","Notes":"Flash di tip"}]),num_rows="dynamic",use_container_width=True,key="trial_manual")
    analysis=st.text_area("🧠 Analisa","Defect utama Flash",height=130,key="tr_analysis"); conclusion=st.text_area("✅ Conclusion","Trial PASS",height=100,key="tr_conc")
    if st.button("📄 Generate Trial Report FINAL",type="primary",use_container_width=True,key="gen_trial"):
        doc=Document(); doc.add_heading(f"TRIAL REPORT",1).alignment=WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph(f"Trial No: {trial_no} | Product: {product} | Date: {trial_date} | PIC: {pic}")
        bio=io.BytesIO(); doc.save(bio); bio.seek(0); st.download_button("📥 Download Trial Report",data=bio,file_name=f"{trial_no}_Trial_Report.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)

def catalog_reject_tab():
    st.subheader("📚 Catalog Reject"); catalog=[{"name":"BLACK DOT","alias":"Titik hitam","level":"CRITICAL","desc":"Titik hitam di part clear","criteria":"REJECT 100% LED 6W"},{"name":"FLASH","alias":"Selaput","level":"MAJOR","desc":"Kelebihan material tipis","criteria":"REJECT jika >0.1mm"},{"name":"SHORT SHOT","alias":"Isi Tidak Penuh","level":"CRITICAL","desc":"Part tidak terisi penuh","criteria":"REJECT 100%"}]
    search=st.text_input("🔍 Cari defect","",key="cat_search")
    for item in catalog:
        if search and search.lower() not in item["name"].lower(): continue
        with st.expander(f"🔴 {item['name']} [{item['level']}]"): st.error(f"Kriteria: {item['criteria']}"); st.write(item["desc"])

def dmaic_tab():
    st.subheader("🌀 DMAIC Generator + Auto Chart")
    c1,c2=st.columns(2)
    with c1: dmaic_no=st.text_input("DMAIC No",f"DMAIC-{datetime.now().strftime('%Y%m%d')}-001",key="dmaic_no"); product=st.text_input("Product / Part","LV Tips NT8",key="dmaic_prod"); issue_type=st.selectbox("Jenis Defect",["BLACK DOT","FLASH","SHORT SHOT","Other"],key="dmaic_issue")
    with c2: dmaic_date=st.date_input("Date",datetime.now(),key="dmaic_date"); team=st.text_input("Team / PIC","Quality, Engineering",key="dmaic_team")
    measure_file=st.file_uploader("Upload data defect (.csv / .xlsx) - DefectType,Qty",type=["csv","xlsx"],key="dmaic_measure_file")
    try:
        if measure_file:
            if measure_file.name.endswith(".csv"): df_m=pd.read_csv(measure_file, encoding='utf-8-sig')
            else: df_m=pd.read_excel(measure_file)
            df_m.columns=[c.strip() for c in df_m.columns]
        else: df_m=pd.DataFrame([{"DefectType":"BLACK DOT","Qty":45},{"DefectType":"FLASH","Qty":30},{"DefectType":"SHORT SHOT","Qty":15}])
        df_m=st.data_editor(df_m,num_rows="dynamic",use_container_width=True,key="dmaic_measure_edit")
        if not df_m.empty and "DefectType" in df_m.columns and "Qty" in df_m.columns:
            df_m["Qty"]=pd.to_numeric(df_m["Qty"],errors='coerce'); df_m=df_m.dropna(subset=["Qty"])
            if not df_m.empty:
                df_pareto=df_m.groupby("DefectType")["Qty"].sum().reset_index().sort_values("Qty",ascending=False)
                df_pareto["Cumulative"]=df_pareto["Qty"].cumsum(); df_pareto["Cumulative%"]=df_pareto["Cumulative"]/df_pareto["Qty"].sum()*100
                fig_combined=make_subplots(specs=[[{"secondary_y": True}]])
                fig_combined.add_trace(go.Bar(x=df_pareto["DefectType"],y=df_pareto["Qty"],name="Qty",marker_color="#ef5350"),secondary_y=False)
                fig_combined.add_trace(go.Scatter(x=df_pareto["DefectType"],y=df_pareto["Cumulative%"],name="Cum %",mode="lines+markers",line=dict(color="#1e88e5",width=3)),secondary_y=True)
                fig_combined.update_layout(title="Pareto 80/20"); st.plotly_chart(fig_combined,use_container_width=True)
    except Exception as e: st.warning(f"Chart error: {e}")

# MAIN
st.title("📊 Quality AI PRO - AUTO DETECT ATTRIBUTE")
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📊 MSA AUTO","👁️ AAA Visual","📄 WI Generator","🖼️ Screenshot Compare","🧪 Trial Report","📚 Catalog Reject","🌀 DMAIC Auto Chart"])
with tab1: msa_tab()
with tab2: aaa_visual_tab()
with tab3: wi_tab()
with tab4: comparator_tab()
with tab5: trial_report_tab()
with tab6: catalog_reject_tab()
with tab7: dmaic_tab()
