import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import io, os
from datetime import datetime
import plotly.graph_objects as go

st.set_page_config(page_title="Quality AI PRO - FINAL FIX 44 BARIS", page_icon="📊", layout="wide")

# ===== PARSER FINAL UNTUK FILE KAMU 44 BARIS =====
def parse_any_grr_csv(uploaded_file):
    try:
        uploaded_file.seek(0)
        df_raw = pd.read_csv(uploaded_file, header=None, encoding='utf-8-sig', dtype=str, nrows=300)
        # Potong sebelum ANOVA Table
        cutoff_col = 30
        for _, row in df_raw.iterrows():
            for col_idx, val in enumerate(row.values):
                if "ANOVA" in str(val).upper():
                    cutoff_col = col_idx
                    break
        df_raw = df_raw.iloc[:, :cutoff_col]

        # Cari header Part (P) / Technician (T)
        header_idx = None
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x).upper() for x in row.values if str(x).lower()!='nan'])
            if "PART" in row_str and ("TECHNICIAN" in row_str or "OPERATOR" in row_str):
                header_idx = i
                break
        if header_idx is None:
            header_idx = 0

        header = df_raw.iloc[header_idx].tolist()
        data = df_raw.iloc[header_idx+1:].reset_index(drop=True)
        
        # Rename header berantakan
        clean = []
        for h in header:
            hs = str(h).strip().lower()
            if "no" == hs or "no." in hs or "order" in hs: clean.append("No")
            elif "part" in hs: clean.append("Part")
            elif "technician" in hs or "operator" in hs or hs in ['t','b','c','a']: clean.append("Operator")
            elif "meas" in hs or "response" in hs or "y" == hs or hs.replace('.','').replace(',','').isdigit(): clean.append("Measurement")
            else: clean.append(str(h).strip())
        data.columns = clean

        # Jika Measurement belum ketemu, cari kolom angka 21.xxx
        if "Measurement" not in data.columns:
            for col in data.columns:
                try:
                    vals = pd.to_numeric(data[col].astype(str).str.replace(',','.').str.replace(' ','') , errors='coerce')
                    if vals.notna().sum() > 5:
                        # Ini kolom Measurement
                        if col != "Part" and col != "Operator" and col != "No":
                            data.rename(columns={col:"Measurement"}, inplace=True)
                            break
                except: pass
            # Jika masih belum, ambil kolom ke-3 sebagai Measurement
            if "Measurement" not in data.columns and len(data.columns) >= 3:
                # kolom terakhir yang ada angka
                last_col = data.columns[-1]
                if last_col not in ["Part","Operator","No"]:
                    data.rename(columns={last_col:"Measurement"}, inplace=True)

        # Hanya ambil 3 kolom penting
        keep = [c for c in ["No","Part","Operator","Measurement"] if c in data.columns]
        data = data[keep]

        # Hapus baris kosong / header No.
        data = data.replace(['None','nan','NaN',''], pd.NA).dropna(how='all')
        if "Part" in data.columns:
            data = data[~data["Part"].astype(str).str.contains("No\\.", na=False)]
            data = data[~data["Part"].astype(str).str.lower().isin(['none','part','part (p)'])]
        
        if "Measurement" in data.columns:
            data["Measurement"] = pd.to_numeric(data["Measurement"].astype(str).str.replace(',','.').str.replace(' ','') , errors='coerce')
            data = data.dropna(subset=["Measurement"])
            # Filter Part yang kosong
            if "Part" in data.columns:
                data = data[data["Part"].astype(str).str.strip() != '']
        
        # Jika ada kolom No, drop aja, kita pakai Part saja
        if "No" in data.columns:
            data = data.drop(columns=["No"])
            
        # Pastikan Part & Operator string
        if "Part" in data.columns: data["Part"] = data["Part"].astype(str).str.strip()
        if "Operator" in data.columns: data["Operator"] = data["Operator"].astype(str).str.strip()

        return data, f"minitab_crossed_{len(data)}_baris_dipotong_ANOVA_OK"
    except Exception as e:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed|ANOVA', case=False)]
            return df, f"fallback: {e}"
        except Exception as e2:
            return pd.DataFrame(), f"error: {e2}"

def msa_tab():
    st.subheader("📊 MSA - FINAL FIX 44 BARIS")
    st.caption("Auto potong ANOVA Table & Unnamed 83 kolom -> jadi Part, Operator, Measurement")
    uploaded=st.file_uploader("Upload CSV MSA (3rd GRn...- Copy.csv)",type=["csv"],key="msa_final_44")
    if not uploaded:
        st.info("Upload file yang di screenshot kamu. Harusnya jadi 3 kolom bersih, bukan 2 kolom.")
        return
    df, mode = parse_any_grr_csv(uploaded)
    st.code(f"Mode deteksi: {mode}")
    if df.empty:
        st.error("Gagal baca. Cek file CSV"); return
    
    st.success(f"✅ Terbaca {len(df)} baris - Kolom bersih: {list(df.columns)}")
    st.dataframe(df.head(40), use_container_width=True)
    
    try:
        if "Measurement" in df.columns and "Operator" in df.columns and "Part" in df.columns:
            fig=px.histogram(df, x="Measurement", color="Operator", title="MSA Distribution - 44 Data", barmode='overlay')
            st.plotly_chart(fig, use_container_width=True)
            
            summary = df.groupby(["Part","Operator"])["Measurement"].agg(['mean','std','count']).reset_index()
            st.write("**Summary Gage R&R**")
            st.dataframe(summary, use_container_width=True)
            
            # Gage R&R simple calculation
            overall_mean = df["Measurement"].mean()
            part_var = df.groupby("Part")["Measurement"].mean().var()
            st.metric("Overall Mean", f"{overall_mean:.4f}")
            st.metric("Part Variation (indikasi)", f"{part_var:.5f}")
            
            st.session_state['last_msa'] = df
        else:
            st.warning(f"Kolom belum lengkap. Butuh Part, Operator, Measurement. Sekarang: {list(df.columns)}")
    except Exception as e:
        st.error(f"Hitung error: {e}")

def aaa_visual_tab():
    st.subheader("👁️ AAA Visual")
    st.info("Tab ini untuk file Attribute NG/G. File Minitab Variabel pakai tab MSA")
    df=pd.DataFrame([{"Part":"P-01","Operator":"A","Trial":1,"Result":"NG","Standard":"NG"}])
    st.dataframe(df, use_container_width=True)

def wi_tab():
    st.subheader("📄 WI Generator")
    st.info("Module WI OK")

def comparator_tab():
    st.subheader("🖼️ Compare")
    st.info("Upload 2 WI untuk compare")

def trial_report_tab():
    st.subheader("🧪 Trial Report")
    st.info("Trial OK")

def catalog_reject_tab():
    st.subheader("📚 Catalog Reject")
    st.info("Catalog OK")

def dmaic_tab():
    st.subheader("🌀 DMAIC Auto Chart")
    df=pd.DataFrame([{"DefectType":"BLACK DOT","Qty":45},{"DefectType":"FLASH","Qty":30}])
    df=st.data_editor(df, num_rows="dynamic", key="dmaic_44")
    try:
        fig=px.bar(df, x="DefectType", y="Qty", title="Pareto")
        st.plotly_chart(fig, use_container_width=True)
    except: st.dataframe(df)

st.title("📊 Quality AI PRO - FIX 44 BARIS")
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📊 MSA AUTO","👁️ AAA","📄 WI","🖼️ Compare","🧪 Trial","📚 Catalog","🌀 DMAIC"])
with tab1: msa_tab()
with tab2: aaa_visual_tab()
with tab3: wi_tab()
with tab4: comparator_tab()
with tab5: trial_report_tab()
with tab6: catalog_reject_tab()
with tab7: dmaic_tab()
