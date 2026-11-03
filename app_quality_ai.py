import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Quality AI PRO - ISOLATED", layout="wide")
st.title("📊 Quality AI PRO - ISOLATED TABS (FIX 44 BARIS)")

# --- INISIALISASI MEMORY TERPISAH BIAR GAK GANGGU ---
if "msa_df" not in st.session_state:
    st.session_state.msa_df = pd.DataFrame()
if "aaa_df" not in st.session_state:
    st.session_state.aaa_df = pd.DataFrame({"Sample":[1,2,3], "Standard":["OK","NG","OK"], "Appraiser A":["OK","NG","OK"], "Appraiser B":["OK","OK","OK"]})
if "trial_df" not in st.session_state:
    st.session_state.trial_df = pd.DataFrame({"Trial":["Trial 1","Trial 2","Trial 3"],"Date":["2026-11-01","2026-11-02","2026-11-03"],"Qty Check":[50,100,100],"NG Qty":[2,1,0],"Status":["NG","NG","OK"],"Remark":["Burr found","Flash","OK to mass pro"]})
if "catalog_df" not in st.session_state:
    st.session_state.catalog_df = pd.DataFrame({"Defect Type":["BLACK DOT","FLASH","BURR","DENT","SCRATCH","SHORT MOLD","WARPAGE","DISCOLOR","GATE MARK","BUBBLE","CRACK","CONTAMINATION","PIN MARK","SPLAY","SINK MARK"],"Kriteria":["Titik hitam >0.2mm","Flash line","Sisa tajam","Cekungan","Goresan >3mm","Short","Deformasi","Warna off","Gate mark","Bubble","Crack","Kotoran","Pin mark","Splay","Sink"],"Qty":[0]*15,"Action":["Sortir","Rework","Scrap","Sortir","Sortir","Scrap","Sortir","Hold","OK/NG","Scrap","Scrap","Sortir","OK/NG","Hold","OK/NG"]})
if "dmaic_df" not in st.session_state:
    st.session_state.dmaic_df = pd.DataFrame({"DefectType":["BLACK DOT","FLASH","BURR","DENT","SCRATCH","SHORT MOLD","WARPAGE","DISCOLOR","GATE MARK","BUBBLE","CRACK","CONTAMINATION","PIN MARK","SPLAY","SINK MARK"],"Qty":[12,8,5,4,3,2,2,1,1,1,0,0,0,0,0]})

tabs = st.tabs(["MSA AUTO", "AAA", "WI", "Compare", "Trial", "Catalog", "DMAIC"])

with tabs[0]:
    st.subheader("MSA AUTO - 44 Baris FIXED - ISOLATED")
    uploaded = st.file_uploader("Upload MSA", type=["csv","xlsx","xls"], key="uploader_msa_isolated")
    if uploaded:
        st.session_state.msa_df = pd.read_csv(uploaded) if uploaded.name.endswith('.csv') else pd.read_excel(uploaded)
        st.write(f"Terbaca {len(st.session_state.msa_df)} baris")
        st.dataframe(st.session_state.msa_df.head(), key="msa_preview")

with tabs[1]:
    st.subheader("AAA - ISOLATED")
    # key berbeda = tidak akan mereset tab lain
    st.session_state.aaa_df = st.data_editor(st.session_state.aaa_df, num_rows="dynamic", key="editor_aaa_isolated")

with tabs[2]:
    st.subheader("WI - ISOLATED")
    st.write("Work Instruction visual - tidak terpengaruh MSA/AAA")

with tabs[3]:
    st.subheader("Compare - ISOLATED")
    st.file_uploader("Upload Foto OK", type=["jpg","png"], key="ok_isolated")
    st.file_uploader("Upload Foto NG", type=["jpg","png"], key="ng_isolated")

with tabs[4]:
    st.subheader("Trial - ISOLATED")
    st.session_state.trial_df = st.data_editor(st.session_state.trial_df, num_rows="dynamic", use_container_width=True, key="editor_trial_isolated")

with tabs[5]:
    st.subheader("Catalog - ISOLATED - FULL 15 Kategori")
    st.session_state.catalog_df = st.data_editor(st.session_state.catalog_df, use_container_width=True, num_rows="dynamic", key="editor_catalog_isolated")
    fig = px.bar(st.session_state.catalog_df, x="Defect Type", y="Qty", title="Pareto Catalog")
    st.plotly_chart(fig, key="chart_catalog_isolated")

with tabs[6]:
    st.subheader("DMAIC - ISOLATED - FULL 15 Kategori")
    st.session_state.dmaic_df = st.data_editor(st.session_state.dmaic_df, num_rows="dynamic", use_container_width=True, key="editor_dmaic_isolated")
    df_p = st.session_state.dmaic_df.sort_values("Qty", ascending=False)
    fig2 = px.bar(df_p, x="DefectType", y="Qty", title="Pareto DMAIC")
    st.plotly_chart(fig2, key="chart_dmaic_isolated", use_container_width=True)

st.success("✅ FIX: MSA/AAA edit tidak akan mereset Catalog/DMAIC/Trial lagi karena pakai session_state + key terpisah")
