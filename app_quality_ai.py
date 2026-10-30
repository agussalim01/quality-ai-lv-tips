import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats

st.set_page_config(page_title="Quality AI Platform", layout="wide", page_icon="⚙️")

class GageRnRTemplate:
    def __init__(self, tolerance):
        self.tol = tolerance
        self.K1 = 0.5908; self.K2 = 0.5231; self.K3 = 0.3146
    def analyze(self, df_long):
        parts = df_long['part'].nunique()
        trials = df_long['trial'].nunique()
        K1 = 0.5908 if trials==3 else 0.8862
        rbars, xbars_op = [], []
        for op in df_long['operator'].unique():
            sub = df_long[df_long['operator']==op]
            ranges = sub.groupby('part')['value'].apply(lambda x: x.max()-x.min())
            rbars.append(ranges.mean())
            xbars_op.append(sub['value'].mean())
        Rbarbar = np.mean(rbars)
        Xbar_diff = max(xbars_op) - min(xbars_op)
        Rp = df_long.groupby('part')['value'].mean().max() - df_long.groupby('part')['value'].mean().min()
        EV = Rbarbar * K1
        AV = np.sqrt(max(0, (Xbar_diff * self.K2)**2 - (EV**2 / (parts*trials))))
        PV = Rp * self.K3
        TV = np.sqrt(EV**2 + AV**2 + PV**2)
        GRR = np.sqrt(EV**2 + AV**2)
        pct_tol = (GRR*5.15/self.tol)*100
        pct_tol6 = (GRR*6/self.tol)*100
        NDC = 1.41 * (PV/GRR) if GRR!=0 else 0
        status = "ACCEPTABLE (<10%)" if pct_tol<10 else ("MARGINAL (10-30%)" if pct_tol<30 else "NOT ACCEPTABLE")
        return {"Rbarbar": Rbarbar, "EV": EV, "AV": AV, "PV": PV, "GRR": GRR, "TV": TV, "%Tol 5.15": pct_tol, "%Tol 6.0": pct_tol6, "NDC": NDC, "Status": status, "rbars": rbars, "xbars_op": xbars_op}

def capability_analysis(data, lsl, usl):
    mean = np.mean(data); std = np.std(data, ddof=1)
    Cp = (usl-lsl)/(6*std) if std!=0 else 0
    Cpu = (usl-mean)/(3*std) if std!=0 else 0
    Cpl = (mean-lsl)/(3*std) if std!=0 else 0
    Cpk = min(Cpu, Cpl)
    return {"Mean": mean, "Std": std, "Cp": Cp, "Cpk": Cpk}

st.sidebar.title("Quality AI Platform - LV Tips NT8")
menu = st.sidebar.selectbox("Pilih Modul:", ["1. GR&R Analyzer", "2. Capability Generator", "3. Trial Report Generator", "4. Data Comparison", "5. AI Defect Assistant", "6. DMAIC Generator"])

if menu.startswith("1"):
    st.title("1. GR&R Analyzer - Universal")
    lsl = st.number_input("LSL", value=21.78); usl = st.number_input("USL", value=21.83); tol = usl-lsl
    uploaded = st.file_uploader("Upload Excel GR&R", type=['xlsx','xls'])
    if st.button("Load Contoh Data LV Tips"):
        example = pd.DataFrame({'part': list(range(1,11))*9,'operator': ['Agus']*30 + ['Heri']*30 + ['Aznen']*30,'trial': ([1]*10 + [2]*10 + [3]*10)*3,'value': [21.799,21.8,21.806,21.793,21.813,21.797,21.795,21.804,21.809,21.773,21.8,21.8,21.808,21.792,21.814,21.797,21.794,21.803,21.81,21.774,21.801,21.801,21.807,21.792,21.813,21.798,21.794,21.803,21.81,21.772,21.8,21.8,21.808,21.793,21.814,21.797,21.796,21.804,21.808,21.774,21.799,21.8,21.806,21.793,21.813,21.798,21.798,21.804,21.809,21.774,21.801,21.801,21.807,21.792,21.813,21.798,21.797,21.804,21.808,21.772,21.8,21.799,21.806,21.793,21.812,21.797,21.797,21.803,21.808,21.774,21.8,21.8,21.806,21.793,21.813,21.797,21.797,21.804,21.808,21.774,21.801,21.799,21.806,21.792,21.813,21.797,21.797,21.803,21.808,21.774]})
        st.session_state['df_grr'] = example
    df = None
    if uploaded:
        try: df = pd.read_excel(uploaded)
        except: df = None
    if 'df_grr' in st.session_state: df = st.session_state['df_grr']
    if df is not None and 'value' in df.columns:
        st.dataframe(df.head(20))
        analyzer = GageRnRTemplate(tolerance=tol); res = analyzer.analyze(df)
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("%GRR Tol", f"{res['%Tol 5.15']:.2f}%"); c2.metric("NDC", f"{res['NDC']:.1f}"); c3.metric("Status", res['Status'])
        fig, ax = plt.subplots(1,2, figsize=(10,4))
        ax[0].bar(range(len(res['rbars'])), res['rbars']); ax[0].set_title("Rbar per Operator")
        ax[1].bar(range(len(res['xbars_op'])), res['xbars_op']); ax[1].set_title("Xbar per Operator")
        st.pyplot(fig)

elif menu.startswith("2"):
    st.title("2. Capability Generator")
    lsl = st.number_input("LSL Cap", value=21.78); usl = st.number_input("USL Cap", value=21.83)
    manual = st.text_area("Paste data:", "21.799,21.8,21.806,21.793,21.813,21.797,21.795,21.804,21.809,21.773")
    if manual:
        data = [float(x.strip()) for x in manual.replace("\n",",").split(',') if x.strip()!='']
        if len(data)>1:
            cap = capability_analysis(np.array(data), lsl, usl)
            st.metric("Cpk", f"{cap['Cpk']:.2f}"); st.metric("Cp", f"{cap['Cp']:.2f}")
            fig, ax = plt.subplots(); ax.hist(data, bins=15); st.pyplot(fig)

elif menu.startswith("3"):
    st.title("3. Trial Report Generator")
    part_no = st.text_input("Part No", "LV TIPS NT8"); machine = st.selectbox("Machine", ["JSW 180Ton", "Sumitomo 180Ton"]); result = st.selectbox("Result", ["OK", "NG"])
    if st.button("Generate Report"):
        st.success(f"Trial Report {part_no} - {machine} - {result} Generated!"); st.download_button("Download", f"# Trial {part_no}\nMachine {machine}\nResult {result}", file_name="Trial.md")

elif menu.startswith("4"):
    st.title("4. Data Comparison JSW vs Sumitomo")
    st.info("Upload 2 file JSW & Sumitomo untuk t-test & boxplot")
    f1 = st.file_uploader("File JSW", key='jsw'); f2 = st.file_uploader("File Sumitomo", key='sumi')
    if f1 and f2:
        df1 = pd.read_excel(f1); df2 = pd.read_excel(f2)
        col1 = df1.select_dtypes(include=[np.number]).iloc[:,0]; col2 = df2.select_dtypes(include=[np.number]).iloc[:,0]
        t,p = stats.ttest_ind(col1, col2, equal_var=False); st.write(f"p-value {p:.5f}")

elif menu.startswith("5"):
    st.title("5. AI Defect Assistant")
    query = st.text_input("Tanya defect:", "Short shot LV Tips")
    if query: st.write(f"AI: Untuk {query}, cek melt temp, injection speed, venting mold. Rekomendasi: +5C melt, +10% speed")

elif menu.startswith("6"):
    st.title("6. DMAIC Generator")
    problem = st.text_input("Problem", "LV Tips Length NG 3.5%")
    if st.button("Generate DMAIC"): st.markdown(f"# DMAIC - {problem}\nDefine, Measure, Analyze, Improve, Control - Auto Generated")
