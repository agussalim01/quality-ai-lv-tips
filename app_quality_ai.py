import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Quality AI Platform", layout="wide", page_icon="⚙️")

st.title("Quality AI - LV Tips Analysis")

class GageRnRTemplate:
    def __init__(self, tolerance):
        self.tol = tolerance
        self.K1 = 0.5908; self.K2 = 0.5231; self.K3 = 0.3146

uploaded = st.file_uploader("Upload CSV (part, trial, value)", type=['csv'])

if uploaded:
    df = pd.read_csv(uploaded)
    st.write(df.head())
    fig = px.box(df, x='part', y='value', title="Variasi per Part")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Upload file CSV atau lihat demo di bawah")
    df_demo = pd.DataFrame({
        'part': np.repeat(range(1,11), 3),
        'value': np.random.normal(10, 0.05, 30)
    })
    fig = px.histogram(df_demo, x='value', title="Demo Distribusi")
    st.plotly_chart(fig, use_container_width=True)
