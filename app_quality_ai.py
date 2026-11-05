# app_quality_ai.py - V6 FINAL FIXED - SIDE BY SIDE GOOGLE DOCS COMPARE INTEGRATED
import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import io
from datetime import datetime
import difflib
import requests

# Safe Import AI Helper
AI_AVAILABLE = False
try:
    import ai_helper
    from ai_helper import analyze_wi_text, analyze_trial_screenshot
    AI_AVAILABLE = True
except Exception:
    pass

st.set_page_config(page_title="Quality AI V6 FINAL V3 HEADER REAL", page_icon="📊", layout="wide")

# =================== CACHING & UTILS ===================
@st.cache_data(show_spinner=False)
def extract_file_content(file_bytes, file_name):
    try:
        if file_name.lower().endswith('.docx'):
            doc = Document(io.BytesIO(file_bytes))
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ''])
        else:
            import fitz
            pdf = fitz.open(stream=file_bytes, filetype="pdf")
            txt = ""
            for pg in pdf:
                txt += pg.get_text() + "\n"
            return txt
    except Exception as e:
        return str(e)

def generate_side_by_side_html(old_text, new_text):
    """Menghasilkan teks HTML terpisah untuk dokumen OLD (Kiri) dan NEW (Kanan)."""
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    
    old_html = []
    new_html = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            word_str = " ".join(old_words[i1:i2])
            old_html.append(word_str)
            new_html.append(word_str)
        elif tag == 'delete':
            deleted = " ".join(old_words[i1:i2])
            old_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
        elif tag == 'insert':
            inserted = " ".join(new_words[j1:j2])
            new_html.append(f'<span style="background-color: #dafbe1; color: #1a7f37; font-weight: bold; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
        elif tag == 'replace':
            deleted = " ".join(old_words[i1:i2])
            inserted = " ".join(new_words[j1:j2])
            old_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
            new_html.append(f'<span style="background-color: #dafbe1; color: #1a7f37; font-weight: bold; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
            
    return " ".join(old_html), " ".join(new_html)

def safe_join(row):
    vals = []
    for v in row.values:
        try:
            if pd.isna(v): continue
            s = str(v).strip()
            if s.lower() not in ['nan', '']: vals.append(s)
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
                    if 'flashing' in rt.lower(): result["define"] += rt + "\n"
                    if 'actual condition' in rt.lower(): result["measure"] += rt + "\n"
                    if 'mechanism' in rt.lower(): result["analyze"] += rt + "\n"
                    if 'main solution' in rt.lower(): result["improve"] += rt + "\n
