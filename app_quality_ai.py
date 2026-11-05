# app_quality_ai.py - V6 FINAL FIXED V3 - FULL 7 TABS WITH IMAGE CATALOG
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
    """Menghasilkan HTML bersisian (Kiri = OLD, Kanan = NEW) ala Google Docs"""
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
            new_html.append(f'<span style="color: red; font-weight: bold; background-color: #dafbe1; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
        elif tag == 'replace':
            deleted = " ".join(old_words[i1:i2])
            inserted = " ".join(new_words[j1:j2])
            old_html.append(f'<span style="background-color: #ffebe9; color: #cf222e; text-decoration: line-through; padding: 2px 4px; border-radius: 3px;">{deleted}</span>')
            new_html.append(f'<span style="color: red; font-weight: bold; background-color: #dafbe1; padding: 2px 4px; border-radius: 3px;">{inserted}</span>')
            
    return " ".join(old_html), " ".join(new_html)

def safe_join(row):
