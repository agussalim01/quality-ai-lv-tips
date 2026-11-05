import os
import streamlit as st
import google.generativeai as genai
from PIL import Image

def init_gemini():
    """Inisialisasi API Key Gemini dari Streamlit Secrets atau Environment Variable."""
    api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
    if not api_key:
        st.error("⚠️ GEMINI_API_KEY belum dikonfigurasi di Streamlit Secrets / Env!")
        st.stop()
    genai.configure(api_key=api_key)

def load_system_prompt(file_path="prompts/system_prompt_reviewer.txt"):
    """Membaca prompt standar dari folder lokal repositori GitHub."""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return """
        Kamu adalah AI Reviewer khusus untuk aplikasi Quality AI PRO V6 (Project: NT8 > LV Tips).
        Tugas utama: Evaluasi dan standarisasi laporan, WI, serta perbandingan dokumen sesuai aturan template_WI_landscape.docx.
        """

def analyze_wi_text(old_text: str, new_text: str, project: str, part: str):
    """Menganalisis perbedaan dokumen OLD vs NEW WI (Tab 7)."""
    init_gemini()
    system_prompt = load_system_prompt()
    
    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        system_instruction=system_prompt
    )
    
    prompt = f"""
    Lakukan analisis Quality Review untuk revisi dokumen WI berikut:
    - Project: {project}
    - Part Name: {part}
    
    --- ISI OLD WI ---
    {old_text[:3000]}
    
    --- ISI NEW WI ---
    {new_text[:3000]}
    
    Berikan evaluasi apakah perubahan ini sudah memenuhi standar penulisan (termasuk konsistensi istilah dan aturan penandaan kritis).
    """
    
    response = model.generate_content(prompt)
    return response.text

def analyze_trial_screenshot(before_img: Image.Image, after_img: Image.Image, project: str, part: str):
    """Menganalisis foto Before vs After pada Tab Screenshot / Trial Report."""
    init_gemini()
    system_prompt = load_system_prompt()
    
    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        system_instruction=system_prompt
    )
    
    prompt = [
        f"Analisis foto Trial / Perbaikan Kualitas untuk Project {project} - Part {part}.",
        "1. Bandingkan kondisi Gambar 1 (Before) dan Gambar 2 (After).",
        "2. Identifikasi area perbaikan atau defect (misal: flashing, alignment).",
        "3. Buatkan draft kalimat deskripsi ringkas untuk dimasukkan ke Trial Report.",
        "Gambar 1 (Before):", before_img,
        "Gambar 2 (After):", after_img
    ]
    
    response = model.generate_content(prompt)
    return response.text
