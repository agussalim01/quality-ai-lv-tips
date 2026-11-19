"""Gemini helper functions for quality engineering reviews."""
import os
import streamlit as st
from google import genai
from google.genai import types

def _client():
    try:
        key = st.secrets.get("GEMINI_API_KEY")
    except Exception:
        key = None
    key = key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise RuntimeError("GEMINI_API_KEY belum diatur di Streamlit Secrets.")
    return genai.Client(api_key=key)

def load_system_prompt(file_path="prompts/system_prompt_reviewer.txt"):
    try:
        with open(file_path, encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return ("Anda adalah reviewer quality engineering. Jangan mengarang data, "
                "parameter, hasil inspeksi, atau persetujuan. Tandai hal yang perlu "
                "diverifikasi engineer.")

def _generate(contents):
    client = _client()
    try:
        response = client.models.generate_content(
            model=os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
            contents=contents,
            config=types.GenerateContentConfig(system_instruction=load_system_prompt()),
        )
        if not response.text:
            raise RuntimeError("Gemini tidak menghasilkan teks.")
        return response.text
    finally:
        client.close()

def analyze_wi_text(old_text: str, new_text: str, project: str, part: str):
    if not old_text.strip() or not new_text.strip():
        raise ValueError("Kedua WI harus memiliki teks.")
    prompt = (
        f"Review perubahan Work Instruction project {project}, part {part}. "
        "Bandingkan langkah, parameter, peralatan, keselamatan, dan kontrol kualitas. "
        "Pisahkan temuan dari asumsi dan sebutkan kebutuhan verifikasi.\n\n"
        f"OLD WI:\n{old_text}\n\nNEW WI:\n{new_text}"
    )
    return _generate(prompt)

def analyze_trial_screenshot(before_img, after_img, project: str, part: str):
    return _generate([
        f"Bandingkan foto before/after trial {project} - {part}. "
        "Deskripsikan hanya perbedaan yang terlihat. Jangan menyimpulkan OK/NG "
        "tanpa kriteria inspeksi.",
        before_img, after_img,
    ])
