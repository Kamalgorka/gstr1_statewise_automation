import streamlit as st
from pathlib import Path

def load_global_css():
    css_path = Path(__file__).parent / "assets" / "global.css"
    if css_path.exists():
        st.markdown(
            f"<style>{css_path.read_text(encoding='utf-8')}</style>",
            unsafe_allow_html=True
        )

