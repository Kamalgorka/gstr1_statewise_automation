import streamlit as st
from ui import load_global_css
load_global_css()

st.set_page_config(page_title="MML Smart Reports", layout="wide")
st.title("📊 MML Smart Reports")
st.write("Select a team from the left sidebar.")

