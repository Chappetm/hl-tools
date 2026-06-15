import streamlit as st

from views.hl_stocktake_converter import render as render_converter
from views.hl_log_errors import render as render_log_errors

st.set_page_config(page_title="H&L Tools", page_icon="🏪", layout="centered")

PAGES = {
    "stocktake_converter": ("📤 Stocktake XLSX → DAT", render_converter),
    "log_errors": ("🔍 Log Error Report", render_log_errors),
}

if "page" not in st.session_state:
    st.session_state.page = "stocktake_converter"

with st.sidebar:
    st.title("🏪 H&L Tools")
    st.divider()
    for key, (label, _) in PAGES.items():
        if st.button(label, use_container_width=True, key=f"nav_{key}"):
            st.session_state.page = key

_, render_fn = PAGES[st.session_state.page]
render_fn()
