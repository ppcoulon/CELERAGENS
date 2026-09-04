import pathlib

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Vigie NPE — Démonstrateur CelerAgens",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Streamlit adds its own padding/margins around the page — strip them so the
# embedded app can use the full viewport, and hide the default chrome.
st.markdown(
    """
    <style>
        .block-container { padding: 0 !important; max-width: 100% !important; }
        header[data-testid="stHeader"] { display: none; }
        #MainMenu, footer { visibility: hidden; }
        iframe { display: block; }
    </style>
    """,
    unsafe_allow_html=True,
)

HTML_PATH = pathlib.Path(__file__).parent / "vigie-npe.html"
html = HTML_PATH.read_text(encoding="utf-8")

components.html(html, height=1400, scrolling=True)
