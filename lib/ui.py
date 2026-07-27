"""Royal Chain branding + CSS matching the React/Tailwind look:
the #1D3587 brand colour, white rounded cards and the top bar."""

from __future__ import annotations

import streamlit as st

BRAND = "#1D3587"
LOGO_URL = (
    "https://sp-ao.shortpixel.ai/client/to_webp,q_glossy,ret_img,w_2048/"
    "https://royalchaingroup.com/wp-content/uploads/2026/02/"
    "Royal-Chain-Limited-3-2048x1413.webp"
)


def inject_css():
    st.markdown(
        f"""
        <style>
        :root {{ --brand: {BRAND}; }}
        .stApp {{ background: #f3f4f6; }}
        .block-container {{ padding-top: 1.5rem; max-width: 860px; }}

        .rc-h1 {{ font-size:1.875rem; font-weight:700; margin:0; color:#111827; }}
        .rc-sub {{ color:#6b7280; font-size:.875rem; margin:.25rem 0 0; }}

        /* Top bar */
        .rc-topbar {{
            display:flex; align-items:center; justify-content:space-between;
            background:#fff; padding:.75rem 1.25rem; border-radius:.75rem;
            box-shadow:0 1px 2px rgba(0,0,0,.05); margin-bottom:1.25rem; margin-top:1.25rem;
        }}
        .rc-topbar img {{ height:40px; }}
        .rc-avatar {{
            width:2rem; height:2rem; border-radius:9999px; background:{BRAND};
            color:#fff; display:flex; align-items:center; justify-content:center;
            font-size:.8rem; font-weight:700;
        }}

        div[data-testid="stMetricValue"] {{ color:{BRAND}; }}
        div[data-testid="stFileUploaderDropzone"] {{ border-radius:.75rem; }}

        </style>
        """,
        unsafe_allow_html=True,
    )


def top_bar():
    st.markdown(
        f'<div class="rc-topbar">'
        f'<img src="{LOGO_URL}" alt="Royal Chain">'
        f'<p style="font-size:.85rem;color:#6b7280;margin:0;">Welcome: '
        f'<span style="font-weight:600;color:#1f2937;">Royal Chain Limited</span></p>'
        f'<span class="rc-avatar">R</span></div>',
        unsafe_allow_html=True,
    )
