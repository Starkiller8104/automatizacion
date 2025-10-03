
import streamlit as st

def inject_base_css():
    st.markdown(
        """
        <style>
        .im-card { 
            padding: 1rem 1.25rem; 
            border-radius: 0.9rem; 
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.08);
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            margin-bottom: 0.75rem;
        }
        .im-title { font-weight: 700; font-size: 1.1rem; margin-bottom: 0.25rem; }
        .im-subtle { opacity: 0.8; font-size: 0.9rem; }
        .stButton > button { border-radius: 0.75rem; padding: 0.6rem 1rem; font-weight: 600; }
        div[data-testid="metric-container"] {
            padding: 0.6rem 0.8rem; border-radius: 0.75rem;
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.08);
        }
        div[data-testid="stMetricDelta"] svg { display: none; }
        </style>
        """,
        unsafe_allow_html=True
    )

def header(logo_path:str, title:str, subtitle:str="", updated:str|None=None):
    cols = st.columns([1, 5, 3])
    with cols[0]:
        try:
            st.image(logo_path, use_container_width=True)
        except Exception:
            st.write("")
    with cols[1]:
        st.markdown(f"### {title}")
        if subtitle:
            st.markdown(f"<span class='im-subtle'>{subtitle}</span>", unsafe_allow_html=True)
    with cols[2]:
        if updated:
            st.markdown(
                f"<div class='im-card'><div class='im-title'>Estado</div>"
                f"<div class='im-subtle'>Actualizado: {updated}</div></div>",
                unsafe_allow_html=True
            )

def section_card(title:str, body_builder):
    st.markdown(f"<div class='im-card'><div class='im-title'>{title}</div>", unsafe_allow_html=True)
    try:
        # IMPORTANT: return the inner component's return value
        return body_builder()
    finally:
        st.markdown("</div>", unsafe_allow_html=True)

def metric_row(items: list[tuple[str, str, str|None]]):
    cols = st.columns(len(items))
    for i, (label, value, delta) in enumerate(items):
        with cols[i]:
            st.metric(label, value, delta if delta else None)
