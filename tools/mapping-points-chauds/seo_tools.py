import streamlit as st

st.set_page_config(
    page_title="🧰 SEO Tools Library",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Pages ──
home = st.Page("tools/home.py", title="Accueil", icon="🏠", default=True)
hotspot = st.Page("tools/hotspot_mapper.py", title="Hotspot Mapper", icon="🔥")
title_opt = st.Page("tools/title_optimizer.py", title="Title Optimizer", icon="🏷️")
reopt = st.Page("tools/reoptimisation.py", title="Ré-optimisation", icon="📊")

pg = st.navigation(
    {
        "": [home],
        "SEO Tools": [hotspot, title_opt, reopt],
    },
    position="sidebar",
)

pg.run()
