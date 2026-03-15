import streamlit as st

st.set_page_config(
    page_title="SEO Tools",
    page_icon="🛠️",
    layout="wide"
)

# Navigation multi-pages
analyse_complete = st.Page("pages/6_analyse_complete.py", title="Analyse Complète", icon="📊")
mapping_points_chauds = st.Page("pages/4_mapping_points_chauds.py", title="Mapping Points Chauds", icon="🔥")
reoptimisation = st.Page("pages/1_reoptimisation.py", title="Ré-optimisation SEO", icon="🔄")
title_optimizer = st.Page("pages/2_title_optimizer.py", title="Title Optimizer", icon="🏷️")
serp_analyzer = st.Page("pages/3_serp_analyzer.py", title="SERP Analyzer", icon="🎯")

pg = st.navigation([analyse_complete, mapping_points_chauds, reoptimisation, title_optimizer, serp_analyzer])
pg.run()
