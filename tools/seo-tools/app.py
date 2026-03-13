import streamlit as st

st.set_page_config(
    page_title="SEO Tools",
    page_icon="🛠️",
    layout="wide"
)

# Navigation multi-pages
reoptimisation = st.Page("pages/1_reoptimisation.py", title="Ré-optimisation SEO", icon="📊")
title_optimizer = st.Page("pages/2_title_optimizer.py", title="Title Optimizer", icon="🏷️")
serp_analyzer = st.Page("pages/3_serp_analyzer.py", title="SERP Analyzer", icon="🎯")
product_scraper = st.Page("pages/4_product_scraper.py", title="Scraper Produits", icon="🛒")
positions = st.Page("pages/5_positions_opportunities.py", title="Positions & Opportunités", icon="🔍")
analyse_complete = st.Page("pages/6_analyse_complete.py", title="Analyse Complète", icon="📊")

pg = st.navigation([analyse_complete, reoptimisation, title_optimizer, serp_analyzer, product_scraper, positions])
pg.run()
