import streamlit as st

st.title("🧰 SEO Tools Library")
st.markdown("**Ta boîte à outils SEO — tous tes outils réunis au même endroit.**")

st.divider()

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(
        """
        <div style="background:#1e1e2e; border-radius:12px; padding:24px; height:280px;">
            <h2 style="color:#ff6b6b;">🔥 Hotspot Mapper</h2>
            <p style="color:#ccc; font-size:14px;">
                Fusionne tes exports <b>Screaming Frog</b> + <b>Google Search Console</b>
                pour identifier les pages à fort potentiel d'optimisation.<br><br>
                Génère un mapping <b>Title / H1 / Meta Description</b> prêt à remplir,
                exportable en Excel et Google Sheets.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("tools/hotspot_mapper.py", label="Ouvrir Hotspot Mapper", icon="🔥", use_container_width=True)

with col2:
    st.markdown(
        """
        <div style="background:#1e1e2e; border-radius:12px; padding:24px; height:280px;">
            <h2 style="color:#ffd43b;">🏷️ Title Optimizer</h2>
            <p style="color:#ccc; font-size:14px;">
                Scrape les <b>balises title de la SERP</b> via DataForSEO,
                analyse les patterns dominants (n-grams), puis génère
                des titles optimisées grâce à un <b>LLM</b> (Claude, GPT, Mistral, Groq).<br><br>
                Basé sur les données réelles de la SERP.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("tools/title_optimizer.py", label="Ouvrir Title Optimizer", icon="🏷️", use_container_width=True)

with col3:
    st.markdown(
        """
        <div style="background:#1e1e2e; border-radius:12px; padding:24px; height:280px;">
            <h2 style="color:#74c0fc;">📊 Ré-optimisation</h2>
            <p style="color:#ccc; font-size:14px;">
                Compare deux extractions <b>GSC</b> (ancien vs actuel) + <b>Ahrefs Top Pages</b>
                pour identifier les pages en perte de positions, clics et CTR.<br><br>
                Génère un <b>Excel formaté</b> avec mise en forme conditionnelle.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("tools/reoptimisation.py", label="Ouvrir Ré-optimisation", icon="📊", use_container_width=True)

with col4:
    st.markdown(
        """
        <div style="background:#1e1e2e; border-radius:12px; padding:24px; height:280px;">
            <h2 style="color:#51cf66;">🔍 Product Analyzer</h2>
            <p style="color:#ccc; font-size:14px;">
                Scrape les produits <b>Shopify</b> via leur API JSON,
                analyse les <b>matières, couleurs, coupes, types</b> et génère
                des combinaisons de <b>mots-clés SEO</b>.<br><br>
                Croise avec <b>Ahrefs</b> pour identifier les trous SEO par matière.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("tools/product_analyzer.py", label="Ouvrir Product Analyzer", icon="🔍", use_container_width=True)

st.divider()
st.caption("SEO Tools Library v1.0 — Utilise la navigation dans la sidebar pour accéder aux outils.")
