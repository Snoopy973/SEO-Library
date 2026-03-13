import streamlit as st
import pandas as pd
from collections import Counter
from io import BytesIO
from datetime import datetime

st.title("🔍 Positions & Opportunités")
st.markdown("Croisez vos produits, mots-clés et pages pour trouver les **trous dans votre stratégie SEO**.")

# ─────────────────────────────────────────────
# PARSEURS CSV AHREFS
# ─────────────────────────────────────────────

def parse_ahrefs_csv(uploaded_file):
    """Parse un CSV Ahrefs (UTF-16 tab-separated ou UTF-8)"""
    content = uploaded_file.read()
    try:
        text = content.decode("utf-16")
        sep = "\t"
    except (UnicodeDecodeError, UnicodeError):
        try:
            text = content.decode("utf-8-sig")
            sep = "\t" if "\t" in text[:500] else ","
        except UnicodeDecodeError:
            text = content.decode("latin-1")
            sep = "\t" if "\t" in text[:500] else ","

    from io import StringIO
    df = pd.read_csv(StringIO(text), sep=sep)
    # Nettoyer les noms de colonnes
    df.columns = [c.strip().strip('"') for c in df.columns]
    return df


def detect_ahrefs_type(df):
    """Détecte le type d'export Ahrefs"""
    cols = set(c.lower() for c in df.columns)
    if "top keyword" in cols or "top keyword: volume" in cols:
        return "top_pages"
    if "keyword" in cols and "volume" in cols:
        return "keywords"
    if "referring page url" in cols or "anchor" in cols:
        return "backlinks"
    return "unknown"


# ─────────────────────────────────────────────
# MATCHING LOGIC
# ─────────────────────────────────────────────

def match_keywords_to_pages(df_keywords, df_pages, combos_with_materials):
    """
    Croise :
    - combos produits (type + matière/couleur) avec les matières associées
    - volumes de mots-clés Ahrefs
    - pages qui rankent déjà
    Retourne un DataFrame complet
    """
    rows = []

    # Index des pages par top keyword
    page_index = {}
    if df_pages is not None and not df_pages.empty:
        for _, row in df_pages.iterrows():
            kw = str(row.get("Top keyword", "")).strip().lower()
            if kw:
                page_index[kw] = {
                    "url": row.get("URL", ""),
                    "position": row.get("Top keyword: Position", ""),
                    "traffic": row.get("Traffic", 0),
                    "keywords_count": row.get("Keywords", 0),
                }

    # Index des volumes par keyword
    vol_index = {}
    if df_keywords is not None and not df_keywords.empty:
        for _, row in df_keywords.iterrows():
            kw = str(row.get("Keyword", "")).strip().lower()
            if kw:
                vol_index[kw] = {
                    "volume": row.get("Volume", 0),
                    "kd": row.get("Difficulty", ""),
                    "cpc": row.get("CPC", ""),
                    "traffic_potential": row.get("Traffic potential", ""),
                }

    # Pour chaque combinaison produit
    for combo_kw, materials in combos_with_materials.items():
        combo_lower = combo_kw.lower().strip()

        # Chercher le volume
        vol_data = vol_index.get(combo_lower, {})
        volume = vol_data.get("volume", "N/A")
        kd = vol_data.get("kd", "N/A")
        cpc = vol_data.get("cpc", "N/A")
        tp = vol_data.get("traffic_potential", "N/A")

        # Chercher si une page ranke
        page_data = page_index.get(combo_lower, {})
        page_url = page_data.get("url", "")
        position = page_data.get("position", "")
        traffic = page_data.get("traffic", 0)

        # Si pas de match exact, chercher un match partiel dans les URLs de pages
        if not page_url and df_pages is not None:
            for _, prow in df_pages.iterrows():
                url = str(prow.get("URL", "")).lower()
                slug = combo_lower.replace(" ", "-")
                if slug in url:
                    page_url = prow.get("URL", "")
                    position = prow.get("Top keyword: Position", "")
                    traffic = prow.get("Traffic", 0)
                    break

        # Déterminer le statut
        if page_url and position:
            try:
                pos_int = int(float(str(position)))
            except (ValueError, TypeError):
                pos_int = 99
            if pos_int <= 3:
                statut = "🟢 Top 3"
            elif pos_int <= 10:
                statut = "🟡 Top 10"
            elif pos_int <= 20:
                statut = "🟠 Page 2"
            else:
                statut = "⚪ > 20"
        else:
            statut = "🔴 Pas de page"

        rows.append({
            "Mot-clé": combo_kw,
            "Matière": ", ".join(materials) if materials else "",
            "Volume": volume,
            "KD": kd,
            "CPC (€)": cpc,
            "Potentiel trafic": tp,
            "Page qui ranke": page_url,
            "Position": position if position else "—",
            "Traffic actuel": traffic if traffic else 0,
            "Statut": statut,
        })

    df = pd.DataFrame(rows)

    # Trier par volume décroissant (N/A en dernier)
    df["_vol_sort"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(-1)
    df = df.sort_values("_vol_sort", ascending=False).drop(columns="_vol_sort")

    return df


# ─────────────────────────────────────────────
# UI — UPLOAD DES FICHIERS
# ─────────────────────────────────────────────

st.markdown("### 1. Charger les données")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**📊 CSV Mots-clés Ahrefs** (Keywords Explorer)")
    uploaded_kw = st.file_uploader("Export Keywords", type=["csv"], key="kw_upload",
                                    help="Export Ahrefs Keywords Explorer avec Volume, KD, CPC...")

with col2:
    st.markdown("**📄 CSV Top Pages Ahrefs** (Site Explorer)")
    uploaded_pages = st.file_uploader("Export Top Pages", type=["csv"], key="pages_upload",
                                      help="Export Ahrefs Top Pages avec URL, Traffic, Top keyword...")

# Charger les fichiers
df_keywords = None
df_pages = None

if uploaded_kw:
    df_keywords = parse_ahrefs_csv(uploaded_kw)
    st.success(f"✅ {len(df_keywords)} mots-clés chargés")

if uploaded_pages:
    df_pages = parse_ahrefs_csv(uploaded_pages)
    st.success(f"✅ {len(df_pages)} pages chargées")


# ─────────────────────────────────────────────
# UI — ANALYSE
# ─────────────────────────────────────────────

# Vérifier qu'on a les données produits
has_products = "product_results" in st.session_state

if not has_products:
    st.info("💡 Aucune donnée produit en mémoire. Va d'abord sur la page **🛒 Scraper Produits** pour importer un catalogue, ou uploade directement les CSV Ahrefs ci-dessus.")

# On peut quand même fonctionner avec juste les CSV Ahrefs
if df_keywords is not None or df_pages is not None:

    st.divider()
    st.markdown("### 2. Analyse Positions & Opportunités")

    # Construire les combinaisons avec matières associées
    combos_with_materials = {}

    if has_products:
        results = st.session_state["product_results"]

        # Depuis les combos type+matière, on connait la matière de chaque combo
        for combo_key, count in results.get("combos_type_mat", {}).items():
            parts = combo_key.split()
            if len(parts) >= 2:
                # La matière est le 2e mot (ou plus)
                mat = " ".join(parts[1:]).capitalize()
                combos_with_materials[combo_key] = [mat]

        # Ajouter les combos type+couleur (sans matière)
        for combo_key in results.get("combos_type_col", {}):
            if combo_key not in combos_with_materials:
                combos_with_materials[combo_key] = []

        # Ajouter les combos type+coupe (sans matière)
        for combo_key in results.get("combos_type_coupe", {}):
            if combo_key not in combos_with_materials:
                combos_with_materials[combo_key] = []

    # Si pas de produits, utiliser les keywords Ahrefs directement
    if not combos_with_materials and df_keywords is not None:
        for _, row in df_keywords.iterrows():
            kw = str(row.get("Keyword", ""))
            combos_with_materials[kw] = []

    if combos_with_materials:
        # Lancer le matching
        df_matched = match_keywords_to_pages(df_keywords, df_pages, combos_with_materials)

        # ── FILTRES ──
        st.markdown("### 🎛️ Filtres")
        filter_col1, filter_col2, filter_col3 = st.columns(3)

        with filter_col1:
            # Filtre par matière
            all_materials = sorted(set(
                m for mats in combos_with_materials.values()
                for m in mats if m
            ))
            selected_materials = st.multiselect(
                "🧵 Filtrer par matière",
                options=all_materials,
                default=[],
                help="Sélectionne une ou plusieurs matières pour voir les opportunités"
            )

        with filter_col2:
            # Filtre par statut
            all_statuts = sorted(df_matched["Statut"].unique())
            selected_statuts = st.multiselect(
                "📌 Filtrer par statut",
                options=all_statuts,
                default=[],
                help="🔴 = pas de page, 🟡 = à optimiser, 🟢 = bien placé"
            )

        with filter_col3:
            # Filtre volume minimum
            vol_min = st.number_input("🔢 Volume minimum", min_value=0, value=0, step=100,
                                       help="N'afficher que les mots-clés avec un volume >= à cette valeur")

        # Appliquer les filtres
        df_display = df_matched.copy()

        if selected_materials:
            df_display = df_display[df_display["Matière"].apply(
                lambda x: any(m.lower() in str(x).lower() for m in selected_materials)
            )]

        if selected_statuts:
            df_display = df_display[df_display["Statut"].isin(selected_statuts)]

        if vol_min > 0:
            df_display = df_display[pd.to_numeric(df_display["Volume"], errors="coerce").fillna(0) >= vol_min]

        # ── KPIs ──
        st.divider()
        total_kw = len(df_display)
        with_page = len(df_display[df_display["Statut"] != "🔴 Pas de page"])
        without_page = len(df_display[df_display["Statut"] == "🔴 Pas de page"])
        top3 = len(df_display[df_display["Statut"] == "🟢 Top 3"])
        to_optimize = len(df_display[df_display["Statut"].isin(["🟡 Top 10", "🟠 Page 2"])])

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total mots-clés", total_kw)
        k2.metric("🟢 Top 3", top3)
        k3.metric("🟡 À optimiser", to_optimize)
        k4.metric("🔴 Sans page", without_page)
        k5.metric("Couverture", f"{round(with_page/max(total_kw,1)*100)}%")

        # ── GRAPHIQUES ──
        st.markdown("### 📊 Vue d'ensemble")
        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            import plotly.express as px
            statut_counts = df_display["Statut"].value_counts()
            fig = px.pie(values=statut_counts.values, names=statut_counts.index,
                        title="Répartition des statuts",
                        color_discrete_map={
                            "🟢 Top 3": "#2ecc71", "🟡 Top 10": "#f1c40f",
                            "🟠 Page 2": "#e67e22", "⚪ > 20": "#bdc3c7",
                            "🔴 Pas de page": "#e74c3c"
                        })
            st.plotly_chart(fig, use_container_width=True)

        with chart_col2:
            if selected_materials or all_materials:
                # Opportunités par matière
                mat_opps = {}
                for _, row in df_matched.iterrows():
                    if row["Statut"] == "🔴 Pas de page" and row["Matière"]:
                        for m in str(row["Matière"]).split(","):
                            m = m.strip()
                            if m:
                                mat_opps[m] = mat_opps.get(m, 0) + 1
                if mat_opps:
                    df_opp = pd.DataFrame(sorted(mat_opps.items(), key=lambda x: -x[1]),
                                           columns=["Matière", "KW sans page"])
                    fig2 = px.bar(df_opp.head(15), x="Matière", y="KW sans page",
                                  title="🔴 Opportunités par matière (KW sans page dédiée)")
                    st.plotly_chart(fig2, use_container_width=True)

        # ── TABLEAU PRINCIPAL ──
        st.markdown("### 📋 Détail des mots-clés")
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Page qui ranke": st.column_config.LinkColumn("Page qui ranke"),
                "Volume": st.column_config.NumberColumn("Volume", format="%d"),
                "Traffic actuel": st.column_config.NumberColumn("Traffic", format="%d"),
            }
        )

        # ── TOP OPPORTUNITÉS ──
        st.markdown("### 🎯 Top 20 Opportunités (volume élevé, pas de page)")
        df_opps = df_display[df_display["Statut"] == "🔴 Pas de page"].copy()
        df_opps["_vol"] = pd.to_numeric(df_opps["Volume"], errors="coerce").fillna(0)
        df_opps = df_opps.sort_values("_vol", ascending=False).drop(columns="_vol").head(20)
        if not df_opps.empty:
            st.dataframe(df_opps, use_container_width=True, hide_index=True)
        else:
            st.success("🎉 Toutes les combinaisons sont couvertes par une page !")

        # ── VUE PAR MATIÈRE ──
        if all_materials:
            st.markdown("### 🧵 Vue détaillée par matière")
            mat_choice = st.selectbox("Choisir une matière", options=["Toutes"] + all_materials)

            if mat_choice != "Toutes":
                df_mat_view = df_matched[df_matched["Matière"].str.contains(mat_choice, case=False, na=False)]
                covered = len(df_mat_view[df_mat_view["Statut"] != "🔴 Pas de page"])
                not_covered = len(df_mat_view[df_mat_view["Statut"] == "🔴 Pas de page"])
                total_mat = len(df_mat_view)

                m1, m2, m3 = st.columns(3)
                m1.metric(f"Mots-clés '{mat_choice}'", total_mat)
                m2.metric("✅ Couverts (page existante)", covered)
                m3.metric("🔴 Non couverts", not_covered)

                if not_covered > 0:
                    st.warning(f"⚠️ {not_covered} mots-clés contenant '{mat_choice}' n'ont PAS de page dédiée !")

                st.dataframe(df_mat_view, use_container_width=True, hide_index=True)

        # ── EXPORT ──
        st.divider()
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_display.to_excel(writer, sheet_name="Positions & Opportunités", index=False)
            if not df_opps.empty:
                df_opps.to_excel(writer, sheet_name="Top Opportunités", index=False)

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        store = st.session_state.get("product_store", "site")
        st.download_button(
            "⬇️ Télécharger l'analyse Excel",
            data=buffer.getvalue(),
            file_name=f"positions_opportunites_{store}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.markdown("---")
    st.info("☝️ Uploade au moins un fichier CSV Ahrefs ci-dessus pour commencer l'analyse.")
