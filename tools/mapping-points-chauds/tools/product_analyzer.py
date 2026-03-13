import io
import json
import re
import ssl
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# st.set_page_config est appelé par home.py (entrypoint)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
DATA_DIR = Path.home() / "Desktop" / "shopify-scraper" / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def detect_site_from_filename(filename):
    m = re.search(r'(www\.\w[\w.-]+\.\w+)', filename)
    if m:
        return m.group(1)
    m = re.search(r'(\w[\w.-]+\.\w+)', filename)
    if m:
        return m.group(1)
    return "inconnu"


def read_ahrefs_csv(uploaded_file):
    raw = uploaded_file.read()
    uploaded_file.seek(0)
    for enc in ["utf-16", "utf-8-sig", "utf-8", "latin-1"]:
        for sep in ["\t", ",", ";"]:
            try:
                df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep=sep)
                if len(df.columns) > 2:
                    return df
            except Exception:
                continue
    return None


def fetch_shopify_products(domain):
    domain = domain.replace("https://", "").replace("http://", "").strip("/")
    if not domain.startswith("www."):
        domain = "www." + domain
    base_url = f"https://{domain}/products.json?limit=250"
    all_products = []
    page = 1
    while True:
        url = f"{base_url}&page={page}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        try:
            with urllib.request.urlopen(req, context=SSL_CTX, timeout=30) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            if page == 1:
                st.error(f"Erreur: {e}")
                return []
            break
        products = data.get("products", [])
        if not products:
            break
        all_products.extend(products)
        page += 1
    return all_products


def parse_products(products):
    rows = []
    mat_counter = Counter()
    type_counter = Counter()
    color_counter = Counter()
    coupe_counter = Counter()
    mat_type = Counter()
    type_color = Counter()
    type_coupe = Counter()
    mat_type_color = Counter()
    taxonomy = defaultdict(set)

    for p in products:
        raw_tags = p.get("tags", "")
        if isinstance(raw_tags, list):
            tags = [t.strip() for t in raw_tags if isinstance(t, str) and t.strip()]
        else:
            tags = [t.strip() for t in str(raw_tags).split(",") if t.strip()]
        product_type = p.get("product_type", "").strip()
        matieres, couleurs, coupes, formes, motifs, collections, saisons, genres = [], [], [], [], [], [], [], []

        for tag in tags:
            tl = tag.lower()
            if tl.startswith("matiere:"):
                val = tag.split(":", 1)[1].strip()
                matieres.append(val)
                taxonomy["Matière"].add(val)
            elif tl.startswith("couleur:"):
                val = tag.split(":", 1)[1].strip()
                couleurs.append(val)
                taxonomy["Couleur"].add(val)
            elif tl.startswith("coupe:"):
                val = tag.split(":", 1)[1].strip()
                coupes.append(val)
                taxonomy["Coupe"].add(val)
            elif tl.startswith("forme:"):
                val = tag.split(":", 1)[1].strip()
                formes.append(val)
                taxonomy["Forme"].add(val)
            elif tl.startswith("motif:"):
                val = tag.split(":", 1)[1].strip()
                motifs.append(val)
                taxonomy["Motif"].add(val)
            elif tl.startswith("collection:"):
                val = tag.split(":", 1)[1].strip()
                collections.append(val)
                taxonomy["Collection"].add(val)
            elif tl.startswith("saison:"):
                val = tag.split(":", 1)[1].strip()
                saisons.append(val)
                taxonomy["Saison"].add(val)
            elif tl.startswith("genre:"):
                val = tag.split(":", 1)[1].strip()
                genres.append(val)
                taxonomy["Genre"].add(val)

        if product_type:
            taxonomy["Type de produit"].add(product_type)

        prices = []
        for v in p.get("variants", []):
            try:
                prices.append(float(v.get("price", 0)))
            except (ValueError, TypeError):
                pass
        price = min(prices) if prices else 0

        mat_str = ", ".join(matieres) if matieres else "Non spécifié"
        col_str = ", ".join(couleurs) if couleurs else "Non spécifié"
        coupe_str = ", ".join(coupes) if coupes else "Non spécifié"

        type_counter[product_type or "Non spécifié"] += 1
        for m in (matieres or ["Non spécifié"]):
            mat_counter[m] += 1
            mat_type[(m, product_type or "Non spécifié")] += 1
            for c in (couleurs or ["Non spécifié"]):
                mat_type_color[(product_type or "Non spécifié", m, c)] += 1
        for c in (couleurs or ["Non spécifié"]):
            color_counter[c] += 1
            type_color[(product_type or "Non spécifié", c)] += 1
        for cp in (coupes or ["Non spécifié"]):
            coupe_counter[cp] += 1
            type_coupe[(product_type or "Non spécifié", cp)] += 1

        rows.append({
            "Titre": p.get("title", ""),
            "Type": product_type,
            "Matières": mat_str,
            "Couleurs": col_str,
            "Coupes": coupe_str,
            "Formes": ", ".join(formes),
            "Motifs": ", ".join(motifs),
            "Prix (€)": price,
            "Collections": ", ".join(collections),
            "URL": p.get("handle", ""),
        })

    return {
        "products_df": pd.DataFrame(rows),
        "mat_counter": mat_counter,
        "type_counter": type_counter,
        "color_counter": color_counter,
        "coupe_counter": coupe_counter,
        "mat_type": mat_type,
        "type_color": type_color,
        "type_coupe": type_coupe,
        "mat_type_color": mat_type_color,
        "taxonomy": taxonomy,
        "total": len(products),
    }


def generate_seo_keywords(taxonomy):
    types = sorted(taxonomy.get("Type de produit", set()))
    combos = []
    for attr_name, attr_key in [
        ("Matière", "Matière"), ("Couleur", "Couleur"),
        ("Coupe", "Coupe"), ("Forme", "Forme"),
        ("Motif", "Motif"), ("Saison", "Saison"),
    ]:
        vals = sorted(taxonomy.get(attr_key, set()))
        for t in types:
            for v in vals:
                combos.append({
                    "Mot-clé": f"{t.lower()} {v.lower()}",
                    "Type": t,
                    "Attribut": attr_name,
                    "Valeur": v,
                })
    return pd.DataFrame(combos)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
st.title("🔍 SEO Product Analyzer")

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Product Analyzer")

mode = st.sidebar.radio(
    "Source des données produits",
    ["📁 Fichier JSON Shopify", "🌐 Scraper en direct"],
    index=0,
    key="pa_mode",
)

st.sidebar.markdown("---")
st.sidebar.subheader("📊 Données Ahrefs (optionnel)")
keywords_file = st.sidebar.file_uploader("CSV Mots-clés (Keywords Explorer)", type=["csv"], key="pa_kw")
top_pages_file = st.sidebar.file_uploader("CSV Top Pages (Site Explorer)", type=["csv"], key="pa_tp")

# ─────────────────────────────────────────────
# MAIN — Load data
# ─────────────────────────────────────────────
parsed = None
site_name = ""

if mode == "🌐 Scraper en direct":
    domain = st.text_input("Domaine Shopify", value="www.balibaris.com", key="pa_domain")
    if st.button("🚀 Lancer le scraping", key="pa_scrape"):
        with st.spinner(f"Scraping de {domain}..."):
            products = fetch_shopify_products(domain)
        if products:
            site_name = domain.replace("www.", "").split(".")[0]
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            json_path = DATA_DIR / f"{site_name}_{ts}.json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(products, f, ensure_ascii=False)
            st.success(f"✅ {len(products)} produits récupérés — sauvegardé: {json_path.name}")
            parsed = parse_products(products)
            st.session_state["pa_parsed"] = parsed
            st.session_state["pa_site"] = site_name
        else:
            st.error("Aucun produit trouvé.")
    elif "pa_parsed" in st.session_state:
        parsed = st.session_state["pa_parsed"]
        site_name = st.session_state.get("pa_site", "")
else:
    json_file = st.file_uploader("Fichier JSON Shopify (/products.json ou export)", type=["json"], key="pa_json")
    if json_file:
        site_name = detect_site_from_filename(json_file.name).replace("www.", "").split(".")[0]
        products = json.load(json_file)
        if isinstance(products, dict) and "products" in products:
            products = products["products"]
        parsed = parse_products(products)
        st.session_state["pa_parsed"] = parsed
        st.session_state["pa_site"] = site_name
        st.success(f"✅ {parsed['total']} produits chargés depuis {json_file.name}")
    elif "pa_parsed" in st.session_state:
        parsed = st.session_state["pa_parsed"]
        site_name = st.session_state.get("pa_site", "")

# Load Ahrefs data
df_keywords = None
df_top_pages = None
if keywords_file:
    df_keywords = read_ahrefs_csv(keywords_file)
    if df_keywords is not None:
        st.session_state["pa_kw_df"] = df_keywords
elif "pa_kw_df" in st.session_state:
    df_keywords = st.session_state["pa_kw_df"]

if top_pages_file:
    df_top_pages = read_ahrefs_csv(top_pages_file)
    if df_top_pages is not None:
        st.session_state["pa_tp_df"] = df_top_pages
elif "pa_tp_df" in st.session_state:
    df_top_pages = st.session_state["pa_tp_df"]


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
if parsed:
    df_prod = parsed["products_df"]

    # Filtre par matière
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔧 Filtres")
    all_mats = sorted(parsed["mat_counter"].keys())
    selected_mats = st.sidebar.multiselect(
        "Filtrer par matière", options=all_mats, default=[], placeholder="Toutes les matières", key="pa_mats",
    )

    if selected_mats:
        mask = df_prod["Matières"].apply(lambda x: any(m in x for m in selected_mats))
        df_filtered = df_prod[mask]
        st.info(f"🔧 Filtre actif: {', '.join(selected_mats)} — {len(df_filtered)}/{len(df_prod)} produits")
    else:
        df_filtered = df_prod

    tabs = st.tabs([
        "📊 Vue d'ensemble", "🧵 Par Matière", "📦 Par Type", "🎨 Par Couleur",
        "📐 Par Coupe", "🗂 Taxonomie", "🔑 Mots-clés SEO", "🏆 Top Combinaisons",
        "🔍 Positions & Opportunités",
    ])

    # ── TAB 0: VUE D'ENSEMBLE ──
    with tabs[0]:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Produits", parsed["total"])
        c2.metric("Types", len(parsed["type_counter"]))
        c3.metric("Matières", len(parsed["mat_counter"]))
        c4.metric("Couleurs", len(parsed["color_counter"]))

        col1, col2 = st.columns(2)
        with col1:
            top_types = parsed["type_counter"].most_common(15)
            fig = px.bar(x=[t[1] for t in top_types], y=[t[0] for t in top_types],
                         orientation="h", title="Top 15 types de produits", labels={"x": "Nombre", "y": "Type"})
            fig.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            top_mats = parsed["mat_counter"].most_common(10)
            fig = px.pie(names=[t[0] for t in top_mats], values=[t[1] for t in top_mats],
                         title="Répartition par matière (Top 10)")
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Tous les produits")
        st.dataframe(df_filtered, use_container_width=True, height=400)

    # ── TAB 1: PAR MATIÈRE ──
    with tabs[1]:
        st.subheader("🧵 Répartition par Matière")
        mat_data = []
        for m, count in parsed["mat_counter"].most_common():
            if selected_mats and m not in selected_mats:
                continue
            prices = df_prod[df_prod["Matières"].str.contains(re.escape(m), na=False)]["Prix (€)"]
            mat_data.append({
                "Matière": m, "Produits": count, "%": f"{count/parsed['total']*100:.1f}%",
                "Prix moyen": f"{prices.mean():.0f}€" if len(prices) else "—",
                "Prix min": f"{prices.min():.0f}€" if len(prices) else "—",
                "Prix max": f"{prices.max():.0f}€" if len(prices) else "—",
            })
        st.dataframe(pd.DataFrame(mat_data), use_container_width=True, hide_index=True)
        fig = px.bar(pd.DataFrame(mat_data).head(20), x="Matière", y="Produits",
                     title="Nombre de produits par matière", color="Produits", color_continuous_scale="Blues")
        st.plotly_chart(fig, use_container_width=True)

    # ── TAB 2: PAR TYPE ──
    with tabs[2]:
        st.subheader("📦 Répartition par Type")
        type_data = [{"Type": t, "Produits": c, "%": f"{c/parsed['total']*100:.1f}%"}
                     for t, c in parsed["type_counter"].most_common()]
        st.dataframe(pd.DataFrame(type_data), use_container_width=True, hide_index=True)
        fig = px.bar(pd.DataFrame(type_data).head(20), x="Type", y="Produits",
                     title="Nombre de produits par type", color="Produits")
        st.plotly_chart(fig, use_container_width=True)

    # ── TAB 3: PAR COULEUR ──
    with tabs[3]:
        st.subheader("🎨 Répartition par Couleur")
        color_data = [{"Couleur": c, "Produits": n, "%": f"{n/parsed['total']*100:.1f}%"}
                      for c, n in parsed["color_counter"].most_common()]
        st.dataframe(pd.DataFrame(color_data), use_container_width=True, hide_index=True, height=500)

    # ── TAB 4: PAR COUPE ──
    with tabs[4]:
        st.subheader("📐 Répartition par Coupe")
        coupe_data = [{"Coupe": c, "Produits": n, "%": f"{n/parsed['total']*100:.1f}%"}
                      for c, n in parsed["coupe_counter"].most_common()]
        st.dataframe(pd.DataFrame(coupe_data), use_container_width=True, hide_index=True)
        if parsed["type_coupe"]:
            cross = [{"Type": t, "Coupe": cp, "Produits": n}
                     for (t, cp), n in parsed["type_coupe"].most_common(30)]
            fig = px.bar(pd.DataFrame(cross), x="Type", y="Produits", color="Coupe",
                         title="Type × Coupe (Top 30)", barmode="group")
            st.plotly_chart(fig, use_container_width=True)

    # ── TAB 5: TAXONOMIE ──
    with tabs[5]:
        st.subheader("🗂 Taxonomie — Valeurs uniques par attribut")
        tax = parsed["taxonomy"]
        max_len = max(len(v) for v in tax.values()) if tax else 0
        tax_df = {}
        for attr, vals in sorted(tax.items()):
            sorted_vals = sorted(vals)
            tax_df[f"{attr} ({len(sorted_vals)})"] = sorted_vals + [""] * (max_len - len(sorted_vals))
        st.dataframe(pd.DataFrame(tax_df), use_container_width=True, height=600)

    # ── TAB 6: MOTS-CLÉS SEO ──
    with tabs[6]:
        st.subheader("🔑 Combinaisons de Mots-clés SEO")
        seo_df = generate_seo_keywords(parsed["taxonomy"])
        if selected_mats:
            seo_df = seo_df[(seo_df["Attribut"] != "Matière") | (seo_df["Valeur"].isin(selected_mats))]
        st.metric("Combinaisons générées", len(seo_df))

        attr_filter = st.multiselect("Filtrer par attribut", options=seo_df["Attribut"].unique().tolist(),
                                     default=seo_df["Attribut"].unique().tolist(), key="pa_attr")
        filtered_seo = seo_df[seo_df["Attribut"].isin(attr_filter)]
        st.dataframe(filtered_seo, use_container_width=True, height=500)

        if df_keywords is not None and "Keyword" in df_keywords.columns:
            st.markdown("---")
            st.subheader("📊 Enrichissement avec volumes Ahrefs")
            kw_map = {}
            for _, row in df_keywords.iterrows():
                kw = str(row.get("Keyword", "")).strip().lower()
                kw_map[kw] = {"Volume": row.get("Volume", 0), "KD": row.get("Difficulty", 0),
                              "CPC": row.get("CPC", 0), "Traffic pot.": row.get("Traffic potential", 0)}
            enriched = []
            for _, row in filtered_seo.iterrows():
                kw = row["Mot-clé"].lower()
                data = kw_map.get(kw, {})
                enriched.append({**row.to_dict(), "Volume": data.get("Volume", "N/A"),
                                 "KD": data.get("KD", "N/A"), "CPC (€)": data.get("CPC", "N/A"),
                                 "Traffic pot.": data.get("Traffic pot.", "N/A")})
            df_enriched = pd.DataFrame(enriched)
            df_with_vol = df_enriched[df_enriched["Volume"] != "N/A"].sort_values("Volume", ascending=False)
            st.metric("Mots-clés avec volume", len(df_with_vol))
            st.dataframe(df_with_vol, use_container_width=True, height=400)

    # ── TAB 7: TOP COMBINAISONS ──
    with tabs[7]:
        st.subheader("🏆 Combinaisons les plus fréquentes")
        sections = [
            ("🧵 Type + Matière", parsed["mat_type"], ["Type", "Matière"]),
            ("🎨 Type + Couleur", parsed["type_color"], ["Type", "Couleur"]),
            ("📐 Type + Coupe", parsed["type_coupe"], ["Type", "Coupe"]),
            ("🔗 Type + Mat + Couleur", parsed["mat_type_color"], ["Type", "Matière", "Couleur"]),
        ]
        for title, counter, cols in sections:
            st.markdown(f"### {title}")
            items = counter.most_common(20)
            rows = []
            for keys, count in items:
                if isinstance(keys, tuple):
                    row = {c: k for c, k in zip(cols, keys)}
                else:
                    row = {cols[0]: keys}
                row["Produits"] = count
                row["%"] = f"{count/parsed['total']*100:.1f}%"
                rows.append(row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── TAB 8: POSITIONS & OPPORTUNITÉS ──
    with tabs[8]:
        st.subheader("🔍 Positions & Opportunités SEO")
        st.markdown("Croise les **mots-clés produits** avec les **pages qui rankent** pour trouver les trous.")

        if df_top_pages is None:
            st.warning("⬆️ Upload le CSV **Top Pages** (Ahrefs Site Explorer) dans la sidebar pour activer cette vue.")
        else:
            page_kw_map = {}
            for _, row in df_top_pages.iterrows():
                kw = str(row.get("Top keyword", "")).strip().lower()
                if kw:
                    page_kw_map[kw] = {"URL": row.get("URL", ""), "Position": row.get("Top keyword: Position", ""),
                                       "Traffic": row.get("Traffic", 0), "Traffic value": row.get("Traffic value", 0)}

            seo_df = generate_seo_keywords(parsed["taxonomy"])
            if selected_mats:
                seo_df = seo_df[(seo_df["Attribut"] != "Matière") | (seo_df["Valeur"].isin(selected_mats))]

            kw_vol_map = {}
            if df_keywords is not None and "Keyword" in df_keywords.columns:
                for _, row in df_keywords.iterrows():
                    kw = str(row.get("Keyword", "")).strip().lower()
                    kw_vol_map[kw] = {"Volume": row.get("Volume", 0), "KD": row.get("Difficulty", 0)}

            results = []
            for _, row in seo_df.iterrows():
                kw = row["Mot-clé"].lower()
                vol_data = kw_vol_map.get(kw, {})
                page_data = page_kw_map.get(kw, {})
                volume = vol_data.get("Volume", "N/A")
                kd = vol_data.get("KD", "N/A")
                url = page_data.get("URL", "")
                position = page_data.get("Position", "")
                traffic = page_data.get("Traffic", 0)

                if url:
                    if isinstance(position, (int, float)) and position <= 3:
                        statut = "🟢 Top 3"
                    elif isinstance(position, (int, float)) and position <= 10:
                        statut = "🟡 Top 10"
                    elif isinstance(position, (int, float)) and position <= 20:
                        statut = "🟠 Page 2"
                    else:
                        statut = "⚪ Position lointaine"
                else:
                    statut = "🔴 Aucune page"

                results.append({"Mot-clé": kw, "Type": row["Type"], "Attribut": row["Attribut"],
                                "Valeur": row["Valeur"], "Volume": volume, "KD": kd,
                                "Page qui ranke": url, "Position": position if position else "—",
                                "Traffic actuel": traffic if traffic else 0, "Statut": statut})

            df_results = pd.DataFrame(results)

            mat_filter_opp = st.selectbox("🧵 Filtrer par matière",
                                          options=["Toutes"] + sorted(parsed["mat_counter"].keys()),
                                          index=0, key="pa_mat_opp")
            if mat_filter_opp != "Toutes":
                df_results = df_results[
                    (df_results["Attribut"] != "Matière") |
                    (df_results["Valeur"].str.lower() == mat_filter_opp.lower())
                ]

            statut_filter = st.multiselect(
                "Filtrer par statut",
                options=["🔴 Aucune page", "🟠 Page 2", "🟡 Top 10", "🟢 Top 3", "⚪ Position lointaine"],
                default=["🔴 Aucune page", "🟠 Page 2", "🟡 Top 10"],
                key="pa_statut",
            )
            if statut_filter:
                df_results = df_results[df_results["Statut"].isin(statut_filter)]

            df_display = df_results.copy()
            df_display["_vol_sort"] = pd.to_numeric(df_display["Volume"], errors="coerce").fillna(-1)
            df_display = df_display.sort_values("_vol_sort", ascending=False).drop(columns=["_vol_sort"])

            total_kw = len(df_results)
            no_page = len(df_results[df_results["Statut"] == "🔴 Aucune page"])
            top3 = len(df_results[df_results["Statut"] == "🟢 Top 3"])
            top10 = len(df_results[df_results["Statut"] == "🟡 Top 10"])

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Mots-clés analysés", total_kw)
            c2.metric("🔴 Sans page", no_page)
            c3.metric("🟢 Top 3", top3)
            c4.metric("🟡 Top 10", top10)

            statut_counts = df_results["Statut"].value_counts()
            fig = px.pie(names=statut_counts.index, values=statut_counts.values,
                         title="Couverture SEO",
                         color_discrete_sequence=["#ff4444", "#ff8800", "#ffcc00", "#44bb44", "#cccccc"])
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 Détail complet")
            st.dataframe(df_display, use_container_width=True, height=600)

            st.markdown("### 🎯 Opportunités par matière (mots-clés sans page dédiée)")
            opp_by_mat = df_results[
                (df_results["Statut"] == "🔴 Aucune page") & (df_results["Attribut"] == "Matière")
            ]
            if len(opp_by_mat):
                opp_summary = opp_by_mat.groupby("Valeur").size().sort_values(ascending=False).reset_index()
                opp_summary.columns = ["Matière", "Mots-clés sans page"]
                st.dataframe(opp_summary, use_container_width=True, hide_index=True)
                fig = px.bar(opp_summary, x="Matière", y="Mots-clés sans page",
                             title="Matières avec le plus de trous SEO",
                             color="Mots-clés sans page", color_continuous_scale="Reds")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("✅ Toutes les combinaisons matière ont une page !")

else:
    st.info("👈 Charge des données produits depuis la sidebar pour commencer.")
    st.markdown("""
    ### Comment utiliser
    1. **Option A** : Upload un fichier JSON Shopify (exporté via `/products.json`)
    2. **Option B** : Entre un domaine Shopify et clique "Scraper"
    3. **(Optionnel)** : Upload les CSV Ahrefs pour enrichir l'analyse SEO

    ### Fichiers Ahrefs attendus
    - **Keywords Explorer** → Export CSV de recherche de mots-clés
    - **Top Pages** → Site Explorer → Top Pages → Export CSV
    """)
