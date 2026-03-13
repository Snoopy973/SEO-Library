import streamlit as st
import pandas as pd
import json
import re
import ssl
import urllib.request
import urllib.error
from collections import Counter, defaultdict
from io import BytesIO, StringIO
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE

st.title("📊 Analyse SEO Complète")
st.markdown("Scrape un site, charge tes exports Ahrefs, et obtiens l'analyse croisée en un seul endroit.")

# ═════════════════════════════════════════════
# FONCTIONS SCRAPING
# ═════════════════════════════════════════════

def fetch_shopify_products(store_domain, progress_bar=None):
    all_products = []
    page = 1
    store_domain = store_domain.replace("https://", "").replace("http://", "").strip("/")
    if not store_domain.startswith("www."):
        store_domain = "www." + store_domain
    base_url = f"https://{store_domain}/products.json?limit=250"
    while True:
        url = f"{base_url}&page={page}"
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
            })
            with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as response:
                data = json.loads(response.read().decode())
            products = data.get("products", [])
            if not products:
                break
            all_products.extend(products)
            if progress_bar:
                progress_bar.progress(min(page * 0.15, 0.95), f"Page {page} — {len(all_products)} produits...")
            page += 1
        except urllib.error.HTTPError as e:
            if e.code == 429:
                import time; time.sleep(2)
                continue
            else:
                st.error(f"Erreur HTTP {e.code}: {e.reason}")
                break
        except Exception as e:
            st.error(f"Erreur: {e}")
            break
    if progress_bar:
        progress_bar.progress(1.0, f"Terminé — {len(all_products)} produits")
    return all_products


def fetch_woo_products(store_domain, progress_bar=None):
    all_products = []
    page = 1
    store_domain = store_domain.replace("https://", "").replace("http://", "").strip("/")
    if not store_domain.startswith("www."):
        store_domain = "www." + store_domain
    base_url = f"https://{store_domain}/wp-json/wc/store/products?per_page=100"
    while True:
        url = f"{base_url}&page={page}"
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
            })
            with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as response:
                data = json.loads(response.read().decode())
            if not data:
                break
            products = data if isinstance(data, list) else data.get("products", data.get("items", []))
            if not products:
                break
            all_products.extend(products)
            if progress_bar:
                progress_bar.progress(min(page * 0.15, 0.95), f"Page {page} — {len(all_products)} produits...")
            page += 1
        except urllib.error.HTTPError as e:
            if e.code in (400, 404):
                break
            elif e.code == 429:
                import time; time.sleep(2)
                continue
            else:
                st.error(f"Erreur HTTP {e.code}: {e.reason}")
                break
        except Exception as e:
            st.error(f"Erreur: {e}")
            break
    if progress_bar:
        progress_bar.progress(1.0, f"Terminé — {len(all_products)} produits")
    return all_products


# ═════════════════════════════════════════════
# FONCTIONS EXTRACTION
# ═════════════════════════════════════════════

def extract_tag_values(tags, prefix):
    if isinstance(tags, list):
        return [t.split(":", 1)[1].strip().capitalize()
                for t in tags if isinstance(t, str) and t.lower().startswith(prefix + ":")]
    return []


def extract_materials_from_description(html_body):
    if not html_body:
        return []
    text = re.sub(r'<[^>]+>', ' ', html_body).replace('&nbsp;', ' ')
    materials = []
    for match in re.findall(r'(\d+)%\s*([\wéèêë]+)', text.lower()):
        mat = match[1].strip().capitalize()
        if mat not in materials:
            materials.append(mat)
    return materials


def extract_composition(html_body):
    if not html_body:
        return ""
    text = re.sub(r'<[^>]+>', ' ', html_body)
    matches = re.findall(r'\d+%\s*[\wéèêë]+', text.lower())
    return ", ".join(matches) if matches else ""


def parse_shopify_product(p, store_domain):
    tags = p.get("tags", [])
    body = p.get("body_html", "")
    product_type = p.get("product_type", "Non défini")
    materials = extract_tag_values(tags, "matiere") or extract_tag_values(tags, "matière")
    if not materials:
        materials = extract_materials_from_description(body)
    colors = extract_tag_values(tags, "couleur")
    coupes = extract_tag_values(tags, "coupe")
    formes = extract_tag_values(tags, "forme")
    collections = [t.split(":", 1)[1].strip() for t in tags
                   if isinstance(t, str) and t.lower().startswith("collection:")]
    variants = p.get("variants", [])
    price = float(variants[0].get("price", 0)) if variants else None
    compare_price = None
    if variants and variants[0].get("compare_at_price"):
        compare_price = float(variants[0]["compare_at_price"])
    return {
        "title": p.get("title", ""),
        "type": product_type,
        "materials": materials,
        "materials_str": ", ".join(materials) if materials else "Non renseigné",
        "composition": extract_composition(body),
        "colors": colors,
        "colors_str": ", ".join(colors) if colors else "Non renseigné",
        "coupes": coupes,
        "coupes_str": ", ".join(coupes) if coupes else "",
        "formes": formes,
        "formes_str": ", ".join(formes) if formes else "",
        "collections": collections,
        "price": price,
        "compare_price": compare_price,
        "url": f"https://{store_domain}/products/{p.get('handle', '')}",
        "tags": tags,
    }


def parse_woo_product(p, store_domain):
    name = p.get("name", p.get("title", ""))
    body = p.get("description", p.get("short_description", ""))
    categories = [c.get("name", "") for c in p.get("categories", [])]
    product_type = categories[0] if categories else "Non défini"
    tags = [t.get("name", "") for t in p.get("tags", [])]
    materials = extract_materials_from_description(body)
    price = None
    prices = p.get("prices", {})
    if prices:
        price_str = prices.get("price", prices.get("regular_price", "0"))
        try:
            price = float(price_str) / 100 if len(str(price_str)) > 4 else float(price_str)
        except (ValueError, TypeError):
            pass
    elif p.get("price"):
        try:
            price = float(p["price"])
        except (ValueError, TypeError):
            pass
    permalink = p.get("permalink", p.get("slug", ""))
    if permalink and not permalink.startswith("http"):
        permalink = f"https://{store_domain}/product/{permalink}"
    return {
        "title": name,
        "type": product_type,
        "materials": materials,
        "materials_str": ", ".join(materials) if materials else "Non renseigné",
        "composition": extract_composition(body),
        "colors": [],
        "colors_str": "Non renseigné",
        "coupes": [],
        "coupes_str": "",
        "formes": [],
        "formes_str": "",
        "collections": categories,
        "price": price,
        "compare_price": None,
        "url": permalink,
        "tags": tags,
    }


def detect_cms(products):
    if not products:
        return "unknown"
    sample = products[0]
    if "variants" in sample and "vendor" in sample:
        return "shopify"
    if "prices" in sample or "price_html" in sample:
        return "woocommerce"
    return "unknown"


# ═════════════════════════════════════════════
# ANALYSE PRODUITS
# ═════════════════════════════════════════════

def analyze_parsed_products(parsed_products):
    results = {
        "products": parsed_products,
        "materials_count": Counter(),
        "type_count": Counter(),
        "color_count": Counter(),
        "coupe_count": Counter(),
        "forme_count": Counter(),
        "collection_count": Counter(),
        "material_by_type": defaultdict(Counter),
        "coupe_by_type": defaultdict(Counter),
        "price_by_material": defaultdict(list),
        "combos_type_mat": Counter(),
        "combos_type_col": Counter(),
        "combos_type_coupe": Counter(),
        "taxonomy": defaultdict(set),
        "total": len(parsed_products),
    }
    for p in parsed_products:
        ptype = p["type"]
        results["type_count"][ptype] += 1
        results["taxonomy"]["Type de produit"].add(ptype)
        for mat in p["materials"]:
            results["materials_count"][mat] += 1
            results["material_by_type"][mat][ptype] += 1
            results["taxonomy"]["Matières"].add(mat)
            if p["price"]:
                results["price_by_material"][mat].append(p["price"])
            results["combos_type_mat"][f"{ptype.lower()} {mat.lower()}"] += 1
        for col in p["colors"]:
            results["color_count"][col] += 1
            results["taxonomy"]["Couleurs"].add(col)
            results["combos_type_col"][f"{ptype.lower()} {col.lower()}"] += 1
        for coupe in p["coupes"]:
            results["coupe_count"][coupe] += 1
            results["coupe_by_type"][coupe][ptype] += 1
            results["taxonomy"]["Coupes"].add(coupe)
            results["combos_type_coupe"][f"{ptype.lower()} {coupe.lower()}"] += 1
        for forme in p["formes"]:
            results["forme_count"][forme] += 1
            results["taxonomy"]["Formes"].add(forme)
        for coll in p["collections"]:
            results["collection_count"][coll] += 1
            results["taxonomy"]["Collections"].add(coll)
    return results


# ═════════════════════════════════════════════
# PARSEURS CSV AHREFS
# ═════════════════════════════════════════════

def parse_ahrefs_csv(uploaded_file):
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
    df = pd.read_csv(StringIO(text), sep=sep)
    df.columns = [c.strip().strip('"') for c in df.columns]
    return df


def detect_ahrefs_type(df):
    cols = set(c.lower() for c in df.columns)
    if "top keyword" in cols or "top keyword: volume" in cols:
        return "top_pages"
    if "keyword" in cols and "volume" in cols:
        return "keywords"
    return "unknown"


# ═════════════════════════════════════════════
# MATCHING KEYWORDS VS PAGES
# ═════════════════════════════════════════════

def match_keywords_to_pages(df_keywords, df_pages, combos_with_materials):
    rows = []
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

    for combo_kw, materials in combos_with_materials.items():
        combo_lower = combo_kw.lower().strip()
        vol_data = vol_index.get(combo_lower, {})
        volume = vol_data.get("volume", "N/A")
        kd = vol_data.get("kd", "N/A")
        cpc = vol_data.get("cpc", "N/A")
        tp = vol_data.get("traffic_potential", "N/A")

        page_data = page_index.get(combo_lower, {})
        page_url = page_data.get("url", "")
        position = page_data.get("position", "")
        traffic = page_data.get("traffic", 0)

        if not page_url and df_pages is not None:
            for _, prow in df_pages.iterrows():
                url = str(prow.get("URL", "")).lower()
                slug = combo_lower.replace(" ", "-")
                if slug in url:
                    page_url = prow.get("URL", "")
                    position = prow.get("Top keyword: Position", "")
                    traffic = prow.get("Traffic", 0)
                    break

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
    if not df.empty:
        df["_vol_sort"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(-1)
        df = df.sort_values("_vol_sort", ascending=False).drop(columns="_vol_sort")
    return df


# ═════════════════════════════════════════════
# EXPORT EXCEL 10 ONGLETS
# ═════════════════════════════════════════════

def build_excel(results, df_matched, store_name):
    buffer = BytesIO()
    wb = Workbook()
    total = results["total"]

    hf = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
    hfont = Font(bold=True, color="FFFFFF", size=11)

    def style_header(ws, max_col, row=1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = hf
            cell.font = hfont
            cell.alignment = Alignment(horizontal="center")

    # 1. Par Matière
    ws = wb.active
    ws.title = "Par Matière"
    ws.append(["Matière", "Nb produits", "% catalogue", "Prix moyen", "Prix min", "Prix max"])
    style_header(ws, 6)
    for mat, count in results["materials_count"].most_common():
        prices = results["price_by_material"].get(mat, [])
        ws.append([mat, count, f"{count/total*100:.1f}%",
                    round(sum(prices)/len(prices), 2) if prices else "",
                    round(min(prices), 2) if prices else "",
                    round(max(prices), 2) if prices else ""])

    # 2. Par Type
    ws2 = wb.create_sheet("Par Type")
    ws2.append(["Type", "Nb produits", "% catalogue"])
    style_header(ws2, 3)
    for t, c in results["type_count"].most_common():
        ws2.append([t, c, f"{c/total*100:.1f}%"])

    # 3. Matière x Type
    ws3 = wb.create_sheet("Matière x Type")
    all_types = sorted(set(t for mat_c in results["material_by_type"].values() for t in mat_c))
    ws3.append(["Matière"] + all_types)
    style_header(ws3, 1 + len(all_types))
    for mat, tc in sorted(results["material_by_type"].items()):
        ws3.append([mat] + [tc.get(t, 0) for t in all_types])

    # 4. Par Couleur
    ws4 = wb.create_sheet("Par Couleur")
    ws4.append(["Couleur", "Nb produits", "% catalogue"])
    style_header(ws4, 3)
    for col, c in results["color_count"].most_common():
        ws4.append([col, c, f"{c/total*100:.1f}%"])

    # 5. Par Coupe
    ws5 = wb.create_sheet("Par Coupe")
    ws5.append(["Coupe", "Nb produits", "% catalogue"])
    style_header(ws5, 3)
    for cp, c in results["coupe_count"].most_common():
        ws5.append([cp, c, f"{c/total*100:.1f}%"])

    # 6. Tous les produits
    ws6 = wb.create_sheet("Tous les produits")
    ws6.append(["Produit", "Type", "Matières", "Composition", "Couleurs", "Coupes", "Formes", "Prix", "Ancien prix", "URL"])
    style_header(ws6, 10)
    for p in results["products"]:
        ws6.append([p["title"], p["type"], p["materials_str"], p["composition"],
                     p["colors_str"], p["coupes_str"], p["formes_str"],
                     p["price"], p.get("compare_price", ""), p["url"]])

    # 7. Taxonomie
    ws7 = wb.create_sheet("Taxonomie")
    taxonomy = results["taxonomy"]
    if taxonomy:
        all_attrs = sorted(taxonomy.keys())
        max_len = max(len(v) for v in taxonomy.values()) if taxonomy else 0
        ws7.append(all_attrs)
        style_header(ws7, len(all_attrs))
        for i in range(max_len):
            row = []
            for attr in all_attrs:
                vals = sorted(taxonomy[attr])
                row.append(vals[i] if i < len(vals) else "")
            ws7.append(row)

    # 8. Mots-clés SEO
    ws8 = wb.create_sheet("Mots-clés SEO")
    headers_seo = ["cat", "cat + matières", "cat + couleurs", "cat + coupes", "cat + formes"]
    ws8.append(headers_seo)
    style_header(ws8, len(headers_seo))
    types_list = sorted(results["type_count"].keys())
    mats_list = sorted(results["materials_count"].keys())
    cols_list = sorted(results["color_count"].keys())
    coupes_list = sorted(results["coupe_count"].keys())
    formes_list = sorted(results["forme_count"].keys())
    max_combos = max(
        len(types_list) * max(len(mats_list), 1),
        len(types_list) * max(len(cols_list), 1),
        len(types_list) * max(len(coupes_list), 1),
        len(types_list) * max(len(formes_list), 1),
        1
    )
    idx = 0
    for t in types_list:
        for mat in (mats_list or [""]):
            row = [t if idx == 0 or mat == mats_list[0] else ""]
            row.append(f"{t.lower()} {mat.lower()}" if mat else "")
            # couleur
            col_val = cols_list[idx % len(cols_list)] if cols_list and idx < len(types_list) * len(cols_list) else ""
            row.append(f"{t.lower()} {col_val.lower()}" if col_val else "")
            # coupe
            cp_val = coupes_list[idx % len(coupes_list)] if coupes_list and idx < len(types_list) * len(coupes_list) else ""
            row.append(f"{t.lower()} {cp_val.lower()}" if cp_val else "")
            # forme
            fm_val = formes_list[idx % len(formes_list)] if formes_list and idx < len(types_list) * len(formes_list) else ""
            row.append(f"{t.lower()} {fm_val.lower()}" if fm_val else "")
            ws8.append(row)
            idx += 1

    # 9. Top Combinaisons
    ws9 = wb.create_sheet("Top Combinaisons")
    ws9.append(["Combinaison", "Nb produits", "% du total"])
    style_header(ws9, 3)
    for combo, count in results["combos_type_mat"].most_common():
        ws9.append([combo, count, f"{count/total*100:.1f}%"])

    # 10. Requêtes vs Pages
    if df_matched is not None and not df_matched.empty:
        ws10 = wb.create_sheet("Requêtes vs Pages")
        cols_export = ["Mot-clé", "Matière", "Volume", "KD", "Potentiel trafic", "CPC (€)",
                        "Statut", "Page qui ranke", "Position", "Traffic actuel"]
        ws10.append(cols_export)
        style_header(ws10, len(cols_export))
        for _, row in df_matched.iterrows():
            ws10.append([row.get(c, "") for c in cols_export])

        # Couleurs par statut
        status_colors = {
            "🟢 Top 3": "27AE60",
            "🟡 Top 10": "F1C40F",
            "🟠 Page 2": "E67E22",
            "⚪ > 20": "BDC3C7",
            "🔴 Pas de page": "E74C3C",
        }
        statut_col_idx = cols_export.index("Statut") + 1
        for row_idx in range(2, ws10.max_row + 1):
            cell = ws10.cell(row=row_idx, column=statut_col_idx)
            color = status_colors.get(str(cell.value), "FFFFFF")
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            if color in ("27AE60", "E74C3C"):
                cell.font = Font(color="FFFFFF", bold=True)

    wb.save(buffer)
    return buffer.getvalue()


# ═════════════════════════════════════════════
# SIDEBAR — INPUTS
# ═════════════════════════════════════════════

with st.sidebar:
    st.header("📥 Sources de données")

    # 1. URL du site
    st.markdown("**1. Site e-commerce**")
    store_url = st.text_input("URL du site", placeholder="www.balibaris.com",
                               help="Shopify ou WooCommerce")
    cms_choice = st.selectbox("CMS", ["Auto-détecter", "Shopify", "WooCommerce"])

    if st.button("🚀 Scraper", type="primary", use_container_width=True):
        if store_url:
            progress = st.progress(0, "Scraping...")
            domain = store_url.replace("https://", "").replace("http://", "").strip("/")
            if cms_choice == "WooCommerce":
                raw = fetch_woo_products(domain, progress)
                cms = "woocommerce"
            elif cms_choice == "Shopify":
                raw = fetch_shopify_products(domain, progress)
                cms = "shopify"
            else:
                raw = fetch_shopify_products(domain, progress)
                cms = "shopify"
                if not raw:
                    st.info("Tentative WooCommerce...")
                    raw = fetch_woo_products(domain, progress)
                    cms = "woocommerce"
            if raw:
                store_clean = domain.replace("www.", "")
                parser = parse_shopify_product if cms == "shopify" else parse_woo_product
                parsed = [parser(p, domain) for p in raw]
                st.session_state["ac_results"] = analyze_parsed_products(parsed)
                st.session_state["ac_store"] = store_clean
                st.success(f"✅ {len(raw)} produits")
            else:
                st.error("❌ Aucun produit trouvé")

    # Statut produits
    if "ac_results" in st.session_state:
        r = st.session_state["ac_results"]
        st.success(f"✅ {r['total']} produits chargés")

    st.divider()

    # 2. CSV Mots-clés
    st.markdown("**2. Mots-clés Ahrefs**")
    uploaded_kw = st.file_uploader("CSV Keywords", type=["csv"], key="ac_kw",
                                    help="Export Ahrefs Keywords Explorer")

    # 3. CSV Top Pages
    st.markdown("**3. Top Pages Ahrefs**")
    uploaded_pages = st.file_uploader("CSV Top Pages", type=["csv"], key="ac_pages",
                                       help="Export Ahrefs Top Pages")

    # Parser les CSV
    df_keywords = None
    df_pages = None

    if uploaded_kw:
        df_keywords = parse_ahrefs_csv(uploaded_kw)
        st.session_state["ac_df_keywords"] = df_keywords
        st.success(f"✅ {len(df_keywords)} mots-clés")
    elif "ac_df_keywords" in st.session_state:
        df_keywords = st.session_state["ac_df_keywords"]

    if uploaded_pages:
        df_pages = parse_ahrefs_csv(uploaded_pages)
        st.session_state["ac_df_pages"] = df_pages
        st.success(f"✅ {len(df_pages)} pages")
    elif "ac_df_pages" in st.session_state:
        df_pages = st.session_state["ac_df_pages"]


# ═════════════════════════════════════════════
# ZONE PRINCIPALE — RÉSULTATS
# ═════════════════════════════════════════════

has_products = "ac_results" in st.session_state
has_ahrefs = df_keywords is not None or df_pages is not None

if not has_products and not has_ahrefs:
    st.info("👈 Utilise la sidebar pour charger tes données : URL du site + exports Ahrefs.")
    st.stop()

# Préparer le matching si on a assez de données
df_matched = pd.DataFrame()
combos_with_materials = {}

if has_products:
    results = st.session_state["ac_results"]
    store = st.session_state.get("ac_store", "site")
    total = results["total"]

    # Construire les combos
    for combo_key in results.get("combos_type_mat", {}):
        parts = combo_key.split()
        if len(parts) >= 2:
            mat = " ".join(parts[1:]).capitalize()
            combos_with_materials[combo_key] = [mat]
    for combo_key in results.get("combos_type_col", {}):
        if combo_key not in combos_with_materials:
            combos_with_materials[combo_key] = []
    for combo_key in results.get("combos_type_coupe", {}):
        if combo_key not in combos_with_materials:
            combos_with_materials[combo_key] = []

if not combos_with_materials and df_keywords is not None:
    for _, row in df_keywords.iterrows():
        kw = str(row.get("Keyword", ""))
        combos_with_materials[kw] = []

if has_ahrefs and combos_with_materials:
    df_matched = match_keywords_to_pages(df_keywords, df_pages, combos_with_materials)


# ── TABS ──
tabs_list = []
if has_products:
    tabs_list += ["📦 Catalogue", "🔀 Croisements", "🔑 Mots-clés SEO"]
if has_ahrefs:
    tabs_list += ["📊 Positions & Gaps", "🎯 Opportunités"]
tabs_list.append("📥 Export Excel")

tabs = st.tabs(tabs_list)
tab_idx = 0

# ── TAB CATALOGUE ──
if has_products:
    with tabs[tab_idx]:
        # KPIs
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Produits", total)
        c2.metric("Types", len(results["type_count"]))
        c3.metric("Matières", len(results["materials_count"]))
        c4.metric("Couleurs", len(results["color_count"]))
        c5.metric("Coupes", len(results["coupe_count"]))

        import plotly.express as px

        sub_t1, sub_t2, sub_t3, sub_t4 = st.tabs(["🧵 Matières", "📦 Types", "🎨 Couleurs", "📐 Coupes"])

        with sub_t1:
            if results["materials_count"]:
                df_mat = pd.DataFrame(results["materials_count"].most_common(), columns=["Matière", "Nb produits"])
                df_mat["% catalogue"] = (df_mat["Nb produits"] / total * 100).round(1)
                prix_data = []
                for mat in df_mat["Matière"]:
                    prices = results["price_by_material"].get(mat, [])
                    prix_data.append({
                        "Prix moyen": round(sum(prices)/len(prices), 2) if prices else None,
                        "Prix min": round(min(prices), 2) if prices else None,
                        "Prix max": round(max(prices), 2) if prices else None,
                    })
                df_mat = pd.concat([df_mat, pd.DataFrame(prix_data)], axis=1)
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.dataframe(df_mat, use_container_width=True, hide_index=True)
                with col2:
                    fig = px.pie(df_mat.head(15), values="Nb produits", names="Matière", title="Top 15 matières")
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)

        with sub_t2:
            if results["type_count"]:
                df_type = pd.DataFrame(results["type_count"].most_common(), columns=["Type", "Nb produits"])
                df_type["% catalogue"] = (df_type["Nb produits"] / total * 100).round(1)
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.dataframe(df_type, use_container_width=True, hide_index=True)
                with col2:
                    fig = px.bar(df_type.head(15), x="Type", y="Nb produits", title="Par type")
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)

        with sub_t3:
            if results["color_count"]:
                df_col = pd.DataFrame(results["color_count"].most_common(), columns=["Couleur", "Nb produits"])
                df_col["% catalogue"] = (df_col["Nb produits"] / total * 100).round(1)
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.dataframe(df_col, use_container_width=True, hide_index=True)
                with col2:
                    fig = px.bar(df_col.head(20), x="Couleur", y="Nb produits", title="Top 20 couleurs")
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)

        with sub_t4:
            if results["coupe_count"]:
                df_coupe = pd.DataFrame(results["coupe_count"].most_common(), columns=["Coupe", "Nb produits"])
                df_coupe["% catalogue"] = (df_coupe["Nb produits"] / total * 100).round(1)
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.dataframe(df_coupe, use_container_width=True, hide_index=True)
                with col2:
                    fig = px.pie(df_coupe, values="Nb produits", names="Coupe", title="Coupes")
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Aucune donnée de coupe trouvée")

    tab_idx += 1

    # ── TAB CROISEMENTS ──
    with tabs[tab_idx]:
        import plotly.express as px

        st.markdown("#### Matière x Type de produit")
        if results["material_by_type"]:
            all_types = sorted(set(t for mat_c in results["material_by_type"].values() for t in mat_c))
            cross_data = []
            for mat, tc in sorted(results["material_by_type"].items()):
                row = {"Matière": mat}
                for t in all_types:
                    row[t] = tc.get(t, 0)
                cross_data.append(row)
            df_cross = pd.DataFrame(cross_data)
            st.dataframe(df_cross, use_container_width=True, hide_index=True)

            df_heat = df_cross.set_index("Matière")
            fig = px.imshow(df_heat.head(20), text_auto=True, aspect="auto",
                           title="Heatmap Matière x Type (top 20)")
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.markdown("#### Taxonomie")
        taxonomy = results["taxonomy"]
        if taxonomy:
            max_len = max(len(v) for v in taxonomy.values())
            tax_data = {}
            for col_name, values in sorted(taxonomy.items()):
                sorted_vals = sorted(values)
                tax_data[f"{col_name} ({len(sorted_vals)})"] = sorted_vals + [""] * (max_len - len(sorted_vals))
            st.dataframe(pd.DataFrame(tax_data), use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("#### Liste complète des produits")
        df_all = pd.DataFrame([{
            "Produit": p["title"], "Type": p["type"], "Matières": p["materials_str"],
            "Couleurs": p["colors_str"], "Coupes": p["coupes_str"],
            "Prix": p["price"], "URL": p["url"],
        } for p in results["products"]])
        st.dataframe(df_all, use_container_width=True, hide_index=True,
                      column_config={"URL": st.column_config.LinkColumn("URL")})

    tab_idx += 1

    # ── TAB MOTS-CLÉS SEO ──
    with tabs[tab_idx]:
        st.markdown("#### Combinaisons type + attribut")

        if results["combos_type_mat"]:
            st.markdown("**Type + Matière** (top 30)")
            df_combos = pd.DataFrame(results["combos_type_mat"].most_common(30),
                                      columns=["Combinaison", "Nb produits"])
            df_combos["% du total"] = (df_combos["Nb produits"] / total * 100).round(1)
            st.dataframe(df_combos, use_container_width=True, hide_index=True)

        if results["combos_type_col"]:
            st.markdown("**Type + Couleur** (top 30)")
            df_cc = pd.DataFrame(results["combos_type_col"].most_common(30),
                                  columns=["Combinaison", "Nb produits"])
            st.dataframe(df_cc, use_container_width=True, hide_index=True)

        if results["combos_type_coupe"]:
            st.markdown("**Type + Coupe** (top 30)")
            df_ccp = pd.DataFrame(results["combos_type_coupe"].most_common(30),
                                   columns=["Combinaison", "Nb produits"])
            st.dataframe(df_ccp, use_container_width=True, hide_index=True)

    tab_idx += 1

# ── TAB POSITIONS & GAPS ──
if has_ahrefs:
    with tabs[tab_idx]:
        if df_matched.empty:
            st.info("Charge au moins un CSV Ahrefs + scrape un site pour voir les positions.")
        else:
            import plotly.express as px

            # Filtres
            st.markdown("### 🎛️ Filtres")
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                all_materials = sorted(set(m for mats in combos_with_materials.values() for m in mats if m))
                sel_mat = st.multiselect("🧵 Matière", options=all_materials, default=[])
            with fc2:
                all_statuts = sorted(df_matched["Statut"].unique())
                sel_stat = st.multiselect("📌 Statut", options=all_statuts, default=[])
            with fc3:
                vol_min = st.number_input("🔢 Volume min", min_value=0, value=0, step=100)

            df_display = df_matched.copy()
            if sel_mat:
                df_display = df_display[df_display["Matière"].apply(
                    lambda x: any(m.lower() in str(x).lower() for m in sel_mat))]
            if sel_stat:
                df_display = df_display[df_display["Statut"].isin(sel_stat)]
            if vol_min > 0:
                df_display = df_display[pd.to_numeric(df_display["Volume"], errors="coerce").fillna(0) >= vol_min]

            # KPIs
            total_kw = len(df_display)
            with_page = len(df_display[df_display["Statut"] != "🔴 Pas de page"])
            without_page = len(df_display[df_display["Statut"] == "🔴 Pas de page"])
            top3 = len(df_display[df_display["Statut"] == "🟢 Top 3"])
            to_opt = len(df_display[df_display["Statut"].isin(["🟡 Top 10", "🟠 Page 2"])])

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Total", total_kw)
            k2.metric("🟢 Top 3", top3)
            k3.metric("🟡 À optimiser", to_opt)
            k4.metric("🔴 Sans page", without_page)
            k5.metric("Couverture", f"{round(with_page/max(total_kw,1)*100)}%")

            # Charts
            ch1, ch2 = st.columns(2)
            with ch1:
                statut_counts = df_display["Statut"].value_counts()
                fig = px.pie(values=statut_counts.values, names=statut_counts.index,
                            title="Répartition des statuts",
                            color_discrete_map={
                                "🟢 Top 3": "#2ecc71", "🟡 Top 10": "#f1c40f",
                                "🟠 Page 2": "#e67e22", "⚪ > 20": "#bdc3c7",
                                "🔴 Pas de page": "#e74c3c"})
                st.plotly_chart(fig, use_container_width=True)
            with ch2:
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
                                  title="🔴 Opportunités par matière")
                    st.plotly_chart(fig2, use_container_width=True)

            # Tableau
            st.markdown("### 📋 Détail")
            st.dataframe(df_display, use_container_width=True, hide_index=True,
                          column_config={
                              "Page qui ranke": st.column_config.LinkColumn("Page"),
                              "Volume": st.column_config.NumberColumn("Volume", format="%d"),
                              "Traffic actuel": st.column_config.NumberColumn("Traffic", format="%d"),
                          })

            # Vue par matière
            if all_materials:
                st.markdown("### 🧵 Vue détaillée par matière")
                mat_choice = st.selectbox("Matière", options=["Toutes"] + all_materials)
                if mat_choice != "Toutes":
                    df_mv = df_matched[df_matched["Matière"].str.contains(mat_choice, case=False, na=False)]
                    covered = len(df_mv[df_mv["Statut"] != "🔴 Pas de page"])
                    not_covered = len(df_mv[df_mv["Statut"] == "🔴 Pas de page"])
                    m1, m2, m3 = st.columns(3)
                    m1.metric(f"Mots-clés '{mat_choice}'", len(df_mv))
                    m2.metric("✅ Couverts", covered)
                    m3.metric("🔴 Non couverts", not_covered)
                    st.dataframe(df_mv, use_container_width=True, hide_index=True)

    tab_idx += 1

    # ── TAB OPPORTUNITÉS ──
    with tabs[tab_idx]:
        if df_matched.empty:
            st.info("Pas encore de données croisées.")
        else:
            st.markdown("### 🎯 Top Opportunités — mots-clés sans page dédiée")
            df_opps = df_matched[df_matched["Statut"] == "🔴 Pas de page"].copy()
            df_opps["_vol"] = pd.to_numeric(df_opps["Volume"], errors="coerce").fillna(0)
            df_opps = df_opps.sort_values("_vol", ascending=False).drop(columns="_vol")

            if not df_opps.empty:
                st.metric("Nombre d'opportunités", len(df_opps))
                st.dataframe(df_opps, use_container_width=True, hide_index=True,
                              column_config={"Volume": st.column_config.NumberColumn("Volume", format="%d")})
            else:
                st.success("🎉 Toutes les combinaisons sont couvertes !")

    tab_idx += 1

# ── TAB EXPORT ──
with tabs[tab_idx]:
    st.markdown("### 📥 Export Excel complet")

    if has_products:
        st.markdown("L'export contiendra jusqu'à **10 onglets** : Par Matière, Par Type, Matière x Type, "
                     "Par Couleur, Par Coupe, Tous les produits, Taxonomie, Mots-clés SEO, "
                     "Top Combinaisons, Requêtes vs Pages.")

        if st.button("💾 Générer le fichier Excel", type="primary", use_container_width=True):
            excel_data = build_excel(results, df_matched if not df_matched.empty else None, store)
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                "⬇️ Télécharger l'Excel",
                data=excel_data,
                file_name=f"analyse_{store}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    elif has_ahrefs and not df_matched.empty:
        if st.button("💾 Exporter les positions", type="primary", use_container_width=True):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_matched.to_excel(writer, sheet_name="Positions", index=False)
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                "⬇️ Télécharger",
                data=buffer.getvalue(),
                file_name=f"positions_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("Charge des données pour pouvoir exporter.")
