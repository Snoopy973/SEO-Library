import streamlit as st
import requests
import pandas as pd
import json
import re
from datetime import datetime
from collections import Counter
import io

# ============================================================================
# CONFIG
# ============================================================================
# st.set_page_config est appelé par seo_tools.py (entrypoint)

# SESSION STATE
if 'results' not in st.session_state:
    st.session_state.results = []

# ============================================================================
# CONSTANTES
# ============================================================================
COUNTRIES = {
    "France": {"location_name": "France", "language_name": "French", "se_domain": "google.fr"},
    "Belgium (FR)": {"location_name": "Belgium", "language_name": "French", "se_domain": "google.be"},
    "Switzerland (FR)": {"location_name": "Switzerland", "language_name": "French", "se_domain": "google.ch"},
    "United States": {"location_name": "United States", "language_name": "English", "se_domain": "google.com"},
    "Germany": {"location_name": "Germany", "language_name": "German", "se_domain": "google.de"},
    "Spain": {"location_name": "Spain", "language_name": "Spanish", "se_domain": "google.es"},
    "Italy": {"location_name": "Italy", "language_name": "Italian", "se_domain": "google.it"},
}

STOPWORDS = {'le', 'la', 'les', 'un', 'une', 'des', 'de', 'du', 'à', 'au', 'en', 'et',
             'ou', 'qui', 'que', 'ce', 'pour', 'par', 'sur', 'avec', 'dans', 'est', 'sont',
             'the', 'a', 'an', 'is', 'are', 'to', 'of', 'in', 'for', 'on', 'with', 'and'}

LLM_PROVIDERS = {
    "Claude (Anthropic)": ["claude-sonnet-4-20250514", "claude-3-5-sonnet-20241022"],
    "GPT (OpenAI)": ["gpt-4o", "gpt-4o-mini", "gpt-3.5-turbo"],
    "Mistral": ["mistral-large-latest", "mistral-small-latest"],
    "Groq": ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"],
}

DEFAULT_PROMPT = """Tu es un expert SEO. Ta mission est d'analyser les balises title RÉELLES de la SERP Google et de proposer une title qui RESPECTE les structures gagnantes.

RÈGLE ABSOLUE: Les titles scrappées ci-dessous sont la VÉRITÉ. Ta proposition DOIT s'en inspirer directement. Tu ne dois PAS inventer une structure qui n'existe pas dans les données.

DONNÉES SCRAPPÉES (SOURCE DE VÉRITÉ):
===========================================
BALISES TITLE GOOGLE (Top {titles_count}):
{titles_list}

PATTERNS LES PLUS FRÉQUENTS (N-grams):
{ngrams_list}
===========================================

CONTEXTE:
- Mot-clé initial: "{keyword}"
- Marque (optionnel): "{brand}"
- Longueur max: {max_length} caractères

MÉTHODE D'ANALYSE OBLIGATOIRE:
1. IDENTIFIER LA STRUCTURE DOMINANTE: Regarde comment les titles commencent (ex: "Comment...", "Les meilleurs...", "[Mot] : ..."). La structure qui revient le plus = la structure à utiliser.
2. IDENTIFIER LES MOTS RÉCURRENTS: Les mots qui apparaissent dans plusieurs titles sont les mots que Google valorise. Utilise-les.
3. IDENTIFIER LE FORMAT: Note les séparateurs (: - |), la présence de dates, de chiffres, etc.

RÈGLES STRICTES:
- Ta title DOIT reprendre la structure dominante des titles scrappées
- Ta title DOIT contenir les mots/expressions qui reviennent le plus souvent
- Tu ne dois PAS inventer des mots ou structures absents des données
- Si 6 titles sur 10 commencent par "Comment", ta title DOIT commencer par "Comment"
- Si "meilleur", "guide", "comparatif" reviennent souvent, utilise-les
- Longueur: 50-{max_length} caractères max

EXEMPLE DE RAISONNEMENT:
Si les titles sont:
- "Comment choisir son matelas"
- "Comment bien choisir un matelas"
- "Comment choisir le bon matelas"
→ Structure dominante = "Comment [bien] choisir [son/un/le bon] matelas"
→ Ta proposition = "Comment bien choisir son matelas : guide complet" (si "guide" apparaît ailleurs)

RÉPONDS UNIQUEMENT EN JSON STRICT:
{{"title_proposed": "Title basée sur la structure dominante", "title_length": 55, "titles_alternatives": ["Variante 1 basée sur les données", "Variante 2 basée sur les données", "Variante 3 basée sur les données"], "structure_dominante": "La structure que tu as identifiée", "mots_cles_recurrents": ["mot1", "mot2", "mot3"], "intention_detectee": "transactionnelle|informationnelle|navigationnelle", "score_confiance": 85, "justification": "J'ai choisi cette structure car elle apparaît dans X titles sur {titles_count}"}}"""

# ============================================================================
# FONCTIONS N-GRAMS
# ============================================================================
def clean_title(title):
    if not title:
        return ""
    title = re.sub(r'\s*[\|\-–—:]\s*[^|\-–—:]{0,40}$', '', title)
    title = re.sub(r'[^\w\sàâäéèêëïîôùûüçœæ\'-]', ' ', title, flags=re.IGNORECASE)
    title = re.sub(r'\s+', ' ', title).strip()
    return title.lower()

def extract_ngrams(text, n):
    words = [w for w in text.split() if w not in STOPWORDS and len(w) > 2]
    if len(words) < n:
        return []
    return [' '.join(words[i:i+n]) for i in range(len(words) - n + 1)]

def analyze_titles(titles):
    all_ngrams = Counter()
    for title in titles:
        cleaned = clean_title(title)
        for n in range(2, 6):
            all_ngrams.update(extract_ngrams(cleaned, n))
    return all_ngrams.most_common(20)

# ============================================================================
# API DATAFORSEO
# ============================================================================
def get_serp_titles(keyword, login, password, config, depth=10):
    url = "https://api.dataforseo.com/v3/serp/google/organic/live/advanced"

    payload = [{
        "keyword": keyword,
        "location_name": config["location_name"],
        "language_name": config["language_name"],
        "se_domain": config["se_domain"],
        "depth": depth,
        "device": "desktop",
        "os": "windows"
    }]

    try:
        response = requests.post(
            url,
            auth=(login, password),
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=120
        )

        if response.status_code != 200:
            return None, f"HTTP {response.status_code}"

        data = response.json()

        if data.get("status_code") != 20000:
            return None, f"API: {data.get('status_message')}"

        tasks = data.get("tasks", [])
        if not tasks:
            return None, "Pas de tasks"

        task = tasks[0]
        if task.get("status_code") != 20000:
            return None, f"Task: {task.get('status_message')}"

        results = task.get("result", [])
        if not results:
            return None, "Pas de résultats"

        items = results[0].get("items", [])
        if not items:
            return None, "Pas d'items"

        titles = []
        for item in items:
            if item.get("type") == "organic":
                title = item.get("title", "")
                if title:
                    titles.append(title)

        if not titles:
            return None, "Pas de titles organiques"

        return titles, None

    except requests.exceptions.Timeout:
        return None, "Timeout"
    except Exception as e:
        return None, str(e)

# ============================================================================
# API LLM
# ============================================================================
def analyze_with_llm(titles, keyword, brand, max_length, top_ngrams, provider, model, api_key, custom_prompt):
    # Préparer les variables pour le prompt
    titles_list = chr(10).join([f"- {t}" for t in titles])
    ngrams_list = chr(10).join([f"- '{ng}': {c} occurrences" for ng, c in top_ngrams[:10]])

    # Remplacer les placeholders dans le prompt
    prompt = custom_prompt.format(
        keyword=keyword,
        brand=brand if brand else "Non spécifié",
        max_length=max_length,
        titles_count=len(titles),
        titles_list=titles_list,
        ngrams_list=ngrams_list
    )

    try:
        if provider == "Claude (Anthropic)":
            resp = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={"Content-Type": "application/json", "x-api-key": api_key, "anthropic-version": "2023-06-01"},
                json={"model": model, "max_tokens": 1024, "messages": [{"role": "user", "content": prompt}]},
                timeout=60
            )
            if resp.status_code != 200:
                return None, f"Claude API: {resp.status_code} - {resp.text}"
            text = resp.json().get("content", [{}])[0].get("text", "")

        elif provider == "GPT (OpenAI)":
            resp = requests.post(
                "https://api.openai.com/v1/chat/completions",
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
                json={"model": model, "messages": [{"role": "user", "content": prompt}], "max_tokens": 1024},
                timeout=60
            )
            if resp.status_code != 200:
                return None, f"OpenAI API: {resp.status_code} - {resp.text}"
            text = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "")

        elif provider == "Mistral":
            resp = requests.post(
                "https://api.mistral.ai/v1/chat/completions",
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
                json={"model": model, "messages": [{"role": "user", "content": prompt}], "max_tokens": 1024},
                timeout=60
            )
            if resp.status_code != 200:
                return None, f"Mistral API: {resp.status_code} - {resp.text}"
            text = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "")

        elif provider == "Groq":
            resp = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
                json={"model": model, "messages": [{"role": "user", "content": prompt}], "max_tokens": 1024},
                timeout=60
            )
            if resp.status_code != 200:
                return None, f"Groq API: {resp.status_code} - {resp.text}"
            text = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        else:
            return None, "Provider non supporté"

        # Parser le JSON
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group()), None
        return None, "JSON non trouvé dans la réponse"

    except json.JSONDecodeError as e:
        return None, f"JSON invalide: {e}"
    except Exception as e:
        return None, str(e)

# ============================================================================
# INTERFACE STREAMLIT
# ============================================================================
st.title("🏷️ Title Optimizer")
st.markdown("**Analyse les balises title de la SERP et génère des titles SEO optimisées**")

# SIDEBAR - Configuration
st.sidebar.header("🔐 Configuration")

st.sidebar.subheader("DataForSEO")
dfs_login = st.sidebar.text_input("Login", key="to_dfs_login")
dfs_password = st.sidebar.text_input("Password", type="password", key="to_dfs_password")

st.sidebar.subheader("LLM")
llm_provider = st.sidebar.selectbox("Provider", list(LLM_PROVIDERS.keys()), key="to_llm_provider")
llm_key = st.sidebar.text_input(f"API Key {llm_provider.split()[0]}", type="password", key="to_llm_key")
llm_model = st.sidebar.selectbox("Modèle", LLM_PROVIDERS[llm_provider], key="to_llm_model")

st.sidebar.subheader("Paramètres SERP")
country = st.sidebar.selectbox("Pays", list(COUNTRIES.keys()), key="to_country")
serp_depth = st.sidebar.slider("Nombre de titles à scrapper", 10, 100, 10, 10, key="to_serp_depth")

st.sidebar.subheader("Paramètres Title")
max_title_length = st.sidebar.number_input("Longueur max de la title", min_value=40, max_value=70, value=60, key="to_max_title_length")
brand_name = st.sidebar.text_input("Nom de marque (facultatif)", placeholder="MaMarque", key="to_brand_name")

# Vérification des credentials
if not dfs_login or not dfs_password:
    st.warning("👈 Entre tes identifiants DataForSEO dans la sidebar")
    st.stop()

if not llm_key:
    st.warning("👈 Entre ta clé API LLM dans la sidebar")
    st.stop()

# ============================================================================
# MAIN - Mots-clés
# ============================================================================
st.header("📝 Mots-clés à analyser")

input_method = st.radio("Mode d'entrée:", ["Textarea", "CSV"], horizontal=True, key="to_input_method")

keywords = []

if input_method == "Textarea":
    kw_text = st.text_area("Un mot-clé par ligne:", height=150, placeholder="chaussures running\nbaskets homme\nsneakers blanches", key="to_kw_text")
    keywords = [k.strip() for k in kw_text.split("\n") if k.strip()]
else:
    uploaded = st.file_uploader("Fichier CSV", type=["csv"], key="to_csv_upload")
    if uploaded:
        df_up = pd.read_csv(uploaded)
        st.dataframe(df_up.head())
        col_kw = st.selectbox("Colonne contenant les mots-clés:", df_up.columns, key="to_col_kw")
        keywords = df_up[col_kw].dropna().astype(str).tolist()

if keywords:
    st.success(f"📋 {len(keywords)} mot(s)-clé(s) prêts à analyser")

# ============================================================================
# PROMPT PERSONNALISABLE
# ============================================================================
st.header("🤖 Prompt LLM")

with st.expander("✏️ Modifier le prompt (optionnel)", expanded=False):
    st.markdown("""
    **Variables disponibles:**
    - `{keyword}` : Le mot-clé analysé
    - `{brand}` : Le nom de marque (si renseigné)
    - `{max_length}` : La longueur max de la title
    - `{titles_count}` : Le nombre de titles scrappées
    - `{titles_list}` : La liste des titles scrappées
    - `{ngrams_list}` : La liste des n-grams détectés
    """)
    custom_prompt = st.text_area("Prompt:", value=DEFAULT_PROMPT, height=400, key="to_custom_prompt")

if 'custom_prompt' not in dir() or not custom_prompt:
    custom_prompt = DEFAULT_PROMPT

# ============================================================================
# BOUTONS ET TRAITEMENT
# ============================================================================
st.markdown("---")
col_a, col_b = st.columns([2, 1])
with col_a:
    run_btn = st.button("🚀 Lancer l'analyse", type="primary", disabled=not keywords, key="to_run_btn")
with col_b:
    if st.button("🗑️ Reset", key="to_reset_btn"):
        st.session_state.results = []
        st.rerun()

# TRAITEMENT
if run_btn:
    config = COUNTRIES[country]

    st.session_state.results = []
    progress_bar = st.progress(0)
    progress_text = st.empty()

    for i, kw in enumerate(keywords):
        progress_bar.progress((i + 1) / len(keywords))
        progress_text.text(f"🔍 [{i+1}/{len(keywords)}] Scraping SERP pour: {kw}")

        # Appel DataForSEO
        titles, error = get_serp_titles(
            keyword=kw,
            login=dfs_login,
            password=dfs_password,
            config=config,
            depth=serp_depth
        )

        # Construire le résultat
        result = {
            "keyword": kw,
            "titles_scrapped": titles if titles else [],
            "titles_count": len(titles) if titles else 0,
            "ngrams": [],
            "title_proposed": "",
            "title_length": 0,
            "titles_alternatives": [],
            "structure_dominante": "",
            "intention": "",
            "patterns": [],
            "score": 0,
            "justification": "",
            "error": error
        }

        if error:
            st.session_state.results.append(result)
            continue

        # Analyse N-grams
        top_ngrams = analyze_titles(titles)
        result["ngrams"] = top_ngrams

        # Analyse LLM
        progress_text.text(f"🤖 [{i+1}/{len(keywords)}] Analyse IA pour: {kw}")

        llm_result, llm_error = analyze_with_llm(
            titles=titles,
            keyword=kw,
            brand=brand_name,
            max_length=max_title_length,
            top_ngrams=top_ngrams,
            provider=llm_provider,
            model=llm_model,
            api_key=llm_key,
            custom_prompt=custom_prompt
        )

        if llm_error:
            result["error"] = f"LLM: {llm_error}"
        else:
            result["title_proposed"] = llm_result.get("title_proposed", "")
            result["title_length"] = llm_result.get("title_length", len(result["title_proposed"]))
            result["titles_alternatives"] = llm_result.get("titles_alternatives", [])
            result["intention"] = llm_result.get("intention_detectee", "")
            result["patterns"] = llm_result.get("mots_cles_recurrents", [])
            result["structure_dominante"] = llm_result.get("structure_dominante", "")
            result["score"] = llm_result.get("score_confiance", 0)
            result["justification"] = llm_result.get("justification", "")

        st.session_state.results.append(result)

    progress_text.text("✅ Terminé!")
    st.rerun()

# ============================================================================
# RÉSULTATS
# ============================================================================
if st.session_state.results:
    st.header("📊 Résultats")

    results = st.session_state.results

    # Stats
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Mots-clés traités", len(results))
    c2.metric("Titles scrappées", sum([r["titles_count"] for r in results]))
    c3.metric("Titles générées", len([r for r in results if r["title_proposed"]]))
    c4.metric("Erreurs", len([r for r in results if r.get("error")]))

    # Affichage détaillé par mot-clé
    st.subheader("🔍 Détail par mot-clé")

    for r in results:
        with st.expander(f"**{r['keyword']}** → {r['title_proposed'][:50]}..." if r['title_proposed'] else f"**{r['keyword']}** (erreur)", expanded=False):

            if r.get("error"):
                st.error(f"❌ Erreur: {r['error']}")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📥 Titles scrappées")
                if r["titles_scrapped"]:
                    for idx, title in enumerate(r["titles_scrapped"], 1):
                        st.markdown(f"{idx}. {title}")
                else:
                    st.warning("Aucune title récupérée")

                st.markdown("### 🔤 Patterns détectés (N-grams)")
                if r["ngrams"]:
                    for ng, count in r["ngrams"][:10]:
                        st.markdown(f"- **{ng}**: {count} occurrences")
                else:
                    st.info("Aucun pattern détecté")

            with col2:
                st.markdown("### ✅ Title recommandée")
                if r["title_proposed"]:
                    st.success(r["title_proposed"])
                    st.caption(f"📏 {r['title_length']} caractères | 🎯 Score: {r['score']}/100")

                    if r.get("structure_dominante"):
                        st.warning(f"📐 Structure dominante détectée: **{r['structure_dominante']}**")

                    if r["intention"]:
                        st.info(f"🧭 Intention détectée: **{r['intention']}**")

                    if r["patterns"]:
                        st.markdown("### 🔑 Mots-clés récurrents")
                        st.write(", ".join(r["patterns"]))

                    if r["titles_alternatives"]:
                        st.markdown("### 🔄 Alternatives")
                        for alt in r["titles_alternatives"]:
                            st.markdown(f"- {alt}")

                    if r["justification"]:
                        st.markdown("### 💡 Justification")
                        st.caption(r["justification"])
                else:
                    st.warning("Aucune title générée")

    # ============================================================================
    # TABLEAU RÉCAPITULATIF
    # ============================================================================
    st.subheader("📋 Tableau récapitulatif")

    df_data = []
    for r in results:
        df_data.append({
            "Mot-clé": r["keyword"],
            "Titles scrappées": "\n".join(r["titles_scrapped"]) if r["titles_scrapped"] else "",
            "Nb titles": r["titles_count"],
            "Structure dominante": r.get("structure_dominante", ""),
            "Title recommandée": r["title_proposed"],
            "Longueur": r["title_length"],
            "Alternatives": "\n".join(r["titles_alternatives"]) if r["titles_alternatives"] else "",
            "Mots-clés récurrents": " | ".join(r["patterns"]) if r["patterns"] else "",
            "Patterns N-grams": " | ".join([f"{ng} ({c})" for ng, c in r["ngrams"][:5]]) if r["ngrams"] else "",
            "Intention": r["intention"],
            "Score": r["score"],
            "Justification": r["justification"],
            "Erreur": r.get("error", "")
        })

    df = pd.DataFrame(df_data)

    # Affichage du tableau
    display_cols = ["Mot-clé", "Titles scrappées", "Nb titles", "Structure dominante", "Title recommandée", "Longueur", "Score"]
    if any(r.get("error") for r in results):
        display_cols.append("Erreur")

    st.dataframe(df[display_cols], use_container_width=True)

    # ============================================================================
    # EXPORT
    # ============================================================================
    st.header("📥 Export")

    col1, col2 = st.columns(2)

    with col1:
        # CSV
        csv = df.to_csv(index=False, sep=";").encode('utf-8')
        st.download_button(
            "📄 Télécharger CSV",
            csv,
            f"title_optimizer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "text/csv",
            key="to_csv_dl",
        )

    with col2:
        # Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Résultats')
        st.download_button(
            "📊 Télécharger Excel",
            buffer.getvalue(),
            f"title_optimizer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="to_excel_dl",
        )

# ============================================================================
# FOOTER
# ============================================================================
st.markdown("---")
st.caption("Title Optimizer v1.0 | DataForSEO + LLM")
