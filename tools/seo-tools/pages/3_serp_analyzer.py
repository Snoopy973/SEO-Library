import streamlit as st
import requests

st.title("🎯 Analyseur SERP")

# ============================================================================
# FONCTION - Test connexion DataForSEO
# ============================================================================

def test_dataforseo_connection(login, password):
    """Teste si la connexion à DataForSEO fonctionne"""
    try:
        url = "https://api.dataforseo.com/v3/serp/google/organic/live/advanced"
        payload = [{
            "keyword": "test",
            "location_code": 2250,
            "language_code": "fr",
            "depth": 1
        }]
        
        response = requests.post(
            url,
            auth=(login, password),
            json=payload,
            headers={'content-type': 'application/json'},
            timeout=10
        )
        
        # Status 20000 = succès
        if response.status_code == 200:
            data = response.json()
            if data.get('tasks') and data['tasks'][0].get('status_code') == 20000:
                return True, "✅ Connecté"
        
        return False, "❌ Identifiants incorrects"
        
    except Exception as e:
        return False, f"❌ Erreur: {str(e)}"

# ============================================================================
# SIDEBAR - Configuration
# ============================================================================

with st.sidebar:
    st.header("⚙️ Configuration")
    
    st.subheader("🔑 DataForSEO")
    dfs_login = st.text_input("Email", placeholder="votre_email@example.com")
    dfs_password = st.text_input("Mot de passe", type="password")
    
    # Test connexion
    if dfs_login and dfs_password:
        if st.button("🔗 Tester la connexion", use_container_width=True):
            is_connected, message = test_dataforseo_connection(dfs_login, dfs_password)
            if is_connected:
                st.success(message)
                st.session_state.dfs_connected = True
            else:
                st.error(message)
                st.session_state.dfs_connected = False
    
    # Afficher le status
    if "dfs_connected" in st.session_state:
        if st.session_state.dfs_connected:
            st.success("✅ DataForSEO - Connecté")
        else:
            st.error("❌ DataForSEO - Déconnecté")
    
    st.divider()
    
    st.subheader("🌍 Paramètres")
    location_code = st.number_input("Location Code (2250=France)", value=2250)
    language_code = st.selectbox("Langue", ["fr", "en", "de", "es"])

# ============================================================================
# MAIN
# ============================================================================

st.markdown("---")
st.info("Application en construction... On construit ça ensemble ! ✨")


