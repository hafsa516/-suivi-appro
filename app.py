import streamlit as st
import tempfile, os
from auth import login_page, logout
from processing import traiter_excel

st.set_page_config(page_title="Suivi Appro", page_icon="📊", layout="centered")

if not login_page():
    st.stop()

user = st.session_state["user"]

col1, col2 = st.columns([4, 1])
with col1:
    st.markdown("## 📊 Suivi Appro")
    st.caption(f"👤 {user['email']}")
with col2:
    if st.button("🚪 Déconnexion"):
        logout()

st.divider()

st.subheader("1️⃣ Charger votre fichier Excel")
uploaded = st.file_uploader("Glissez votre fichier", type=["xlsx", "xlsm"])

if uploaded:
    st.success(f"✅ **{uploaded.name}** chargé")
    st.divider()
    st.subheader("2️⃣ Lancer le traitement")

    if st.button("▶ Exécuter", type="primary", use_container_width=True):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        logs = []
        bar  = st.progress(0)
        zone = st.empty()

        def log_fn(msg):
            logs.append(msg)
            zone.code("\n".join(logs), language=None)

        def progress_fn(val):
            bar.progress(val / 100)

        try:
            traiter_excel(tmp_path, log_fn=log_fn, progress_fn=progress_fn)
            st.divider()
            st.subheader("3️⃣ Télécharger le résultat")
            with open(tmp_path, "rb") as f:
                st.download_button(
                    "⬇️ Télécharger",
                    data=f,
                    file_name=f"traite_{uploaded.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
        finally:
            os.unlink(tmp_path)