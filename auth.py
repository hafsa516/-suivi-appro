import streamlit as st
import smtplib
import secrets
import time
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def _email_valide(email: str) -> bool:
    # Option 1 — tout email valide
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", email))

    # Option 2 — uniquement ton entreprise (décommente si besoin)
    # return email.strip().endswith("@tonentreprise.com")


_tokens = {}


def _send_magic_link(email: str, token: str):
    app_url = st.secrets["app"]["url"]
    link = f"{app_url}?token={token}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "🔐 Votre lien de connexion — Suivi Appro"
    msg["From"] = st.secrets["gmail"]["sender"]
    msg["To"] = email

    html = f"""
    <div style="font-family:sans-serif;max-width:480px;margin:auto;
                padding:32px;border-radius:12px;background:#f9f9f9;">
        <h2 style="color:#1a1a2e">📊 Suivi Appro</h2>
        <p>Cliquez sur le bouton ci-dessous pour vous connecter.</p>
        <p>Ce lien est valable <b>10 minutes</b>.</p>
        <a href="{link}"
           style="display:inline-block;margin:24px 0;padding:14px 32px;
                  background:#0078d4;color:white;text-decoration:none;
                  border-radius:8px;font-weight:bold;">
           ✅ Me connecter
        </a>
        <p style="color:#999;font-size:12px">
            Si vous n'avez pas demandé ce lien, ignorez cet email.
        </p>
    </div>
    """
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(
            st.secrets["gmail"]["sender"],
            st.secrets["gmail"]["app_password"]
        )
        server.send_message(msg)


def _create_token(email: str) -> str:
    token = secrets.token_urlsafe(32)
    _tokens[token] = {
        "email": email,
        "expires": time.time() + 600
    }
    return token


def _verify_token(token: str):
    data = _tokens.get(token)
    if not data:
        return None
    if time.time() > data["expires"]:
        del _tokens[token]
        return None
    del _tokens[token]
    return data["email"]


def login_page() -> bool:
    if st.session_state.get("authenticated"):
        return True

    token = st.query_params.get("token")
    if token:
        email = _verify_token(token)
        st.query_params.clear()
        if email:
            st.session_state["authenticated"] = True
            st.session_state["user"] = {"email": email}
            st.rerun()
        else:
            st.error("❌ Lien invalide ou expiré.")
        return False

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("## 📊 Suivi Appro")
        st.markdown("Entrez votre email pour recevoir un lien de connexion.")
        email = st.text_input("Email", placeholder="prenom.nom@entreprise.com")
        if st.button("📧 Recevoir mon lien", type="primary", use_container_width=True):
            if not email:
                st.warning("Entrez votre email.")
            elif not _email_valide(email.strip()):
                st.error("❌ Format d'email invalide.")
            else:
                token = _create_token(email.strip())
                try:
                    _send_magic_link(email.strip(), token)
                    st.success("✅ Lien envoyé ! Vérifiez votre boîte mail.")
                    st.info("⏱ Valable 10 minutes.")
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
    return False


def logout():
    st.session_state.clear()
    st.rerun()