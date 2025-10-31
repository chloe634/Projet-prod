# pages/_01_Reset_password.py
from __future__ import annotations
import streamlit as st
from common.auth_reset import verify_token, consume_token_and_set_password

st.set_page_config(page_title="Réinitialisation du mot de passe", page_icon="🔐", layout="centered")

def main():
    st.title("Réinitialisation du mot de passe")

    # Récupération du token via l’URL
    qp = st.query_params
    token = qp.get("token")
    if isinstance(token, list):
        token = token[0]
    token = token or st.text_input("Code reçu par e-mail", type="password", help="Le lien reçu contient ce code automatiquement.")

    if not token:
        st.stop()

    check = verify_token(token)
    if not check:
        st.error("Lien invalide ou expiré. Refaite une demande depuis « Mot de passe oublié ».")
        st.stop()

    with st.form("reset_form"):
        pwd1 = st.text_input("Nouveau mot de passe", type="password")
        pwd2 = st.text_input("Confirmer le mot de passe", type="password")
        ok = st.form_submit_button("Mettre à jour mon mot de passe", type="primary")

    if ok:
        if len(pwd1) < 8:
            st.warning("Le mot de passe doit faire au moins 8 caractères.")
            st.stop()
        if pwd1 != pwd2:
            st.warning("Les deux mots de passe ne correspondent pas.")
            st.stop()
        try:
            consume_token_and_set_password(check["reset_id"], check["user_id"], pwd1)
            st.success("Mot de passe mis à jour ✅")
            st.page_link("pages/_00_Auth.py", label="➡️ Retour à la connexion")
        except Exception:
            st.error("Une erreur est survenue. Réessayez plus tard.")

if __name__ == "__main__":
    main()
