# pages/_01_Reset_password.py
from __future__ import annotations
import streamlit as st
from common.auth_reset import verify_token, consume_token_and_set_password

st.set_page_config(page_title="Réinitialisation du mot de passe", page_icon="🔐", layout="centered")

def main():
    st.title("Réinitialisation du mot de passe")
    qp = st.query_params
    token = qp.get("token", [""])[0] if isinstance(qp.get("token"), list) else qp.get("token", "")

    if not token:
        token = st.text_input("Collez ici le code reçu par e-mail", type="password", help="Le lien dans l’e-mail inclut ce code automatiquement.")
        st.info("Le lien envoyé par e-mail contient le code. Sinon, copiez/collez-le ici.")
        if not token:
            st.stop()

    check = verify_token(token)
    if not check:
        st.error("Lien invalide ou expiré. Refaite une demande depuis « Mot de passe oublié ».")
        st.stop()

    with st.form("reset_form", clear_on_submit=False):
        pwd1 = st.text_input("Nouveau mot de passe", type="password")
        pwd2 = st.text_input("Confirmer le mot de passe", type="password")
        ok = st.form_submit_button("Mettre à jour mon mot de passe")

    if ok:
        if len(pwd1) < 8:
            st.warning("Le mot de passe doit faire au moins 8 caractères.")
            st.stop()
        if pwd1 != pwd2:
            st.warning("Les deux mots de passe ne correspondent pas.")
            st.stop()
        try:
            consume_token_and_set_password(check["reset_id"], check["user_id"], pwd1)
            st.success("Mot de passe mis à jour. Vous pouvez maintenant vous connecter.")
            st.page_link("_00_Auth", label="Aller à la connexion")
        except Exception as e:
            st.error("Une erreur est survenue. Réessayez plus tard.")

if __name__ == "__main__":
    main()
