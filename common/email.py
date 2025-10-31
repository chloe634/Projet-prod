# common/email.py — version diagnostique (Brevo + messages d'erreur explicites)
from __future__ import annotations
import os, json, http.client

BREVO_API_KEY = os.getenv("BREVO_API_KEY")
SENDER_EMAIL  = os.getenv("SENDER_EMAIL", "station.ferment@gmail.com")
SENDER_NAME   = os.getenv("SENDER_NAME", "Symbiose")

class EmailSendError(RuntimeError):
    pass

def _require_env():
    missing = []
    if not BREVO_API_KEY:
        missing.append("BREVO_API_KEY")
    if not SENDER_EMAIL:
        missing.append("SENDER_EMAIL")
    if missing:
        raise EmailSendError(f"Variables d'environnement manquantes: {', '.join(missing)}")

def send_reset_email(to_email: str, reset_url: str) -> dict:
    """
    Envoie l'email de réinitialisation via Brevo.
    Retourne un dict avec {status, provider_msg_id?, response?}.
    Lève EmailSendError avec détail en cas d'erreur.
    """
    _require_env()

    html = f"""
    <p>Bonjour,</p>
    <p>Vous avez demandé à réinitialiser votre mot de passe.</p>
    <p><a href="{reset_url}">Réinitialiser mon mot de passe</a></p>
    <p>Ce lien expire dans 60 minutes. Si vous n’êtes pas à l’origine de cette demande, ignorez ce message.</p>
    """

    payload = {
        "sender": {"name": SENDER_NAME, "email": SENDER_EMAIL},
        "to": [{"email": to_email}],
        "subject": "Réinitialisation de votre mot de passe",
        "htmlContent": html
    }

    try:
        conn = http.client.HTTPSConnection("api.brevo.com", timeout=15)
        headers = {
            "api-key": BREVO_API_KEY,
            "accept": "application/json",
            "content-type": "application/json",
        }
        body = json.dumps(payload)
        conn.request("POST", "/v3/smtp/email", body=body, headers=headers)
        resp = conn.getresponse()
        raw = resp.read().decode("utf-8", errors="replace")
    except Exception as e:
        raise EmailSendError(f"Echec connexion Brevo: {e}") from e

    # Exemple succès: 202 + {"messageId":"<202...@smtp-relay.mailin.fr>"}
    if resp.status not in (200, 201, 202):
        raise EmailSendError(f"Brevo HTTP {resp.status} — réponse: {raw}")

    try:
        data = json.loads(raw) if raw else {}
    except Exception:
        data = {"raw": raw}

    return {
        "status": "sent",
        "provider_msg_id": data.get("messageId"),
        "response": data,
    }
