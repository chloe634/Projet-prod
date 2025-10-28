# common/email.py
from __future__ import annotations
import os, base64, logging
from typing import List, Optional, Tuple

try:
    import streamlit as st
except Exception:
    st = None

logger = logging.getLogger(__name__)

# --------- Lecture de secrets (ENV prioritaire, fallback st.secrets en local) ---------
def _get(key: str, default: Optional[str] = None) -> Optional[str]:
    v = os.environ.get(key)
    if v is None and st is not None:
        try:
            v = st.secrets.get(key) or None
        except Exception:
            v = None
    return v or default

def _get_ns(ns: str, key: str, default: Optional[str] = None) -> Optional[str]:
    env_key = f"{ns.upper()}_{key.upper()}"
    v = os.environ.get(env_key)
    if v is None and st is not None:
        try:
            v = st.secrets.get(ns, {}).get(key)
        except Exception:
            v = None
    return v or default

# --------- Signature HTML (logos inline base64) ---------
def html_signature(body_html: str) -> str:
    import pathlib
    def b64img(relpath: str) -> str:
        p = pathlib.Path("assets/signature") / relpath
        data = p.read_bytes()
        return "data:image/png;base64," + base64.b64encode(data).decode("ascii")
    logo_symbiose = b64img("logo_symbiose.png")
    logo_niko = b64img("NIKO_Logo.png")
    return f"""
    <div style="font-family:Inter,Arial,sans-serif;font-size:14px;line-height:1.5;">
      {body_html}
      <hr style="margin:16px 0;border:none;border-top:1px solid #e5e7eb" />
      <table role="presentation" style="border-collapse:collapse">
        <tr>
          <td style="padding-right:8px">
            <img src="{logo_symbiose}" alt="Symbiose" style="height:28px"/>
          </td>
          <td>
            <img src="{logo_niko}" alt="NIKO" style="height:28px"/>
          </td>
        </tr>
      </table>
    </div>
    """

# --------- Backends ---------
class EmailBackend:
    def send(self, subject: str, html_body: str, recipients: List[str],
             attachments: Optional[List[Tuple[str, bytes]]] = None) -> None:
        raise NotImplementedError

class SendGridBackend(EmailBackend):
    def __init__(self):
        self.api_key = _get("SENDGRID_API_KEY")
        self.sender = _get_ns("email", "sender") or _get("EMAIL_SENDER")
        if not (self.api_key and self.sender):
            raise RuntimeError("SENDGRID_API_KEY et EMAIL_SENDER requis pour SendGrid.")

    def send(self, subject: str, html_body: str, recipients: List[str],
             attachments: Optional[List[Tuple[str, bytes]]] = None) -> None:
        import requests, json
        url = "https://api.sendgrid.com/v3/mail/send"
        data = {
            "personalizations": [{"to": [{"email": r} for r in recipients]}],
            "from": {"email": self.sender},
            "subject": subject,
            "content": [{"type": "text/html", "value": html_body}],
        }
        if attachments:
            data["attachments"] = [{
                "content": base64.b64encode(content).decode("ascii"),
                "filename": filename,
                "type": "application/pdf",
                "disposition": "attachment",
            } for (filename, content) in attachments]

        resp = requests.post(url, headers={
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }, data=json.dumps(data), timeout=20)
        if resp.status_code >= 400:
            logger.error("SendGrid error %s: %s", resp.status_code, resp.text)
            raise RuntimeError(f"SendGrid API error {resp.status_code}: {resp.text}")

class MailgunBackend(EmailBackend):
    def __init__(self):
        self.domain = _get("MAILGUN_DOMAIN")
        self.api_key = _get("MAILGUN_API_KEY")
        self.sender = _get_ns("email", "sender") or _get("EMAIL_SENDER")
        if not (self.domain and self.api_key and self.sender):
            raise RuntimeError("MAILGUN_DOMAIN, MAILGUN_API_KEY et EMAIL_SENDER requis pour Mailgun.")

    def send(self, subject: str, html_body: str, recipients: List[str],
             attachments: Optional[List[Tuple[str, bytes]]] = None) -> None:
        import requests
        url = f"https://api.mailgun.net/v3/{self.domain}/messages"
        files = []
        if attachments:
            for (filename, content) in attachments:
                files.append(("attachment", (filename, content, "application/pdf")))
        data = {
            "from": self.sender,
            "to": recipients,
            "subject": subject,
            "html": html_body,
        }
        resp = requests.post(url, auth=("api", self.api_key), data=data, files=files, timeout=20)
        if resp.status_code >= 400:
            logger.error("Mailgun error %s: %s", resp.status_code, resp.text)
            raise RuntimeError(f"Mailgun API error {resp.status_code}: {resp.text}")

class SMTPBackend(EmailBackend):
    def __init__(self):
        import smtplib
        self.host   = _get_ns("email", "host") or _get("EMAIL_HOST")
        self.port   = int(_get_ns("email", "port") or _get("EMAIL_PORT") or 587)
        self.user   = _get_ns("email", "user") or _get("EMAIL_USER")
        self.passwd = _get_ns("email", "password") or _get("EMAIL_PASSWORD")
        self.sender = _get_ns("email", "sender") or _get("EMAIL_SENDER") or self.user
        self._smtplib = smtplib
        if not (self.host and self.port and self.user and self.passwd and self.sender):
            raise RuntimeError("SMTP incomplet: host, port, user, password, sender requis.")

    def send(self, subject: str, html_body: str, recipients: List[str],
             attachments: Optional[List[Tuple[str, bytes]]] = None) -> None:
        from email.message import EmailMessage
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = self.sender
        msg["To"] = ", ".join(recipients)
        msg.set_content("Version texte indisponible.")
        msg.add_alternative(html_body, subtype="html")
        for (filename, content) in (attachments or []):
            msg.add_attachment(content, maintype="application", subtype="pdf", filename=filename)

        if self.port == 465:
            with self._smtplib.SMTP_SSL(self.host, self.port) as smtp:
                smtp.login(self.user, self.passwd)
                smtp.send_message(msg)
        else:
            with self._smtplib.SMTP(self.host, self.port) as smtp:
                smtp.ehlo()
                try:
                    smtp.starttls()
                except Exception:
                    pass
                smtp.login(self.user, self.passwd)
                smtp.send_message(msg)

# --------- Sélection auto ---------
def get_backend() -> EmailBackend:
    provider = (_get("EMAIL_PROVIDER") or "").lower()
    if provider == "sendgrid" or _get("SENDGRID_API_KEY"):
        return SendGridBackend()
    if provider == "mailgun" or (_get("MAILGUN_DOMAIN") and _get("MAILGUN_API_KEY")):
        return MailgunBackend()
    return SMTPBackend()

# --------- API publique ---------
def send_html_with_pdf(subject: str, html_body: str, recipients: List[str],
                       pdf_bytes: Optional[bytes], pdf_name: str = "fiche_de_ramasse.pdf") -> None:
    attachments = [(pdf_name, pdf_bytes)] if pdf_bytes else None
    backend = get_backend()
    logger.info("Envoi email via backend: %s", backend.__class__.__name__)
    backend.send(subject, html_body, recipients, attachments)
