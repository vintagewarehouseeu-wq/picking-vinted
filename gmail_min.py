"""Client Gmail minimal et autonome (pour le déploiement cloud, sans le gros pipeline).

Fournit exactement ce que build_pick_list utilise : get_gmail_service, _extract_html_body,
_html_to_text. Reprend le token OAuth existant (scopes drive + gmail.readonly).
"""
from __future__ import annotations
import base64
import re
from html import unescape
from pathlib import Path

from bs4 import BeautifulSoup
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/gmail.readonly",
]


def get_gmail_service(token_path: Path = Path("token.json")):
    """Construit le client Gmail à partir du token OAuth existant (refresh auto)."""
    token_path = Path(token_path)
    if not token_path.exists():
        raise RuntimeError(f"Pas de token OAuth à {token_path}.")
    creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            token_path.write_text(creds.to_json(), encoding="utf-8")
        else:
            raise RuntimeError("Token OAuth invalide.")
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def _extract_html_body(payload: dict) -> str:
    """Récupère la partie HTML (ou text/plain) d'un message Gmail."""
    if "body" in payload and payload["body"].get("data"):
        return base64.urlsafe_b64decode(payload["body"]["data"]).decode("utf-8", errors="replace")
    parts = payload.get("parts", [])
    for part in parts:
        if part.get("mimeType") == "text/html" and part.get("body", {}).get("data"):
            return base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", errors="replace")
    for part in parts:
        if part.get("mimeType") == "text/plain" and part.get("body", {}).get("data"):
            return base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", errors="replace")
    for part in parts:
        if part.get("parts"):
            nested = _extract_html_body(part)
            if nested:
                return nested
    return ""


def _html_to_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()
    text = soup.get_text(separator="\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return unescape(text).strip()
