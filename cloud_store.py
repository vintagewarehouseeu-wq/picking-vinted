"""Stockage persistant dans Google Drive (pour la version cloud, système de fichiers éphémère).

Réutilise le token.json (scopes gmail.readonly + drive). En local sans secrets,
is_cloud() renvoie False et l'app garde ses fichiers locaux.

Tout est rangé dans un dossier Drive dédié "VJS_Picking".
"""
from __future__ import annotations
import json, io
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FOLDER = "VJS_Picking"


def _secrets():
    """Renvoie les secrets en dict simple, ou {} si aucun fichier secrets (mode local)."""
    try:
        import streamlit as st
        return dict(st.secrets)   # déclenche le parse ; lève si pas de secrets.toml
    except Exception:
        return {}


def bootstrap_token():
    """Cloud : matérialise token.json depuis les secrets pour que le code existant marche."""
    s = _secrets()
    tok = s.get("google_token") if s else None
    if not tok:
        return
    p = ROOT / "token.json"
    if not p.exists():
        p.write_text(tok if isinstance(tok, str) else json.dumps(dict(tok)), encoding="utf-8")


def is_cloud() -> bool:
    s = _secrets()
    return bool(s and s.get("google_token"))


def _creds():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    info = json.loads((ROOT / "token.json").read_text(encoding="utf-8"))
    c = Credentials.from_authorized_user_info(info)
    if not c.valid:
        c.refresh(Request())
    return c


def _drive():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=_creds(), cache_discovery=False)


def _folder_id(svc) -> str:
    q = ("mimeType='application/vnd.google-apps.folder' and name='%s' and trashed=false" % FOLDER)
    r = svc.files().list(q=q, spaces="drive", fields="files(id)").execute().get("files", [])
    if r:
        return r[0]["id"]
    meta = {"name": FOLDER, "mimeType": "application/vnd.google-apps.folder"}
    return svc.files().create(body=meta, fields="id").execute()["id"]


def _find(svc, fid, name):
    q = "name='%s' and '%s' in parents and trashed=false" % (name, fid)
    r = svc.files().list(q=q, spaces="drive", fields="files(id)").execute().get("files", [])
    return r[0]["id"] if r else None


def read_text(name: str, default: str = "") -> str:
    """Lit un fichier texte du dossier Drive (ou default s'il n'existe pas)."""
    from googleapiclient.http import MediaIoBaseDownload
    svc = _drive()
    fid = _find(svc, _folder_id(svc), name)
    if not fid:
        return default
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, svc.files().get_media(fileId=fid))
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue().decode("utf-8")


def write_text(name: str, text: str) -> None:
    """Écrit (ou remplace) un fichier texte dans le dossier Drive."""
    from googleapiclient.http import MediaIoBaseUpload
    svc = _drive()
    fid_folder = _folder_id(svc)
    fid = _find(svc, fid_folder, name)
    media = MediaIoBaseUpload(io.BytesIO(text.encode("utf-8")),
                              mimetype="application/json", resumable=False)
    if fid:
        svc.files().update(fileId=fid, media_body=media).execute()
    else:
        svc.files().create(body={"name": name, "parents": [fid_folder]},
                           media_body=media, fields="id").execute()


def read_json(name: str, default):
    try:
        txt = read_text(name, "")
        return json.loads(txt) if txt else default
    except Exception:
        return default


def write_json(name: str, obj) -> None:
    write_text(name, json.dumps(obj, ensure_ascii=False))
