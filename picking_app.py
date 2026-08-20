"""App de picking — web, mobile, sur ton téléphone dans l'entrepôt.

Lancer :  .\.venv\Scripts\python.exe -m streamlit run picking_app.py
Puis ouvre l'URL "Network" affichée (http://192.168.x.x:8601) sur ton téléphone
(même wifi que le PC). Un bouton "Charger", et ta tournée par box s'affiche.
"""
from __future__ import annotations
import sys, re, json, time
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "scripts"))
import build_pick_list as B  # noqa: E402
import cloud_store as CS  # noqa: E402

CS.bootstrap_token()          # cloud : écrit token.json depuis les secrets
CLOUD = CS.is_cloud()         # True sur Streamlit Cloud, False en local

st.set_page_config(page_title="Picking", page_icon="📦", layout="centered")
PP = "#7c3aed"


st.markdown(f"<h2 style='margin:0'>📦 Picking du jour</h2>", unsafe_allow_html=True)

# ---------- 1. INVENTAIRE : où tu donnes tes fichiers avec les box ----------
# Cloud : catalogue {vjs -> rec} rangé dans Drive. Local : fichiers Excel + cache.
INVENTAIRE = ROOT / "inventaire"
CATALOG_NAME = "catalog_byvjs.json"


def get_catalog():
    """(master {titre->[recs]}, n_items, n_box) selon le mode (Drive ou local)."""
    if CLOUD:
        by = CS.read_json(CATALOG_NAME, {})
        n_box = sum(1 for r in by.values() if r.get("box"))
        return B.title_map(by), len(by), n_box
    master, n_items, n_box = B.load_catalog()
    return master or {}, n_items, n_box


if not CLOUD:
    INVENTAIRE.mkdir(exist_ok=True)
existing = list(INVENTAIRE.glob("*.xlsx")) if not CLOUD else []
master, n_items, n_box = get_catalog()

with st.expander(f"📁 Mon inventaire — {n_items} article(s) avec box", expanded=(n_items == 0)):
    st.caption("Dépose ici tes fichiers Excel d'inventaire (ceux qui contiennent les n° de box). "
               "Ils restent enregistrés — tu ne les remets qu'une fois (ou quand tu en ajoutes).")
    ups = st.file_uploader("Fichiers Excel (glisse-les d'un coup)",
                           type=["xlsx"], accept_multiple_files=True)
    if ups:
        with st.spinner("Intégration de l'inventaire…"):
            if CLOUD:
                by = CS.read_json(CATALOG_NAME, {})
                for u in ups:
                    B.merge_rows(by, B.rows_from_bytes(u.getvalue()))
                CS.write_json(CATALOG_NAME, by)
            else:
                for u in ups:
                    (INVENTAIRE / u.name).write_bytes(u.getbuffer())
                (ROOT / "catalog_cache.json").unlink(missing_ok=True)
        st.success(f"✅ {len(ups)} fichier(s) intégré(s).")
        st.rerun()
    if n_items:
        pct = 100 * n_box // max(n_items, 1)
        st.metric("Articles avec une box localisée", f"{n_box} / {n_items}", f"{pct}%")
        if pct < 50:
            st.warning("Couverture box faible : vérifie que tes fichiers contiennent bien le n° de box "
                       "(dans une colonne 'box' ou en fin de description 'BOX 39').")
        if st.button("🗑️ Vider l'inventaire (repartir de zéro)"):
            if CLOUD:
                CS.write_json(CATALOG_NAME, {})
            else:
                for f in existing:
                    f.unlink()
                (ROOT / "catalog_cache.json").unlink(missing_ok=True)
            st.rerun()

st.divider()
col1, col2 = st.columns([2, 1])
days = col2.number_input("Jours", 1, 14, 3, help="Fenêtre Gmail")
if col1.button("🔄 Charger les ventes (Gmail)", use_container_width=True, type="primary"):
    with st.spinner("Lecture de tes ventes Gmail (titres, box, photos, bordereaux)…"):
        picks, problems, n_pdf = B.gmail_picks(days=int(days), master=master)
    st.session_state.picks = picks
    st.session_state.problems = problems
    st.session_state.done = set()
    st.success(f"{len(picks)} pièces · {n_pdf} bordereaux PDF prêts")

# --- état "expédié" PARTAGÉ + PERSISTANT (Drive en cloud, fichier en local) ---
SHIPPED = ROOT / "shipped.json"


def load_shipped() -> set:
    if CLOUD:
        return set(CS.read_json("shipped.json", []))
    try:
        return set(json.loads(SHIPPED.read_text(encoding="utf-8")))
    except Exception:
        return set()


def _save_shipped(s: set):
    if CLOUD:
        CS.write_json("shipped.json", list(s))
    else:
        SHIPPED.write_text(json.dumps(list(s)), encoding="utf-8")


def mark_shipped(pid: str):
    s = load_shipped()
    s.add(pid)
    _save_shipped(s)


def unmark_shipped(pid: str):
    s = load_shipped()
    s.discard(pid)
    _save_shipped(s)


def _photo(col, p):
    """Vignette Vinted cliquable (→ photo en grand). Sinon 👕 (article non catalogué)."""
    url = p.get("photo") or ""
    if "vinted.net" in url:
        col.markdown(
            f'<a href="{url}" target="_blank" title="Voir en grand">'
            f'<img src="{url}" width="64" style="border-radius:8px;object-fit:cover"></a>',
            unsafe_allow_html=True)
    else:
        col.markdown("<div style='font-size:28px;text-align:center'>👕</div>", unsafe_allow_html=True)


def _lot_badge(p) -> str:
    """Badge ⭐ LOT n : les articles d'un même lot partent ensemble avec UN seul bordereau."""
    lot = p.get("lot")
    if not lot:
        return ""
    return (f"<span style='background:#7c3aed;color:#fff;border-radius:6px;"
            f"padding:1px 7px;font-size:12px;font-weight:700'>⭐ LOT {lot}</span> ")


picks = st.session_state.get("picks")
if not picks:
    st.info("👆 Appuie sur **Charger** pour lire les ventes du jour.")
    st.stop()

shipped = load_shipped()
todo = [p for p in picks if p.get("id") not in shipped]
done = [p for p in picks if p.get("id") in shipped]
# tri par URGENCE : la plus vieille vente d'abord (deadline Vinted = 5 jours)
todo.sort(key=lambda p: p.get("ts", 0))
done.sort(key=lambda p: p.get("ts", 0), reverse=True)  # derniers expédiés en haut

now = time.time()
tab_todo, tab_done = st.tabs([f"📦 À expédier ({len(todo)})", f"✅ Expédiés ({len(done)})"])

with tab_todo:
    if not todo:
        st.success("✅ Tout est expédié !")

    # --- REGROUPEMENT PAR BOX : on va à une box, on prend tout ce qu'elle contient ---
    groups = {}
    for p in todo:
        key = str(p["box"]) if p.get("box") else None
        groups.setdefault(key, []).append(p)
    # box la plus urgente d'abord (= celle qui contient la vente la plus vieille) ; box ? à la fin
    order = sorted(groups.keys(),
                   key=lambda k: (k is None, min(x.get("ts", now) for x in groups[k])))

    idx = 0
    for key in order:
        items = groups[key]
        oldest = min(x.get("ts", now) for x in items)
        left = max(0, 5 - int((now - oldest) / 86400))     # Vinted : 5 jours pour expédier
        urg = "🔴" if left <= 1 else ("🟠" if left <= 2 else "🟢")
        titre = f"📦 BOX {key} — {len(items)} article(s)" if key else f"❓ Box inconnue — {len(items)} article(s)"
        st.markdown(f"#### {urg} {titre} · J-{left}")
        for p in items:
            pid = p.get("id") or p.get("vjs") or f"row{idx}"
            c1, c2, c3 = st.columns([1.1, 4, 1.5])
            _photo(c1, p)
            c2.markdown(
                f"{_lot_badge(p)}{p.get('titre','')[:60]}  \n"
                f"<span style='color:#888;font-size:12px'>{p.get('vjs','')} · {p.get('taille','')} · {p.get('marque','')}</span>",
                unsafe_allow_html=True)
            bpath = p.get("bordereau", "")
            if bpath and Path(bpath).exists():
                c3.download_button("🧾 Bordereau", data=Path(bpath).read_bytes(),
                                   file_name=Path(bpath).name, mime="application/pdf",
                                   key=f"bord_{pid}_{idx}", use_container_width=True)
            if c3.button("✅ Expédié", key=f"ship_{pid}_{idx}", use_container_width=True):
                mark_shipped(pid)
                st.rerun()
            idx += 1
        st.divider()

with tab_done:
    if not done:
        st.info("Rien d'expédié pour l'instant.")
    for i, p in enumerate(done):
        pid = p.get("id") or p.get("vjs") or f"row{i}"
        box = f"**BOX {p['box']}**" if p.get("box") else "box ?"
        c1, c2, c3 = st.columns([1.1, 4, 1.5])
        _photo(c1, p)
        c2.markdown(
            f"✅ {box} {_lot_badge(p)}  \n{p.get('titre','')[:60]}  \n"
            f"<span style='color:#888;font-size:12px'>{p.get('vjs','')} · {p.get('taille','')} · {p.get('marque','')}</span>",
            unsafe_allow_html=True)
        if c3.button("↩️ Annuler", key=f"unship_{pid}_{i}", use_container_width=True):
            unmark_shipped(pid)
            st.rerun()
        st.divider()

problems = st.session_state.get("problems", [])
if problems:
    with st.expander(f"⚠️ {len(problems)} ventes non localisées (à trouver à la main)"):
        for why, t in problems:
            st.write(f"[{why}] {t}")
