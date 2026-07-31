"""Connecteur pick-list : ventes Vinted → match sur Import_VJS → page de picking.

Deux sources de "titres vendus" :
  - Gmail (--gmail) : lit les mails "Bordereau d'envoi Vinted", extrait le titre
    (après "pour …"), télécharge le PDF du bordereau dans bordereaux/ (prêt pour
    Label Life). Réutilise l'OAuth de vinted_pipeline.gmail_sales.
  - Fichier texte (--sold ventes_du_jour.txt) : 1 titre vendu par ligne.

La box est extraite de la fin de la DESCRIPTION de l'annonce ("… BOX 22").
Sortie : picking_du_jour.html (double-clic, sans serveur).
"""
from __future__ import annotations
import sys, json, argparse, unicodedata, re, base64
from pathlib import Path

for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8")
    except Exception: pass

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))


def _norm(s: str) -> str:
    s = "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def _find_col(cols, *needles):
    for c in cols:
        n = _norm(c)
        if any(_norm(x) in n for x in needles):
            return c
    return None


BOX_RE = re.compile(r"\bbox\s*0*(\d+)\b", re.IGNORECASE)
# Sujet : "Bordereau d'envoi Vinted - à utiliser avant le … pour <TITRE>"
POUR_RE = re.compile(r"\bpour\s+(.+?)\s*$", re.IGNORECASE)


def _extract_box(desc) -> str | None:
    if not desc or str(desc) == "nan":
        return None
    m = list(BOX_RE.finditer(str(desc)))
    return m[-1].group(1) if m else None


CATALOG_CACHE = ROOT / "catalog_cache.json"
INVENTAIRE = ROOT / "inventaire"  # dossier où l'utilisateur dépose ses fichiers box


def find_all_masters() -> list[Path]:
    """UNION : inventaire/ (tes fichiers avec box) + tous les Import_VJS de Downloads.
    Le dédoublonnage par VJS garde la version AVEC box → matching maximal + box
    depuis là où elle existe."""
    files = []
    if INVENTAIRE.exists():
        files += sorted(INVENTAIRE.glob("*.xlsx"))
    dl = Path.home() / "Downloads"
    if dl.exists():
        files += sorted(dl.glob("Import_VJS*.xlsx"))
    return files


def _box_from_row(row, c_box, c_desc):
    """Box = colonne dédiée si présente, sinon extraite de la description."""
    if c_box:
        val = str(row[c_box])
        if val and val != "nan":
            m = re.search(r"\d+", val)
            if m:
                return m.group()
    return _extract_box(row[c_desc]) if c_desc else None


def _rows_from_file(path):
    df = pd.read_excel(path, sheet_name=0)  # 1re feuille (peu importe son nom)
    yield from _rows_from_df(df)


def rows_from_bytes(data: bytes):
    """Parse un .xlsx reçu en mémoire (upload cloud, pas de fichier disque)."""
    import io as _io
    df = pd.read_excel(_io.BytesIO(data), sheet_name=0)
    yield from _rows_from_df(df)


def _rows_from_df(df):
    cols = list(df.columns)
    c_ref = _find_col(cols, "reference", "ref")
    c_nom = _find_col(cols, "nom", "titre")
    c_photo = _find_col(cols, "photos", "photo")
    c_taille = _find_col(cols, "taille")
    c_marque = _find_col(cols, "marque")
    c_desc = _find_col(cols, "description")
    c_box = _find_col(cols, "box", "emplacement", "casier", "bac")  # colonne box éventuelle
    if not c_ref or not c_nom:
        return
    for _, row in df.iterrows():
        nom = str(row[c_nom]).strip()
        if not nom or nom == "nan":
            continue
        photos = str(row[c_photo]) if c_photo else ""
        yield {
            "vjs": str(row[c_ref]).strip(),
            "box": _box_from_row(row, c_box, c_desc),
            "titre": nom,
            "taille": str(row[c_taille]).strip() if c_taille else "",
            "marque": str(row[c_marque]).strip() if c_marque else "",
            "photo": photos.split(",")[0].strip() if photos and photos != "nan" else "",
        }


def _title_map(recs_by_vjs):
    master = {}
    for rec in recs_by_vjs.values():
        master.setdefault(_norm(rec["titre"]), []).append(rec)
    return master


def title_map(by_vjs):
    """Public : {titre normalisé -> [recs]} depuis un dict {vjs -> rec} (cloud)."""
    return _title_map(by_vjs)


def merge_rows(by_vjs: dict, rows) -> dict:
    """Fusionne des lignes dans un catalogue {vjs -> rec}, en préférant la version AVEC box."""
    for rec in rows:
        v = rec["vjs"]
        if v and (v not in by_vjs or (rec["box"] and not by_vjs[v]["box"])):
            by_vjs[v] = rec
    return by_vjs


def build_catalog(paths):
    """Fusionne tous les exports, dédoublonne par VJS (garde la version avec box)."""
    by_vjs = {}
    for p in paths:
        try:
            for rec in _rows_from_file(p):
                v = rec["vjs"]
                if v and (v not in by_vjs or (rec["box"] and not by_vjs[v]["box"])):
                    by_vjs[v] = rec
        except Exception:
            continue
    return by_vjs


def load_catalog():
    """Catalogue complet (tous les Import_VJS fusionnés) avec cache disque."""
    paths = find_all_masters()
    if not paths:
        return None, 0, 0
    newest = max(p.stat().st_mtime for p in paths)
    if CATALOG_CACHE.exists() and CATALOG_CACHE.stat().st_mtime >= newest:
        by_vjs = json.loads(CATALOG_CACHE.read_text(encoding="utf-8"))
    else:
        by_vjs = build_catalog(paths)
        CATALOG_CACHE.write_text(json.dumps(by_vjs, ensure_ascii=False), encoding="utf-8")
    master = _title_map(by_vjs)
    n_box = sum(1 for r in by_vjs.values() if r["box"])
    return master, len(by_vjs), n_box


def load_single(xlsx_path: str):
    by_vjs = {r["vjs"]: r for r in _rows_from_file(xlsx_path) if r["vjs"]}
    n_box = sum(1 for r in by_vjs.values() if r["box"])
    return _title_map(by_vjs), len(by_vjs), n_box


def build(master, sold_titles):
    """Match par titre normalisé ; fallback préfixe (sujets Gmail parfois tronqués)."""
    keys = list(master.keys())
    out, problems = [], []
    for t in sold_titles:
        k = _norm(t)
        hits = master.get(k)
        if not hits:  # fallback préfixe
            cands = [kk for kk in keys if kk.startswith(k) or k.startswith(kk)]
            if len(cands) == 1:
                hits = master[cands[0]]
            elif len(cands) > 1:
                problems.append((f"AMBIGU ({len(cands)})", t))
                hits = master[cands[0]]
        if not hits:
            problems.append(("NON TROUVÉ", t))
            continue
        # préfère la version AVEC box (re-listings : plusieurs VJS pour un titre)
        with_box = [h for h in hits if h["box"]]
        out.append(with_box[0] if with_box else hits[0])
    return out, problems


# ============================================================
# Source Gmail : mails bordereau → titres + téléchargement des PDF
# ============================================================
def _safe_name(title):
    return (re.sub(r"[^\w\-]+", "_", _norm(title))[:60] or "bordereau")


def gmail_bordereaux(days: int = 3):
    try:
        from vinted_pipeline import gmail_sales as G
    except Exception:
        import gmail_min as G  # cloud : client Gmail autonome
    svc = G.get_gmail_service(ROOT / "token.json")  # chemin absolu (cwd indep.)
    q = f"from:vinted.fr subject:bordereau newer_than:{days}d"
    r = svc.users().messages().list(userId="me", q=q, maxResults=300).execute()
    ids = [m["id"] for m in r.get("messages", [])]
    ddir = ROOT / "bordereaux"; ddir.mkdir(exist_ok=True)
    titles, n_new, n_skip = [], 0, 0
    for mid in ids:
        # 1) métadonnées seules (rapide) → titre
        meta = svc.users().messages().get(
            userId="me", id=mid, format="metadata", metadataHeaders=["Subject"]).execute()
        subject = next((h["value"] for h in meta["payload"].get("headers", [])
                        if h["name"].lower() == "subject"), "")
        m = POUR_RE.search(subject)
        if not m:
            continue
        title = m.group(1).strip()
        titles.append(title)
        path = ddir / f"{_safe_name(title)}.pdf"
        if path.exists():           # déjà téléchargé → on ne refait pas l'appel lourd
            n_skip += 1
            continue
        # 2) message complet + téléchargement du PDF (seulement si nouveau)
        full = svc.users().messages().get(userId="me", id=mid, format="full").execute()
        if _download_pdf(svc, mid, full, path):
            n_new += 1
    return titles, n_new, n_skip, ddir


def _download_pdf(svc, mid, msg, path):
    def walk(parts):
        for p in parts:
            body = p.get("body", {})
            if (p.get("filename") or "").lower().endswith(".pdf") and body.get("attachmentId"):
                att = svc.users().messages().attachments().get(
                    userId="me", messageId=mid, id=body["attachmentId"]).execute()
                path.write_bytes(base64.urlsafe_b64decode(att["data"]))
                return True
            if p.get("parts") and walk(p["parts"]):
                return True
        return False
    return walk(msg.get("payload", {}).get("parts", []))


# ============================================================
# Photos : depuis les mails "Vendu" (image CDN Vinted, s'affiche sans ORB)
# ============================================================
def gmail_vendus(days: int = 3):
    """Depuis les mails 'Ton article s'est vendu' : [(titre, photo_Vinted)].
    Titre + photo viennent de la MÊME balise <img alt=titre src=photo> → alignés."""
    try:
        from vinted_pipeline import gmail_sales as G
    except Exception:
        import gmail_min as G  # cloud : client Gmail autonome
    from bs4 import BeautifulSoup
    svc = G.get_gmail_service(ROOT / "token.json")
    r = svc.users().messages().list(
        userId="me", q=f"from:vinted.fr subject:vendu newer_than:{days}d", maxResults=300).execute()
    out = []
    for m in r.get("messages", []):
        msg = svc.users().messages().get(userId="me", id=m["id"], format="full").execute()
        html = G._extract_html_body(msg["payload"])
        soup = BeautifulSoup(html, "html.parser")
        title = photo = None
        for img in soup.find_all("img"):
            src = img.get("src", "")
            alt = (img.get("alt") or "").strip()
            if "images1.vinted.net" in src and len(alt) >= 8:
                title, photo = alt, src
                break
        if not title:  # secours : la ligne juste après "a acheté"
            lines = [l.strip() for l in G._html_to_text(html).split("\n") if l.strip()]
            for i, l in enumerate(lines):
                if l == "a acheté" and i + 1 < len(lines):
                    title = lines[i + 1]
                    break
        if title:
            ts = int(msg.get("internalDate", 0)) // 1000  # date de vente (unix sec)
            out.append({"id": m["id"], "title": title, "photo": photo or "", "ts": ts})
    return out


def match_one(master, title):
    """Un titre → le meilleur article du catalogue (préfère celui AVEC box)."""
    k = _norm(title)
    hits = master.get(k)
    if not hits:
        cands = [kk for kk in master if kk.startswith(k) or k.startswith(kk)]
        hits = master[cands[0]] if cands else None
    if not hits:
        return None
    with_box = [h for h in hits if h["box"]]
    return with_box[0] if with_box else hits[0]


def gmail_picks(days: int = 3, master=None):
    """Pipeline : mails vendu (titre+photo) → box depuis catalogue + PDF bordereaux.

    master : catalogue {titre normalisé -> [recs]} déjà construit (mode cloud, depuis
    Drive). Si None, on lit le catalogue local (fichiers Excel)."""
    if master is None:
        master, n_items, n_box = load_catalog()
    if not master:
        return [], [], 0
    picks, problems = [], []
    for v in gmail_vendus(days):
        rec = match_one(master, v["title"])
        if rec:
            picks.append({**rec, "id": v["id"], "ts": v["ts"],
                          "titre": v["title"], "photo": v["photo"] or rec["photo"]})
        else:
            problems.append(("NON TROUVÉ", v["title"]))
    titles, n_new, n_skip, ddir = gmail_bordereaux(days)  # télécharge les PDF bordereaux
    # rattache le PDF bordereau à chaque pick (fichier nommé d'après le titre)
    for p in picks:
        cand = ddir / f"{_safe_name(p['titre'])}.pdf"
        p["bordereau"] = str(cand) if cand.exists() else ""
    return picks, problems, n_new + n_skip


# ============================================================
# Sortie HTML autonome
# ============================================================
def write_standalone_html(picks) -> Path:
    tpl = (ROOT / "pick_list.html").read_text(encoding="utf-8")
    data = json.dumps(picks, ensure_ascii=False)
    html = re.sub(r"const SAMPLE = \[.*?\];", f"const SAMPLE = {data};", tpl, count=1, flags=re.DOTALL)
    out = ROOT / "picking_du_jour.html"
    out.write_text(html, encoding="utf-8")
    return out


def read_sold(path: Path):
    return [l.strip() for l in path.read_text(encoding="utf-8").splitlines()
            if l.strip() and not l.strip().startswith("#")]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?", help="Import_VJS.xlsx (auto: dernier de Downloads)")
    ap.add_argument("--gmail", action="store_true", help="lit les ventes depuis Gmail (mails bordereau)")
    ap.add_argument("--days", type=int, default=7, help="fenêtre Gmail en jours (défaut 7)")
    ap.add_argument("--sold", help="fichier txt de titres vendus (défaut: ventes_du_jour.txt)")
    ap.add_argument("--out", default=str(ROOT / "pick_list_today.json"))
    args = ap.parse_args()

    if args.gmail:
        try:
            picks, problems, n_pdf = gmail_picks(days=args.days)
        except Exception as e:
            print(f"❌ Gmail: {type(e).__name__} — {str(e)[:160]}"); return
        print(f"Gmail : {len(picks)} pièces · {n_pdf} bordereaux PDF prêts")
    else:
        if args.xlsx:
            if not Path(args.xlsx).exists():
                print("❌ Fichier introuvable :", args.xlsx); return
            master, n_items, n_box = load_single(args.xlsx)
        else:
            master, n_items, n_box = load_catalog()
            if not master:
                print("❌ Aucun Import_VJS.xlsx dans Downloads."); return
        print(f"Catalogue : {n_items} articles · box {n_box}/{n_items}")
        sold_path = Path(args.sold) if args.sold else (ROOT / "ventes_du_jour.txt")
        sold = read_sold(sold_path) if sold_path.exists() else []
        if not sold:
            sold = [v[0]["titre"] for v in list(master.values())[:3]]
            print(f"[DÉMO] {sold_path.name} vide → 3 titres d'exemple.")
        picks, problems = build(master, sold)
    Path(args.out).write_text(json.dumps(picks, ensure_ascii=False, indent=2), encoding="utf-8")
    html_out = write_standalone_html(picks)

    print(f"\n✓ {len(picks)} pièces à récupérer :")
    by_box = {}
    for p in picks:
        by_box.setdefault(p["box"] or "?", []).append(p["vjs"])
    for box in sorted(by_box, key=lambda b: (b == "?", int(b) if b.isdigit() else 0)):
        print(f"   BOX {box} : {', '.join(by_box[box])}")
    if problems:
        print(f"\n⚠️ {len(problems)} à vérifier à la main :")
        for why, t in problems[:15]:
            print(f"   [{why}] {t[:60]}")
    print(f"\n👉 Ouvre : {html_out}")


if __name__ == "__main__":
    main()
