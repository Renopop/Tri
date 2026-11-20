#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tri_llm.py — Classification d'événements moteur par deux LLM, puis comparaison.

Fichier UNIQUE, sans interface graphique. Pour chaque ligne du classeur Excel :

  1. lit les colonnes D « Description », E « Immediate Action », F « Long term » ;
  2. demande à mistral-small-4 (passerelle LiteLLM, même mécanisme que CompareDB)
     dans quelle catégorie classer l'événement :
         - engine shutdown commanded
         - engine power modulated
         - engine forced to idle
         - non classé  (si aucune ne convient)
     -> résultat en colonne H « LLM Mistral » ;
  3. même question à gemma-4 -> colonne I « LLM Gemma » ;
  4. colonne K : OK (blanc sur vert) si H et I sont identiques, KO (blanc sur
     rouge) sinon ; ERREUR (blanc sur orange) si l'un des deux appels a échoué.

Lancement :
    python tri_llm.py                          -> demande le nom du fichier
    python tri_llm.py mon_fichier.xlsx         -> fichier donné
    python tri_llm.py mon_fichier.xlsx --limit 20   (test sur 20 lignes)
    python tri_llm.py mon_fichier.xlsx --force      (recommence tout)

Le résultat est écrit dans <nom>_LLM.xlsx à côté de l'original (jamais
l'original). Si ce fichier existe déjà, le script REPREND : les lignes déjà
classées sont conservées, seules les lignes manquantes ou en erreur sont
renvoyées aux modèles (sauf --force). Sauvegarde toutes les 10 lignes.

CLÉS API — fichier Tri_keys.env à côté du script (modèle : Tri_keys.env.example)
    TRI_MISTRAL_KEY=...     clé du modèle mistral-small-4
    TRI_GEMMA_KEY=...       clé du modèle gemma-4
La passerelle délivre UNE clé PAR modèle. Une variable déjà définie dans
l'environnement n'est jamais écrasée par le fichier. TRI_KEYS_FILE force le
chemin du fichier. Les clés CompareDB/RIVAGE pour mistral-small
(COMPAREDB_LLM_KEY, RIVAGE_SYNTH_KEY) sont acceptées en repli.

Options (variables d'environnement, ou dans Tri_keys.env) :
    TRI_API_BASE        passerelle (défaut https://litellm.dassault-aviation.pro/v1)
    TRI_MISTRAL_MODEL   défaut mistral-small-4
    TRI_GEMMA_MODEL     défaut gemma-4
    TRI_MISTRAL_RPM     requêtes/minute autorisées (défaut 20)
    TRI_GEMMA_RPM       requêtes/minute autorisées (défaut 20, limite LiteMaaS)
    TRI_SHEET           nom de la feuille (défaut : la première)
    DALLEM_CA_BUNDLE    bundle CA interne (.pem) si CERTIFICATE_VERIFY_FAILED
    DALLEM_VERIFY_SSL   false pour désactiver la vérification TLS (repli)
    DALLEM_TIMEOUT      timeout HTTP en secondes (défaut 120)

Dépendances : pip install httpx openpyxl
"""

import os
import re
import sys
import time
import random
import logging
import argparse
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from typing import Optional, Tuple, List, Dict

try:
    import httpx
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
except ImportError as e:
    print(f"Dépendance manquante ({e.name}). Installez : pip install httpx openpyxl")
    sys.exit(1)

# =============================================================================
#  PARAMÈTRES
# =============================================================================

DEFAULT_FILE = "ACES_FT_classified_simplified.xlsx"
OUT_SUFFIX = "_LLM"

CATEGORIES = [
    "engine shutdown commanded",
    "engine power modulated",
    "engine forced to idle",
]
UNCLASSIFIED = "non classé"

COL_DESC, COL_IMM, COL_LONG = "D", "E", "F"
COL_MISTRAL, COL_GEMMA, COL_CMP = "H", "I", "K"
HEADERS = {COL_MISTRAL: "LLM Mistral", COL_GEMMA: "LLM Gemma", COL_CMP: "Comparaison"}

MAX_CHARS_PER_FIELD = 6000      # sécurité contexte (max observé ≈ 3000)
SAVE_EVERY = 10                 # lignes entre deux sauvegardes
RETRIABLE_STATUS = (429, 502, 503, 504)
MAX_ATTEMPTS = 6
RETRY_CAP_S = 60.0

FILL_OK = PatternFill("solid", fgColor="00B050")      # vert
FILL_KO = PatternFill("solid", fgColor="FF0000")      # rouge
FILL_ERR = PatternFill("solid", fgColor="ED7D31")     # orange
FONT_WHITE = Font(color="FFFFFF", bold=True)
FONT_BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center")

SYSTEM_PROMPT = (
    "You are an aeronautical engine event analyst. You classify in-service "
    "engine event reports into exactly ONE of these categories:\n"
    "1. engine shutdown commanded — the crew deliberately shut the engine down "
    "(in-flight shutdown, IFSD, precautionary shutdown, engine secured).\n"
    "2. engine power modulated — the crew reduced or adjusted engine power/thrust "
    "without shutting it down and without forcing it to idle (power reduction, "
    "throttle back, thrust limited, derated).\n"
    "3. engine forced to idle — the engine was brought to or held at idle "
    "(forced idle, retarded to idle, kept at idle for the rest of the flight).\n"
    "4. non classé — none of the above applies (no crew action on engine power, "
    "aborted take-off, ground event, information insufficient).\n"
    "Answer with the category text ONLY, exactly as written above, in lowercase, "
    "without explanation."
)

logger = logging.getLogger("tri_llm")


# =============================================================================
#  CONFIGURATION : fichier de clés + variables d'environnement
# =============================================================================

def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _parse_env_line(line: str):
    """KEY=VALUE (sémantique dotenv : `export` toléré, ` #` = commentaire de
    fin de ligne sur valeur non quotée, guillemets retirés)."""
    line = line.strip()
    if not line or line.startswith("#") or "=" not in line:
        return None, None
    if line.startswith("export "):
        line = line[7:].strip()
    key, _, value = line.partition("=")
    key = key.strip()
    if not key or any(c.isspace() for c in key):
        return None, None
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
        value = value[1:-1]
    else:
        hash_at = value.find(" #")
        if hash_at >= 0:
            value = value[:hash_at].rstrip()
    return key, value


def load_keys_file() -> Optional[Path]:
    """Charge Tri_keys.env dans l'environnement SANS écraser l'existant.
    Cherché : TRI_KEYS_FILE, à côté du script, dossier parent, répertoire
    courant. Retourne le chemin chargé ou None."""
    candidates = []
    forced = os.getenv("TRI_KEYS_FILE")
    if forced:
        candidates.append(Path(forced))
    base = _app_dir()
    candidates += [base / "Tri_keys.env", base.parent / "Tri_keys.env",
                   Path.cwd() / "Tri_keys.env"]
    for path in candidates:
        if not path.is_file():
            continue
        try:
            text = path.read_bytes()
        except OSError:
            continue
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                content = text.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        for line in content.splitlines():
            k, v = _parse_env_line(line)
            if k:
                os.environ.setdefault(k, v)
        return path
    return None


def _env_first(*names: str, default: Optional[str] = None) -> Optional[str]:
    for n in names:
        v = os.getenv(n)
        if v is not None and v.strip():
            return v.strip()
    return default


def _env_float(name: str, default: float) -> float:
    try:
        v = os.getenv(name)
        return float(v) if v and v.strip() else default
    except ValueError:
        logger.warning(f"{name}={os.getenv(name)!r} non numérique — défaut {default}")
        return default


def resolve_verify():
    """TLS : DALLEM_CA_BUNDLE (bundle .pem) > DALLEM_VERIFY_SSL=false > système."""
    ca = os.getenv("DALLEM_CA_BUNDLE")
    if ca:
        if os.path.isfile(ca):
            return ca
        logger.warning(f"DALLEM_CA_BUNDLE={ca!r} introuvable — ignoré.")
    v = os.getenv("DALLEM_VERIFY_SSL")
    if v is None:
        return True
    return v.strip().lower() in ("1", "true", "yes", "oui", "on")


def mask(key: Optional[str]) -> str:
    if not key:
        return "<absente>"
    return "***" if len(key) <= 8 else key[:4] + "…" + key[-3:]


# =============================================================================
#  CLIENT LLM (chat/completions) avec limitation de débit et retries
# =============================================================================

class RateLimiter:
    """Espace les appels d'un même modèle : au plus `rpm` par minute."""

    def __init__(self, rpm: float):
        self.interval = 60.0 / rpm if rpm and rpm > 0 else 0.0
        self._lock = threading.Lock()
        self._next = 0.0

    def wait(self):
        if self.interval <= 0:
            return
        with self._lock:
            now = time.monotonic()
            delay = self._next - now
            if delay > 0:
                time.sleep(delay)
                now = time.monotonic()
            self._next = now + self.interval


def _retry_after_s(resp) -> Optional[float]:
    try:
        raw = resp.headers.get("Retry-After")
        return max(0.0, float(raw)) if raw else None
    except Exception:
        return None


def normalize_label(raw: str) -> str:
    """Ramène la réponse brute du modèle à une des catégories exactes."""
    s = (raw or "").strip().lower()
    s = s.replace("modullated", "modulated").replace("’", "'")
    s = re.sub(r"[^a-zàâçéèêëîïôûùüÿœ\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    for cat in CATEGORIES:
        if cat in s:
            return cat
    # Variantes fréquentes
    if "shutdown" in s or "shut down" in s or "ifsd" in s:
        return CATEGORIES[0]
    if "idle" in s:
        return CATEGORIES[2]
    if "modulat" in s or "power reduc" in s or "thrust reduc" in s or "throttle" in s:
        return CATEGORIES[1]
    return UNCLASSIFIED


class ChatModel:
    """Un modèle de la passerelle, avec SA clé et SON débit."""

    def __init__(self, label: str, model: str, api_key: str, api_base: str,
                 http: httpx.Client, rpm: float, timeout: float):
        self.label = label
        self.model = model
        self.api_key = api_key
        self.url = api_base.rstrip("/") + "/chat/completions"
        self.http = http
        self.limiter = RateLimiter(rpm)
        self.timeout = timeout
        self.calls = 0
        self.errors = 0

    def classify(self, description: str, immediate: str, long_term: str) -> Tuple[str, str]:
        """Retourne (catégorie normalisée, réponse brute). En cas d'échec
        définitif : ('ERREUR: …', '')."""
        user = (
            f"Description:\n{description or '(empty)'}\n\n"
            f"Immediate Action:\n{immediate or '(empty)'}\n\n"
            f"Long term:\n{long_term or '(empty)'}\n\n"
            "Category:"
        )
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user},
            ],
            "temperature": 0.0,
            "max_tokens": 24,
        }
        headers = {"Authorization": f"Bearer {self.api_key}",
                   "Content-Type": "application/json"}
        last_err = ""
        for attempt in range(1, MAX_ATTEMPTS + 1):
            self.limiter.wait()
            try:
                self.calls += 1
                resp = self.http.post(self.url, headers=headers, json=payload,
                                      timeout=self.timeout)
            except Exception as e:
                msg = str(e).lower()
                if "certificate" in msg or "ssl" in msg:
                    self.errors += 1
                    return ("ERREUR: certificat TLS non validé — renseignez "
                            "DALLEM_CA_BUNDLE ou DALLEM_VERIFY_SSL=false", "")
                last_err = f"{type(e).__name__}: {e}"
                delay = min(RETRY_CAP_S, 2.0 * (2 ** (attempt - 1))) * random.uniform(0.7, 1.3)
                logger.warning(f"[{self.label}] réseau ({last_err}) — tentative "
                               f"{attempt}/{MAX_ATTEMPTS}, reprise dans {delay:.0f}s")
                time.sleep(delay)
                continue

            if resp.status_code in RETRIABLE_STATUS:
                delay = min(RETRY_CAP_S, 3.0 * (2 ** (attempt - 1))) * random.uniform(0.7, 1.3)
                ra = _retry_after_s(resp)
                if ra:
                    delay = min(RETRY_CAP_S, max(delay, ra))
                last_err = f"HTTP {resp.status_code}"
                logger.warning(f"[{self.label}] {last_err} (passerelle saturée / quota) — "
                               f"tentative {attempt}/{MAX_ATTEMPTS}, reprise dans {delay:.0f}s")
                time.sleep(delay)
                continue
            if resp.status_code >= 400:
                self.errors += 1
                body = " ".join((resp.text or "").split())[:200]
                hint = " — clé invalide ?" if resp.status_code == 401 else ""
                hint = " — modèle inconnu de la passerelle ?" if resp.status_code == 404 else hint
                return (f"ERREUR: HTTP {resp.status_code}{hint} {body}".strip(), "")
            try:
                data = resp.json()
                choice = data["choices"][0]
                raw = ((choice.get("message") or {}).get("content") or "").strip()
            except Exception as e:
                self.errors += 1
                return (f"ERREUR: réponse illisible ({e})", "")
            if not raw:
                self.errors += 1
                return ("ERREUR: réponse vide du modèle", "")
            return normalize_label(raw), raw
        self.errors += 1
        return (f"ERREUR: {last_err} après {MAX_ATTEMPTS} tentatives", "")


# =============================================================================
#  EXCEL
# =============================================================================

def cell_text(ws, col: str, row: int) -> str:
    v = ws[f"{col}{row}"].value
    if v is None:
        return ""
    s = str(v).strip()
    return s[:MAX_CHARS_PER_FIELD]


def is_error(value) -> bool:
    return isinstance(value, str) and value.startswith("ERREUR")


def write_comparison(ws, row: int, h: str, i: str):
    cell = ws[f"{COL_CMP}{row}"]
    cell.font = FONT_WHITE
    cell.alignment = CENTER
    if is_error(h) or is_error(i) or not h or not i:
        cell.value = "ERREUR"
        cell.fill = FILL_ERR
    elif h == i:
        cell.value = "OK"
        cell.fill = FILL_OK
    else:
        cell.value = "KO"
        cell.fill = FILL_KO


def ensure_headers(ws):
    for col, title in HEADERS.items():
        c = ws[f"{col}1"]
        c.value = title
        c.font = FONT_BOLD
        c.alignment = CENTER
    for col, width in ((COL_MISTRAL, 28), (COL_GEMMA, 28), (COL_CMP, 14)):
        if (ws.column_dimensions[col].width or 0) < width:
            ws.column_dimensions[col].width = width


def save_workbook(wb, out_path: Path):
    """Sauvegarde atomique (fichier temporaire puis remplacement) : un arrêt
    brutal pendant l'écriture ne corrompt pas le résultat précédent."""
    tmp = out_path.with_suffix(out_path.suffix + ".tmp")
    wb.save(tmp)
    for attempt in range(5):
        try:
            os.replace(tmp, out_path)
            return
        except PermissionError:
            # Fichier ouvert dans Excel : on réessaie un peu puis on prévient.
            time.sleep(0.5 * (attempt + 1))
    logger.error(f"Impossible d'écrire {out_path} (ouvert dans Excel ?). "
                 f"Résultat conservé dans {tmp}")


# =============================================================================
#  PROGRAMME PRINCIPAL
# =============================================================================

def setup_logging(log_path: Path):
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s"))
    logger.handlers[:] = [ch, fh]


def ask_file(default: str) -> str:
    try:
        answer = input(f"Fichier Excel à analyser [{default}] : ").strip()
    except EOFError:
        answer = ""
    return answer or default


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Classe chaque ligne d'un Excel avec mistral-small-4 (col. H) et "
                    "gemma-4 (col. I), puis compare (col. K).")
    ap.add_argument("fichier", nargs="?", help=f"classeur Excel (défaut : demandé, {DEFAULT_FILE})")
    ap.add_argument("--out", help="fichier de sortie (défaut : <nom>_LLM.xlsx à côté de l'original)")
    ap.add_argument("--force", action="store_true", help="recommence toutes les lignes, même déjà classées")
    ap.add_argument("--limit", type=int, default=0, help="ne traite que les N premières lignes utiles (test)")
    ap.add_argument("--sheet", default=os.getenv("TRI_SHEET"), help="nom de la feuille (défaut : la première)")
    args = ap.parse_args()

    keys_file = load_keys_file()

    src = Path(args.fichier or ask_file(DEFAULT_FILE)).expanduser()
    if not src.is_file():
        print(f"Fichier introuvable : {src}")
        return 1
    out = Path(args.out).expanduser() if args.out else src.with_name(f"{src.stem}{OUT_SUFFIX}{src.suffix}")
    if out.resolve() == src.resolve():
        print("Le fichier de sortie ne peut pas être le fichier d'origine (il est préservé).")
        return 1

    setup_logging(out.with_suffix(".log"))
    logger.info("=== tri_llm — classification par deux LLM ===")
    logger.info(f"Source  : {src}")
    logger.info(f"Sortie  : {out}")
    keys_desc = str(keys_file) if keys_file else "aucun fichier Tri_keys.env (variables d'environnement seules)"
    logger.info(f"Clés    : {keys_desc}")

    # --- Configuration des modèles -------------------------------------------
    api_base = _env_first("TRI_API_BASE", "COMPAREDB_API_BASE", "RIVAGE_API_BASE",
                          default="https://litellm.dassault-aviation.pro/v1")
    mistral_model = _env_first("TRI_MISTRAL_MODEL", default="mistral-small-4")
    gemma_model = _env_first("TRI_GEMMA_MODEL", default="gemma-4")
    mistral_key = _env_first("TRI_MISTRAL_KEY", "COMPAREDB_LLM_KEY", "RIVAGE_SYNTH_KEY")
    gemma_key = _env_first("TRI_GEMMA_KEY")
    mistral_rpm = _env_float("TRI_MISTRAL_RPM", 20.0)
    gemma_rpm = _env_float("TRI_GEMMA_RPM", 20.0)
    timeout = _env_float("DALLEM_TIMEOUT", 120.0)
    verify = resolve_verify()

    logger.info(f"Passerelle : {api_base} | TLS={verify if isinstance(verify, bool) else 'bundle ' + verify}")
    logger.info(f"Mistral : {mistral_model} | clé {mask(mistral_key)} | {mistral_rpm:.0f} req/min")
    logger.info(f"Gemma   : {gemma_model} | clé {mask(gemma_key)} | {gemma_rpm:.0f} req/min")
    missing = [n for n, k in (("TRI_MISTRAL_KEY", mistral_key), ("TRI_GEMMA_KEY", gemma_key)) if not k]
    if missing:
        logger.error(f"Clé(s) manquante(s) : {', '.join(missing)}. Renseignez Tri_keys.env "
                     f"(modèle : Tri_keys.env.example) à côté de {Path(__file__).name}.")
        return 1

    http = httpx.Client(timeout=timeout, verify=verify)
    mistral = ChatModel("Mistral", mistral_model, mistral_key, api_base, http, mistral_rpm, timeout)
    gemma = ChatModel("Gemma", gemma_model, gemma_key, api_base, http, gemma_rpm, timeout)

    # --- Classeur : reprise depuis la sortie si elle existe ------------------
    resume = out.is_file() and not args.force
    wb = openpyxl.load_workbook(out if resume else src)
    if resume:
        logger.info("Reprise : la sortie existe déjà, les lignes classées sont conservées (--force pour tout refaire).")
    if args.sheet:
        if args.sheet not in wb.sheetnames:
            logger.error(f"Feuille {args.sheet!r} absente. Feuilles : {wb.sheetnames}")
            return 1
        ws = wb[args.sheet]
    else:
        ws = wb.worksheets[0]
    logger.info(f"Feuille : {ws.title}")
    ensure_headers(ws)

    # Lignes utiles : au moins une des colonnes D/E/F renseignée.
    rows: List[int] = [r for r in range(2, ws.max_row + 1)
                       if any(cell_text(ws, c, r) for c in (COL_DESC, COL_IMM, COL_LONG))]
    if args.limit > 0:
        rows = rows[:args.limit]
    todo = [r for r in rows
            if args.force
            or not ws[f"{COL_MISTRAL}{r}"].value or is_error(ws[f"{COL_MISTRAL}{r}"].value)
            or not ws[f"{COL_GEMMA}{r}"].value or is_error(ws[f"{COL_GEMMA}{r}"].value)]
    logger.info(f"Lignes utiles : {len(rows)} | à traiter : {len(todo)} | déjà faites : {len(rows) - len(todo)}")
    if todo:
        per_row = max(60.0 / mistral_rpm, 60.0 / gemma_rpm)
        logger.info(f"Durée estimée : ~{len(todo) * per_row / 60:.0f} min (limite de débit des clés)")

    # --- Traitement : les deux modèles en parallèle sur chaque ligne ----------
    t0 = time.time()
    done_since_save = 0
    counts: Dict[str, Dict[str, int]] = {"Mistral": {}, "Gemma": {}}
    cmp_counts = {"OK": 0, "KO": 0, "ERREUR": 0}
    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            for n, r in enumerate(todo, 1):
                d, e, f = cell_text(ws, COL_DESC, r), cell_text(ws, COL_IMM, r), cell_text(ws, COL_LONG, r)
                need_m = args.force or not ws[f"{COL_MISTRAL}{r}"].value or is_error(ws[f"{COL_MISTRAL}{r}"].value)
                need_g = args.force or not ws[f"{COL_GEMMA}{r}"].value or is_error(ws[f"{COL_GEMMA}{r}"].value)
                fut_m = pool.submit(mistral.classify, d, e, f) if need_m else None
                fut_g = pool.submit(gemma.classify, d, e, f) if need_g else None
                if fut_m is not None:
                    label_m, raw_m = fut_m.result()
                    ws[f"{COL_MISTRAL}{r}"].value = label_m
                    if raw_m and raw_m.strip().lower() != label_m:
                        logger.debug(f"L{r} Mistral brut : {raw_m!r} -> {label_m}")
                else:
                    label_m = ws[f"{COL_MISTRAL}{r}"].value
                if fut_g is not None:
                    label_g, raw_g = fut_g.result()
                    ws[f"{COL_GEMMA}{r}"].value = label_g
                    if raw_g and raw_g.strip().lower() != label_g:
                        logger.debug(f"L{r} Gemma brut : {raw_g!r} -> {label_g}")
                else:
                    label_g = ws[f"{COL_GEMMA}{r}"].value

                write_comparison(ws, r, label_m, label_g)
                verdict = ws[f"{COL_CMP}{r}"].value
                ident = ws[f"A{r}"].value
                logger.info(f"[{n}/{len(todo)}] L{r} id={ident} | Mistral: {label_m} | Gemma: {label_g} | {verdict}")

                done_since_save += 1
                if done_since_save >= SAVE_EVERY:
                    save_workbook(wb, out)
                    done_since_save = 0
    except KeyboardInterrupt:
        logger.warning("Interruption clavier — sauvegarde de l'état courant, relancez pour reprendre.")
    finally:
        # Comparaison recalculée sur TOUTES les lignes utiles (reprise comprise).
        for r in rows:
            h, i = ws[f"{COL_MISTRAL}{r}"].value, ws[f"{COL_GEMMA}{r}"].value
            if h or i:
                write_comparison(ws, r, h, i)
                cmp_counts[ws[f"{COL_CMP}{r}"].value] += 1
                for name, v in (("Mistral", h), ("Gemma", i)):
                    key = "ERREUR" if is_error(v) else (v or "(vide)")
                    counts[name][key] = counts[name].get(key, 0) + 1
        save_workbook(wb, out)
        http.close()

    # --- Bilan -----------------------------------------------------------------
    logger.info("=== Bilan ===")
    for name in ("Mistral", "Gemma"):
        detail = ", ".join(f"{k}: {v}" for k, v in sorted(counts[name].items(), key=lambda kv: -kv[1]))
        logger.info(f"{name:8s}: {detail}")
    logger.info(f"Comparaison : OK={cmp_counts['OK']} | KO={cmp_counts['KO']} | ERREUR={cmp_counts['ERREUR']}")
    logger.info(f"Appels : Mistral {mistral.calls} ({mistral.errors} erreur(s)) | "
                f"Gemma {gemma.calls} ({gemma.errors} erreur(s)) | {time.time() - t0:.0f}s")
    if cmp_counts["ERREUR"]:
        logger.warning("Des lignes sont en ERREUR : relancez le script, elles seules seront renvoyées aux modèles.")
    logger.info(f"Résultat : {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
