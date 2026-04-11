# -*- coding: utf-8 -*-
"""
cache.py — Gestion du cache de données PULSE.

Architecture :
    CACHE[(section, flux)]  → { dates, reel, prev_headers, prev_vals }
    STRUCT[section]         → { flux_name: { col_start, prev_headers, … } }
    TOKENS[section]         → [ (flux_name, token_col), … ]
    YEAR_INDEX[(section, flux)] → { years: { year: { row_idx, prof_idx, headers } } }
    PREV_UNION              → ensemble de tous les headers de prévision connus

Point d'entrée : init_full_load()
"""
from __future__ import annotations

import os
import re
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from ..config import (
    FICHIER_EXCEL_DIR,
    FICHIER_CONFIG_SECTIONS,
    FEUILLE_CONFIG_SECTIONS,
    COL_DEST,
    COL_SOURCE,
    COL_PREV,
)
from .loader import robust_load_workbook, diag_path

# ---------------------------------------------------------------------------
# État global
# ---------------------------------------------------------------------------
CACHE: dict[tuple[str, str], dict[str, Any]] = {}
STRUCT: dict[str, OrderedDict] = {}
TOKENS: dict[str, list[tuple[str, int]]] = {}
PREV_UNION: set[str] = set()
YEAR_INDEX: dict[tuple[str, str], dict] = {}

# Rempli après init_full_load()
sections: dict[str, str] = {}          # { nom_logique: nom_feuille_excel }
SECTIONS_CONFIG: list[dict] = []       # [ { dest, source, prev }, … ]
FEUILLE_REFERENCE: str = ""

# ---------------------------------------------------------------------------
# Regex
# ---------------------------------------------------------------------------
_MENSUEL_RE = re.compile(r"^Historique_prev_reel_filiales_(\d{4})_(\d{2})\.xlsx$")
_RE_PREV = re.compile(r"^Prévision\s+\d{2}/\d{2}", re.I)


def _is_prev(h: Any) -> bool:
    return isinstance(h, str) and bool(_RE_PREV.search(h.strip()))


# ---------------------------------------------------------------------------
# Parsing de dates Excel
# ---------------------------------------------------------------------------

def _parse_excel_date(v: Any) -> pd.Timestamp | None:
    """Convertit tout format de date Excel (float, datetime, str…) en Timestamp normalisé."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass

    if isinstance(v, (datetime, date)):
        return pd.Timestamp(v).normalize()

    if isinstance(v, (int, float)):
        if 10_000 <= v <= 60_000:
            try:
                return (
                    pd.Timestamp("1899-12-30") + pd.to_timedelta(int(v), unit="D")
                ).normalize()
            except Exception:
                return None
        return None

    if isinstance(v, str):
        ts = pd.to_datetime(v, dayfirst=True, errors="coerce")
        return ts.normalize() if pd.notna(ts) else None

    ts = pd.to_datetime(v, errors="coerce")
    return ts.normalize() if pd.notna(ts) else None


# ---------------------------------------------------------------------------
# Helpers labels profils
# ---------------------------------------------------------------------------

def _clean_profil_label(raw: Any, i: int) -> str:
    """Normalise le label d'un profil de prévision."""
    if raw is None:
        return f"Profil {i + 1}"
    s = str(raw).replace("(K€)", "").replace("Prévision", "Profil").strip()
    return s if s else f"Profil {i + 1}"


def _flux_name_from_token(section: str, token_col: int) -> str | None:
    """Retrouve le nom du flux associé à une colonne token dans TOKENS."""
    for name, tok in TOKENS.get(section, []):
        if tok == token_col:
            return name
    return None


# ---------------------------------------------------------------------------
# Chargement de la configuration sections
# ---------------------------------------------------------------------------

def charger_sections_depuis_cells(
    fichier_config: str = FICHIER_CONFIG_SECTIONS,
) -> list[dict[str, str]]:
    """
    Lit 'Filiales Analysées.xlsx' et retourne la liste des sections configurées.

    Returns:
        Liste de { "dest": …, "source": …, "prev": … }
    """
    if not os.path.isfile(fichier_config):
        return []

    try:
        wb = robust_load_workbook(fichier_config)
    except Exception as exc:
        diag_path(fichier_config)
        raise RuntimeError(f"Impossible d'ouvrir la config sections : {exc}") from exc

    if FEUILLE_CONFIG_SECTIONS not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[FEUILLE_CONFIG_SECTIONS]
    mapping: list[dict[str, str]] = []
    row = 2
    while True:
        dest = ws.cell(row=row, column=COL_DEST).value
        if dest is None or str(dest).strip() == "":
            break
        src = ws.cell(row=row, column=COL_SOURCE).value
        prev = ws.cell(row=row, column=COL_PREV).value
        dest_s = str(dest).strip()
        mapping.append({
            "dest":   dest_s,
            "source": str(src).strip() if src else dest_s,
            "prev":   str(prev).strip() if prev else dest_s,
        })
        row += 1

    wb.close()
    return mapping


def charger_noms_feuilles_depuis_cells(
    fichier_config: str = FICHIER_CONFIG_SECTIONS,
) -> list[str]:
    """Retourne uniquement les noms de feuilles destination (colonne A), ordonnés et uniques."""
    if not os.path.isfile(fichier_config):
        # Fallback: récupère les noms de feuilles depuis un fichier mensuel
        files = _lister_fichiers_mensuels()
        if files:
            try:
                wb = robust_load_workbook(files[-1][2])
                noms = [name for name in wb.sheetnames if name and name != "Divers"]
                wb.close()
                return noms
            except Exception:
                pass
        return []

    try:
        wb = robust_load_workbook(fichier_config)
    except Exception as exc:
        diag_path(fichier_config)
        # Fallback sur fichier mensuel
        files = _lister_fichiers_mensuels()
        if files:
            try:
                wb = robust_load_workbook(files[-1][2])
                noms = [name for name in wb.sheetnames if name and name != "Divers"]
                wb.close()
                return noms
            except Exception:
                pass
        return []

    if FEUILLE_CONFIG_SECTIONS not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[FEUILLE_CONFIG_SECTIONS]
    noms: list[str] = []
    seen: set[str] = set()
    row = 2
    while True:
        val = ws.cell(row=row, column=COL_DEST).value
        if val is None or str(val).strip() == "":
            break
        name = str(val).strip()
        if name not in seen:
            noms.append(name)
            seen.add(name)
        row += 1

    wb.close()
    return noms


# ---------------------------------------------------------------------------
# Lister les fichiers mensuels
# ---------------------------------------------------------------------------

def _lister_fichiers_mensuels() -> list[tuple[int, int, str]]:
    """
    Parcourt récursivement FICHIER_EXCEL_DIR pour trouver tous les fichiers
    'Historique_prev_reel_filiales_YYYY_MM.xlsx'.
    En cas de doublon (même année/mois), conserve le plus récent (mtime).

    Returns:
        Liste triée de (annee, mois, chemin_absolu).
    """
    if not os.path.isdir(FICHIER_EXCEL_DIR):
        return []

    best: dict[tuple[int, int], dict] = {}
    for root, _, files in os.walk(FICHIER_EXCEL_DIR):
        for fname in files:
            m = _MENSUEL_RE.match(fname)
            if not m:
                continue
            y, mth = int(m.group(1)), int(m.group(2))
            fullpath = os.path.join(root, fname)
            mtime = 0.0
            try:
                mtime = os.path.getmtime(fullpath)
            except OSError:
                pass
            key = (y, mth)
            if key not in best or mtime > best[key]["mtime"]:
                best[key] = {"path": fullpath, "mtime": mtime}

    return sorted(
        (y, mth, info["path"]) for (y, mth), info in best.items()
    )


# ---------------------------------------------------------------------------
# Tri des headers de prévision
# ---------------------------------------------------------------------------

def _parse_prev_header_sort_key(h: Any) -> tuple[int, int, str]:
    s = str(h)
    m = re.search(r"(\d{2})/(\d{2,4})", s)
    if not m:
        return (9999, 99, s)
    mm = int(m.group(1))
    yy = int(m.group(2))
    if yy < 100:
        yy += 2000
    return (yy, mm, s)


# ---------------------------------------------------------------------------
# Réconciliation des headers
# ---------------------------------------------------------------------------

def _reconcile_headers(B: dict, new_headers: list[str]) -> None:
    """
    Fusionne les headers existants de B avec new_headers en conservant l'ordre
    chronologique, et rétro-remplit None pour les nouvelles séries.
    """
    if "prev_headers" not in B or "prev_vals" not in B:
        B["prev_headers"] = []
        B["prev_vals"] = []

    old_headers: list[str] = list(B["prev_headers"])
    seen: dict[str, None] = {}
    for h in old_headers + list(new_headers):
        if h not in seen:
            seen[h] = None
    union = sorted(seen.keys(), key=_parse_prev_header_sort_key)

    if union == old_headers:
        return

    old_idx = {h: i for i, h in enumerate(old_headers)}
    n_rows = len(B.get("dates", []))

    new_prev_vals: list[list] = []
    for h in union:
        if h in old_idx:
            new_prev_vals.append(B["prev_vals"][old_idx[h]])
        else:
            new_prev_vals.append([None] * n_rows)

    B["prev_headers"] = union
    B["prev_vals"] = new_prev_vals


# ---------------------------------------------------------------------------
# Lecture structure de référence
# ---------------------------------------------------------------------------

def _lire_structure_reference(path_ref: str) -> None:
    """
    Lit la structure des flux et colonnes depuis le dernier fichier mensuel.
    Remplit STRUCT, TOKENS et PREV_UNION.
    """
    df_head = pd.read_excel(
        path_ref, sheet_name=FEUILLE_REFERENCE, header=None, nrows=3
    )
    row2 = df_head.iloc[1].tolist()
    row3 = df_head.iloc[2].tolist()

    ref_od: OrderedDict = OrderedDict()
    ref_tokens: list[tuple[str, int]] = []
    col = 2       # colonne C (0-based)
    token_col = 3

    while col < len(row2):
        flux_name = row2[col]
        if pd.isna(flux_name):
            break
        flux_name = str(flux_name).strip()

        prev_headers: list[str] = []
        prev_cols: list[int] = []
        c = col + 1
        while c < len(row3):
            h = row3[c]
            if isinstance(h, str) and "Prévision" in h:
                prev_headers.append(h.strip())
                prev_cols.append(c)
                c += 2
                continue
            break

        if prev_headers:
            ref_od[flux_name] = {
                "col_start":    col + 1,
                "prev_headers": prev_headers,
                "prev_cols":    [x + 1 for x in prev_cols],
                "dates_col":    col,
                "reel_col":     col + 1,
            }
            ref_tokens.append((flux_name, token_col))
            PREV_UNION.update(prev_headers)

        col += 2 + 2 * len(prev_headers) + 1
        token_col += 2 + 2 * len(prev_headers) + 1

    for sec in sections:
        STRUCT[sec] = ref_od.copy()
        TOKENS[sec] = ref_tokens.copy()


# ---------------------------------------------------------------------------
# Buckets & accumulation
# ---------------------------------------------------------------------------

def _ensure_flux_bucket(section: str, flux_name: str) -> dict:
    key = (section, flux_name)
    if key not in CACHE:
        CACHE[key] = {
            "dates":        [],
            "reel":         [],
            "prev_headers": [],
            "prev_vals":    [],
        }
    return CACHE[key]


def _accumuler_valeurs_tous_mois(files_list: list[tuple[int, int, str]]) -> None:
    """
    Lit chaque fichier mensuel et alimente CACHE[(section, flux)].

    Optimisations :
      - Lecture parallèle des fichiers (ThreadPoolExecutor, I/O-bound).
      - Seules les feuilles utiles sont lues (sheet_name ciblé).
      - Extraction vectorisée par colonne (plus de df.iat ligne par ligne).
      - Moteur calamine utilisé si python-calamine est installé.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    n = len(files_list)
    needed_sheets = list(set(sections.values()))

    # Moteur le plus rapide disponible
    try:
        import python_calamine  # noqa: F401
        _ENGINE: str | None = "calamine"
    except ImportError:
        _ENGINE = None

    def _read_one(args: tuple[int, str]):
        idx, path = args
        kw: dict = {"header": None}
        if _ENGINE:
            kw["engine"] = _ENGINE
        try:
            data = pd.read_excel(path, sheet_name=needed_sheets, **kw)
            if isinstance(data, pd.DataFrame):
                data = {needed_sheets[0]: data}
            return idx, path, data
        except Exception:
            pass
        # Fallback feuille par feuille (feuille absente dans ce fichier)
        result: dict = {}
        for sheet in needed_sheets:
            try:
                result[sheet] = pd.read_excel(path, sheet_name=sheet, header=None)
            except Exception:
                pass
        return idx, path, result if result else None

    # Lecture parallèle (I/O bound → threads)
    file_results: dict[int, tuple] = {}
    with ThreadPoolExecutor(max_workers=min(6, max(1, n))) as pool:
        futures = {
            pool.submit(_read_one, (i, path)): i
            for i, (_, _, path) in enumerate(files_list, 1)
        }
        for fut in as_completed(futures):
            idx, path, sheets = fut.result()
            file_results[idx] = (path, sheets)

    # Accumulation séquentielle (ordre chronologique préservé)
    for seq in range(1, n + 1):
        path, all_sheets = file_results.get(seq, (None, None))
        if not all_sheets:
            continue

        for sec, feuille in sections.items():
            if feuille not in all_sheets:
                continue
            df = all_sheets[feuille]
            if df.shape[0] < 5:
                continue

            row2 = df.iloc[1].tolist()
            row3 = df.iloc[2].tolist()
            col = 2

            while col < len(row2):
                flux_val = row2[col]
                if pd.isna(flux_val):
                    break
                flux_name = str(flux_val).strip()

                prev_headers: list[str] = []
                prev_cols: list[int] = []
                c = col + 1
                while c < len(row3):
                    h = row3[c]
                    if isinstance(h, str) and "Prévision" in h:
                        prev_headers.append(h.strip())
                        prev_cols.append(c)
                        c += 2
                        continue
                    break

                if not prev_headers:
                    col += 10
                    continue

                B = _ensure_flux_bucket(sec, flux_name)
                _reconcile_headers(B, prev_headers)
                header_to_col = {h: prev_cols[j] for j, h in enumerate(prev_headers)}

                # Extraction vectorisée (une colonne entière d'un coup)
                date_series = df.iloc[3:, col - 1]
                notna_mask  = date_series.notna()
                if not notna_mask.any():
                    col += 2 + 2 * len(prev_headers) + 1
                    continue

                # Utiliser directement les valeurs non-nulles et leurs positions
                good_positions = []  # positions dans date_series (0-based)
                good_dates = []      # dates parsées correspondantes
                
                for pos, (idx, date_val) in enumerate(date_series.items()):
                    if pd.isna(date_val):
                        continue
                    parsed_date = _parse_excel_date(date_val)
                    if parsed_date is not None:
                        good_positions.append(pos + 3)  # +3 car date_series commence à row 3
                        good_dates.append(parsed_date)
                
                dates_batch = good_dates

                if not dates_batch:
                    col += 2 + 2 * len(prev_headers) + 1
                    continue

                reel_raw = [df.iloc[p, col] for p in good_positions]

                B["dates"].extend(dates_batch)
                B["reel"].extend([v if pd.notna(v) else 0 for v in reel_raw])

                n_batch    = len(dates_batch)
                prev_h_lst = B["prev_headers"]
                prev_v_lst = B["prev_vals"]

                for k, h in enumerate(prev_h_lst):
                    try:
                        if h in header_to_col:
                            col_idx = header_to_col[h]
                            if col_idx >= len(df.columns):
                                prev_v_lst[k].extend([None] * n_batch)
                            else:
                                vals = [df.iloc[p, col_idx] for p in good_positions]
                                prev_v_lst[k].extend([v if pd.notna(v) else None for v in vals])
                        else:
                            prev_v_lst[k].extend([None] * n_batch)
                    except (IndexError, KeyError):
                        # Gérer les erreurs d'index
                        prev_v_lst[k].extend([None] * n_batch)

                col += 2 + 2 * len(prev_headers) + 1


# ---------------------------------------------------------------------------
# Index annuel
# ---------------------------------------------------------------------------

def _build_year_index() -> None:
    """
    Construit YEAR_INDEX pour un accès efficace par année :
    { (section, flux): { years: { year: { row_idx, prof_idx, headers } } } }
    """
    YEAR_INDEX.clear()
    for (section, flux_name), B in CACHE.items():
        dates = B.get("dates", [])
        prev_vals = B.get("prev_vals", [])
        headers = B.get("prev_headers", [])

        rows_by_year: dict[int, list[int]] = {}
        for i, d in enumerate(dates):
            rows_by_year.setdefault(d.year, []).append(i)

        years_map: dict[int, dict] = {}
        for y, row_idx in rows_by_year.items():
            prof_idx: list[int] = []
            for k, serie in enumerate(prev_vals):
                active = any(
                    i < len(serie) and serie[i] is not None and _nonzero(serie[i])
                    for i in row_idx
                )
                if active:
                    prof_idx.append(k)

            years_map[y] = {
                "row_idx":  row_idx,
                "prof_idx": prof_idx,
                "headers":  [
                    _clean_profil_label(headers[k] if k < len(headers) else None, k)
                    for k in prof_idx
                ],
            }

        YEAR_INDEX[(section, flux_name)] = {"years": years_map}


def _nonzero(v: Any) -> bool:
    try:
        return float(v) != 0.0
    except Exception:
        return True


# ---------------------------------------------------------------------------
# Initialisation complète
# ---------------------------------------------------------------------------

def init_full_load() -> None:
    """
    Point d'entrée principal : résout la configuration, lit tous les fichiers
    mensuels et construit les caches CACHE, STRUCT, TOKENS, YEAR_INDEX.
    Doit être appelé une seule fois au démarrage de l'application.
    
    Si aucune donnée n'est trouvée, charge automatiquement des données de test.
    """
    global sections, SECTIONS_CONFIG, FEUILLE_REFERENCE

    try:
        # 1) Charger la configuration des sections
        print("[DEBUG] Chargement configuration sections…")
        SECTIONS_CONFIG = charger_sections_depuis_cells()
        print(f"[DEBUG] Config sections chargée: {len(SECTIONS_CONFIG)} sections")
        
        dest_names = charger_noms_feuilles_depuis_cells()
        print(f"[DEBUG] Noms feuilles: {dest_names}")
        
        # IMPORTANT: Modifier le dict au lieu de le remplacer (pour que les imports existants se mettent à jour)
        sections.clear()
        sections.update({name: name for name in dest_names})
        
        if dest_names:
            FEUILLE_REFERENCE = dest_names[0]
            print(f"[DEBUG] FEUILLE_REFERENCE= {FEUILLE_REFERENCE}")

        # 2) Lister les fichiers mensuels
        files = _lister_fichiers_mensuels()
        print(f"[DEBUG] Fichiers mensuels trouvés: {len(files)}")
        
        # Vérifier que nous avons des données
        if not files or not FEUILLE_REFERENCE:
            raise RuntimeError(f"Données insuffisantes: files={len(files) if files else 0}, feuille_ref='{FEUILLE_REFERENCE}'")

        print(f"[DEBUG] Dernier fichier: {files[-1]}")
        
        # 3) Lire la structure depuis le dernier fichier
        print("[DEBUG] Lecture structure de référence…")
        _lire_structure_reference(files[-1][2])
        print(f"[DEBUG] Structure lue: STRUCT={len(STRUCT)}, TOKENS={sum(len(v) for v in TOKENS.values())}")

        # 4) Accumuler toutes les données
        print("[DEBUG] Accumulation des valeurs…")
        _accumuler_valeurs_tous_mois(files)
        print(f"[DEBUG] Données accumulées: CACHE={len(CACHE)} entrées")

        # 5) Construire l'index annuel
        print("[DEBUG] Construction index annuel…")
        _build_year_index()
        print(f"[DEBUG] Index construit: YEAR_INDEX={len(YEAR_INDEX)} entrées")
        
    except Exception as e:
        print(f"[WARN] Chargement normal échoué ({e}), chargement du mode test…")
        import traceback
        traceback.print_exc()
        
        # Charge les données de test en fallback
        try:
            from .test_data import load_test_data
            load_test_data()
            print("[OK] Mode test activé - données fictives chargées")
        except Exception as test_err:
            print(f"[ERROR] Chargement test échoué aussi : {test_err}")
            import traceback
            traceback.print_exc()
            raise
