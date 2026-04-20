# -*- coding: utf-8 -*-
"""
api/clustering_3d.py — Surface 3D de performance PULSE.
Axes : Filiale (Y) × Mois (X) → Écart Réel/Prévision (Z).

Surface lissée par interpolation cosinus (x4) pour un rendu arrondi.

GET /api/clustering_3d
Query params:
    annee      : int  (optionnel)
    filiale    : str  (optionnel)
    flux_type  : str  "enc" | "dec" | ""
    metrique   : str  "perf" | "reel" | "precision"
"""
from __future__ import annotations

import math
from collections import defaultdict

from flask import Blueprint, jsonify, request

from pulse_v2.config import FLUX_DECAISSEMENTS, FLUX_ENCAISSEMENTS
from pulse_v2.data.cache import TOKENS
from pulse_v2.data.extractor import extraire_valeurs

bp = Blueprint("clustering_3d", __name__, url_prefix="/api")

MOIS_LABELS = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
INTERP_STEPS = 8   # 12 mois → (12-1)×8 + 1 = 89 points interpolés → surface très lisse


# ── Interpolation cosinus pour surface arrondie ────────────────────────────

def _cosine_interp(a: float, b: float, t: float) -> float:
    """Interpolation cosinus entre a et b pour le paramètre t ∈ [0,1]."""
    t2 = (1.0 - math.cos(t * math.pi)) / 2.0
    return a * (1.0 - t2) + b * t2


def _smooth_row(row: list, steps: int) -> list:
    """Interpole cosinus entre chaque paire de valeurs adjacentes."""
    out = []
    for i in range(len(row) - 1):
        a, b = row[i], row[i + 1]
        for s in range(steps):
            t = s / steps
            if a is None and b is None:
                out.append(None)
            elif a is None:
                out.append(b)
            elif b is None:
                out.append(a)
            else:
                out.append(round(_cosine_interp(a, b, t), 3))
    out.append(row[-1])
    return out


def _smooth_x_axis(steps: int):
    """Positions X numériques et labels pour les axes Plotly."""
    x_vals = []
    for i in range(len(MOIS_LABELS) - 1):
        for s in range(steps):
            x_vals.append(round(i + s / steps, 4))
    x_vals.append(float(len(MOIS_LABELS) - 1))

    tick_vals = list(range(len(MOIS_LABELS)))
    tick_text = MOIS_LABELS
    return x_vals, tick_vals, tick_text


def _sens_favorable(flux: str, ecart_pct: float) -> float:
    if flux in FLUX_DECAISSEMENTS:
        return -ecart_pct
    return ecart_pct


def _mean(lst):
    return round(sum(lst) / len(lst), 3) if lst else None


@bp.route("/clustering_3d")
def get_surface_3d():
    try:
        annee_s   = request.args.get("annee",     "").strip()
        filiale_f = request.args.get("filiale",   "").strip()
        flux_type = request.args.get("flux_type", "").strip()
        metrique  = request.args.get("metrique",  "perf").strip()

        annee_filter = int(annee_s) if annee_s else None

        acc_perf      = defaultdict(list)
        acc_reel      = defaultdict(list)
        acc_precision = defaultdict(list)
        annees_set    = set()

        for section, flux_list in TOKENS.items():
            if filiale_f and section != filiale_f:
                continue

            for flux_name, _col in flux_list:
                if flux_type == "enc" and flux_name not in FLUX_ENCAISSEMENTS:
                    continue
                if flux_type == "dec" and flux_name not in FLUX_DECAISSEMENTS:
                    continue

                dates, reel_serie, previsions, _ = extraire_valeurs(
                    section, flux_name, 0, annee=None
                )

                for i, date in enumerate(dates):
                    y    = getattr(date, "year",  None)
                    mois = getattr(date, "month", 1) or 1

                    if annee_filter and y != annee_filter:
                        continue
                    if i >= len(reel_serie) or reel_serie[i] is None:
                        continue

                    r = float(reel_serie[i])
                    prev_vals = [
                        float(p[i]) for p in previsions
                        if i < len(p) and p[i] is not None
                    ]
                    if not prev_vals:
                        continue

                    prev_mean = sum(prev_vals) / len(prev_vals)
                    if r == 0 and prev_mean == 0:
                        continue

                    denom     = prev_mean if prev_mean != 0 else 1.0
                    ecart_pct = (r - prev_mean) / abs(denom) * 100
                    # Cap à ±200% pour éviter les outliers sur prévisions quasi-nulles
                    ecart_pct = max(-200.0, min(200.0, ecart_pct))
                    perf      = _sens_favorable(flux_name, ecart_pct)
                    precision = max(0.0, 100.0 - abs(ecart_pct))

                    key = (section, mois)
                    acc_perf[key].append(perf)
                    acc_reel[key].append(r)
                    acc_precision[key].append(precision)
                    if y:
                        annees_set.add(y)

        filiales = sorted(set(k[0] for k in acc_perf.keys()))
        if not filiales:
            return jsonify({
                "filiales": [], "x_vals": [], "tick_vals": [], "tick_text": [],
                "z": [], "annees": [], "total": 0, "kpis": {},
                "metrique_label": ""
            })

        # Choix métrique
        if metrique == "reel":
            acc = acc_reel
            metrique_label = "Volume réel (k€)"
        elif metrique == "precision":
            acc = acc_precision
            metrique_label = "Précision (%)"
        else:
            acc = acc_perf
            metrique_label = "Performance (%)"

        # Matrice brute [filiale][mois 1..12]
        z_raw = []
        for filiale in filiales:
            row = [_mean(acc.get((filiale, m), [])) for m in range(1, 13)]
            z_raw.append(row)

        # ── Remplissage des None par interpolation linéaire dans la rangée ──
        def fill_row(row):
            """Comble les None par interpolation cosinus entre valeurs connues."""
            n = len(row)
            result = row[:]
            known = [(i, v) for i, v in enumerate(result) if v is not None]
            if not known:
                return result
            # Extrapoler aux extrémités
            for i in range(known[0][0]):
                result[i] = known[0][1]
            for i in range(known[-1][0] + 1, n):
                result[i] = known[-1][1]
            # Interpolation cosinus entre les valeurs connues
            for ki in range(len(known) - 1):
                i0, v0 = known[ki]
                i1, v1 = known[ki + 1]
                for i in range(i0 + 1, i1):
                    t = (i - i0) / (i1 - i0)
                    result[i] = round(_cosine_interp(v0, v1, t), 3)
            return result

        z_filled = [fill_row(row) for row in z_raw]

        # ── Lissage cosinus : 12 → 45 points ──
        z_smooth = [_smooth_row(row, INTERP_STEPS) for row in z_filled]
        x_vals, tick_vals, tick_text = _smooth_x_axis(INTERP_STEPS)

        # ── KPIs ──
        all_perf = [v for vals in acc_perf.values() for v in vals]
        n_total  = sum(len(v) for v in acc_perf.values())

        def _filiale_mean(f):
            flat = [v for m in range(1, 13) for v in acc_perf.get((f, m), [])]
            return _mean(flat) or 0.0

        def _mois_mean(m):
            flat = [v for f in filiales for v in acc_perf.get((f, m), [])]
            return _mean(flat) or 0.0

        best_filiale  = max(filiales, key=_filiale_mean)
        worst_filiale = min(filiales, key=_filiale_mean)
        best_mois_idx  = max(range(1, 13), key=_mois_mean)
        worst_mois_idx = min(range(1, 13), key=_mois_mean)

        return jsonify({
            "filiales":        filiales,
            "x_vals":          x_vals,
            "tick_vals":       tick_vals,
            "tick_text":       tick_text,
            "z":               z_smooth,
            "metrique_label":  metrique_label,
            "annees":          sorted(annees_set),
            "total":           n_total,
            "kpis": {
                "perf_globale":   round(_mean(all_perf) or 0, 1),
                "best_filiale":   best_filiale,
                "worst_filiale":  worst_filiale,
                "best_mois":      MOIS_LABELS[best_mois_idx - 1],
                "worst_mois":     MOIS_LABELS[worst_mois_idx - 1],
                "n_filiales":     len(filiales),
            }
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500
