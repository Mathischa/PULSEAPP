# -*- coding: utf-8 -*-
from __future__ import annotations

from collections import defaultdict
from itertools import zip_longest
import datetime as dt
import statistics
import math

from flask import Blueprint, jsonify, request

from pulse_v2.data.cache import TOKENS
from pulse_v2.data.extractor import extraire_valeurs

bp = Blueprint("tendance", __name__, url_prefix="/api")

SEUIL_MIN_FLUX = 10000
KMEANS_MAX_K = 3


# =========================================================
# HELPERS GÉNÉRAUX
# =========================================================
def _to_number(x):
    if x is None:
        return None
    if isinstance(x, str):
        s = x.strip().replace("\xa0", " ").replace(" ", "")
        if s in {"", "-", "—", "NA", "N/A"}:
            return None
        s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return None
    try:
        return float(x)
    except Exception:
        return None


def _to_date(x):
    if x is None:
        return None
    if hasattr(x, "year") and hasattr(x, "month") and hasattr(x, "day"):
        try:
            return dt.date(x.year, x.month, x.day)
        except Exception:
            return None
    if isinstance(x, str):
        txt = x.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d"):
            try:
                return dt.datetime.strptime(txt, fmt).date()
            except Exception:
                pass
    return None


def _is_business_day(d):
    return d is not None and d.weekday() < 5


def _safe_mean(vals):
    vals = [v for v in vals if v is not None]
    return statistics.mean(vals) if vals else 0.0


def _safe_stdev(vals):
    vals = [v for v in vals if v is not None]
    if len(vals) < 2:
        return 0.0
    try:
        return statistics.stdev(vals)
    except Exception:
        return 0.0


def _pct_vs(base, value):
    if base is None or base == 0:
        return 0.0
    return (value - base) / abs(base) * 100


def _last_business_day_of_month(d):
    if d is None:
        return None
    if d.month == 12:
        next_month = dt.date(d.year + 1, 1, 1)
    else:
        next_month = dt.date(d.year, d.month + 1, 1)
    last_day = next_month - dt.timedelta(days=1)
    while last_day.weekday() >= 5:
        last_day -= dt.timedelta(days=1)
    return last_day


def _score_tendance(moyennes, moyenne_globale, counts):
    if not moyennes or moyenne_globale == 0:
        return 0.0
    vals_non_vides = [
        abs(_pct_vs(moyenne_globale, v))
        for v, c in zip(moyennes, counts)
        if c > 0
    ]
    if not vals_non_vides:
        return 0.0
    intensite = max(vals_non_vides)
    couverture = sum(1 for c in counts if c > 0) / len(counts) if counts else 0.0
    return round(intensite * couverture, 1)


def _score_saisonnalite(moyennes, moyenne_globale, counts):
    if not moyenne_globale or not any(counts):
        return 0.0
    deviations = [
        abs(_pct_vs(moyenne_globale, v))
        for v, c in zip(moyennes, counts)
        if c > 0
    ]
    return round(sum(deviations) / len(deviations), 1) if deviations else 0.0


def _niveau_risque(score):
    if score >= 20:
        return "Élevé"
    if score >= 10:
        return "Modéré"
    return "Faible"


def _couleur_risque(score):
    if score >= 20:
        return "#C0392B"
    if score >= 10:
        return "#D68910"
    return "#1E8449"


def _label_stabilite(cv):
    if cv < 10:
        return "Très stable"
    if cv < 20:
        return "Stable"
    if cv < 35:
        return "Variable"
    return "Très variable"


def _tag_stabilite(cv):
    if cv < 20:
        return "stable"
    if cv < 35:
        return "variable"
    return "tres_variable"


def _stats_stabilite(vals):
    vals = [v for v in vals if v is not None]
    n = len(vals)

    if n == 0:
        return {
            "n": 0,
            "mean": 0.0,
            "median": 0.0,
            "min": 0.0,
            "max": 0.0,
            "stdev": 0.0,
            "cv": 0.0,
            "ic_low": 0.0,
            "ic_high": 0.0,
            "ic_margin": 0.0,
            "amplitude": 0.0,
            "label": "Aucune donnée",
            "tag": "tres_variable",
        }

    mean_v = statistics.mean(vals)
    median_v = statistics.median(vals)
    min_v = min(vals)
    max_v = max(vals)
    stdev_v = statistics.stdev(vals) if n >= 2 else 0.0
    cv = (stdev_v / abs(mean_v) * 100) if mean_v not in (0, None) else 0.0
    margin = 1.96 * (stdev_v / math.sqrt(n)) if n >= 2 else 0.0

    return {
        "n": n,
        "mean": mean_v,
        "median": median_v,
        "min": min_v,
        "max": max_v,
        "stdev": stdev_v,
        "cv": cv,
        "ic_low": mean_v - margin,
        "ic_high": mean_v + margin,
        "ic_margin": margin,
        "amplitude": max_v - min_v,
        "label": _label_stabilite(cv),
        "tag": _tag_stabilite(cv),
    }


def _pick_peak_and_trough(valid_idx, stats_list, orientation):
    if not valid_idx:
        return None, None
    if orientation < 0:
        idx_peak = min(valid_idx, key=lambda i: stats_list[i]["mean"])
        idx_trough = max(valid_idx, key=lambda i: stats_list[i]["mean"])
    else:
        idx_peak = max(valid_idx, key=lambda i: stats_list[i]["mean"])
        idx_trough = min(valid_idx, key=lambda i: stats_list[i]["mean"])
    return idx_peak, idx_trough


def _pick_peak_idx(valid_idx, stats_list, orientation):
    if not valid_idx:
        return None
    if orientation < 0:
        return min(valid_idx, key=lambda i: stats_list[i]["mean"])
    return max(valid_idx, key=lambda i: stats_list[i]["mean"])


def _calculer_indices_radar(all_values_par_mois):
    mois_labels = [
        "Jan", "Fév", "Mar", "Avr", "Mai", "Jun",
        "Jul", "Aoû", "Sep", "Oct", "Nov", "Déc"
    ]

    periodes_dispo = sorted(all_values_par_mois.keys())
    if not periodes_dispo:
        return {}, []

    moyennes = {
        m: (_safe_mean(all_values_par_mois[m]) if all_values_par_mois[m] else 0.0)
        for m in periodes_dispo
    }
    moy_globale = _safe_mean(list(moyennes.values()))
    if moy_globale == 0:
        moy_globale = 1.0

    indices = {m: round((moyennes[m] / moy_globale) * 100, 1) for m in periodes_dispo}
    labels = [mois_labels[m - 1] for m in periodes_dispo if 1 <= m <= 12]
    return indices, labels


# =========================================================
# HELPERS K-MEANS 1D
# =========================================================
def _cluster_name(rank, k):
    if k <= 1:
        return "Régime unique"
    if k == 2:
        return ["Bas", "Haut"][rank]
    return ["Bas", "Moyen", "Haut"][rank]


def _cluster_color(rank, k):
    if k <= 1:
        return "#AAB7B8"
    palette = ["#5DADE2", "#F5B041", "#EC7063"]
    return palette[min(rank, len(palette) - 1)]


def _kmeans_1d(vals, k, max_iter=100):
    vals = [float(v) for v in vals if v is not None]
    n = len(vals)

    if n == 0:
        return {
            "k": 0,
            "values": [],
            "assignments": [],
            "centers": [],
            "clusters": [],
            "inertia": 0.0,
        }

    unique_vals = sorted(set(vals))
    k = max(1, min(k, len(unique_vals), n))

    if k == 1:
        center = _safe_mean(vals)
        s = _stats_stabilite(vals)
        return {
            "k": 1,
            "values": vals[:],
            "assignments": [0] * n,
            "centers": [center],
            "clusters": [{
                **s,
                "cluster_index": 0,
                "center": center,
                "name": _cluster_name(0, 1),
                "color": _cluster_color(0, 1),
                "values": vals[:],
            }],
            "inertia": sum((v - center) ** 2 for v in vals),
        }

    unique_sorted = unique_vals[:]
    positions = [round(i * (len(unique_sorted) - 1) / (k - 1)) for i in range(k)]
    centers = [unique_sorted[pos] for pos in positions]

    dedup_centers = []
    for c in centers:
        if c not in dedup_centers:
            dedup_centers.append(c)
    for candidate in unique_sorted:
        if len(dedup_centers) >= k:
            break
        if candidate not in dedup_centers:
            dedup_centers.append(candidate)
    centers = dedup_centers[:k]

    assignments = [0] * n

    for _ in range(max_iter):
        for i, v in enumerate(vals):
            assignments[i] = min(range(len(centers)), key=lambda j: (abs(v - centers[j]), j))

        new_centers = []
        for j in range(len(centers)):
            cluster_vals = [v for v, a in zip(vals, assignments) if a == j]
            if cluster_vals:
                new_centers.append(_safe_mean(cluster_vals))
            else:
                farthest_point = max(vals, key=lambda v: min(abs(v - c) for c in centers))
                new_centers.append(farthest_point)

        if len(new_centers) == len(centers) and all(abs(a - b) < 1e-9 for a, b in zip(new_centers, centers)):
            centers = new_centers
            break

        centers = new_centers

    order = sorted(range(len(centers)), key=lambda j: centers[j])
    remap = {old_idx: new_idx for new_idx, old_idx in enumerate(order)}
    centers_sorted = [centers[j] for j in order]
    assignments_sorted = [remap[a] for a in assignments]

    clusters = []
    inertia = 0.0

    for new_j, center in enumerate(centers_sorted):
        cluster_vals = [v for v, a in zip(vals, assignments_sorted) if a == new_j]
        if not cluster_vals:
            continue

        s = _stats_stabilite(cluster_vals)
        inertia += sum((v - center) ** 2 for v in cluster_vals)
        clusters.append({
            **s,
            "cluster_index": new_j,
            "center": center,
            "name": _cluster_name(new_j, len(centers_sorted)),
            "color": _cluster_color(new_j, len(centers_sorted)),
            "values": cluster_vals[:],
        })

    if len(clusters) != len(centers_sorted):
        new_centers = [c["center"] for c in clusters]
        new_assignments = []
        for v in vals:
            if not new_centers:
                new_assignments.append(0)
            else:
                new_assignments.append(
                    min(range(len(new_centers)), key=lambda j: (abs(v - new_centers[j]), j))
                )
        centers_sorted = new_centers
        assignments_sorted = new_assignments

    final_k = len(clusters)
    for rank, cluster in enumerate(clusters):
        cluster["cluster_index"] = rank
        cluster["name"] = _cluster_name(rank, final_k)
        cluster["color"] = _cluster_color(rank, final_k)

    return {
        "k": final_k,
        "values": vals[:],
        "assignments": assignments_sorted,
        "centers": centers_sorted,
        "clusters": clusters,
        "inertia": inertia,
    }


def _best_kmeans_1d(vals, max_k=3):
    vals = [v for v in vals if v is not None]
    n = len(vals)
    unique_count = len(set(vals))

    if n == 0:
        return {
            "k": 0,
            "values": [],
            "assignments": [],
            "centers": [],
            "clusters": [],
            "inertia": 0.0,
        }

    if n < 6 or unique_count <= 1:
        return _kmeans_1d(vals, 1)

    k_upper = min(max_k, unique_count, 3 if n >= 12 else 2)
    models = {k: _kmeans_1d(vals, k) for k in range(1, k_upper + 1)}

    chosen_k = 1

    if 2 in models and models[1]["inertia"] > 0:
        gain_2 = (models[1]["inertia"] - models[2]["inertia"]) / models[1]["inertia"]
        min_cluster_size_2 = min((c["n"] for c in models[2]["clusters"]), default=0)
        if gain_2 >= 0.18 and min_cluster_size_2 >= 2:
            chosen_k = 2

    if chosen_k == 2 and 3 in models and models[2]["inertia"] > 0:
        gain_3 = (models[2]["inertia"] - models[3]["inertia"]) / models[2]["inertia"]
        min_cluster_size_3 = min((c["n"] for c in models[3]["clusters"]), default=0)
        if gain_3 >= 0.10 and min_cluster_size_3 >= 2:
            chosen_k = 3

    return models[chosen_k]


def _cluster_dominance_metrics(km):
    total_n = len(km["values"]) if km else 0
    if not km or km["k"] == 0 or total_n == 0:
        return {
            "share": 0.0,
            "dominant_n": 0,
            "dominant_cluster": None,
            "label": "Aucune donnée",
            "tag": "tres_variable",
            "color": "#C0392B",
        }

    dominant_cluster = max(km["clusters"], key=lambda c: c["n"])
    dominant_n = dominant_cluster["n"]
    share = dominant_n / total_n

    if share >= 0.60:
        label_v, tag_v, color_v = "Stable", "stable", "#27AE60"
    elif share >= 0.45:
        label_v, tag_v, color_v = "Variable", "variable", "#D68910"
    else:
        label_v, tag_v, color_v = "Très variable", "tres_variable", "#C0392B"

    return {
        "share": share,
        "dominant_n": dominant_n,
        "dominant_cluster": dominant_cluster,
        "label": label_v,
        "tag": tag_v,
        "color": color_v,
    }


# =========================================================
# CALCUL PRINCIPAL
# =========================================================
def calculer_analyse_tendance(section, flux_nom, annee, tokens, extraire_valeurs_fn):
    if not section or not flux_nom:
        raise ValueError("Paramètres 'section' et 'flux' requis")

    col_start = None
    for name, col in tokens.get(section, []):
        if name == flux_nom:
            col_start = col
            break

    if col_start is None:
        raise LookupError(f"Flux '{flux_nom}' introuvable dans '{section}'")

    dates, reel, previsions, noms_profils = extraire_valeurs_fn(
        section, flux_nom, 0, annee=None
    )

    dates = list(dates) if dates is not None else []
    reel = list(reel) if reel is not None else []

    weekly_data = defaultdict(list)
    monthly_day_data = defaultdict(list)
    yearly_month_data = defaultdict(list)
    radar_month_data = defaultdict(list)
    month_position_data = {
        "Début de mois": [],
        "Milieu de mois": [],
        "Fin de mois": [],
    }

    total_debug = {
        "raw_pairs": 0,
        "zip_missing": 0,
        "date_ok": 0,
        "reel_ok": 0,
        "annee_ok": 0,
        "weekend_exclus": 0,
        "seuil_exclus": 0,
        "kept": 0,
    }

    all_values = []

    for d_raw, r_raw in zip_longest(dates, reel, fillvalue=None):
        total_debug["raw_pairs"] += 1

        if d_raw is None or r_raw is None:
            total_debug["zip_missing"] += 1
            continue

        d = _to_date(d_raw)
        if d is None:
            continue
        total_debug["date_ok"] += 1

        r = _to_number(r_raw)
        if r is None:
            continue
        total_debug["reel_ok"] += 1

        radar_month_data[d.month].append(r)

        if annee is not None and d.year != annee:
            continue
        total_debug["annee_ok"] += 1

        if not _is_business_day(d):
            total_debug["weekend_exclus"] += 1
            continue

        if abs(r) < SEUIL_MIN_FLUX:
            total_debug["seuil_exclus"] += 1
            continue

        total_debug["kept"] += 1

        all_values.append(r)
        weekly_data[d.weekday()].append(r)
        monthly_day_data[d.day].append(r)
        yearly_month_data[d.month].append(r)

        last_bd = _last_business_day_of_month(d)
        if last_bd:
            if d.day <= 5:
                month_position_data["Début de mois"].append(r)
            elif (last_bd - d).days <= 4:
                month_position_data["Fin de mois"].append(r)
            else:
                month_position_data["Milieu de mois"].append(r)

    if not all_values:
        return {
            "section": section,
            "flux": flux_nom,
            "annee": annee,
            "error": "Aucune donnée exploitable pour la combinaison sélectionnée.",
            "debug": total_debug,
        }

    moyenne_globale = _safe_mean(all_values)
    orientation_flux = -1 if moyenne_globale < 0 else 1

    jours_semaine = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
    idx_jours_semaine = [0, 1, 2, 3, 4]

    stats_week = []
    moy_week = []
    count_week = []

    for i in idx_jours_semaine:
        s = _stats_stabilite(weekly_data[i])
        stats_week.append(s)
        moy_week.append(s["mean"])
        count_week.append(s["n"])

    valid_week_idx = [i for i, s in enumerate(stats_week) if s["n"] > 0]
    if not valid_week_idx:
        return {
            "section": section,
            "flux": flux_nom,
            "annee": annee,
            "error": "Aucune donnée exploitable pour la combinaison sélectionnée.",
            "debug": total_debug,
        }

    idx_week_peak, idx_week_trough = _pick_peak_and_trough(valid_week_idx, stats_week, orientation_flux)
    score_hebdo = _score_tendance(moy_week, moyenne_globale, count_week)
    score_saisonnalite = _score_saisonnalite(moy_week, moyenne_globale, count_week)
    risque_global = max(score_hebdo, score_saisonnalite)

    weekday_clusters = {}
    weekday_cluster_metrics = {}
    for i in idx_jours_semaine:
        weekday_clusters[i] = _best_kmeans_1d(weekly_data[i], max_k=KMEANS_MAX_K)
        weekday_cluster_metrics[i] = _cluster_dominance_metrics(weekday_clusters[i])

    idx_kmeans_max = max(
        valid_week_idx,
        key=lambda i: (
            weekday_clusters[i]["k"],
            (
                weekday_clusters[i]["clusters"][-1]["center"] - weekday_clusters[i]["clusters"][0]["center"]
                if weekday_clusters[i]["k"] >= 2 else 0.0
            )
        )
    )

    idx_stab_km_max = max(valid_week_idx, key=lambda i: weekday_cluster_metrics[i]["share"])
    idx_stab_km_min = min(valid_week_idx, key=lambda i: weekday_cluster_metrics[i]["share"])

    jours_mois = list(range(1, 32))
    stats_month_day = [_stats_stabilite(monthly_day_data[j]) for j in jours_mois]
    valid_month_day_idx = [i for i, s in enumerate(stats_month_day) if s["n"] > 0]

    rolling_window_centers = []
    rolling_window_stats = []
    for center_day in range(1, 32):
        vals = []
        for neighbor in [center_day - 1, center_day, center_day + 1]:
            if 1 <= neighbor <= 31:
                vals.extend(monthly_day_data[neighbor])
        rolling_window_centers.append(center_day)
        rolling_window_stats.append(_stats_stabilite(vals))

    valid_roll_idx = [i for i, s in enumerate(rolling_window_stats) if s["n"] > 0]
    idx_roll_peak = _pick_peak_idx(valid_roll_idx, rolling_window_stats, orientation_flux) if valid_roll_idx else None

    mois_annee = ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"]
    stats_year_month = [_stats_stabilite(yearly_month_data[m]) for m in range(1, 13)]
    valid_year_month_idx = [i for i, s in enumerate(stats_year_month) if s["n"] > 0]

    if valid_year_month_idx:
        idx_year_month_peak, idx_year_month_trough = _pick_peak_and_trough(
            valid_year_month_idx, stats_year_month, orientation_flux
        )
    else:
        idx_year_month_peak, idx_year_month_trough = None, None

    debut_stats = _stats_stabilite(month_position_data["Début de mois"])
    milieu_stats = _stats_stabilite(month_position_data["Milieu de mois"])
    fin_stats = _stats_stabilite(month_position_data["Fin de mois"])

    indices_radar, labels_radar = _calculer_indices_radar(radar_month_data)
    periodes_radar = sorted(indices_radar.keys()) if indices_radar else []

    moy_jours_valides = [stats_week[i]["mean"] for i in valid_week_idx]
    moy_globale_hebdo = _safe_mean(moy_jours_valides) if moy_jours_valides else 0.0

    blocs_jour = []
    for start in range(1, 32, 3):
        end = min(start + 2, 31)
        vals_bloc = []
        for j in range(start, end + 1):
            vals_bloc.extend(monthly_day_data[j])
        if vals_bloc:
            blocs_jour.append((start, end, vals_bloc))

    moy_glob_jour = 0.0
    if len(blocs_jour) >= 3:
        moy_blocs = [_safe_mean(b[2]) for b in blocs_jour]
        moy_glob_jour = _safe_mean(moy_blocs) if moy_blocs else 0.0

    return {
        "section": section,
        "flux": flux_nom,
        "annee": annee,
        "debug": total_debug,
        "global": {
            "moyenne_globale": round(moyenne_globale, 2),
            "orientation_flux": orientation_flux,
            "score_hebdo": score_hebdo,
            "score_saisonnalite": score_saisonnalite,
            "risque_global": risque_global,
            "niveau_risque": _niveau_risque(risque_global),
            "couleur_risque": _couleur_risque(risque_global),
            "nb_points": len(all_values),
        },
        "hebdo": {
            "labels": jours_semaine,
            "stats": stats_week,
            "valid_idx": valid_week_idx,
            "idx_peak": idx_week_peak,
            "idx_trough": idx_week_trough,
            "cluster_metrics": weekday_cluster_metrics,
            "clusters": weekday_clusters,
            "idx_kmeans_max": idx_kmeans_max,
            "idx_stab_km_max": idx_stab_km_max,
            "idx_stab_km_min": idx_stab_km_min,
        },
        "mensuel": {
            "jours": jours_mois,
            "stats_jour": stats_month_day,
            "valid_idx": valid_month_day_idx,
            "rolling_centers": rolling_window_centers,
            "rolling_stats": rolling_window_stats,
            "valid_roll_idx": valid_roll_idx,
            "idx_roll_peak": idx_roll_peak,
            "positions": {
                "Début de mois": debut_stats,
                "Milieu de mois": milieu_stats,
                "Fin de mois": fin_stats,
            },
        },
        "annuel": {
            "labels": mois_annee,
            "stats_mois": stats_year_month,
            "valid_idx": valid_year_month_idx,
            "idx_peak": idx_year_month_peak,
            "idx_trough": idx_year_month_trough,
        },
        "radars": {
            "mensuel": {
                "indices": indices_radar,
                "labels": labels_radar,
                "periodes": periodes_radar,
            },
            "hebdo": {
                "moy_globale": round(moy_globale_hebdo, 2),
                "labels": [jours_semaine[i] for i in valid_week_idx],
                "values": [
                    round((stats_week[i]["mean"] / moy_globale_hebdo) * 100, 1)
                    for i in valid_week_idx
                ] if moy_globale_hebdo != 0 and len(valid_week_idx) >= 3 else [],
            },
            "intra_mensuel": {
                "moy_globale": round(moy_glob_jour, 2),
                "labels": [
                    f"{b[0]}" if b[0] == b[1] else f"{b[0]}-{b[1]}"
                    for b in blocs_jour
                ] if moy_glob_jour != 0 else [],
                "values": [
                    round((_safe_mean(b[2]) / moy_glob_jour) * 100, 1)
                    for b in blocs_jour
                ] if moy_glob_jour != 0 else [],
                "blocs": [
                    {
                        "start": b[0],
                        "end": b[1],
                        "stats": _stats_stabilite(b[2]),
                    }
                    for b in blocs_jour
                ],
            },
        },
        "previsions": [
            {
                "label": noms_profils[i] if i < len(noms_profils) else f"Profil {i + 1}",
                "values": [
                    round(float(v), 2) if v is not None else None
                    for v in prev_serie
                ],
            }
            for i, prev_serie in enumerate(previsions or [])
        ],
        "reel_serie": [
            round(float(v), 2) if v is not None else None
            for v in reel
        ],
        "dates_serie": [
            d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)
            for d in dates
        ],
    }


# =========================================================
# ROUTE API
# =========================================================
@bp.route("/tendance")
def get_tendance():
    section = request.args.get("section", "").strip()
    flux_nom = request.args.get("flux", "").strip()
    annee_s = request.args.get("annee", "").strip()

    if not section or not flux_nom:
        return jsonify({"error": "Paramètres 'section' et 'flux' requis"}), 400

    annee = int(annee_s) if annee_s.isdigit() else None

    try:
        payload = calculer_analyse_tendance(
            section=section,
            flux_nom=flux_nom,
            annee=annee,
            tokens=TOKENS,
            extraire_valeurs_fn=extraire_valeurs,
        )
    except LookupError as e:
        return jsonify({"error": str(e)}), 404
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Erreur interne: {e}"}), 500

    if payload.get("error"):
        return jsonify(payload), 404

    return jsonify(payload)