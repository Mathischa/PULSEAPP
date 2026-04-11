# -*- coding: utf-8 -*-
"""
api/ml_ecarts.py — Analyse ML des écarts (clustering + explication).

POST /api/ml_ecarts/analyse
    body: { section, flux, annee, algo }
    → points, clusters, summary, global_stats, algo_info

POST /api/ml_ecarts/explication
    body: { section, flux, annee }
    → feature_importance
"""
from __future__ import annotations
import math
from flask import Blueprint, jsonify, request
from pulse_v2.data.cache import CACHE

bp = Blueprint("ml_ecarts", __name__, url_prefix="/api/ml_ecarts")

CLUSTER_COLORS = ["#00CC66", "#FFD700", "#FFA500", "#EF4444", "#A855F7", "#38BDF8"]
NOISE_COLOR    = "#888888"


# ─────────────────────────────────────────────────────────────
#  COLLECTE DES POINTS
# ─────────────────────────────────────────────────────────────
def _collect(section_filter: str, flux_filter: str, annee_filter: int | None) -> list[dict]:
    points = []
    for (section, flux_name), bucket in CACHE.items():
        if section_filter and section != section_filter:
            continue
        if flux_filter and flux_filter not in ("", "Tous les flux") and flux_name != flux_filter:
            continue

        dates     = bucket.get("dates",     [])
        reel      = bucket.get("reel",      [])
        prev_vals = bucket.get("prev_vals", [])

        for i, d in enumerate(dates):
            if annee_filter is not None:
                if not (hasattr(d, "year") and d.year == annee_filter):
                    continue

            r = reel[i] if i < len(reel) else None
            if r is None or not isinstance(r, (int, float)) or math.isnan(r):
                continue

            for prev_serie in prev_vals:
                pv = prev_serie[i] if i < len(prev_serie) else None
                if pv is None or not isinstance(pv, (int, float)) or math.isnan(pv):
                    continue

                denom = pv if pv != 0 else (r if r != 0 else None)
                if denom is None:
                    continue
                try:
                    ecart_pct = (r - pv) / abs(denom) * 100.0
                except ZeroDivisionError:
                    continue

                if not math.isfinite(ecart_pct):
                    continue

                points.append({
                    "x":       round(ecart_pct, 4),
                    "y":       round(r - pv, 2),
                    "section": section,
                    "flux":    flux_name,
                    "date":    d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d),
                })
    return points


# ─────────────────────────────────────────────────────────────
#  ENDPOINT ANALYSE (clustering)
# ─────────────────────────────────────────────────────────────
@bp.route("/analyse", methods=["POST"])
def analyse():
    import numpy as np
    from sklearn.cluster import KMeans, DBSCAN
    from sklearn.preprocessing import RobustScaler
    from sklearn.ensemble import IsolationForest
    from sklearn.metrics import silhouette_score

    body        = request.get_json(force=True) or {}
    section     = body.get("section", "").strip()
    flux        = body.get("flux", "").strip()
    annee_s     = body.get("annee", "")
    algo        = body.get("algo", "kmeans")
    annee       = int(annee_s) if str(annee_s).isdigit() else None

    if not section:
        return jsonify({"error": "Paramètre section requis"}), 400

    raw = _collect(section, flux, annee)
    if len(raw) < 4:
        return jsonify({"error": f"Pas assez de points ({len(raw)}) pour l'analyse."}), 422

    X = np.array([[p["x"], p["y"]] for p in raw])

    # Scaling
    scaler = RobustScaler()
    Xs     = scaler.fit_transform(X)

    # Anomaly detection
    contamination = min(0.08, max(0.02,
        np.sum(np.abs(X[:, 1]) > np.quantile(np.abs(X[:, 1]), 0.95)) / max(len(X), 1)
    ))
    iso    = IsolationForest(contamination=contamination, random_state=42)
    is_out = iso.fit_predict(Xs) == -1  # True = anomalie

    # Clustering
    algo_info   = ""
    labels      = None
    centers_inv = None

    if algo == "dbscan":
        db          = DBSCAN(eps=0.8, min_samples=5)
        labels      = db.fit_predict(Xs)
        unique_cl   = sorted(set(labels))
        order       = [cl for cl in unique_cl if cl != -1] + ([-1] if -1 in unique_cl else [])
        color_map   = {}
        ci          = 0
        for cl in order:
            if cl == -1:
                color_map[cl] = NOISE_COLOR
            else:
                color_map[cl] = CLUSTER_COLORS[ci % len(CLUSTER_COLORS)]
                ci += 1
        algo_info = f"DBSCAN (eps=0.8, min_samples=5) — {len([c for c in order if c != -1])} cluster(s)"
    else:
        # KMeans auto-k
        n_samples = Xs.shape[0]
        best_k, best_sil = 2, None
        for k in range(2, min(7, n_samples)):
            try:
                km  = KMeans(n_clusters=k, n_init=10, random_state=42, algorithm="lloyd")
                lbl = km.fit_predict(Xs)
                if len(set(lbl)) < 2:
                    continue
                s = silhouette_score(Xs, lbl)
                if best_sil is None or s > best_sil:
                    best_k, best_sil = k, s
            except Exception:
                continue

        km          = KMeans(n_clusters=best_k, n_init=10, random_state=42, algorithm="lloyd")
        labels      = km.fit_predict(Xs)
        centers_inv = scaler.inverse_transform(km.cluster_centers_).tolist()

        # Ordonner par impact absolu moyen croissant
        impact     = {cl: float(np.abs(X[labels == cl, 1]).mean()) for cl in range(best_k)}
        order      = sorted(impact, key=lambda c: impact[c])
        color_map  = {cl: CLUSTER_COLORS[i % len(CLUSTER_COLORS)] for i, cl in enumerate(order)}
        sil_str    = f", silhouette={best_sil:.2f}" if best_sil else ""
        algo_info  = f"KMeans auto-k — k={best_k}{sil_str}"

    # Nom d'affichage
    def cl_name(cl):
        if algo == "dbscan":
            return "Bruit" if cl == -1 else f"Cluster {cl}"
        return f"Cluster {cl + 1}"

    # Construction des points enrichis
    points_out = []
    for i, p in enumerate(raw):
        cl = int(labels[i])
        points_out.append({
            **p,
            "cluster":  cl,
            "color":    color_map.get(cl, NOISE_COLOR),
            "outlier":  bool(is_out[i]),
            "cl_name":  cl_name(cl),
        })

    # Clusters metadata (pour légende)
    seen = {}
    for p in points_out:
        cl = p["cluster"]
        if cl not in seen:
            seen[cl] = {"id": cl, "label": p["cl_name"], "color": p["color"], "count": 0}
        seen[cl]["count"] += 1
    clusters = list(seen.values())
    if centers_inv:
        for i, cl_meta in enumerate(clusters):
            orig_cl = cl_meta["id"]
            cl_meta["cx"] = round(centers_inv[orig_cl][0], 2)
            cl_meta["cy"] = round(centers_inv[orig_cl][1], 2)

    # Répartition
    n_total = len(points_out)
    for cl_meta in clusters:
        cl_meta["pct"] = round(cl_meta["count"] / n_total * 100, 1)

    # Summary par cluster
    import statistics
    summary_rows = []
    for cl_meta in clusters:
        cl_pts = [p for p in points_out if p["cluster"] == cl_meta["id"]]
        xs     = [p["x"] for p in cl_pts]
        ys     = [p["y"] for p in cl_pts]
        anom   = sum(1 for p in cl_pts if p["outlier"])
        summary_rows.append({
            "label":     cl_meta["label"],
            "color":     cl_meta["color"],
            "count":     len(cl_pts),
            "pct":       cl_meta["pct"],
            "mean_pct":  round(statistics.mean(xs), 2) if xs else 0,
            "mean_valo": round(statistics.mean(ys), 1) if ys else 0,
            "sum_valo":  round(sum(ys), 0),
            "anomalies": anom,
            "anom_pct":  round(anom / max(len(cl_pts), 1) * 100, 1),
        })

    # Stats globales
    all_y   = [p["y"] for p in points_out]
    n_anom  = sum(1 for p in points_out if p["outlier"])
    total_y = sum(all_y)

    # Top risques
    from collections import defaultdict
    sec_risk  = defaultdict(float)
    flux_risk = defaultdict(float)
    for p in points_out:
        sec_risk [p["section"]] += abs(p["y"])
        flux_risk[p["flux"]]    += abs(p["y"])

    top_section = max(sec_risk,  key=sec_risk.get)  if sec_risk  else "N/A"
    top_flux    = max(flux_risk, key=flux_risk.get) if flux_risk else "N/A"

    return jsonify({
        "points":      points_out,
        "clusters":    clusters,
        "summary":     summary_rows,
        "algo_info":   algo_info,
        "global_stats": {
            "n_total":     n_total,
            "n_anomalies": n_anom,
            "anom_pct":    round(n_anom / max(n_total, 1) * 100, 1),
            "total_impact": round(total_y, 0),
            "top_section": top_section,
            "top_section_val": round(sec_risk.get(top_section, 0), 0),
            "top_flux":    top_flux,
            "top_flux_val": round(flux_risk.get(top_flux, 0), 0),
        },
    })


# ─────────────────────────────────────────────────────────────
#  ENDPOINT EXPLICATION (feature importance)
# ─────────────────────────────────────────────────────────────
@bp.route("/explication", methods=["POST"])
def explication():
    import numpy as np
    import pandas as pd
    from sklearn.feature_selection import SelectKBest, f_classif

    body    = request.get_json(force=True) or {}
    section = body.get("section", "").strip()
    flux    = body.get("flux",    "").strip()
    annee_s = body.get("annee",   "")
    annee   = int(annee_s) if str(annee_s).isdigit() else None

    if not section:
        return jsonify({"error": "Paramètre section requis"}), 400

    raw = _collect(section, flux, annee)
    if len(raw) < 10:
        return jsonify({"error": f"Pas assez de points ({len(raw)}) pour l'analyse explicative."}), 422

    df = pd.DataFrame(raw)
    df["abs_y"] = df["y"].abs()

    if df["abs_y"].max() == 0:
        return jsonify({"error": "Écarts trop faibles pour l'analyse explicative."}), 422

    seuil  = float(df["abs_y"].quantile(0.75))
    df["target"] = (df["abs_y"] >= seuil).astype(int)

    if df["target"].sum() == 0 or df["target"].sum() == len(df):
        return jsonify({"error": "Impossible de distinguer les gros écarts."}), 422

    df["year"]       = pd.to_datetime(df["date"], errors="coerce").dt.year.fillna(0)
    df["month"]      = pd.to_datetime(df["date"], errors="coerce").dt.month.fillna(0)
    df["section_id"] = df["section"].astype("category").cat.codes
    df["flux_id"]    = df["flux"].astype("category").cat.codes
    df["abs_pct"]    = df["x"].abs()

    features = ["x", "abs_pct", "year", "month", "section_id", "flux_id"]
    labels_fr = {
        "x":          "Écart (%)",
        "abs_pct":    "Écart absolu (%)",
        "year":       "Année",
        "month":      "Mois",
        "section_id": "Filiale / Section",
        "flux_id":    "Flux",
    }

    X = df[features].fillna(0).values
    y = df["target"].values

    sel    = SelectKBest(score_func=f_classif, k=len(features))
    sel.fit(X, y)
    scores = sel.scores_

    feat_imp = sorted(
        [{"feature": f, "label": labels_fr[f], "score": round(float(s), 2)}
         for f, s in zip(features, scores) if math.isfinite(s)],
        key=lambda r: r["score"], reverse=True
    )

    top = feat_imp[0] if feat_imp else {}
    n_gros = int(df["target"].sum())

    return jsonify({
        "features":   feat_imp,
        "seuil":      round(seuil, 0),
        "n_gros":     n_gros,
        "n_total":    len(df),
        "top_feature": top.get("label", "N/A"),
        "top_score":   top.get("score", 0),
    })
