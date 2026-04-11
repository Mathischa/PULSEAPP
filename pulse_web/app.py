# -*- coding: utf-8 -*-
"""
app.py — Point d'entrée de l'application PULSE Web (Flask).

Lancer avec :
    cd pulse_web && python app.py
Puis ouvrir http://localhost:5000
"""
from __future__ import annotations
import os, sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, render_template, jsonify

# Flag global de disponibilité des données
_data_ready = False
_data_error = None


def create_app() -> Flask:
    global _data_ready, _data_error
    app = Flask(__name__)

    import sys
    import threading
    sys.stdout.flush()
    sys.stderr.flush()

    from pulse_v2.data.cache import init_full_load, CACHE, TOKENS, sections

    print("[PULSE] Initialisation du cache…", flush=True)
    sys.stdout.flush()

    def load_data_async():
        global _data_ready, _data_error
        try:
            print("[PULSE-ASYNC] Démarrage chargement données réelles...", flush=True)
            init_full_load()
            _data_ready = True
            print(f"[PULSE-ASYNC] ✓ Cache prêt: {len(CACHE)} entrées, {sum(len(v) for v in TOKENS.values())} flux", flush=True)
        except Exception as e:
            _data_error = str(e)
            print(f"[PULSE-ASYNC] ERREUR: {e}", flush=True)
            import traceback
            traceback.print_exc()

    # Charge les données en background (non-bloquant)
    bg_thread = threading.Thread(target=load_data_async, daemon=True)
    bg_thread.start()

    # Pour la première requête, vérifier si les données sont chargées
    # Sinon, les API attendront ou retourneront des données partielles
    print("[PULSE] Server prêt pour les requêtes", flush=True)
    sys.stdout.flush()

    @app.route("/api/status")
    def api_status():
        # Vérification double : flag global + taille du cache (plus fiable en mode debug)
        from pulse_v2.data.cache import CACHE as _C
        ready = _data_ready or len(_C) > 0
        return jsonify({"ready": ready, "cache_size": len(_C)})

    from api.accueil    import bp as bp_accueil
    from api.ecarts     import bp as bp_ecarts
    from api.catalogue  import bp as bp_catalogue
    from api.tendance   import bp as bp_tendance
    from api.repartition import bp as bp_repartition
    from api.repartition_flux import bp as bp_repartition_flux
    from api.visualisation import bp as bp_visualisation
    from api.visualisation_flux import bp as bp_visualisation_flux
    from api.prevision_repartition import bp as bp_prevision_repartition
    from api.ml_ecarts import bp as bp_ml_ecarts
    from api.heatmap   import bp as bp_heatmap
    from api.heatmap_ecarts import bp as bp_heatmap_ecarts
    from api.import_profils import bp as bp_import_profils

    for bp in (bp_accueil, bp_ecarts, bp_catalogue, bp_tendance, bp_repartition, bp_repartition_flux, bp_visualisation, bp_visualisation_flux, bp_prevision_repartition, bp_ml_ecarts, bp_heatmap, bp_heatmap_ecarts, bp_import_profils):
        app.register_blueprint(bp)

    @app.route("/")
    def page_accueil():
        return render_template("accueil.html")

    @app.route("/ecarts")
    def page_ecarts():
        return render_template("ecarts.html")

    @app.route("/tendance")
    def page_tendance():
        return render_template("tendance.html")

    @app.route("/repartition")
    def page_repartition():
        return render_template("repartition.html")

    @app.route("/repartition_flux")
    def page_repartition_flux():
        return render_template("repartition_flux.html")

    @app.route("/visualisation")
    def page_visualisation():
        return render_template("visualisation.html")

    @app.route("/visualisation_flux")
    def page_visualisation_flux():
        return render_template("visualisation_flux.html")

    @app.route("/prevision_repartition")
    def page_prevision_repartition():
        return render_template("prevision_repartition.html")

    @app.route("/ml_ecarts")
    def page_ml_ecarts():
        return render_template("ml_ecarts.html")

    @app.route("/heatmap")
    def page_heatmap():
        return render_template("heatmap.html")

    @app.route("/heatmap_ecarts")
    def page_heatmap_ecarts():
        return render_template("heatmap_ecarts.html")

    return app


# Créer l'instance app pour production (gunicorn)
app = create_app()

if __name__ == "__main__":
    app.run(debug=True, port=5000, use_reloader=False)
