
from __future__ import annotations

import threading
import sys
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware


# =============================================================================
# Configuration du chemin Python
# =============================================================================

# Le back est lance depuis cadi/Python. Cette securite permet aussi de l'importer
# depuis un autre dossier tout en gardant l'acces au package metier `vte`.
# Exemple attendu:
#   py -m uvicorn apps.cadi_web.back.main:app --reload --port 8001
#
# Si le serveur est lance depuis `apps/cadi_web/back`, Python ne trouve pas
# automatiquement `vte`. On ajoute donc `cadi/Python` au sys.path avant d'importer
# les routeurs qui appellent le code metier.
PYTHON_ROOT = Path(__file__).resolve().parents[3]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from .routers import evalstat, iris


# =============================================================================
# Application FastAPI
# =============================================================================

app = FastAPI(
    title="CADI Web API",
    description="API locale entre l'interface React CADI et le code Python VTE.",
    version="0.1.0",
)


# =============================================================================
# CORS local pour Vite / React
# =============================================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:5173",
        "http://localhost:5173",
        "http://127.0.0.1:5174",
        "http://localhost:5174",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =============================================================================
# Routes metier
# =============================================================================

# Tous les routeurs metier sont montes sous `/api`.
# Le routeur IRIS expose donc:
#   /api/iris/dumps
#   /api/iris/validate-structure
#   /api/iris/traiter
#   /api/iris/job/{job_id}
app.include_router(iris.router, prefix="/api")
app.include_router(evalstat.router, prefix="/api")


@app.on_event("startup")
def warm_backend_caches() -> None:
    """Lance les chargements longs en arriere-plan au demarrage de CADI."""
    threading.Thread(target=evalstat.warm_evalstat_cache, daemon=True).start()


@app.get("/api/health")
def health() -> dict[str, str]:
    """Retour minimal pour verifier que le serveur back repond."""
    return {"status": "ok"}


@app.get("/api/dashboard/")
def dashboard() -> dict[str, object]:
    """Donnees minimales pour eviter le fallback erreur sur la page d'accueil."""
    # Pour l'instant ces donnees sont statiques: elles permettent au front de
    # fonctionner pendant que les vrais endpoints dashboard ne sont pas encore
    # branches sur les services metier.
    return {
        "kpis": {
            "weekly_treatments": 18,
            "active_sessions": 7,
            "unresolved_errors": 0,
        },
        "recent_activity": [
            {
                "date": "01/06/2026",
                "module": "IRIS",
                "formation": "Exports locaux",
                "statut": "ok",
                "fichier": "Extracts originaux",
            }
        ],
        "alerts": [],
    }
