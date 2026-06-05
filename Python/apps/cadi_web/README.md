# CADI Web

Application web CADI avec un front React/Vite et un back FastAPI qui appelle le code Python `vte`.

## Structure

```text
cadi_web/
├── front/            # Interface React + TypeScript + Vite
├── back/             # API FastAPI
└── requirements.txt  # Dependances Python du back
```

## Lancer le back

Depuis le dossier `cadi/Python` :

```powershell
py -m pip install -r apps/cadi_web/requirements.txt
py -m uvicorn apps.cadi_web.back.main:app --reload --host 127.0.0.1 --port 8001
```

Adresses utiles :

```text
API        : http://127.0.0.1:8001/api/health
Swagger    : http://127.0.0.1:8001/docs
IRIS dumps : http://127.0.0.1:8001/api/iris/dumps
```

## Lancer le front

Depuis le dossier `cadi/Python/apps/cadi_web/front` :

```powershell
npm install
npm run dev
```

Puis ouvrir :

```text
http://127.0.0.1:5173/
```

Si `node` et `npm` ne sont pas dans le `PATH`, utiliser le Node portable :

```powershell
$env:Path = "$PWD\.tools\node-v22.22.3-win-x64;" + $env:Path
.\.tools\node-v22.22.3-win-x64\npm.cmd install
.\.tools\node-v22.22.3-win-x64\npm.cmd run dev -- --host 127.0.0.1
```

## Connexion IRIS front / back

La page IRIS appelle maintenant ces routes :

```text
GET  /api/iris/dumps
POST /api/iris/update-dumps
POST /api/iris/validate-structure
POST /api/iris/run-treatment
```

Le back lit les vrais dossiers configures dans `vte/core/config.py` :

```text
REPERTOIRE_EXTRACT_IRIS_LOCAL
REPERTOIRE_EXTRACT_IRIS_GED
REPERTOIRE_EXCEL_IRIS_OUTPUT
```

## Tests rapides

Front :

```powershell
.\.tools\node-v22.22.3-win-x64\node.exe .\node_modules\typescript\bin\tsc --noEmit -p tsconfig.app.json
```

Back :

```powershell
py -c "from apps.cadi_web.back.main import app; print(app.title)"
```
