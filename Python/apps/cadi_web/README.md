# CADI Web

Application web CADI avec un front React/Vite et un back FastAPI qui appelle le
code Python `vte`.

## Structure

```text
cadi_web/
|-- front/            # Interface React + TypeScript + Vite
|-- back/             # API FastAPI
`-- requirements.txt  # Dependances Python du back
```

## Installation Une Fois

PowerShell :

```powershell
cd "C:\Users\MG288147\Desktop\cadi_cea\cadi\Python"
py -m pip install -r apps/cadi_web/requirements.txt

cd "C:\Users\MG288147\Desktop\cadi_cea\cadi\Python\apps\cadi_web\front"
$env:Path = "$PWD\.tools\node-v22.22.3-win-x64;" + $env:Path
npm install
```

Ce bloc installe seulement les dependances. Il ne lance pas l'application.

Ne pas lancer `npm audit fix --force` sur ce projet : cette commande peut
installer une version de Vite trop recente pour la configuration actuelle.

## Tout Arreter

Si un port est deja occupe, liberer le back et le front :

```powershell
Get-NetTCPConnection -LocalPort 8010,5173,5174 -ErrorAction SilentlyContinue | ForEach-Object {
  if ($_.OwningProcess -ne 0) {
  Stop-Process -Id $_.OwningProcess -Force
  }
}
```

## Lancer Le Back

PowerShell 1 :

```powershell
cd "C:\Users\MG288147\Desktop\cadi_cea\cadi\Python"
py -m uvicorn apps.cadi_web.back.main:app --host 127.0.0.1 --port 8010
```

Garder ce terminal ouvert.

Verifier :

```text
http://127.0.0.1:8010/api/health
```

## Lancer Le Front

PowerShell 2 :

```powershell
cd "C:\Users\MG288147\Desktop\cadi_cea\cadi\Python\apps\cadi_web\front"
$env:Path = "$PWD\.tools\node-v22.22.3-win-x64;" + $env:Path
npm run dev
```

Laisser ce terminal ouvert.

Ouvrir l'URL affichee par Vite, par exemple :

```text
http://localhost:5173/
```

ou :

```text
http://localhost:5174/
```

Sans `--reload` sur le back : les jobs IRIS sont stockes en memoire (`_jobs`),
un redemarrage du process en perdrait le suivi pendant un traitement en cours.

URLs utiles :

```text
Front       : URL affichee par Vite, souvent http://localhost:5173/ ou http://localhost:5174/
API         : http://127.0.0.1:8010/api/health
Swagger     : http://127.0.0.1:8010/docs
IRIS dumps  : http://127.0.0.1:8010/api/iris/dumps
```

## Connexion IRIS Front / Back

La page IRIS appelle ces routes :

```text
GET    /api/iris/dumps
POST   /api/iris/update-dumps
POST   /api/iris/validate-structure
POST   /api/iris/select-manual-files
POST   /api/iris/traiter
POST   /api/iris/traiter-manuel-chemins
GET    /api/iris/job/{job_id}
DELETE /api/iris/job/{job_id}
POST   /api/iris/open-path
```

Le back lit les dossiers configures dans `vte/core/config.py` :

```text
REPERTOIRE_EXTRACT_IRIS_LOCAL
REPERTOIRE_EXTRACT_IRIS_GED
REPERTOIRE_EXCEL_IRIS_OUTPUT
```

## EvalStat

Etat actuel :

```text
Mode Par CSV              : interface branchee au backend
Selection CSV             : ouvre une fenetre Windows cote backend
Verification CSV          : controle extension, existence, taille, code IRIS dans le nom
Traitement CSV            : bloque tant que l'analyse CSV n'est pas valide
Mode Par codes IRIS       : verification des codes branchee, traitement multiple a faire
Mode EvalStat filtre : lit le dernier extract IRIS Sessions, l'utilisateur choisit une formation et une annee, le back scanne FORMATIONS_C pour trouver les CSV stagiaires, puis seules les sessions correspondantes sont affichees.

Parcours utilisateur EvalStat :

```text
1. Choisir une formation
2. Choisir une annee ou toutes les annees
3. Choisir un etat CSV : tous, prets, CSV manquant, CSV ambigu
4. Cliquer Rechercher
5. Cocher les sessions pretes
6. Traiter les sessions cochees
```
```

Routes backend :

```text
GET  /api/evalstat/health
POST /api/evalstat/select-csv
POST /api/evalstat/verify-csv
POST /api/evalstat/process-csv
POST /api/evalstat/verify-codes
GET  /api/evalstat/config
POST /api/evalstat/process-config-csv
POST /api/evalstat/open-path
```

## Verifications Rapides

Ces commandes ne lancent pas l'application. Elles servent seulement a verifier
que le back et le front compilent correctement.

Front :

```powershell
cd "C:\Users\MG288147\Desktop\cadi_cea\cadi\Python\apps\cadi_web\front"
.\.tools\node-v22.22.3-win-x64\node.exe .\node_modules\typescript\bin\tsc --noEmit -p tsconfig.app.json
```

Back :

```powershell
cd "C:\Users\MG288147\Desktop\cadi_cea\cadi\Python"
py -c "from apps.cadi_web.back.main import app; print(app.title)"
```
