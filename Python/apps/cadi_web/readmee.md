# CADI Web

CADI Web est une application locale composee d'un frontend React/Vite et d'un backend FastAPI.
Elle sert d'interface moderne autour des traitements Python existants du package `vte`, en particulier le module IRIS.

L'objectif de l'application est de centraliser les outils utiles a la gestion des donnees de formation :

- exports IRIS ;
- evaluations stagiaires ;
- bilans ;
- formations ;
- parametres utilisateur ;
- messagerie ou historique interne ;
- pages d'aide et mentions du footer.

> Etat actuel important : le module IRIS est le plus avance et reellement branche au backend. Certains autres modules existent cote frontend mais leurs routeurs backend ne sont pas encore implementes.

---

## Vue d'ensemble

```text
cadi/Python/apps/cadi_web/
|
|-- back/                  # Backend FastAPI local
|-- front/                 # Frontend React + Vite + TypeScript
|-- requirements.txt       # Dependances Python du backend
|-- README.md              # Documentation courte historique
`-- readme                 # Documentation detaillee de l'architecture
```

Architecture generale :

```text
Utilisateur
  |
  v
Frontend React
  |
  v
API FastAPI locale
  |
  v
Code metier Python vte
  |
  v
Fichiers Excel / Word / dossiers IRIS
```

---

## Architecture Detaillee Commentee

Cette section decrit le projet sous forme d'arborescence commentee.
Elle sert a comprendre rapidement quel fichier fait quoi.

```text
cadi/Python/apps/cadi_web/                 # Dossier racine de l'application web CADI
|
|-- back/                                  # Backend local FastAPI
|   |                                      # Role : recevoir les demandes React,
|   |                                      # appeler le code Python metier `vte`,
|   |                                      # lire/ecrire/ouvrir des fichiers locaux.
|   |
|   |-- main.py                            # Point d'entree du backend.
|   |                                      # Cree l'application FastAPI.
|   |                                      # Configure le CORS pour le front Vite.
|   |                                      # Monte les routeurs disponibles.
|   |                                      # Expose aussi /api/health et /api/dashboard/.
|   |
|   `-- routers/                           # Dossier des routes API.
|       |
|       |-- __init__.py                    # Rend le dossier importable en Python.
|       |
|       |-- iris.py                        # Routeur principal actuellement fonctionnel.
|       |                                  # Gere :
|       |                                  # - affichage des dumps IRIS ;
|       |                                  # - mise a jour depuis GED/local ;
|       |                                  # - validation des fichiers ;
|       |                                  # - selection manuelle via Tkinter ;
|       |                                  # - traitement auto ;
|       |                                  # - traitement manuel ;
|       |                                  # - suivi des jobs ;
|       |                                  # - ouverture securisee de fichiers/dossiers.
|       |
|       |-- bilans.py                      # Fichier reserve pour le futur routeur Bilans.
|       |                                  # Actuellement : commentaire/stub seulement.
|       |                                  # A implementer si on veut servir /api/bilans/...
|       |
|       |-- evalstat.py                    # Fichier reserve pour le futur routeur EvalStat.
|       |                                  # Actuellement : commentaire/stub seulement.
|       |                                  # A implementer si on veut servir /api/evalstat/...
|       |
|       `-- fdc.py                         # Fichier reserve pour les fiches de couts.
|                                          # Actuellement : commentaire/stub seulement.
|
|-- front/                                 # Frontend React + Vite + TypeScript
|   |                                      # Role : interface utilisateur, dashboard,
|   |                                      # navigation, formulaires, appels API.
|   |
|   |-- index.html                         # Page HTML racine.
|   |                                      # Vite injecte l'application React dedans.
|   |
|   |-- package.json                       # Dependances frontend et scripts npm.
|   |                                      # Scripts importants :
|   |                                      # - npm run dev
|   |                                      # - npm run typecheck
|   |                                      # - npm run build
|   |
|   |-- vite.config.ts                     # Configuration Vite.
|   |                                      # Sert au serveur de dev React.
|   |
|   |-- tsconfig.json                      # Configuration TypeScript generale.
|   |-- tsconfig.app.json                  # Configuration TypeScript de l'application.
|   |
|   |-- public/                            # Ressources publiques.
|   |   |
|   |   `-- image.png                      # Logo CADI.
|   |                                      # Utilise dans Header et pages publiques.
|   |
|   `-- src/                               # Code source React principal.
|       |
|       |-- main.tsx                       # Point d'entree React.
|       |                                  # Monte <App /> dans le DOM.
|       |
|       |-- App.tsx                        # Orchestrateur principal.
|       |                                  # Gere :
|       |                                  # - l'authentification locale ;
|       |                                  # - l'ecran public actif ;
|       |                                  # - l'ecran dashboard actif ;
|       |                                  # - la sidebar reduite ou ouverte ;
|       |                                  # - les pages info du footer ;
|       |                                  # - l'utilisateur courant ;
|       |                                  # - la deconnexion.
|       |
|       |
|       |-- lib/                           # Configuration et helpers partages.
|       |   |
|       |   `-- api.ts                     # Base URL de l'API backend.
|       |                                  # Valeur actuelle :
|       |                                  # http://127.0.0.1:8001/api
|       |
|       |-- components/                    # Composants reutilisables.
|       |   |
|       |   |-- Sidebar.tsx                # Navigation principale du dashboard.
|       |   |                              # Contient les pages metier :
|       |   |                              # Accueil, IRIS, EvalStat, Bilans,
|       |   |                              # Formations, Parametres, Messagerie.
|       |   |
|       |   |-- Header.tsx                 # Barre superieure fixe.
|       |   |                              # Affiche :
|       |   |                              # - icone de la page active ;
|       |   |                              # - titre de la page ;
|       |   |                              # - logo CADI a droite.
|       |   |
|       |   |-- Footer.tsx                 # Footer compact du dashboard.
|       |   |                              # Ne contient pas la navigation metier.
|       |   |                              # Contient seulement :
|       |   |                              # Conditions, Confidentialite,
|       |   |                              # Mentions legales, FAQ, Contact.
|       |   |
|       |   `-- ui.tsx                     # Mini design system local.
|       |                                  # Contient :
|       |                                  # Badge, Card, CardHeader, Field,
|       |                                  # Input, Select, Btn, PageHeader,
|       |                                  # THead, TRow, TD, UploadZone.
|       |
|       `-- pages/                         # Pages completes de l'application.
|           |
|           |-- Index.tsx                  # Page publique d'accueil.
|           |                              # Affiche la presentation CADI avant connexion.
|           |
|           |-- Login.tsx                  # Page de connexion.
|           |-- Register.tsx               # Page d'inscription.
|           |-- ForgotPassword.tsx         # Page mot de passe oublie.
|           |
|           |-- Accueil.tsx                # Tableau de bord.
|           |                              # Appelle /api/dashboard/.
|           |                              # Peut fonctionner avec donnees statiques.
|           |
|           |-- IRIS.tsx                   # Page metier IRIS.
|           |                              # Page frontend la plus branchee au backend.
|           |                              # Appelle notamment :
|           |                              # - /api/iris/dumps
|           |                              # - /api/iris/select-manual-files
|           |                              # - /api/iris/validate-structure
|           |                              # - /api/iris/traiter
|           |                              # - /api/iris/traiter-manuel-chemins
|           |                              # - /api/iris/job/{job_id}
|           |
|           |-- EvalStat.tsx               # Page frontend EvalStat.
|           |                              # Appelle des endpoints /api/evalstat/...
|           |                              # Backend correspondant pas encore implemente.
|           |
|           |-- Bilans.tsx                 # Page frontend Bilans.
|           |                              # Appelle des endpoints /api/bilans/...
|           |                              # Backend correspondant pas encore implemente.
|           |
|           |-- Formations.tsx             # Page frontend Formations.
|           |                              # Appelle /api/formations/.
|           |                              # Route backend absente actuellement.
|           |
|           |-- Parametres.tsx             # Page parametres utilisateur.
|           |-- Messagerie.tsx             # Page messagerie / historique.
|           |
|           `-- InfoPages.tsx              # Pages d'information du footer.
|                                          # Contient dans un seul fichier :
|                                          # - TermsPage ;
|                                          # - PrivacyPage ;
|                                          # - LegalPage ;
|                                          # - FAQPage ;
|                                          # - ContactPage.
|
|-- requirements.txt                       # Dependances Python du backend.
|-- README.md                              # README court.
`-- readme                                 # Documentation detaillee actuelle.
```

Lien avec le code metier historique :

```text
cadi/Python/
|
|-- apps/cadi_web/                         # Application web
|   |
|   |-- front/src/pages/IRIS.tsx           # Interface utilisateur IRIS
|   `-- back/routers/iris.py               # API HTTP IRIS
|
`-- vte/                                   # Package metier existant
    |
    |-- domain/
    |   `-- iris.py                        # Logique metier IRIS principale
    |                                      # Classes IRIS et IRIS_natif.
    |
    |-- services/
    |   `-- iris_dumps_services.py         # Recherche, copie, archivage
    |                                      # des dumps IRIS.
    |
    `-- utils/
        `-- office.py                      # Fonctions liees a Office,
                                           # Excel, Word, COM, fichiers.
```

---

## Backend FastAPI

```text
back/
|
|-- main.py
`-- routers/
    |
    |-- __init__.py
    |-- iris.py
    |-- bilans.py
    |-- evalstat.py
    `-- fdc.py
```

### `back/main.py`

Point d'entree du backend.

Il fait actuellement :

- creation de l'application FastAPI ;
- configuration CORS pour Vite ;
- ajout de `/api/health` ;
- ajout de `/api/dashboard/` avec donnees statiques minimales ;
- montage du routeur IRIS sous `/api/iris`.

Routeurs reellement montes actuellement :

```python
from .routers import iris
app.include_router(iris.router, prefix="/api")
```

Donc aujourd'hui, le backend expose vraiment :

```text
GET  /api/health
GET  /api/dashboard/
GET  /api/iris/dumps
POST /api/iris/update-dumps
POST /api/iris/validate-structure
POST /api/iris/open-path
POST /api/iris/select-manual-files
POST /api/iris/traiter
POST /api/iris/traiter-manuel-chemins
GET  /api/iris/job/{job_id}
DEL  /api/iris/job/{job_id}
```

### `back/routers/iris.py`

Routeur backend principal et le plus complet.

Son role :

- lire les dumps IRIS locaux ;
- afficher l'etat des fichiers par code ;
- ouvrir un fichier ou dossier autorise via Windows ;
- ouvrir une fenetre Tkinter pour la selection manuelle ;
- verifier les fichiers avant traitement ;
- lancer un traitement auto ou manuel ;
- suivre un traitement avec un `job_id` ;
- renvoyer au frontend les fichiers generes.

Codes IRIS geres :

```text
R04110 -> Sessions
R0304  -> Formations
R04301 -> Ventes
R04500 -> Inscriptions
```

Le routeur IRIS ne contient pas toute la logique metier lourde.
Il sert surtout de facade HTTP entre React et le package `vte`.

Il appelle notamment :

```text
vte/domain/iris.py
vte/services/iris_dumps_services.py
vte/utils/office.py
```

### `back/routers/bilans.py`, `evalstat.py`, `fdc.py`

Ces fichiers existent, mais ils ne sont pas encore de vrais routeurs FastAPI.
Ils contiennent actuellement des commentaires de description.

Consequence :

- le frontend peut deja avoir des pages Bilans ou EvalStat ;
- mais les endpoints `/api/bilans/...`, `/api/evalstat/...`, `/api/formations/...` ne sont pas encore servis par le backend actuel ;
- des erreurs `404 Not Found` sont donc normales tant que ces routeurs ne sont pas implementes et montes dans `main.py`.

---

## Frontend React

```text
front/
|
|-- index.html
|-- package.json
|-- vite.config.ts
|-- tsconfig.json
|-- tsconfig.app.json
|-- public/
|   `-- image.png
`-- src/
    |
    |-- main.tsx
    |-- App.tsx
    |-- styles.css
    |
    |-- lib/
    |   `-- api.ts
    |
    |-- components/
    |   |-- Sidebar.tsx
    |   |-- Header.tsx
    |   |-- Footer.tsx
    |   `-- ui.tsx
    |
    `-- pages/
        |-- Index.tsx
        |-- Login.tsx
        |-- Register.tsx
        |-- ForgotPassword.tsx
        |-- Accueil.tsx
        |-- IRIS.tsx
        |-- EvalStat.tsx
        |-- Bilans.tsx
        |-- Formations.tsx
        |-- Parametres.tsx
        |-- Messagerie.tsx
        `-- InfoPages.tsx
```

### `src/main.tsx`

Point d'entree React.

Il monte le composant principal `App` dans la page HTML.

### `src/App.tsx`

C'est le coeur de l'application frontend.

Il gere :

- l'etat de connexion ;
- l'ecran public actif avant connexion ;
- l'ecran metier actif apres connexion ;
- la page info active depuis le footer ;
- l'ouverture ou reduction de la sidebar ;
- l'utilisateur courant ;
- le toast de deconnexion.

Etats principaux :

```ts
auth       // index, login, register, forgot
loggedIn   // true / false
screen     // page dashboard active
infoScreen // page info active ou null
collapsed  // sidebar reduite ou non
```

Navigation metier :

```ts
type Screen =
  | 'accueil'
  | 'iris'
  | 'evalstat'
  | 'bilans'
  | 'formations'
  | 'parametres'
  | 'messagerie';
```

Navigation info :

```ts
type InfoPage = 'terms' | 'privacy' | 'legal' | 'faq' | 'contact';
```

Regle importante :

- quand on clique dans la sidebar, `infoScreen` est remis a `null` ;
- quand on clique dans le footer, `infoScreen` prend une valeur ;
- `renderScreen()` affiche soit une page metier, soit une page info.

---

## Layout Dashboard

Quand l'utilisateur est connecte, la structure est :

```tsx
<Sidebar />
<Header />
<main>
  {renderScreen()}
</main>
<Footer />
```

### `components/Sidebar.tsx`

Menu lateral du dashboard.

Il contient la navigation metier :

```text
Accueil
IRIS
EvalStat
Bilans
Formations
Parametres
Messagerie
```

Il gere aussi :

- l'etat actif ;
- le mode reduit ;
- les badges ;
- la zone utilisateur ;
- le bouton deconnexion.

Largeurs :

```ts
SIDEBAR_W = 220
SIDEBAR_W_COLLAPSED = 64
```

### `components/Header.tsx`

Barre superieure du dashboard.

Elle affiche :

- une icone associee a la page active ;
- le titre ;
- le logo CADI a droite.

Le titre vient de `App.tsx`.
Si `infoScreen` est actif, le header affiche le titre de la page info.
Sinon, il affiche le titre de la page metier.

Exemples :

```text
IRIS -> icone Database
Conditions -> icone FileText
Confidentialite -> icone ShieldCheck
Mentions legales -> icone Scale
FAQ -> icone HelpCircle
Contact -> icone Mail
```

### `components/Footer.tsx`

Footer compact du dashboard.

Il ne contient pas la navigation metier, car cette navigation est deja dans la sidebar.

Il contient uniquement les pages d'aide et mentions :

```text
Conditions
Confidentialite
Mentions legales
FAQ
Contact
```

Chaque icone appelle :

```ts
onNavigateInfo(page)
```

Donc le footer modifie `infoScreen` dans `App.tsx`.

### `components/ui.tsx`

Mini design-system local.

Il contient les composants reutilisables :

```text
Badge
Card
CardHeader
SectionLabel
Field
Input
Select
Btn
PageHeader
THead
TRow
TD
UploadZone
```

Ces composants evitent de repeter trop de styles dans les pages metier.

---

## Pages Frontend

### Pages publiques

```text
Index.tsx
Login.tsx
Register.tsx
ForgotPassword.tsx
```

Ces pages sont affichees quand `loggedIn = false`.

### Pages dashboard metier

```text
Accueil.tsx
IRIS.tsx
EvalStat.tsx
Bilans.tsx
Formations.tsx
Parametres.tsx
Messagerie.tsx
```

Ces pages sont affichees quand `loggedIn = true` et `infoScreen = null`.

### Pages info

```text
InfoPages.tsx
```

Ce fichier contient plusieurs pages :

```text
TermsPage    -> Conditions
PrivacyPage  -> Confidentialite
LegalPage    -> Mentions legales
FAQPage      -> FAQ
ContactPage  -> Contact
```

Il contient aussi des composants internes :

```text
InfoHero       -> titre + icone principale
InfoCard       -> carte de contenu
AccordionCard  -> carte de questions/reponses
```

Ces pages ne sont pas des fichiers HTML statiques.
Elles sont rendues directement par React dans le dashboard.

---

## Couche API Frontend

```text
front/src/lib/api.ts
```

Contenu actuel :

```ts
export const API_BASE_URL = 'http://127.0.0.1:8001/api';
```

Toutes les pages qui appellent le backend utilisent cette base.

Exemple IRIS :

```ts
fetch(`${API_BASE_URL}/iris/dumps`)
```

---

## Flux IRIS Detaille

### Chargement de la page IRIS

```text
IRIS.tsx
  |
  v
GET http://127.0.0.1:8001/api/iris/dumps
  |
  v
back/routers/iris.py
  |
  v
Lecture des dossiers configures dans vte/core/config.py
  |
  v
Retour JSON vers React
  |
  v
Affichage du tableau IRIS
```

### Selection manuelle

```text
Utilisateur clique "manuel"
  |
  v
IRIS.tsx appelle /api/iris/select-manual-files
  |
  v
FastAPI ouvre une fenetre Tkinter
  |
  v
Utilisateur choisit des fichiers Excel
  |
  v
FastAPI renvoie les chemins au frontend
```

### Traitement manuel

```text
IRIS.tsx
  |
  v
POST /api/iris/traiter-manuel-chemins
  |
  v
Creation d'un job_id
  |
  v
Traitement en arriere-plan
  |
  v
IRIS.tsx poll /api/iris/job/{job_id}
  |
  v
Affichage progression + resultat
```

### Traitement automatique

```text
IRIS.tsx
  |
  v
POST /api/iris/validate-structure
  |
  v
POST /api/iris/traiter
  |
  v
Creation d'un job_id
  |
  v
Traitement IRIS via vte/domain/iris.py
  |
  v
Suivi avec GET /api/iris/job/{job_id}
```

---

## Dossiers et configuration metier IRIS

Le backend lit les dossiers configures dans le package `vte`, notamment :

```text
vte/core/config.py
```

Variables importantes :

```text
REPERTOIRE_EXTRACT_IRIS_LOCAL
REPERTOIRE_EXTRACT_IRIS_GED
REPERTOIRE_EXCEL_IRIS_OUTPUT
```

Ces chemins servent a :

- trouver les extracts IRIS locaux ;
- copier certains dumps depuis la GED ;
- lister les fichiers complets generes ;
- ouvrir les fichiers autorises depuis l'interface.

---

## Etat Reel Des Modules

| Module | Frontend | Backend | Etat |
| --- | --- | --- | --- |
| Accueil | Oui | `/api/dashboard/` minimal | Fonctionnel en donnees statiques |
| IRIS | Oui | Oui | Module principal fonctionnel |
| EvalStat | Oui | Stub seulement | Backend a implementer |
| Bilans | Oui | Stub seulement | Backend a implementer |
| Formations | Oui | Route absente | Backend a implementer |
| Parametres | Oui | Pas de route dediee | Fonctionnement local/front |
| Messagerie | Oui | Route absente ou non montee | Backend a implementer |
| Pages info | Oui | Pas besoin backend | Fonctionnel cote React |

---

## Lancer Le Backend

Depuis le dossier :

```text
cadi/Python
```

Commande :

```powershell
py -m pip install -r apps/cadi_web/requirements.txt
py -m uvicorn apps.cadi_web.back.main:app --reload --host 127.0.0.1 --port 8001
```

URLs utiles :

```text
API health : http://127.0.0.1:8001/api/health
Swagger    : http://127.0.0.1:8001/docs
IRIS dumps : http://127.0.0.1:8001/api/iris/dumps
```

---

## Lancer Le Frontend

Depuis :

```text
cadi/Python/apps/cadi_web/front
```

Si `npm` est disponible :

```powershell
npm install
npm run dev
```

Puis ouvrir :

```text
http://127.0.0.1:5173/
```

Si Node/NPM ne sont pas dans le PATH, utiliser le Node portable :

```powershell
$env:Path = "$PWD\.tools\node-v22.22.3-win-x64;" + $env:Path
.\.tools\node-v22.22.3-win-x64\npm.cmd install
.\.tools\node-v22.22.3-win-x64\npm.cmd run dev -- --host 127.0.0.1
```

---

## Verifications Rapides

### Frontend

Depuis `front/` :

```powershell
$env:Path = "$PWD\.tools\node-v22.22.3-win-x64;" + $env:Path
npm run typecheck
```

### Backend

Depuis `cadi/Python` :

```powershell
py -c "from apps.cadi_web.back.main import app; print(app.title)"
```

### API

```powershell
Invoke-RestMethod http://127.0.0.1:8001/api/health
Invoke-RestMethod http://127.0.0.1:8001/api/iris/dumps
```

---

## Points A Surveiller

### 1. Routes backend manquantes

Le frontend appelle deja certains endpoints non implementes.

Exemples :

```text
/api/evalstat/...
/api/bilans/...
/api/formations/...
/api/messagerie/...
```

Tant que les routeurs correspondants ne sont pas implementes et montes dans `main.py`, ces appels peuvent retourner `404`.

### 2. IRIS utilise des chemins locaux Windows

Certaines actions ouvrent des fichiers ou dossiers via le backend.
Cela suppose :

- un environnement Windows ;
- des chemins configures correctement ;
- des droits d'acces aux dossiers reseau ;
- Office/Excel/Word disponibles selon les traitements.

### 3. Jobs en memoire

Les traitements IRIS sont suivis via `_jobs` dans `routers/iris.py`.
Ces jobs sont stockes en memoire.

Consequence :

- refresh frontend : possible si le serveur reste lance ;
- redemarrage backend : les jobs en cours sont perdus.

---

## Prochaines Evolutions Logiques

1. Implementer les routeurs backend manquants :
   - `bilans.py`
   - `evalstat.py`
   - `formations.py`
   - `messagerie.py`

2. Monter ces routeurs dans `back/main.py`.

3. Centraliser les schemas API avec Pydantic.

4. Ajouter une vraie page de statut backend :
   - backend lance ;
   - dossiers IRIS accessibles ;
   - Office disponible ;
   - routes actives.

5. Ajouter des tests rapides backend pour IRIS :
   - normalisation des codes ;
   - validation des fichiers ;
   - conversion `file://` vers chemin Windows ;
   - securite d'ouverture des chemins.
