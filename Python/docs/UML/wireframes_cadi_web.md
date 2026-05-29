# Wireframes CADI Web adaptes aux UML

Ce dossier complete les diagrammes UML existants avec des maquettes basse fidelite.
Les wireframes sont ecrits en PlantUML/Salt pour rester dans le meme ecosysteme que
les fichiers `use_case_*.puml` et `uml_*.puml`.

## Correspondance UML vers ecrans

| Wireframe | Cas d'utilisation UML | Classes / services concernes | Objectif UI |
| --- | --- | --- | --- |
| `wireframe_00_dashboard.puml` | `use_case_general.puml` | Modules EvalStat, IRIS, Bilans, FdC, Parametres | Donner une vue d'ensemble, les raccourcis et l'etat des traitements. |
| `wireframe_01_iris.puml` | `use_case_iris.puml`, `uml_09_sequence_iris.puml` | `IRIS`, `IRIS_natif`, `ConfigExportIRIS`, `InfosExportsIRIS` | Importer un export IRIS, choisir son type, verifier puis produire le fichier traite. |
| `wireframe_02_evalstat.puml` | `use_case_evalstat.puml`, `uml_03_sequence_evalstat.puml` | Services EvalStat et fichiers Excel d'evaluation | Saisir trigramme/code IRIS, selectionner les fiches, lancer le traitement. |
| `wireframe_03_bilans.puml` | `use_case_bilans.puml`, `uml_04_sequence_bilansessions.puml`, `uml_06_bilans_detail.puml` | Bilans sessions/formation, IRIS traite, EvalStat, FdC, Specs | Verifier les sources et generer le document Word. |
| `wireframe_04_fdc_parametres.puml` | `use_case_fdc.puml`, `use_case_parametres.puml`, `uml_10_sequence_fdc.puml` | Fiche de couts, configuration, roles | Lire une FdC et administrer chemins, modeles et roles. |

## Regles de conception reprises des UML

- L'acteur principal reste visible dans l'en-tete : Agent INSTN, Responsable pedagogique ou Administrateur fonctionnel.
- Chaque ecran reprend la logique des `include` UML : import/selection, verification, execution, resultat.
- Les modules metier partagent la meme structure : navigation laterale, zone de parametres, zone de verification, zone de resultat.
- Les retours de statut sont toujours visibles, car tous les use cases incluent la consultation des messages de succes ou d'erreur.
- IRIS ne met pas le telechargement au centre : l'ecran indique surtout l'emplacement du fichier traite.

## Generation des images

Depuis le dossier `cadi/Python/docs/UML`, generer les PNG avec PlantUML :

```powershell
java -jar plantuml.jar wireframe_*.puml
```

Si PlantUML est deja configure dans l'IDE, ouvrir un fichier `wireframe_*.puml` puis utiliser l'aperçu ou l'export PNG.

Une version HTML autonome est aussi disponible ici :

```text
cadi/Python/docs/wireframes_cadi_uml.html
```

Elle permet de consulter les maquettes directement dans un navigateur, sans generer les PNG.

## Liste des fichiers ajoutes

- `wireframe_00_dashboard.puml`
- `wireframe_01_iris.puml`
- `wireframe_02_evalstat.puml`
- `wireframe_03_bilans.puml`
- `wireframe_04_fdc_parametres.puml`
- `../wireframes_cadi_uml.html`
