# Diagrammes PlantUML - CADI / VTE

Ce dossier contient une version PlantUML propre de chaque diagramme technique.

## Fichiers generes

| Diagramme | Fichier PlantUML | Role |
|---|---|---|
| 1 - Classes du domaine | `uml_01_classes_domaine.puml` | Structure principale : Formation, Session, EvalStat, Bilans, FdC, Specs |
| 2 - Infrastructure | `uml_02_infrastructure.puml` | Classes techniques : Excel, Word, IRIS, Config |
| 3 - Sequence EvalStat | `uml_03_sequence_evalstat.puml` | Scenario de traitement des evaluations |
| 4 - Sequence BilanSessions | `uml_04_sequence_bilansessions.puml` | Scenario de generation d'un bilan de sessions |
| 5 - IRIS detail | `uml_05_iris_detail.puml` | Organisation des exports IRIS |
| 6 - Bilans detail | `uml_06_bilans_detail.puml` | Classes de generation Word des bilans |
| 7 - FdC facade + strategie | `uml_07_fdc_facade_strategie.puml` | Lecture des fiches de couts selon leur version |
| 8 - Vue globale | `uml_08_vue_globale_domaine.puml` | Vue synthetique de tout le domaine |
| 9 - Sequence IRIS | `uml_09_sequence_iris.puml` | Scenario de traitement d'un export IRIS |
| 10 - Sequence FdC | `uml_10_sequence_fdc.puml` | Scenario de lecture d'une fiche de couts |

## Explication de chaque diagramme

### 1. Classes du domaine

Ce diagramme montre les classes metier principales du projet VTE.  
La classe centrale est `Formation`. Elle contient des sessions, une evaluation de formation, une fiche de couts, des specifications pedagogiques et des bilans.

Il permet d'expliquer comment les objets metier sont relies entre eux :

- une formation possede plusieurs sessions ;
- une session possede une evaluation statistique ;
- les classes `EvalStat_formation` et `EvalStat_session` heritent de `EvalStat` ;
- les bilans utilisent la formation et les donnees d'evaluation ;
- `FdC` est une facade vers un lecteur specialise.

### 2. Infrastructure

Ce diagramme montre les classes techniques utilisees par le domaine.  
Il separe les traitements metier des outils d'acces aux fichiers.

Les classes principales sont :

- `FichierExcel` pour lire et ecrire les fichiers Excel ;
- `FichierWord` pour generer ou modifier les documents Word ;
- `FichierGenerique` comme base commune ;
- `IRIS`, `IRIS_natif`, `IRIS_traite`, `IRIS_sessions` et `IRIS_ventes` pour gerer les exports IRIS ;
- `Config` pour les chemins et parametres.

### 3. Sequence EvalStat

Ce diagramme est un scenario dynamique.  
Il montre l'ordre des appels quand l'utilisateur lance le traitement EvalStat.

Le flux est le suivant :

1. l'utilisateur cree une formation avec des codes IRIS ;
2. la formation cree les sessions ;
3. le fichier EvalStat de formation est ouvert ou cree ;
4. chaque session ouvre ou traite son evaluation ;
5. les donnees sont mises a jour dans le fichier Excel ;
6. le fichier est sauvegarde.

### 4. Sequence BilanSessions

Ce diagramme montre le scenario de generation d'un bilan de sessions.  
Il part d'une demande utilisateur et finit par la creation du document Word.

Le traitement passe par :

- la creation d'un objet `BilanSessions` ;
- la verification des donnees ;
- le traitement des evaluations ;
- le calcul des statistiques ;
- la creation d'un generateur Word ;
- la fusion des champs dans le modele Word ;
- l'ouverture ou l'affichage du document produit.

### 5. IRIS detail

Ce diagramme detaille le sous-systeme IRIS.  
Il montre la difference entre les exports natifs et les exports traites.

Les points importants :

- `IRIS_natif` sert a traiter un export brut ;
- `IRIS_traite` represente un fichier deja exploitable ;
- `IRIS_sessions` et `IRIS_ventes` sont des specialisations ;
- les exports Formations et Inscriptions ne sont pas des classes concretes dediees dans le code actuel.

### 6. Bilans detail

Ce diagramme detaille les classes liees a la generation des bilans.  
Il montre deux familles :

- `BilanSessions` pour les bilans de sessions ;
- `BilanFormation` pour les bilans annuels de formation.

Chaque bilan utilise un generateur Word abstrait.  
Les classes `Bilan_V3_Sessions` et `Bilan_V3_Formation` sont des noms de clarification dans le diagramme. Dans le code, les deux classes concretes s'appellent `Bilan_V3`, mais elles se trouvent dans deux fichiers differents.

### 7. FdC facade + strategie

Ce diagramme montre le design pattern utilise pour lire les fiches de couts.

`FdC` est une facade : elle fournit une interface simple au reste du programme.  
Elle delegue le vrai travail a un lecteur compatible :

- `FdC_Lecteur_V6_1` ;
- `FdC_Lecteur_V1_4` ;
- `FdC_Lecteur_Defaut`.

Ce choix permet de supporter plusieurs versions de fiches Excel sans changer le reste du code.

### 8. Vue globale du domaine

Ce diagramme resume toutes les grandes relations du domaine.  
Il est moins detaille que les autres, mais il permet de presenter rapidement l'architecture globale.

Il est utile dans un rapport pour donner une vision d'ensemble avant d'entrer dans les diagrammes plus detailles.

### 9. Sequence IRIS

Ce diagramme montre le scenario complet de traitement d'un export IRIS.  
L'utilisateur importe un fichier natif et choisit le type d'export. L'application cree un export traite, sauvegarde le resultat dans Excel, puis recharge les donnees pour les rendre exploitables par les autres modules.

### 10. Sequence FdC

Ce diagramme montre la lecture d'une fiche de couts.  
La detection de la version de la fiche et le choix du lecteur compatible sont automatiques. L'utilisateur fournit seulement le fichier Excel, puis l'application extrait les informations utiles.

## Scenarios

Les scenarios en diagrammes de sequence sont :

- `uml_03_sequence_evalstat.puml`
- `uml_04_sequence_bilansessions.puml`
- `uml_09_sequence_iris.puml`
- `uml_10_sequence_fdc.puml`

Les use cases sont dans :

- `use_case_general.puml`
- `use_case_evalstat.puml`
- `use_case_iris.puml`
- `use_case_bilans.puml`
- `use_case_fdc.puml`
- `use_case_parametres.puml`

## Logiciel conseille

Utilise **PlantText** : https://www.planttext.com/

Tu peux aussi ouvrir chaque fichier `.puml` dans VS Code avec l'extension PlantUML, puis lancer :

`PlantUML: Preview Current Diagram`
