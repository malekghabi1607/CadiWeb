# Mission actuelle - Mise a jour des dumps IRIS

Ce document organise la mission de mise a jour des dumps IRIS. Il explique ce qu'il faut faire aujourd'hui, ce qui est manuel, ce qui peut etre automatise, et comment transformer cette procedure en fonctionnalite dans CADI.

## 1. Objectif de la mission

Le but est de maintenir les exports IRIS a jour pour que les autres modules de CADI utilisent les bonnes donnees.

Ces donnees servent ensuite a :

- retrouver les sessions par code IRIS ;
- connaitre les formations existantes ;
- recuperer les ventes ;
- recuperer les inscriptions ;
- alimenter les bilans de sessions et les bilans de formation ;
- eviter de travailler avec des fichiers anciens ou redondants.

La mission consiste donc a mettre a jour les fichiers sources dans le dossier des extracts originaux, puis a lancer le traitement CADI qui concatene ces fichiers dans des fichiers complets.

## 2. Dossiers importants

### Dossier GED IRIS

Chemin :

```text
\\instnt\HOME\REFERENC\IRIS - rapports de synthese
```

Role :

- contient des dumps automatiques venant d'IRIS ;
- sert de source pour certains exports ;
- peut contenir les derniers fichiers mis a disposition automatiquement.

Dans le code, ce chemin correspond a :

```python
REPERTOIRE_EXTRACT_IRIS_GED
```

### Dossier local des extracts originaux

Chemin :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

Chemin reseau equivalent dans le code :

```text
\\harmonie\instn\uem\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

Role :

- c'est le dossier de travail utilise par CADI pour lire les fichiers IRIS source ;
- on y copie les exports recents ;
- on y garde les fichiers historiques ;
- on y supprime ou remplace les dumps redondants.

Dans le code, ce chemin correspond a :

```python
REPERTOIRE_EXTRACT_IRIS_LOCAL
```

### Dossier des extracts complets

Chemin :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts complets
```

Role :

- contient les fichiers finaux generes par CADI ;
- ces fichiers sont le resultat de la concatenation ;
- les modules Bilans et EvalStat peuvent s'appuyer dessus.

Dans le code, ce chemin correspond a :

```python
REPERTOIRE_EXCEL_IRIS_OUTPUT
```

## 3. Les quatre types d'exports IRIS

CADI gere quatre familles d'exports :

| Type CADI | Code IRIS | Role |
|---|---|---|
| Sessions | R04110 | Liste des sessions, dates, lieux, codes IRIS, trigrammes, statuts |
| Formations | R0304 | Referentiel des formations |
| Ventes | R04301 | Donnees de vente par formation ou session |
| Inscriptions | R04500 | Donnees d'inscription des stagiaires |

Chaque famille a ses propres fichiers source, puis CADI produit un fichier complet.

## 4. Partie automatique depuis la GED

### Ce qui est demande

Aller dans :

```text
\\instnt\HOME\REFERENC\IRIS - rapports de synthese
```

Puis copier automatiquement les derniers dumps vers :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

Les fichiers concernes dans l'exemple sont :

```text
R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx
R04500_Sessions-Inscriptions-filtre sur FC2026 au 2026-04-01.xlsx
```

### Ce que ca veut dire

Ces fichiers correspondent aux exports :

- R04301 : Ventes ;
- R04500 : Inscriptions.

Ils changent avec la date. Quand un nouveau fichier arrive, il faut garder le dernier et enlever l'ancien fichier equivalent pour eviter les doublons.

### Probleme actuel

Si on garde deux fichiers de la meme famille et de la meme annee, par exemple :

```text
R04301_Sessions-Ventes-filtre sur FC2026 au 2026-02-02 LG.xlsx
R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx
```

CADI risque de lire deux fois des donnees proches ou redondantes. Le bilan peut alors utiliser des donnees anciennes ou dupliquees.

### Regle a appliquer

Pour une meme famille et une meme annee fiscale, on garde seulement le fichier le plus recent.

Exemple :

- ancien fichier : `R04301_Sessions-Ventes-filtre sur FC2026 au 2026-02-02 LG.xlsx`
- nouveau fichier : `R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx`
- action : supprimer ou archiver l'ancien, puis garder le nouveau.

## 5. Partie manuelle depuis IRIS

### Ce qui reste manuel

Certains fichiers ne viennent pas correctement ou directement depuis la GED. Il faut donc aller dans IRIS, generer les rapports, recuperer les fichiers dans le dossier Telechargements, puis les coller dans le dossier des extracts originaux.

Dossier de telechargement exemple :

```text
C:\Users\vt238770\Downloads
```

Dossier de destination :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

### Export Formations

Dans IRIS :

```text
01_Referentiel -> R0304_Ref_Formation-Listedesformations
```

Fichier attendu apres telechargement :

```text
R0304_Ref_Formation-Listedesformations.xlsx
```

Nom final conseille :

```text
R0304_Ref_Formation-Listedesformations-AAAA.MM.JJ.xlsx
```

Exemple :

```text
R0304_Ref_Formation-Listedesformations-2026.03.03.xlsx
```

Regle :

- on ajoute la date du jour au nom ;
- on supprime l'ancien dump equivalent ;
- on met a jour la liste dans `config_extractsIRIS.py`.

### Export Sessions

Dans IRIS :

```text
05_Planification -> R04110_Sessions
```

Plage de dates a utiliser :

```text
du 1er janvier de l'annee en cours jusqu'a l'annee n+10
```

Exemple pour 2026 :

```text
du 01/01/2026 jusqu'au 31/12/2036
```

Fichier telecharge possible :

```text
R04110_Sessions.xlsx
R04110_Sessions (1).xlsx
R04110_Sessions (2).xlsx
```

Nom final conseille :

```text
R04110_Sessions-AAAA au AAAA.MM.JJ.xlsx
```

Exemple :

```text
R04110_Sessions-2026 au 2026.03.03.xlsx
```

Regle :

- le premier `2026` indique l'annee couverte ;
- la date finale indique la date de generation du dump ;
- on supprime l'ancien dump equivalent de l'annee en cours ;
- on met a jour la liste dans `config_extractsIRIS.py`.

## 6. Ce qu'il faut faire aujourd'hui, etape par etape

### Etape 1 : reperer la date du jour

Choisir le format utilise dans les noms de fichiers :

```text
AAAA.MM.JJ
```

Exemple :

```text
2026.05.22
```

Pour les fichiers GED deja nommes avec tirets, le format peut rester :

```text
AAAA-MM-JJ
```

Il faut garder le format deja present dans le nom source.

### Etape 2 : recuperer les fichiers GED recents

Aller dans :

```text
\\instnt\HOME\REFERENC\IRIS - rapports de synthese
```

Chercher les fichiers les plus recents pour :

```text
R04301_Sessions-Ventes-filtre sur FC2026 ...
R04500_Sessions-Inscriptions-filtre sur FC2026 ...
```

Copier ces fichiers dans :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

### Etape 3 : supprimer les anciens fichiers redondants

Dans le dossier destination, chercher les anciens fichiers de la meme famille.

Exemple pour ventes :

```text
R04301_Sessions-Ventes-filtre sur FC2026 au 2026-02-02 LG.xlsx
R04301_Sessions-Ventes-filtre sur FC2026 au 2026-04-01 LG.xlsx
```

Garder seulement le plus recent.

Faire pareil pour inscriptions :

```text
R04500_Sessions-Inscriptions-filtre sur FC2026 au 2026-02-02.xlsx
R04500_Sessions-Inscriptions-filtre sur FC2026 au 2026-04-01.xlsx
```

Garder seulement le plus recent.

### Etape 4 : generer le fichier Formations depuis IRIS

Dans IRIS, lancer :

```text
01_Referentiel -> R0304_Ref_Formation-Listedesformations
```

Le fichier arrive dans Telechargements.

Le renommer avec la date du jour :

```text
R0304_Ref_Formation-Listedesformations-AAAA.MM.JJ.xlsx
```

Puis le copier dans :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

Ensuite supprimer l'ancien fichier R0304 equivalent.

### Etape 5 : generer le fichier Sessions depuis IRIS

Dans IRIS, lancer :

```text
05_Planification -> R04110_Sessions
```

Mettre la periode :

```text
01/01/annee_en_cours -> 31/12/annee_en_cours+10
```

Renommer le fichier telecharge :

```text
R04110_Sessions-AAAA au AAAA.MM.JJ.xlsx
```

Exemple :

```text
R04110_Sessions-2026 au 2026.05.22.xlsx
```

Puis le copier dans :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts originaux
```

Ensuite supprimer l'ancien fichier R04110 equivalent de l'annee en cours.

### Etape 6 : mettre a jour `config_extractsIRIS.py`

Fichier :

```text
cadi/Python/vte/core/config_extractsIRIS.py
```

Ce fichier contient les listes de fichiers a concatener.

Il faut mettre a jour :

```python
_tSessions
_tFormations
_tVentes
_tInscriptions
```

Exemple :

```python
_tSessions = (
    'R04110_Sessions-2011 a 2014 FINAL.xlsx',
    'R04110_Sessions-2015 FINAL.xlsx',
    ...
    'R04110_Sessions-2026 au 2026.05.22.xlsx',
)
```

Important :

- garder les anciens fichiers historiques `FINAL` ;
- remplacer seulement le fichier courant de l'annee en cours ;
- ne pas laisser deux fichiers recents pour la meme famille et la meme annee.

### Etape 7 : lancer le traitement CADI

Application :

```text
cadi/Python/apps/gestionIRIS/gestionIRIS.py
```

Le menu permet de choisir :

- Tout traiter ;
- Traiter sessions ;
- Traiter formations ;
- Traiter ventes ;
- Traiter inscriptions.

Mode de selection :

- `0` : selection manuelle avec filedialog ;
- `1` : selection automatique avec `config_extractsIRIS.py`.

Pour ta mission, le plus logique est d'utiliser :

```text
selection automatique avec config_extractsIRIS.py
```

Comme ca, CADI lit les fichiers listes dans la config.

### Etape 8 : verifier les fichiers complets produits

Apres traitement, verifier le dossier :

```text
R:\_Echanges\VTE\Prog\IRIS\Extracts complets
```

Tu dois obtenir des fichiers du type :

```text
R04110_Sessions-COMPLET-AAAA.MM.JJ.xlsx
R0304_Formations-COMPLET-AAAA.MM.JJ.xlsx
R04301_Ventes-COMPLET-AAAA.MM.JJ.xlsx
R04500_Inscriptions-COMPLET-AAAA.MM.JJ.xlsx
```

Verifier :

- le fichier existe ;
- la date est recente ;
- le fichier s'ouvre ;
- le nombre de lignes semble coherent ;
- les colonnes principales sont presentes.

## 7. Comment automatiser proprement

### Automatisation 1 : copier les derniers fichiers GED

Objectif :

- chercher automatiquement le dernier fichier `R04301` ;
- chercher automatiquement le dernier fichier `R04500` ;
- les copier dans le dossier local ;
- supprimer ou archiver l'ancien fichier equivalent.

Logique :

1. Lire tous les fichiers du dossier GED.
2. Filtrer ceux qui commencent par `R04301_Sessions-Ventes-filtre sur FC2026`.
3. Trier par date dans le nom ou par date de modification.
4. Prendre le plus recent.
5. Copier dans le dossier local.
6. Supprimer les anciens `R04301` du meme FC2026 dans le dossier local.
7. Refaire la meme chose pour `R04500`.

### Automatisation 2 : eviter de modifier `config_extractsIRIS.py` a la main

Aujourd'hui, le probleme est que les noms de fichiers sont ecrits directement dans :

```text
config_extractsIRIS.py
```

Donc quand un fichier change, on doit modifier le code.

Solution plus propre :

- detecter automatiquement les fichiers presents dans le dossier ;
- garder les fichiers historiques `FINAL` ;
- selectionner automatiquement le dernier fichier courant ;
- retourner la liste au traitement IRIS.

Comme ca, on n'a plus besoin de copier les noms a la main dans le fichier Python.

### Automatisation 3 : ajouter une fonction de nettoyage

Fonction attendue :

```text
nettoyer_anciens_dumps(type_export, annee)
```

Elle ferait :

- identifier les fichiers du meme type ;
- identifier ceux de la meme annee ;
- garder le plus recent ;
- deplacer les autres dans un dossier `Archives` ou les supprimer selon la regle choisie.

Il vaut mieux archiver au debut, puis supprimer plus tard quand on est sur.

### Automatisation 4 : interface web future

Dans la maquette CADI Web, le module IRIS pourrait avoir une section "Mise a jour des dumps".

Ecran propose :

- bouton "Scanner GED" ;
- tableau des fichiers trouves ;
- colonne "dernier fichier detecte" ;
- colonne "fichier local actuel" ;
- bouton "Copier les derniers dumps" ;
- bouton "Nettoyer les anciens dumps" ;
- bouton "Lancer la concatenation" ;
- zone de resultat.

## 8. Checklist rapide pour toi

Avant traitement :

- verifier que tu es connecte aux lecteurs reseau ;
- verifier que le dossier GED est accessible ;
- verifier que le dossier `Extracts originaux` est accessible ;
- verifier la date du jour ;
- verifier les fichiers recents R04301 et R04500 ;
- generer R0304 et R04110 manuellement depuis IRIS si necessaire.

Pendant le traitement :

- copier les nouveaux fichiers ;
- supprimer ou archiver les anciens redondants ;
- mettre a jour `config_extractsIRIS.py` si l'automatisation n'est pas encore faite ;
- lancer `gestionIRIS.py` ;
- choisir le mode automatique ;
- traiter tous les exports ou seulement ceux modifies.

Apres traitement :

- verifier les fichiers dans `Extracts complets` ;
- ouvrir un fichier pour verifier les colonnes ;
- verifier que les modules Bilans peuvent utiliser les nouvelles donnees ;
- noter les erreurs si un fichier est mal structure.

## 9. Resume tres court

Ta mission, c'est :

1. Recuperer les derniers dumps IRIS.
2. Les mettre dans le bon dossier.
3. Supprimer les anciens doublons.
4. Generer manuellement les fichiers qui ne viennent pas automatiquement.
5. Renommer les fichiers correctement.
6. Mettre a jour la config si necessaire.
7. Lancer le traitement CADI.
8. Verifier les fichiers complets produits.

## 10. Ce qu'il faut expliquer dans ton rapport ou a l'oral

Phrase possible :

> La mission de mise a jour des dumps IRIS consiste a maintenir les sources de donnees CADI a jour. Certains exports sont recuperes automatiquement depuis la GED, tandis que d'autres doivent encore etre generes manuellement depuis IRIS. Les fichiers sont ensuite renommes, copies dans le repertoire des extracts originaux, nettoyes pour eviter les doublons, puis concatènes par le module IRIS de CADI afin de produire des extracts complets exploitables par les modules EvalStat et Bilans.

