# Use cases - CADI Web

Ce dossier contient les cas d'utilisation CADI Web en fichiers PlantUML separes.

## Logiciel conseille

- PlantText : https://www.planttext.com/
- VS Code avec l'extension PlantUML

Pour voir un diagramme dans VS Code :

1. ouvrir un fichier `.puml` ;
2. `Ctrl+Shift+P` ;
3. lancer `PlantUML: Preview Current Diagram`.

## Fichiers

| Scenario | Fichier | Image |
|---|---|---|
| General CADI Web | `use_case_general.puml` | `use_case_general.png` |
| EvalStat | `use_case_evalstat.puml` | `use_case_evalstat.png` |
| IRIS | `use_case_iris.puml` | `use_case_iris.png` |
| Bilans | `use_case_bilans.puml` | `use_case_bilans.png` |
| FdC | `use_case_fdc.puml` | `use_case_fdc.png` |
| Comptes, parametres et messagerie | `use_case_parametres.puml` | `use_case_parametres.png` |

## Explications

### Use case general

La vue generale montre les grandes fonctionnalites de CADI Web :

- traiter les evaluations EvalStat ;
- traiter les exports IRIS ;
- generer les bilans ;
- gerer les comptes et parametres ;
- utiliser la messagerie ;
- consulter les messages de resultat ;
- telecharger les fichiers produits quand le scenario le prevoit.

L'action "lire une fiche de couts" a ete retiree de la vue generale, car elle est traitee comme un scenario specialise.

### EvalStat

Le scenario EvalStat demande maintenant :

- le trigramme de la formation ;
- le code IRIS ;
- les fiches d'evaluation a traiter.

Ensuite l'application verifie les fichiers, lance le traitement, met a jour le fichier Excel de formation et affiche un message de resultat.

### IRIS

Le scenario IRIS permet d'importer un export natif, de choisir le type d'export, puis de lancer le traitement.

L'action "telecharger le fichier traite" a ete retiree du diagramme, car elle n'est pas consideree comme une action principale du module.

### Bilans

Les trois premieres actions sont placees au meme niveau :

- saisir le trigramme ;
- saisir l'annee et la periode ;
- choisir le type de bilan.

Elles ne sont pas en chaine : elles representent les informations initiales a fournir avant la verification des donnees.

### FdC

L'utilisateur selectionne seulement une fiche de couts Excel.

La detection de la version de la fiche et le choix du lecteur compatible sont automatiques. Ils ne sont donc plus representes comme des actions utilisateur.

### Comptes, parametres et messagerie

Ce scenario contient :

- connexion ;
- inscription ;
- mot de passe oublie ;
- modification du profil ;
- nom, prenom, mail ;
- unite, site et genre ;
- administration des utilisateurs ;
- modification des chemins et parametres ;
- messagerie vers le chef d'unite uniquement ;
- modification du message par defaut ;
- modification de l'objet et du titre ;
- ajout eventuel d'une piece jointe.

