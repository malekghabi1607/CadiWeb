# Description detaillee pour la maquette CADI Web

Ce document relie les cas d'utilisation, les modules fonctionnels et les ecrans attendus de l'application CADI Web. Il sert de base pour construire une maquette detaillee dans Figma, Canva, Draw.io, Whimsical ou un autre outil de maquettage.

## 1. Idee generale de l'application

CADI Web est une application interne destinee a simplifier les traitements lies aux formations INSTN. L'utilisateur ne doit plus lancer directement des scripts Python ni chercher les fichiers manuellement dans plusieurs dossiers. Il utilise une interface web qui centralise les modules principaux :

- EvalStat : traiter les fiches d'evaluation des stagiaires.
- IRIS : traiter les exports natifs issus d'IRIS.
- Bilans : generer les bilans de sessions et les bilans de formation.
- Fiches de couts : lire une fiche de couts Excel et afficher les informations utiles.
- Parametres : gerer les comptes, profils, chemins et droits.
- Messagerie : preparer un mail au chef d'unite avec un message par defaut modifiable.

L'application doit donner une impression de tableau de bord professionnel : simple, claire, dense mais lisible. Chaque ecran doit guider l'utilisateur avec des champs bien nommes, des boutons explicites et un retour de traitement visible.

## 2. Acteurs principaux

### Agent INSTN

L'agent INSTN est l'utilisateur principal. Il lance les traitements EvalStat, IRIS, Bilans et Fiches de couts. Il consulte les messages de succes ou d'erreur, et peut utiliser la messagerie.

### Responsable pedagogique

Le responsable pedagogique consulte ou produit surtout les bilans. Il peut avoir besoin de verifier les informations de formation, les donnees de session et les documents produits.

### Administrateur fonctionnel

L'administrateur gere les comptes, les roles, les chemins de configuration, les parametres et le message par defaut de la messagerie.

### Chef d'unite

Le chef d'unite n'est pas un utilisateur direct du traitement. Il est surtout destinataire des mails prepares depuis l'application.

## 3. Navigation globale de la maquette

La maquette peut etre organisee avec une barre laterale gauche et une zone de travail principale.

### Barre laterale

Elements proposes :

- Logo CADI / INSTN en haut.
- Accueil.
- EvalStat.
- IRIS.
- Bilans.
- Fiches de couts.
- Messagerie.
- Parametres.
- Profil utilisateur.
- Deconnexion.

### En-tete de page

Elements proposes :

- titre de la page active ;
- nom de l'utilisateur connecte ;
- role de l'utilisateur ;
- bouton d'aide ou de documentation ;
- indicateur de statut si un traitement est en cours.

### Zone principale

Chaque module doit avoir :

- un bloc de formulaire ;
- un bloc de verification ou de resume ;
- un bloc de resultat ;
- une zone de messages d'erreur ou d'avertissement.

## 4. Use case general

### Description

Le use case general montre les grandes fonctionnalites accessibles depuis CADI Web. L'agent INSTN peut traiter les evaluations, traiter les exports IRIS, generer un bilan et utiliser la messagerie. Le responsable pedagogique intervient surtout sur les bilans. L'administrateur gere les parametres et les comptes.

### Fonctionnalites a representer dans la maquette

- Acces aux modules depuis une navigation commune.
- Affichage d'un tableau de bord d'accueil.
- Lancement de traitements metier.
- Consultation des messages de resultat.
- Acces aux fichiers produits lorsque le module le permet.
- Gestion des comptes et parametres pour l'administrateur.

### Ecran Accueil

L'ecran Accueil doit presenter rapidement l'etat de l'application.

Contenu conseille :

- cartes de raccourci vers EvalStat, IRIS, Bilans et Fiches de couts ;
- derniers traitements lances ;
- statut : succes, erreur, avertissement, en cours ;
- raccourci vers la messagerie ;
- message d'information si un chemin reseau ou un modele est indisponible.

## 5. Use case EvalStat

### Objectif du module

Le module EvalStat permet de traiter les fiches d'evaluation des stagiaires. L'utilisateur doit saisir un trigramme de formation, saisir le code IRIS de la session et selectionner les fiches d'evaluation. Le systeme verifie les fichiers, lance le traitement et met a jour le fichier Excel de formation.

### Actions utilisateur

1. Saisir le trigramme de formation.
2. Saisir le code IRIS.
3. Selectionner une ou plusieurs fiches d'evaluation.
4. Lancer le traitement EvalStat.
5. Consulter le resultat.
6. Ouvrir ou telecharger le fichier Excel genere si disponible.

### Actions automatiques du systeme

- Verifier que le trigramme est renseigne.
- Verifier que le code IRIS est present et conforme.
- Verifier le format des fichiers selectionnes.
- Detecter si l'evaluation est deja traitee.
- Creer ou mettre a jour le fichier Excel de session.
- Mettre a jour l'evaluation globale de formation.
- Afficher un message de resultat.

### Ecran de maquette EvalStat

Champs a afficher :

- Trigramme de formation.
- Code IRIS.
- Zone de selection des fiches d'evaluation.
- Liste des fichiers selectionnes.

Boutons :

- Verifier les fichiers.
- Lancer le traitement.
- Reinitialiser.
- Ouvrir le fichier resultat, si disponible.

Zone de resultat :

- statut du traitement ;
- nombre de fichiers lus ;
- nombre de fiches traitees ;
- fichier Excel mis a jour ;
- erreurs ou avertissements.

### Messages possibles

- Traitement termine avec succes.
- Evaluation deja traitee pour ce code IRIS.
- Code IRIS invalide.
- Fichier manquant ou format non reconnu.
- Fichier Excel verrouille.

## 6. Use case IRIS

### Objectif du module

Le module IRIS permet de transformer un export natif en fichier traite exploitable par CADI. L'utilisateur selectionne le fichier source, choisit le type d'export et lance le traitement.

### Actions utilisateur

1. Importer un export IRIS natif.
2. Choisir le type d'export : sessions, formations, ventes ou inscriptions.
3. Lancer le traitement IRIS.
4. Consulter le message de resultat.

### Actions automatiques du systeme

- Verifier la structure du fichier.
- Appliquer la configuration du type d'export.
- Nettoyer les donnees.
- Reorganiser les colonnes.
- Produire le fichier IRIS traite.
- Afficher l'emplacement du fichier produit.

### Ecran de maquette IRIS

Champs a afficher :

- Type d'export.
- Fichier IRIS source.
- Resume du fichier choisi : nom, taille, date si disponible.

Boutons :

- Selectionner un fichier.
- Lancer le traitement.
- Reinitialiser.

Zone de resultat :

- statut ;
- type d'export ;
- nombre de lignes lues ;
- nombre de lignes conservees ;
- emplacement du fichier traite.

Important : l'action "telecharger le fichier traite" ne doit pas etre l'action centrale du module. Le module doit surtout signaler que le fichier est pret et indiquer son emplacement.

## 7. Use case Bilans

### Objectif du module

Le module Bilans permet de generer un bilan de sessions ou un bilan de formation. Il rassemble les donnees EvalStat, IRIS, FdC et Specs selon le type de bilan choisi.

### Point important pour la maquette

Les trois premieres informations doivent etre au meme niveau :

- trigramme de formation ;
- annee et periode ;
- type de bilan.

Elles ne doivent pas etre presentees comme trois etapes separees. Ce sont les donnees de depart du traitement.

### Actions utilisateur

1. Saisir le trigramme de formation.
2. Saisir l'annee et la periode si necessaire.
3. Choisir le type de bilan : bilan de sessions ou bilan de formation.
4. Lancer la verification.
5. Generer le document Word.
6. Ouvrir ou telecharger le bilan produit.

### Actions automatiques du systeme

- Verifier les donnees necessaires.
- Charger les evaluations associees.
- Charger les donnees IRIS.
- Charger FdC et Specs si le bilan de formation en a besoin.
- Construire le document Word.
- Afficher les avertissements.

### Ecran de maquette Bilans

Bloc "Parametres du bilan" :

- Trigramme de formation.
- Annee.
- Periode.
- Type de bilan.
- Eventuellement liste de codes IRIS si le mode manuel est prevu.

Bloc "Verification des donnees" :

- EvalStat disponible ou manquant.
- IRIS disponible ou manquant.
- Fiche de couts disponible ou manquante.
- Specifications disponibles ou manquantes.

Bloc "Generation" :

- bouton Generer le bilan ;
- barre de progression ou statut ;
- lien vers le document Word produit ;
- message d'erreur si generation impossible.

### Messages possibles

- Donnees verifiees avec succes.
- Certaines evaluations sont manquantes.
- Fiche de couts introuvable.
- Export IRIS incomplet.
- Bilan genere avec succes.

## 8. Use case Fiches de couts

### Objectif du module

Le module Fiches de couts permet de lire une fiche de couts Excel et d'afficher les informations utiles de la formation.

### Actions utilisateur

1. Selectionner une fiche de couts Excel.
2. Consulter les informations extraites.

### Actions automatiques du systeme

- Detecter la version de la fiche.
- Choisir le lecteur compatible.
- Extraire les champs principaux.
- Signaler les champs manquants.

Important : la detection de version et le choix du lecteur ne doivent pas apparaitre comme des actions utilisateur.

### Ecran de maquette Fiches de couts

Champs / zones :

- Zone de depot ou selection du fichier Excel.
- Resume du fichier choisi.
- Tableau des informations extraites.

Informations a afficher :

- nom de la formation ;
- nombre de participants prevus ;
- nombre de jours ;
- cout total ;
- intervenants ;
- chemin du fichier.

Messages possibles :

- Fiche lue avec succes.
- Version de fiche non reconnue.
- Champ obligatoire manquant.
- Fichier Excel verrouille.

## 9. Use case Parametres, comptes et messagerie

### Objectif du module

Ce module regroupe les fonctions de connexion, inscription, gestion du profil, administration des utilisateurs, configuration et messagerie.

### Connexion et compte

Ecrans a prevoir :

- Connexion.
- Inscription.
- Mot de passe oublie.
- Profil utilisateur.

Champs du profil :

- nom ;
- prenom ;
- mail ;
- unite ;
- site : Marcoule, Grenoble, Saclay ou autre ;
- genre ou civilite ;
- role : utilisateur ou administrateur.

### Administration

Fonctionnalites administrateur :

- consulter les utilisateurs ;
- modifier un role ;
- activer ou desactiver un compte ;
- consulter les chemins de configuration ;
- modifier un chemin ou un parametre ;
- verifier la validite des chemins ;
- sauvegarder la configuration ;
- modifier le message par defaut de la messagerie.

### Messagerie

Objectif : permettre l'envoi d'un mail au chef d'unite uniquement.

Elements de maquette :

- destinataire affiche automatiquement : chef d'unite ;
- objet du mail ;
- titre du message ;
- message par defaut charge automatiquement ;
- zone de modification du message ;
- ajout d'une piece jointe ;
- bouton preparer/envoyer le mail ;
- message de confirmation.

Messages possibles :

- Message prepare avec succes.
- Destinataire chef d'unite introuvable.
- Piece jointe trop volumineuse.
- Envoi impossible.

## 10. Enchainement complet entre les modules

L'enchainement logique de CADI Web peut etre represente ainsi :

1. L'utilisateur se connecte.
2. Il verifie ou complete son profil.
3. Il traite les exports IRIS si les donnees ne sont pas a jour.
4. Il traite les evaluations EvalStat pour les sessions concernees.
5. Il charge ou controle les fiches de couts.
6. Il genere un bilan.
7. Il utilise la messagerie pour preparer un mail au chef d'unite.
8. Il consulte les fichiers produits et les messages de statut.

Cet enchainement peut apparaitre dans la maquette sous forme de tableau de bord ou de parcours guide.

## 11. Proposition de pages pour la maquette

### Page 1 : Connexion

Contenu :

- logo CADI / INSTN ;
- champ mail ;
- champ mot de passe ;
- bouton connexion ;
- lien inscription ;
- lien mot de passe oublie.

### Page 2 : Accueil

Contenu :

- titre "Tableau de bord CADI" ;
- cartes modules ;
- derniers traitements ;
- alertes ;
- raccourcis.

### Page 3 : EvalStat

Contenu :

- formulaire trigramme, code IRIS et fiches ;
- tableau des fichiers selectionnes ;
- bouton de traitement ;
- zone resultat.

### Page 4 : IRIS

Contenu :

- selection du type d'export ;
- selection du fichier source ;
- bouton lancer ;
- resume du traitement.

### Page 5 : Bilans

Contenu :

- trigramme, annee, periode et type de bilan au meme niveau ;
- panneau de verification des donnees ;
- bouton generer ;
- resultat Word.

### Page 6 : Fiches de couts

Contenu :

- selection de la fiche Excel ;
- tableau des donnees extraites ;
- messages de champs manquants.

### Page 7 : Messagerie

Contenu :

- chef d'unite comme destinataire ;
- objet ;
- titre ;
- message par defaut modifiable ;
- piece jointe ;
- bouton envoyer.

### Page 8 : Parametres

Contenu :

- profil utilisateur ;
- gestion des comptes ;
- chemins de configuration ;
- message par defaut ;
- roles et droits.

## 12. Regles de design pour la maquette

La maquette doit etre claire, professionnelle et orientee travail.

Recommandations :

- utiliser une barre laterale fixe ;
- garder les formulaires courts ;
- regrouper les champs obligatoires ;
- utiliser des tableaux pour les resultats ;
- utiliser des badges de statut : succes, erreur, avertissement, en cours ;
- afficher les erreurs pres du champ concerne ;
- eviter les pages trop decoratives ;
- garder les couleurs sobres avec une couleur principale proche du style CEA/INSTN.

## 13. Texte court a mettre dans la maquette

### EvalStat

"Traiter les fiches d'evaluation d'une session a partir du trigramme de formation, du code IRIS et des fichiers selectionnes."

### IRIS

"Transformer un export IRIS natif en fichier traite exploitable par CADI."

### Bilans

"Generer un bilan de sessions ou un bilan de formation a partir des donnees EvalStat, IRIS, FdC et Specs."

### Fiches de couts

"Lire une fiche de couts Excel et afficher automatiquement les informations principales de la formation."

### Parametres

"Gerer les comptes, les profils, les droits, les chemins de configuration et les messages par defaut."

### Messagerie

"Preparer un mail destine au chef d'unite avec un message par defaut modifiable et une piece jointe optionnelle."

