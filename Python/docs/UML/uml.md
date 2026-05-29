# UML - Projet VTE (INSTN / CEA)

Ce fichier regroupe les diagrammes Mermaid corriges d'apres le code actuel.

---

## Diagramme 1 - Classes du Domaine

```mermaid
classDiagram

    class Formation {
        -str _trigramme_formation
        -EvalStat_formation _eval
        -FdC _fdc
        -Specs _specs
        -list~Session~ _sessions
        -dict _bilans_sessions
        -dict _bilans_formation
        +avec_ajout_sessions()$ Formation
        +avec_ouverture_ou_creation_evalStat_formation()$ Formation
        +pour_traitement_bilanSessions_depuis_codesIRIS()$ Formation
        +pour_traitement_bilanSessions_depuis_periode()$ Formation
        +pour_traitement_bilanFormation_depuis_annee()$ Formation
        +ajout_sessions()
        +ouvrir_ou_creer_eval_formation()
        +ouvrir_ou_traiter_eval_sessions()
        +ouvrir_fdc()
        +ouvrir_specs()
        +ajout_bilan_sessions_avec_traitement()
        +ajout_bilan_formation_avec_traitement()
    }

    class Session {
        -Formation_protocol _formation
        -int _code_IRIS
        -EvalStat_session _eval
        +code_IRIS int
        +trigramme_formation str
        +eval EvalStat_session
        +eval_formation EvalStat_formation
        +avec_ouverture_ou_traitement_evalStat()$ Session
    }

    class EvalStat {
        -FichierExcel _fe
        +df_csv DataFrame
        +df_stagiaires DataFrame
        +fe FichierExcel
        #_ecrit_df_et_sauve()
    }

    class EvalStat_formation {
        -Formation_protocol _formation
        -str _df_initial_hash
        -dict _stats_stagiaires
        +avec_ouverture_ou_creation()$ EvalStat_formation
        +ecrit_et_sauve_df_siModif()
        +calculer_stats_criteres()
        +verifier_traitement_evalStat_session()
        +chemin_eval_formation Path
    }

    class EvalStat_session {
        -Session_protocol _session
        -Path _chemin_csv
        -str _statut
        -dict mapping_statuts$
        +avec_ouverture_ou_traitement()$ EvalStat_session
        +traiter_eval()
        +ouvrir_eval()
        +ouvrir_ou_traiter_eval()
        +statut str
        +code_IRIS int
        +eval_formation EvalStat_formation
    }

    class BilanSessions {
        -Formation_protocol _formation
        -int _annee
        -str _periode
        -str _periode_pour_titre
        -dict _statuts
        -list _codes_IRIS
        +depuis_codesIRIS()$ BilanSessions
        +depuis_periode()$ BilanSessions
        +verifier_traitement_bilan() bool
        +annee int
        +periode str
        +codes_IRIS list
        +eval_formation EvalStat_formation
    }

    class BilanFormation {
        -Formation_protocol _formation
        -int _annee
        -list _codes_IRIS
        -dict _statuts
        +depuis_codesIRIS()$ BilanFormation
        +depuis_annee()$ BilanFormation
        +verifier_traitement_bilan() bool
        +annee int
        +codes_IRIS list
        +eval_formation EvalStat_formation
    }

    class FdC {
        -FdC_Lecteur _lecteur
        +ouvrir()$ FdC
        +max_participants() int
        +chemin Path
        +__getattr__()
    }

    class Specs {
        -Formation_protocol _formation
        +depuis_chemin()$ Specs
    }

    Formation "1" *-- "0..*" Session : _sessions
    Formation "1" *-- "0..1" EvalStat_formation : _eval
    Formation "1" *-- "0..1" FdC : _fdc
    Formation "1" *-- "0..1" Specs : _specs
    Formation "1" *-- "0..*" BilanSessions : _bilans_sessions
    Formation "1" *-- "0..*" BilanFormation : _bilans_formation

    Session "1" *-- "0..1" EvalStat_session : _eval
    Session --> Formation : _formation protocol

    EvalStat_formation --|> EvalStat
    EvalStat_session --|> EvalStat

    EvalStat_session --> EvalStat_formation : eval_formation
    EvalStat_formation --> Formation : _formation protocol
    EvalStat_session --> Session : _session protocol

    BilanSessions --> Formation : _formation protocol
    BilanFormation --> Formation : _formation protocol
    FdC --> FdC_Lecteur : facade
    Specs --> Formation : _formation protocol
```

---

## Diagramme 2 - Infrastructure (utils / core)

```mermaid
classDiagram

    class FichierExcel {
        -Path _chemin_fichier
        -Workbook _wb
        -dict _tableaux
        +depuis_fichier()$ FichierExcel
        +depuis_modele()$ FichierExcel
        +get_tableau()
        +get_df_tableau()
        +set_df_tableau()
        +save()
        +close()
        +actualiser_TCD()
    }

    class FichierWord {
        -Path _chemin
        +depuisFichier()$ FichierWord
        +mail_merge()
        +save()
        +close()
    }

    class FichierGenerique {
        <<abstract>>
        -Path _chemin
        +chemin Path
        +ouvrir()$*
    }

    class IRIS {
        +verifier_code_IRIS()$
        +extraire_code_IRIS_depuis_chemin()$
    }

    class IRIS_natif {
        +avec_traitement()$ IRIS_natif
    }

    class IRIS_traite {
        -DataFrame _df
        +df DataFrame
        +get_champ_depuis_codes_IRIS()
        +get_periode_depuis_codes_IRIS()
    }

    class IRIS_sessions {
        +depuis_fichier()$ IRIS_sessions
    }

    class IRIS_ventes {
        +depuis_fichier()$ IRIS_ventes
        +prix_formation_annee_str() str
        +df_prix_formation_annee() DataFrame
    }

    class Config {
        +CHEMIN_MODELE_EXCEL_EVALUATIONS_STAGIAIRES$
        +CHEMIN_EXCEL_EVALUATIONS_FORMATION$
        +REPERTOIRE_CSV_EVALUATIONS$
    }

    IRIS_natif --|> IRIS
    IRIS_traite --|> IRIS
    IRIS_sessions --|> IRIS_traite
    IRIS_ventes --|> IRIS_traite

    EvalStat --> FichierExcel : _fe
    BilanSessions --> FichierWord : genere
    BilanFormation --> FichierWord : genere
    IRIS --> FichierExcel : _fe
    FdC_Lecteur --|> FichierGenerique
```

---

## Diagramme 3 - Sequence EvalStat

```mermaid
sequenceDiagram
    participant U as Utilisateur
    participant F as Formation
    participant S as Session
    participant ES as EvalStat_session
    participant EF as EvalStat_formation
    participant FE as FichierExcel

    U->>F: avec_ajout_sessions(trig, codes_IRIS)
    F->>S: new Session(formation, code_IRIS)
    U->>F: ouvrir_ou_traiter_eval_sessions()
    F->>EF: avec_ouverture_ou_creation(formation)
    EF->>FE: depuis_fichier() ou depuis_modele()
    loop pour chaque session
        F->>S: session.eval
        S->>ES: avec_ouverture_ou_traitement()
        alt evaluation deja traitee
            ES->>ES: ouvrir_eval()
        else evaluation non traitee
            ES->>ES: traiter_eval(chemin_csv)
            ES->>FE: depuis_modele()
            ES->>EF: mise a jour des donnees formation
        end
    end
    F->>EF: ecrit_et_sauve_df_siModif()
    EF->>FE: save() + close()
```

---

## Diagramme 4 - Sequence BilanSessions

```mermaid
sequenceDiagram
    participant U as Utilisateur
    participant F as Formation
    participant BS as BilanSessions
    participant ES as EvalStat_services
    participant EF as EvalStat_formation
    participant G as BilanSessions_generateur_word
    participant W as FichierWord

    U->>F: pour_traitement_bilanSessions_depuis_periode(trig, annee, periode)
    F->>BS: depuis_periode(formation, annee, periode)
    BS->>BS: verifier_traitement_bilan()
    BS->>BS: _traiter()
    BS->>ES: ouvrir_ou_traiter_evalStat_depuis_liste_codes_IRIS()
    BS->>EF: calculer_stats_criteres(codes_IRIS)
    BS->>G: _get_generateur_word()
    BS->>G: construire()
    G->>G: _construire_champs()
    G->>G: _fusionner_word()
    G->>W: depuisFichier()
    BS->>W: affichage / ouverture du document
    BS->>BS: _envoyer_mail_chef_unite()
```

---

## Diagramme 5 - IRIS detail complet

```mermaid
classDiagram

    class InfosExportsIRIS {
        <<dataclass>>
        +Path repertoire
        +Path chemin_fichier
        +str nom_onglet
        +int nbLignes_avantET
        +list ordre_colonne
    }

    class ConfigExportIRIS {
        -str _nom_typeExport
        -str _codeExport
        -InfosExportsIRIS _input
        -InfosExportsIRIS _modele
        -InfosExportsIRIS _output
        +__str__()
    }

    class IRIS {
        -str _typeExport
        -FichierExcel _fe
        -ConfigExportIRIS _SESSIONS$
        -ConfigExportIRIS _FORMATIONS$
        -ConfigExportIRIS _VENTES$
        -ConfigExportIRIS _INSCRIPTIONS$
        -dict DICT_EXPORTS_IRIS$
        +extraire_infos_numSessionIRIS()$
        +verifier_code_IRIS()$
        +extraire_code_IRIS_depuis_chemin()$
    }

    class IRIS_natif {
        +avec_traitement()$ IRIS_natif
        -_creer_export_IRIS()
        -_ecrire_et_sauver_df_dans_excel()
        -_concatener_fichiers_input()
    }

    class IRIS_traite {
        <<abstract>>
        -DataFrame _df
        +df DataFrame
        +get_champ_depuis_codes_IRIS()
        +get_periode_depuis_codes_IRIS()
        +df_filtre_codes_IRIS()
        +df_filtre_periode()
        +affiche_df_colonnes_principales()
    }

    class IRIS_sessions {
        +depuis_fichier()$ IRIS_sessions
    }

    class IRIS_ventes {
        +depuis_fichier()$ IRIS_ventes
        +prix_formation_annee_str() str
        +df_prix_formation_annee() DataFrame
    }

    note for IRIS_traite "Formations et Inscriptions sont lus via IRIS_traite ; pas de classes IRIS_formations / IRIS_inscriptions dans le code actuel."

    IRIS_natif --|> IRIS
    IRIS_traite --|> IRIS
    IRIS_sessions --|> IRIS_traite
    IRIS_ventes --|> IRIS_traite

    ConfigExportIRIS *-- InfosExportsIRIS : _input / _modele / _output
    IRIS *-- ConfigExportIRIS : _SESSIONS / _FORMATIONS / _VENTES / _INSCRIPTIONS
    IRIS --> FichierExcel : _fe
```

---

## Diagramme 6 - Bilans detail complet

```mermaid
classDiagram

    class BilanSessions {
        -Formation_protocol _formation
        -int _annee
        -str _periode
        -str _periode_pour_titre
        -dict _statuts
        -list _codes_IRIS
        +depuis_codesIRIS()$ BilanSessions
        +depuis_periode()$ BilanSessions
        +verifier_traitement_bilan() bool
        +annee int
        +periode str
        +codes_IRIS list
        +eval_formation EvalStat_formation
        -_traiter()
        -_construire_word()
        -_envoyer_mail_chef_unite()
        -_get_generateur_word()
    }

    class BilanSessions_generateur_word {
        <<abstract>>
        -BilanSessions _bilanSessions
        -dict _champs
        +construire()
        #_construire_champs()*
        #_fusionner_word()*
        #_post_traitement()*
    }

    class Bilan_V3_Sessions {
        #_construire_champs()
        #_construire_entete()
        #_construire_commentaires()
        #_construire_stats()
        #_fusionner_word()
        #_post_traitement()
    }

    class BilanFormation {
        -Formation_protocol _formation
        -int _annee
        -list _codes_IRIS
        -dict _statuts
        +depuis_annee()$ BilanFormation
        +depuis_codesIRIS()$ BilanFormation
        +annee int
        +codes_IRIS list
        +verifier_traitement_bilan() bool
        +eval_formation EvalStat_formation
        +iris_sessions IRIS_sessions
        +iris_ventes IRIS_ventes
        -_traiter()
        -_construire_word()
        -_get_generateur_word()
    }

    class BilanFormation_generateur_word {
        <<abstract>>
        -BilanFormation _bilanFormation
        -dict _champs
        +construire()
        #_construire_champs()*
        #_evaluer_recapDonnees()*
        #_fusionner_word()*
        #_post_traitement()*
    }

    class Bilan_V3_Formation {
        #_construire_champs()
        #_construire_entete()
        #_construire_stats_iris()
        #_construire_stats_eval_n()
        #_construire_stats_eval_nm1()
        #_fusionner_word()
    }

    note for Bilan_V3_Sessions "Nom reel dans le code : Bilan_V3 dans bilanSessions.py"
    note for Bilan_V3_Formation "Nom reel dans le code : Bilan_V3 dans bilanFormation.py"

    Bilan_V3_Sessions --|> BilanSessions_generateur_word
    Bilan_V3_Formation --|> BilanFormation_generateur_word

    BilanSessions --> BilanSessions_generateur_word : cree via _get_generateur_word()
    BilanFormation --> BilanFormation_generateur_word : cree via _get_generateur_word()

    BilanSessions --> EvalStat_formation : eval_formation
    BilanSessions --> IRIS_sessions : iris_sessions
    BilanFormation --> EvalStat_formation : eval_formation
    BilanFormation --> FdC : formation.fdc
    BilanFormation --> Specs : formation.specs
    BilanFormation --> IRIS_sessions : iris_sessions
    BilanFormation --> IRIS_ventes : iris_ventes
```

---

## Diagramme 7 - FdC (Facade + Strategie)

```mermaid
classDiagram

    class FdC {
        -FdC_Lecteur _lecteur
        +ouvrir()$ FdC
        +max_participants() int
        +chemin Path
        +__getattr__()
        -_resoudre_chemin()$
    }

    class FdC_Lecteur {
        <<abstract>>
        -Path _chemin
        -FichierExcel _fe
        -Formation_protocol _formation
        +ouvrir()$ FdC_Lecteur
        +nom_formation str*
        +nb_participants_prevus int*
        +nb_jours_formation float*
        +cout_total float*
        #_charger()*
        #_lire_champs()*
    }

    class FdC_Lecteur_V6_1 {
        +nom_formation str
        +nb_participants_prevus int
        +nb_jours_formation float
        +cout_total float
        #_charger()
        #_lire_champs()
    }

    class FdC_Lecteur_V1_4 {
        +nom_formation str
        +nb_participants_prevus int
        +est_compatible()$ bool
        #_charger()
    }

    class FdC_Lecteur_Defaut {
        +nom_formation str
        +nb_participants_prevus int
        #_charger()
    }

    class FichierGenerique {
        <<abstract>>
        -Path _chemin
        +chemin Path
        +ouvrir()$*
    }

    FdC_Lecteur_V6_1 --|> FdC_Lecteur
    FdC_Lecteur_V1_4 --|> FdC_Lecteur
    FdC_Lecteur_Defaut --|> FdC_Lecteur
    FdC_Lecteur --|> FichierGenerique

    FdC *-- FdC_Lecteur : _lecteur
    FdC_Lecteur --> FichierExcel : _fe
```

---

## Diagramme 8 - Vue globale domaine complet

```mermaid
classDiagram
    direction TB

    Formation *-- Session
    Formation *-- EvalStat_formation
    Formation *-- FdC
    Formation *-- Specs
    Formation *-- BilanSessions
    Formation *-- BilanFormation

    Session *-- EvalStat_session
    EvalStat_formation --|> EvalStat
    EvalStat_session --|> EvalStat

    EvalStat --> FichierExcel
    BilanSessions --> FichierWord
    BilanFormation --> FichierWord

    FdC *-- FdC_Lecteur
    FdC_Lecteur_V6_1 --|> FdC_Lecteur
    FdC_Lecteur_V1_4 --|> FdC_Lecteur
    FdC_Lecteur_Defaut --|> FdC_Lecteur

    Bilan_V3_Sessions --|> BilanSessions_generateur_word
    Bilan_V3_Formation --|> BilanFormation_generateur_word
    BilanSessions --> BilanSessions_generateur_word
    BilanFormation --> BilanFormation_generateur_word

    IRIS_natif --|> IRIS
    IRIS_traite --|> IRIS
    IRIS_sessions --|> IRIS_traite
    IRIS_ventes --|> IRIS_traite

    EvalStat --> IRIS_sessions : lit sessions
    BilanSessions --> IRIS_sessions : lit sessions
    BilanFormation --> IRIS_sessions : lit sessions
    BilanFormation --> IRIS_ventes : lit ventes
```

---

# Use cases

Les cas d'utilisation sont maintenant dans un fichier separe, plus propre et plus adapte :

➡ Voir `use_cases.md`

Ils sont ecrits en PlantUML, qui est le meilleur choix pour les diagrammes de cas d'utilisation.

---

# Versions PlantUML separees

Les diagrammes techniques de ce fichier existent aussi en fichiers PlantUML separes :

- `uml_01_classes_domaine.puml`
- `uml_02_infrastructure.puml`
- `uml_03_sequence_evalstat.puml`
- `uml_04_sequence_bilansessions.puml`
- `uml_05_iris_detail.puml`
- `uml_06_bilans_detail.puml`
- `uml_07_fdc_facade_strategie.puml`
- `uml_08_vue_globale_domaine.puml`

Le fichier `uml_00_index_plantuml.md` explique le role de chaque diagramme.
