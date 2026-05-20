from vte.domain.formation import Formation
from vte.domain.fdc import *

import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))
from conftest import tel_data

# Répertoire des fichiers pour les tests
DATA_DIR = Path(__file__).parent / "data"

tel = tel_data()

"""
# Données constantes des tests
chemin_fdc= DATA_DIR / r"Fiche de coûts V6.1 - TEL.xlsx"
nom_onglet = "Fiche de coûts"
trigramme_formation = "TEL"


formation = Formation(trigramme_formation)
fe_avecChemin = FichierExcel(chemin_fichier=chemin_fdc)
fe_complet = FichierExcel.depuis_fichier(chemin_fichier=chemin_fdc, nom_onglet=nom_onglet)

# Cas 1 - OK mais relou car popup
#fdc = FdC(formation=formation)
#print(fdc.min_participants_cea)


# Cas 2
# fdc = FdC(formation=formation, fe=fe_avecChemin)
# print(fdc.min_participants_cea)


# Cas 3
# fdc = FdC(formation=formation, fe=fe_complet)
# print(fdc.min_participants_cea)

"""





# ---------- FIXTURES (données d'entrée aux tests) ----------

DATASETS_VERSIONS_FDC = {
    "6.1": {
        "chemin": chemin_vers_unc(Path(r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data\Fiche de coûts V6.1 - TEL.xlsx")),
        "dossier_plan_classement_fdc":chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation")),
        "version": "6.1",
        "lecteur": FdC_Lecteur_V6_1,
        "date_derniere_modification_windows": "29/01/2026",
        "nom_onglet": "Fiche de coûts",
        "trigramme_formation": "TEL",
        "nom_formation":"Télémanipulation",
        "annee_creationFormation":1999,
        "date_creationFormation":date(1999, 1, 1),
        "nb_participants_prevus":6,
        "nb_participants_min_cea":3,
        "nb_participants_min_ee":5,
        "cout_T1":4947,
        "cout_T2":7657,
        "cout_T3":7868,
        "prix_cea":1428,
        "prix_ee":1587,
        "prix_min_cea":1298,
        "prix_min_ee":1443,
        "prix_preconise_ee":1587,
        },

    "1.4": {
        "chemin": chemin_vers_unc(Path(r"C:\Users\vt238770\Documents\_CEA\Prog\Python\tests\vte\data\Fiche de coûts V1.4 - TEL.xlsx")),
        "dossier_plan_classement_fdc":chemin_vers_unc(Path(r"P:\FORMATIONS_C\TEL\P05-P06-dossier-conception-referentiel\fiche-de-cout-et-code-de-formation")),
        "version": "1.4",
        "lecteur": FdC_Lecteur_V1_4,
        "date_derniere_modification_windows": "28/04/2026",
        "nom_onglet": "Fiche de coûts",
        "trigramme_formation": "TEL",
        "nom_formation":"Télémanipulation",
        "annee_creationFormation":1999,
        "date_creationFormation":date(1999, 1, 1),
        "nb_participants_prevus":7,
        "nb_participants_min_cea":6,
        "nb_participants_min_ee":3,
        "cout_T1":6334,
        "cout_T2":8073,
        "cout_T3":8401,
        "prix_cea":1251,
        "prix_ee":1390,
        "prix_min_cea":995,
        "prix_min_ee":1320,
        "prix_preconise_ee":1452,
        },
}

# Je dois stocker une fonction car ça dépend de chemin
DATASETS_DECLARATION_FE = {
    "fe = none":
        lambda chemin: None,

    "fe appelé depuis chemin et nom_onglet":
        lambda chemin: FichierExcel.depuis_fichier(
            chemin_fichier=chemin,
            nom_onglet="Fiche de coûts"
        ),
}

DATASETS_RESOUDRE_CHEMIN = {
    "chemin_utilisateur_valide": {
        "chemin_utilisateur": tel["chemin_fdc"],
        "fe": None,
        "retour_auto": None,
        "retour_dialogue": None,
        "attendu": tel["chemin_fdc"],
        "erreur_attendue": False,
    },

    "chemin_invalide_fallback_fe": {
        "chemin_utilisateur": Path("C:/toto.xlsx"),
        "fe": FichierExcel.depuis_fichier(
                chemin_fichier=tel["chemin_fdc"],
                nom_onglet="Fiche de coûts"
            ),
        "retour_auto": None,
        "retour_dialogue": None,
        "attendu": tel["chemin_fdc"],
        "erreur_attendue": True,
    },

    "recherche_automatique": {
        "chemin_utilisateur": Path("C:/toto.xlsx"),
        "fe": None,
        "retour_auto": tel["chemin_fdc"],
        "retour_dialogue": None,
        "attendu": tel["chemin_fdc"],
        "erreur_attendue": True,
    },

    "filedialog": {
        "chemin_utilisateur": Path("C:/toto.xlsx"),
        "fe": None,
        "retour_auto": None,
        "retour_dialogue": tel["chemin_fdc"],
        "attendu": tel["chemin_fdc"],
        "erreur_attendue": True,
    },
}



"""
DATASETS_DECLARATION_FE_BAK = {
    "fe = none": None,
    #"fe appelé depuis chemin": FichierExcel(chemin_fichier=chemin_fdc),  # Ne marche pas car wb non chargé, mais cas ne devrait pas arriver
    "fe appelé depuis chemin et nom_onglet": FichierExcel.depuis_fichier(
            chemin_fichier=tel["chemin_fdc"],
            nom_onglet="Fiche de coûts"
        ),
}
"""



# ---------- TESTS ----------
def test_fdc_init():
    fdc = FdC()
    print(fdc)

    assert isinstance(fdc, FdC)
    assert fdc._lecteur is None

@pytest.mark.parametrize(
    "version, data_version",  # 2 noms car quand on fait DATASETS_VERSIONS_FDC.items(), on obtient un tuple (version, dataset)
    DATASETS_VERSIONS_FDC.items(),
    ids=DATASETS_VERSIONS_FDC.keys()
)
@pytest.mark.parametrize(
        "fe_builder",  # Nom de la variable que l'on va employer pour ce paramètre. Ici on a un dictionnaire qui renvoit la fonction qui construit fe
        DATASETS_DECLARATION_FE.values(),  
        ids=DATASETS_DECLARATION_FE.keys()
        )
def test_fdc_ouvrir(version, data_version, fe_builder):
    formation = Formation(data_version["trigramme_formation"])  # Récupéré depuis conftest

    # Définition du fichier Excel de la fiche de coûts (permet de tester les différents cas)
    fe = fe_builder(data_version["chemin"])

    # On ouvre la FdC
    fdc = FdC.ouvrir(
        chemin=data_version["chemin"],
        formation=formation,
        fe=fe
        )
    print(fdc)

    # Tests basiques
    assert isinstance(fdc, FdC)
    assert isinstance(fdc._lecteur, FdC_Lecteur)
    assert isinstance(fdc._lecteur, data_version["lecteur"])
    assert fdc.chemin == data_version["chemin"]
    assert fdc.version == data_version["version"]
    assert fdc.date_derniere_modification_windows == data_version["date_derniere_modification_windows"]

    # Tests sur les @property de FdC_Lecteur
    assert fdc.trigramme_formation == data_version["trigramme_formation"]
    assert isinstance(fdc.fe, FichierExcel)
    assert fdc.chemin_fe == data_version["chemin"]
    assert isinstance(fdc.tableau, FichierExcel._TableauExcel)
    assert fdc.tableau.nom_onglet == data_version["nom_onglet"]
    assert fdc.dossier_plan_classement == data_version["dossier_plan_classement_fdc"]


    # Tests sur les @property de FdC_Lecteur_versionné
    assert fdc.nom_onglet == data_version["nom_onglet"]
    assert fdc.nom_formation == data_version["nom_formation"]
    assert fdc.annee_creationFormation == data_version["annee_creationFormation"]
    assert fdc.date_creationFormation == data_version["date_creationFormation"]
    assert fdc.nb_participants_prevus == data_version["nb_participants_prevus"]
    assert fdc.nb_participants_min_cea == data_version["nb_participants_min_cea"]
    assert fdc.nb_participants_min_ee == data_version["nb_participants_min_ee"]


@pytest.mark.parametrize(
    "nom_cas, cas",
    DATASETS_RESOUDRE_CHEMIN.items(),
    ids=DATASETS_RESOUDRE_CHEMIN.keys()
)
def AAAtest_resoudre_chemin_BAK(monkeypatch, tel, nom_cas, cas):

    formation = Formation(tel["trigramme_formation"])

    monkeypatch.setattr(
        "vte.domain.fdc.selectionner_fichier_dans_repertoire",
        lambda **kwargs: cas["retour_auto"]
    )

    monkeypatch.setattr(
        "vte.domain.fdc.FdC_Lecteur._choisir_fdc",
        lambda self: cas["retour_dialogue"]
    )

    # Pour n'aller tester que la partie qui m'intéresse. Sinon je peux avoir des pb de workbook ou autres à cause d'une mauvaise instanciation de FdC.ouvrir()
    fdc = FdC_Lecteur(
        chemin=cas["chemin_utilisateur"],
        formation=formation,
        fe=cas["fe"]
    )

    result = fdc._resoudre_chemin()

    assert result == cas["attendu"]

@pytest.mark.parametrize(
    "cas",
    DATASETS_RESOUDRE_CHEMIN.values(),
    ids=DATASETS_RESOUDRE_CHEMIN.keys()
)
def test_resoudre_chemin(monkeypatch, tel, capsys, cas):

    formation = Formation(tel["trigramme_formation"])

    # ---------------------------------------------------------
    # Définition du chemin d'ouverture
    # ---------------------------------------------------------


    #chemin_ouverture = cas["chemin_utilisateur"]

    # ---------------------------------------------------------
    # Construction éventuelle du FE
    # ---------------------------------------------------------

    #fe = cas["fe"]

    # ---------------------------------------------------------
    # Mock recherche automatique
    # ---------------------------------------------------------

    monkeypatch.setattr(
        "vte.domain.fdc.selectionner_fichier_dans_repertoire",
        lambda **kwargs: cas["retour_auto"]
    )

    # ---------------------------------------------------------
    # Mock filedialog
    # ---------------------------------------------------------

    monkeypatch.setattr(
        "vte.domain.fdc.FdC_Lecteur._choisir_fdc",
        lambda self: cas["retour_dialogue"]
    )

    # ---------------------------------------------------------
    # Ouverture réelle
    # ---------------------------------------------------------

    fdc = FdC.ouvrir(
        chemin=cas["chemin_utilisateur"],
        formation=formation,
        fe=cas["fe"]
    )

    # ---------------------------------------------------------
    # Action
    # ---------------------------------------------------------

    #chemin = fdc._resoudre_chemin()

    # ---------------------------------------------------------
    # Assertions
    # ---------------------------------------------------------

    assert fdc.chemin == cas["attendu"]

    # ---------------------------------------------------------
    # Vérification logs
    # ---------------------------------------------------------

    if cas["erreur_attendue"]:

        captured = capsys.readouterr()  # Pour récupérer ce qui est print
        assert (
            "Le chemin donné par l'utilisateur n'existe pas"
            in captured.out
        )


def verif_IHM_resoudre_chemin(tel):
    formation = Formation(tel["trigramme_formation"])  # Récupéré depuis conftest
    
    fe = FichierExcel.depuis_fichier(
            chemin_fichier=tel["chemin_fdc"],
            nom_onglet="Fiche de coûts"
        )
    
    """
    fdc = FdC.ouvrir(
        chemin=tel["chemin_fdc"],
        formation=formation,
        fe=None
        )
    print(fdc) 
    """

    print("\nCas avec chemin donné par l'utilisateur ok") 
    fdc = FdC.ouvrir(
        chemin=tel["chemin_fdc"],
        formation=formation,
        fe=None
        )
    print(fdc) 
    print(fdc._resoudre_chemin())
    assert fdc.chemin == tel["chemin_fdc"]


    print("\n\nCas avec chemin donné par l'utilisateur erroné") 
    fdc = FdC.ouvrir(
        chemin=Path("C:/aa.toto"),
        formation=formation,
        fe=None
        )
    print(fdc)
    fdc.chemin = fdc._resoudre_chemin()
    print(fdc)
    # Doit aller dans vlog.log_erreur("Le chemin donné par l'utilisateur n'existe pas : {self.chemin}.\nSélection du fichier par une autre méthode.", continuer=True)
    # Comment le faire en pytest ?
    

    print("\n\nCas avec fe")
    fdc = FdC.ouvrir(
        chemin=Path("C:/aa.toto"),
        formation=formation,
        fe=fe
        )
    print(fdc)
    fdc.chemin = fdc._resoudre_chemin()
    print(fdc)
    # Doit aller dans vlog.log_erreur("Le chemin donné par l'utilisateur n'existe pas : {self.chemin}.\nSélection du fichier par une autre méthode.", continuer=True)
    # Comment le faire en pytest ?
    assert fdc.chemin == tel["chemin_fdc"]
    

    print("\n\nCas recherche automatique")
    fdc = FdC.ouvrir(
        chemin=Path("C:/aa.toto"),
        formation=formation,
        fe=None
        )
    print(fdc)
    fdc.chemin = fdc._resoudre_chemin()
    print(fdc)
    # Doit aller dans vlog.log_erreur("Le chemin donné par l'utilisateur n'existe pas : {self.chemin}.\nSélection du fichier par une autre méthode.", continuer=True)
    # Comment le faire en pytest ?
    # la fonction selectionner_fichier_dans_repertoire(...) renvoie None (dans utils et appelé depuis FdC_Lecteur)
    # appelle self._choisir_fdc() (dans FdC_Lecteur et appelé depuis FdC_Lecteur)





# TODO : rajouter constructeur depuis_chemin(...)

def main():
    verif_IHM_resoudre_chemin(tel)

if __name__ == "__main__":
    main()




"""
C'était une V1 avec fe_factory décalré directement
@pytest.mark.parametrize(
    "fe_factory",
    [
        lambda chemin: None,
        lambda chemin: FichierExcel(chemin_fichier=chemin),
        lambda chemin: FichierExcel.depuis_fichier(
            chemin_fichier=chemin,
            nom_onglet="Fiche de coûts"
        ),
    ],
)
def test_constructeur(fe_factory, formation_tel, chemin_fdc):
    fe = fe_factory(chemin_fdc)
    fdc = FdC(formation=formation_tel, fe=fe)

    assert fdc.min_participants_cea == 3
"""

# Lancer les tests :
#    - pytest : tous les tests depuis la racine du projet
#    - pytest -v : plus verbeux
#    - cyblé : python -m pytest -v tests/vte/test_fdc.py






