from openpyxl import Workbook

class TableauExcel():
    def __init__(self, nom_onglet:str=None, nom_tableau:str=None, nbLignes_avantET:int=None):
        self._nom_onglet = nom_onglet
        self._ws = None
        
        # Informations si tableau structuré (peuvent être être None si on a un tableau normal)
        self._nom_tableau = nom_tableau
        self._table = None  # Ne contient que les méta-données du tableau, pas les données elles-mêmes

        # Informations si tableau normal (peuvent être être None si on a un tableau structuré)
        self._nbLignes_avantET = nbLignes_avantET

        # Dimensions et références tableau
        self._ref_tableau = None  # Pour mémoriser la référence (ex: A1:D12)
        self._min_row = None
        self._max_row = None
        self._min_col = None
        self._max_col = None

    @classmethod
    def depuis_tableau_structure(cls, nom_onglet:str, nom_tableau:str = None, wb:Workbook=None):
        print()
    @classmethod
    def depuis_tableau_normal(cls, nom_onglet:str, nbLignes_avantET:int=0, wb:Workbook=None):
        print()

    def _initialiser_dimensions_structured_table(self):
        print()

    def _initialiser_dimensions_tableau_normal(self):
        print()
    
    def charger_workbook(self, wb:Workbook):
        print()
        