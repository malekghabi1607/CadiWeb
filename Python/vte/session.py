# ======================================================================================
# CLASSE SESSION
# ======================================================================================
class Session:
    def __init__(self, code_IRIS: int):
        self._code_IRIS:int = code_IRIS
        
        # Une session a une évaluation stagiaire de la session /!\ Faire distinction entre EvalStat natif et le mien → On va prendre le mien
        self._eval:Optional[EvalStat] = None

