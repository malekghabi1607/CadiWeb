import pytest
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))

from conftest import tel_data
#from tests.conftest import tel

from vte.domain.formation import Formation
#from vte.tests.VTE.conftest import *

tel = tel_data()
print(tel["trigramme_formation"])

