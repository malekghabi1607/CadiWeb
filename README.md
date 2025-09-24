# Pour installer tous les packages : lancer depuis le répertoire prog dans un terminal


Get-ChildItem -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force
Get-ChildItem -Path . -Recurse -Include *.pyc -Force | Remove-Item -Force
pip install -e .

pytest -v tests/test_fichierExcel.py
pytest -v -s 