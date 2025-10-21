# Pour installer tous les packages : lancer depuis le répertoire prog dans un terminal

# Supprimer les fichiers temporaires
Get-ChildItem -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force
Get-ChildItem -Path . -Recurse -Include *.pyc -Force | Remove-Item -Force
pip install -e .

# Lancer les tests
pytest -v tests/test_fichierExcel.py
pytest -v -s 

# Git
# Sauvegarde ton environnement actuel : 
pip freeze > requirements.txt

# Ajoute tout et commit initial :
git add .
git commit -m "Initial commit: ajout du projet avec modules vte (utils, office, instn)"

# Committer la nouvelle version
git add vte/office.py vte/__init__.py
git commit -m "utils: ajout fonction soustraction (v0.1.1)"

# Tester les applis
pytest
python main.py

# Voir la liste des évolutions dans git :
git log --oneline