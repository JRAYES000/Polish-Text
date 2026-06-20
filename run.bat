@echo off
REM Lance l'application directement (mode developpement, sans generer d'exe).
REM Installe les dependances au premier lancement si besoin.
python -m pip install -r requirements.txt >nul 2>&1
start "" pythonw main.py
