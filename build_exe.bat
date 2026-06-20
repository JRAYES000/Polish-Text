@echo off
REM ============================================================
REM  Genere TextEnhancerAI.exe (application autonome, sans console)
REM  A lancer une seule fois. L'exe apparaitra dans le dossier "dist".
REM ============================================================
echo.
echo === Installation des dependances ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

echo.
echo === Generation de l'executable ===
pyinstaller --noconfirm --onefile --windowed --name "TextEnhancerAI" main.py

echo.
echo ============================================================
echo  Termine. L'executable se trouve dans :  dist\TextEnhancerAI.exe
echo  Double-clique dessus pour lancer l'application.
echo ============================================================
pause
