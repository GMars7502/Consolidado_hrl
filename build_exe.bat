@echo off
REM Genera ConsolidadorTramas.exe en esta carpeta (usa el venv local si existe)
cd /d "%~dp0"
if exist venv\Scripts\python.exe (set PY=venv\Scripts\python.exe) else (set PY=python)
%PY% -m pip install -r requirements.txt pyinstaller
%PY% -m PyInstaller --noconfirm --onefile --windowed --name ConsolidadorTramas ^
  --exclude-module matplotlib --exclude-module scipy --exclude-module IPython ^
  consolidador_app.py
copy /Y dist\ConsolidadorTramas.exe ConsolidadorTramas.exe
echo.
echo Listo: ConsolidadorTramas.exe (junto a MAESTROS, TRAMAS y SALIDA)
pause
