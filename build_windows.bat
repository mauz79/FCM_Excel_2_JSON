@echo off
setlocal EnableExtensions EnableDelayedExpansion
set APP_NAME=FCM_Excel_2_JSON

echo ===========================================
echo  Build "%APP_NAME%" (solo EXE, no ZIP)
echo ===========================================
echo.

REM 1) venv
if not exist .venv (
  echo [STEP] Creo ambiente virtuale .venv ...
  py -3 -m venv .venv || (echo [ERRORE] Creazione venv fallita & exit /b 1)
)
echo [STEP] Attivo .venv ...
call .venv\Scripts\activate.bat || (echo [ERRORE] Attivazione venv fallita & exit /b 1)

REM 2) deps
echo [STEP] Aggiorno pip ...
python -m pip install --upgrade pip || goto :err

if exist requirements.txt (
  echo [STEP] Installo dipendenze da requirements.txt ...
  pip install -r requirements.txt || goto :err
) else (
  echo [WARN] requirements.txt non trovato. Installo pacchetti minimi...
  pip install pandas==2.2.2 openpyxl==3.1.5 xlrd==2.0.1 FreeSimpleGUI==5.2.0.post1 || goto :err
)

echo [STEP] Installo PyInstaller ...
pip install pyinstaller || goto :err

REM 3) pulizia build precedenti
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM 4) build EXE
echo [STEP] Compilo eseguibile ...
pyinstaller --noconfirm --clean --onefile --windowed app.py --name "%APP_NAME%" || goto :err

if exist "dist\%APP_NAME%.exe" (
  echo.
  echo [OK] Eseguibile creato: "dist\%APP_NAME%.exe"
  echo.
  echo Apro la cartella dist...
  explorer.exe "%cd%\dist"
  exit /b 0
) else (
  echo [ERRORE] Eseguibile non trovato nella cartella dist.
  exit /b 1
)

:err
echo.
echo [BUILD FALLITA] Controlla i messaggi sopra.
exit /b 1
