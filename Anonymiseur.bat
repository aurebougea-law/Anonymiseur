@echo off
rem Lance Anonymiseur (version script, pour le cabinet) avec le Python du cabinet.
rem Les confreres recoivent l'executable construit par _build\Construire_EXE.bat.
cd /d "%~dp0"
rem Python du cabinet (_python.bat) s'il existe, sinon celui installe sur le poste.
set PY=
if exist "%~dp0..\_python.bat" call "%~dp0..\_python.bat"
if not defined PY (where py >nul 2>&1 && set "PY=py -3")
if not defined PY (where python >nul 2>&1 && set "PY=python")
if not defined PY (
  echo Python 3 introuvable : installez-le depuis python.org, ou utilisez Anonymiseur.exe.
  pause & exit /b 1)
%PY% "%~dp0anonymiseur.py"
if errorlevel 1 (
  echo.
  echo ---- Le programme s'est arrete avec le code %errorlevel% ----
  pause
)
