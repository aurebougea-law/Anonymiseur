@echo off
rem ====================================================================
rem  Construire_EXE.bat : fabrique le paquet a transmettre aux confreres.
rem
rem  Double-cliquez, attendez, c'est tout. A la fin :
rem     _a_transmettre\Anonymiseur\            le dossier pret a copier
rem     _a_transmettre\Anonymiseur_<version>.zip   le zip a envoyer
rem
rem  Etapes : composants -> icone -> recettes -> executable -> paquet
rem           (+ Tesseract) -> verification du paquet -> composants tiers
rem           et empreintes -> zip (+ empreinte du zip a publier).
rem  Le script s'arrete au premier probleme et le dit en clair.
rem ====================================================================
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
set VERSION=2026.10.1
if "%LANGUE_BUILD%"=="" set LANGUE_BUILD=fr
set NOM=Anonymiseur
set GUIDE=Mode d'emploi.pdf
if "%LANGUE_BUILD%"=="en" set NOM=Anonymizator
if "%LANGUE_BUILD%"=="en" set GUIDE=User guide.pdf
set LICENCE=LICENCE.txt
if "%LANGUE_BUILD%"=="en" set LICENCE=LICENSE.txt
set DOSSIER_SOURCE=Code source
if "%LANGUE_BUILD%"=="en" set DOSSIER_SOURCE=Source code
rem Licences, composants tiers et empreintes : ranges dans un sous-dossier.
set DOSSIER_DOC=Documentation technique
if "%LANGUE_BUILD%"=="en" set DOSSIER_DOC=Technical documentation
set SOURCE=%~dp0..
set PAQUET=%~dp0..\_a_transmettre\%NOM%
set ZIP=%~dp0..\_a_transmettre\%NOM%_%VERSION%.zip

echo.
echo ==== %NOM% %VERSION% (%LANGUE_BUILD%) : construction du paquet ====
echo.

rem Python : celui du cabinet (_python.bat) s'il existe, sinon celui installe
rem sur le poste (py ou python). Un tiers peut donc reconstruire l'executable
rem avec le seul contenu du zip, sans _python.bat.
set PY=
if exist "%~dp0..\..\_python.bat" call "%~dp0..\..\_python.bat"
if not defined PY (where py >nul 2>&1 && set "PY=py -3")
if not defined PY (where python >nul 2>&1 && set "PY=python")
if not defined PY (
  echo ERREUR : Python 3 introuvable. Installez-le depuis python.org en cochant "Add python.exe to PATH", puis relancez.
  pause & exit /b 1)

echo [1/8] Composants Python (PyInstaller, python-docx, openpyxl, PyMuPDF, Pillow)...
rem PyMuPDF reste en 1.28 : la 2.0 embarque une analyse de mise en page (modeles
rem onnx, onnxruntime, spacy...) dont le programme ne se sert pas, et qui
rem triplait l'executable (139 Mo au lieu de 39) le 19 octobre 2026.
%PY% -m pip install --quiet --upgrade pyinstaller python-docx openpyxl lxml "pymupdf==1.28.2" pillow
if errorlevel 1 (echo ERREUR : installation des composants. Verifiez la connexion internet. & pause & exit /b 1)

echo [2/8] Icone...
%PY% "%~dp0creer_icone.py" || (echo ERREUR : icone. & pause & exit /b 1)

echo [3/8] Recettes du code (3 series de tests, environ une minute)...
pushd "%~dp0.."
%PY% test_recette.py   > "%~dp0journal_recette.txt" 2>&1
if errorlevel 1 (echo ERREUR : test_recette.py a echoue. Voir _build\journal_recette.txt & popd & pause & exit /b 1)
%PY% test_detection.py >> "%~dp0journal_recette.txt" 2>&1
if errorlevel 1 (echo ERREUR : test_detection.py a echoue. Voir _build\journal_recette.txt & popd & pause & exit /b 1)
%PY% test_paquet.py    >> "%~dp0journal_recette.txt" 2>&1
if errorlevel 1 (echo ERREUR : test_paquet.py a echoue. Voir _build\journal_recette.txt & popd & pause & exit /b 1)
popd
echo        Toutes les recettes sont passees.

echo [4/8] Executable (PyInstaller, deux a cinq minutes)...
%PY% -m PyInstaller --noconfirm --clean --log-level WARN "%~dp0anonymiseur.spec"
if errorlevel 1 (echo ERREUR : PyInstaller. & pause & exit /b 1)
if not exist "%~dp0dist\%NOM%.exe" (echo ERREUR : dist\%NOM%.exe absent. & pause & exit /b 1)

echo [5/8] Assemblage du paquet...
rem Un Anonymiseur.exe encore ouvert (essai precedent) verrouillerait le dossier.
taskkill /f /im %NOM%.exe >nul 2>&1
timeout /t 2 /nobreak >nul
if exist "%PAQUET%" rmdir /s /q "%PAQUET%"
if exist "%PAQUET%" (
  echo ERREUR : impossible de vider %PAQUET%
  echo         Fermez Anonymiseur, l'Explorateur et Excel ouverts sur ce dossier, puis relancez.
  pause & exit /b 1)
mkdir "%PAQUET%"
copy /y "%~dp0dist\%NOM%.exe" "%PAQUET%\" >nul
if errorlevel 1 (echo ERREUR : copie de l'executable impossible. & pause & exit /b 1)
fc /b "%~dp0dist\%NOM%.exe" "%PAQUET%\%NOM%.exe" >nul
if errorlevel 1 (echo ERREUR : l'executable copie ne correspond pas a celui construit. & pause & exit /b 1)
if exist "%~dp0%GUIDE%" (
  copy /y "%~dp0%GUIDE%" "%PAQUET%\" >nul
) else (
  echo        ATTENTION : _build\%GUIDE% absent, le paquet part sans mode d'emploi.
)
mkdir "%PAQUET%\%DOSSIER_DOC%"
copy /y "%~dp0%LICENCE%" "%PAQUET%\%DOSSIER_DOC%\" >nul
if errorlevel 1 (echo ERREUR : _build\%LICENCE% absent : le paquet doit contenir la licence. & pause & exit /b 1)
copy /y "%~dp0LICENSE_AGPL-3.0.txt" "%PAQUET%\%DOSSIER_DOC%\" >nul
if errorlevel 1 (echo ERREUR : _build\LICENSE_AGPL-3.0.txt absent : le paquet doit contenir le texte de la licence. & pause & exit /b 1)
rem Code source correspondant (obligation de la licence AGPL) : les modules du
rem programme, les reglages et les scripts de fabrication. Jamais les tables,
rem journaux, tests sur documents reels ni fichiers personnels.
set SRC=%PAQUET%\%DOSSIER_SOURCE%
mkdir "%SRC%" "%SRC%\_build" "%SRC%\_build\langue\fr" "%SRC%\_build\langue\en"
rem detection_config.json : la copie NETTOYEE des chemins de cette machine
rem (produite par anonymiseur.spec a l'etape 4), jamais celle du cabinet.
copy /y "%~dp0langue\fr\detection_config.json" "%SRC%\detection_config.json" >nul
if errorlevel 1 (echo ERREUR : code source : _build\langue\fr\detection_config.json manquant. & pause & exit /b 1)
for %%f in (anonymiseur.py anonymisation_core.py detection_core.py detection_roles.py detection_classeur.py modele_onglet.py classeur_modele.py textes.py convertir_onglets.py test_paquet.py test_recette.py test_detection.py Anonymiseur.bat Creer_raccourci.bat) do (
  copy /y "%SOURCE%\%%f" "%SRC%\" >nul
  if errorlevel 1 (echo ERREUR : code source : %%f manquant. & pause & exit /b 1)
)
for %%f in (anonymiseur.spec Construire_EXE.bat Build_EXE.bat embarquer_tesseract.py inventaire_paquet.py creer_icone.py version_windows.txt version_windows_en.txt _vide.txt mode_emploi.html user_guide.html icone.png icone.ico LICENCE.txt LICENSE.txt LICENSE_AGPL-3.0.txt) do (
  copy /y "%~dp0%%f" "%SRC%\_build\" >nul
  if errorlevel 1 (echo ERREUR : code source : _build\%%f manquant. & pause & exit /b 1)
)
copy /y "%~dp0langue\fr\langue.txt" "%SRC%\_build\langue\fr\" >nul
copy /y "%~dp0langue\en\langue.txt" "%SRC%\_build\langue\en\" >nul
copy /y "%~dp0langue\en\detection_config.json" "%SRC%\_build\langue\en\" >nul
if exist "%SRC%\_canary_info.txt" (echo ERREUR : _canary_info.txt ne doit jamais partir. & pause & exit /b 1)
%PY% "%~dp0embarquer_tesseract.py" "%PAQUET%\Tesseract"
if errorlevel 1 (echo ERREUR : Tesseract. & pause & exit /b 1)
echo        Paquet assemble avec l'executable construit a l'instant.

echo [6/8] Verification du paquet (l'executable rejoue la recette dans un dossier temporaire)...
"%PAQUET%\%NOM%.exe" --verifier "%~dp0verification_paquet_%LANGUE_BUILD%.txt"
if errorlevel 1 (echo ERREUR : la verification du paquet a echoue. Voir _build\verification_paquet_%LANGUE_BUILD%.txt & pause & exit /b 1)
echo        Paquet verifie.

echo [7/8] Composants tiers et empreintes des fichiers...
%PY% "%~dp0inventaire_paquet.py" paquet "%PAQUET%" %LANGUE_BUILD%
if errorlevel 1 (echo ERREUR : inventaire du paquet. & pause & exit /b 1)

echo [8/8] Zip...
if exist "%ZIP%" del /q "%ZIP%"
if exist "%ZIP%" (echo ERREUR : l'ancien zip est ouvert ailleurs et ne peut pas etre remplace. & pause & exit /b 1)
powershell -NoProfile -Command "Compress-Archive -Path '%PAQUET%' -DestinationPath '%ZIP%' -CompressionLevel Optimal"
if not exist "%ZIP%" (echo ERREUR : zip non cree. Le dossier %PAQUET% est pret, zippez-le a la main. & pause & exit /b 1)
%PY% "%~dp0inventaire_paquet.py" zip "%ZIP%" %LANGUE_BUILD%
if errorlevel 1 (echo ERREUR : empreinte du zip. & pause & exit /b 1)

echo.
echo ==== TERMINE ====
echo   Dossier : %PAQUET%
echo   Zip     : %ZIP%
echo   Empreinte du zip a publier avec le lien : %ZIP%.sha256.txt
echo.
echo   Le zip contient le code source (licence AGPL), mais ni vos tables, ni vos journaux.
echo   Le confrere le decompresse ou il veut et double-clique sur %NOM%.exe.
echo.
pause
