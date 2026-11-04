@echo off
rem Cree un raccourci "Anonymiseur" avec l'icone du programme (bouclier), qui
rem lance Anonymiseur.bat. Le raccourci est pose dans ce dossier ET sur le Bureau.
rem Un .bat ne peut pas porter d'icone : le raccourci, si.
cd /d "%~dp0"
if not exist "%~dp0_build\icone.ico" (
  echo L'icone n'existe pas encore : lancez d'abord _build\Construire_EXE.bat.
  pause & exit /b 1
)
powershell -NoProfile -Command ^
  "$w = New-Object -ComObject WScript.Shell;" ^
  "foreach ($d in @('%~dp0', [Environment]::GetFolderPath('Desktop'))) {" ^
  "  $s = $w.CreateShortcut((Join-Path $d 'Anonymiseur.lnk'));" ^
  "  $s.TargetPath = '%~dp0Anonymiseur.bat';" ^
  "  $s.WorkingDirectory = '%~dp0';" ^
  "  $s.IconLocation = '%~dp0_build\icone.ico,0';" ^
  "  $s.Description = 'Anonymiseur : anonymisation reversible avant depot dans une IA';" ^
  "  $s.WindowStyle = 7;" ^
  "  $s.Save(); Write-Host ('Raccourci cree : ' + $s.FullName) }"
echo.
pause
