@echo off
rem ====================================================================
rem  Build_EXE.bat : fabrique la version ANGLAISE (Anonymizator).
rem  Meme chaine que Construire_EXE.bat, en anglais : interface, classeur,
rem  balises, guide (User guide.pdf). Resultat :
rem     _a_transmettre\Anonymizator\  et  _a_transmettre\Anonymizator_<version>.zip
rem ====================================================================
set LANGUE_BUILD=en
call "%~dp0Construire_EXE.bat"
