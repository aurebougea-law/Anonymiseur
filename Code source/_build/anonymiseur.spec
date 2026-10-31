# -*- mode: python ; coding: utf-8 -*-
# Recette PyInstaller de l'executable unique.
# Lancement : voir Construire_EXE.bat (il se place dans le dossier _build).

import os
import sys

ICI = os.path.abspath(SPECPATH)                 # dossier _build
SOURCE = os.path.dirname(ICI)                   # dossier du programme (les .py)
# Langue de la construction : variable d'environnement LANGUE_BUILD ("fr" par
# defaut, "en" pour Anonymizator). Build_EXE.bat la positionne.
LANGUE_BUILD = os.environ.get("LANGUE_BUILD", "fr").lower()
NOM = "Anonymizator" if LANGUE_BUILD == "en" else "Anonymiseur"
# Reglages de detection : francais = ceux du programme ; anglais = _build/langue/en/
CONFIG_DETECTION = (os.path.join(ICI, "langue", "en", "detection_config.json") if LANGUE_BUILD == "en"
                    else os.path.join(SOURCE, "detection_config.json"))

# Les reglages francais sont ceux du cabinet : avant de les embarquer, on en
# retire les chemins propres a cette machine (Tesseract, tessdata). La copie
# nettoyee, _build/langue/fr/detection_config.json, est aussi celle qui part
# dans le code source du paquet.
if LANGUE_BUILD != "en":
    import json as _json
    with open(CONFIG_DETECTION, encoding="utf-8") as _f:
        _reglages = _json.load(_f)
    for _cle in ("tesseract", "tessdata"):
        if isinstance(_reglages.get("ocr"), dict):
            _reglages["ocr"][_cle] = ""
    # Noms des avocats du cabinet : jamais dans l'executable ni dans le code
    # source diffuse (8 octobre 2026).
    if isinstance(_reglages.get("cabinet"), dict):
        _reglages["cabinet"]["avocats"] = []
    CONFIG_DETECTION = os.path.join(ICI, "langue", "fr", "detection_config.json")
    os.makedirs(os.path.dirname(CONFIG_DETECTION), exist_ok=True)
    with open(CONFIG_DETECTION, "w", encoding="utf-8") as _f:
        _json.dump(_reglages, _f, ensure_ascii=False, indent=2)

FICHIER_VERSION = os.path.join(ICI, "version_windows_en.txt" if LANGUE_BUILD == "en"
                               else "version_windows.txt")

# python-docx lit ses gabarits (default.docx, default-header.xml...) sur le
# disque : il faut les livrer dans l'executable.
import docx
from PyInstaller.utils.hooks import collect_all
# PyMuPDF (lecture des PDF) : on embarque tout le paquet, bibliotheques
# natives comprises, sans dependre du crochet de PyInstaller.
PYMUPDF_DATAS, PYMUPDF_BIN, PYMUPDF_HIDDEN = collect_all("pymupdf")
DOCX_TEMPLATES = os.path.join(os.path.dirname(docx.__file__), "templates")

a = Analysis(
    [os.path.join(SOURCE, "anonymiseur.py")],
    pathex=[SOURCE],
    binaries=PYMUPDF_BIN,
    datas=[
        (CONFIG_DETECTION, "."),
        (os.path.join(ICI, "langue", LANGUE_BUILD, "langue.txt"), "."),
        (os.path.join(ICI, "icone.png"), "."),
        (os.path.join(ICI, "icone.ico"), "."),
        (DOCX_TEMPLATES, os.path.join("docx", "templates")),
        # python-docx atteint ses gabarits par « parts/../templates » : le
        # dossier parts doit exister sur le disque, d'ou ce fichier vide.
        (os.path.join(ICI, "_vide.txt"), os.path.join("docx", "parts")),
    ] + PYMUPDF_DATAS,
    hiddenimports=[
        "anonymisation_core", "detection_core", "detection_roles",
        "detection_classeur", "modele_onglet", "classeur_modele", "test_paquet",
        "docx", "openpyxl", "lxml", "lxml.etree", "pymupdf",
        "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox",
        "tkinter.simpledialog",
    ] + PYMUPDF_HIDDEN,
    hookspath=[],
    runtime_hooks=[],
    # Hors de l'executable : les recettes de developpement, le convertisseur
    # d'anciens onglets, et les lecteurs PDF de secours de detection_core
    # (pdfplumber, pypdf...) que PyInstaller embarquerait sinon par precaution,
    # alors que PyMuPDF est livre et suffit (8 octobre 2026).
    excludes=["test_recette", "test_detection", "convertir_onglets",
              "pdfplumber", "pdfminer", "pypdf", "pypdfium2",
              "cryptography", "bs4",
              "matplotlib", "numpy", "pandas", "PIL",
              "pytest", "unittest", "IPython", "jupyter"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=NOM,
    debug=False,
    strip=False,
    upx=False,
    console=False,                                   # pas de fenetre noire
    icon=os.path.join(ICI, "icone.ico") if sys.platform == "win32" else None,
    version=FICHIER_VERSION if sys.platform == "win32" else None,
)
