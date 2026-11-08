"""
embarquer_tesseract.py  <dossier_cible>
---------------------------------------
Copie Tesseract (lecture des PDF scannes) dans le paquet a transmettre, pour
que les confreres n'aient rien a installer.

Cherche tesseract.exe et le pack de langue francais aux endroits habituels,
puis copie seulement ce qui sert : l'executable, ses DLL, fra.traineddata
(et osd/eng s'ils sont la), et la licence. Tesseract est sous licence
Apache 2.0, qui autorise cette redistribution a condition de joindre la
licence : c'est fait ici.

Si rien n'est trouve, le script demande le chemin.
"""

import os
import sys
import json
import shutil
import pathlib

ICI = pathlib.Path(__file__).parent
SOURCE = ICI.parent

NOTICE_APACHE = """Tesseract OCR
=============
Ce dossier contient une copie de Tesseract OCR (https://github.com/tesseract-ocr/tesseract),
utilisee par Anonymiseur pour lire les PDF scannes. Tesseract est un logiciel
libre distribue sous licence Apache 2.0, dont le texte est reproduit ci-dessous.
Les packs de langue (tessdata) sont egalement sous licence Apache 2.0.
Anonymiseur n'en est ni l'auteur ni l'editeur : il l'appelle tel quel.

"""


def chercher_tesseract():
    candidats = []
    try:
        cfg = json.load(open(SOURCE / "detection_config.json", encoding="utf-8"))
        candidats.append(cfg.get("ocr", {}).get("tesseract", ""))
    except Exception:
        cfg = {}
    candidats.append(shutil.which("tesseract") or "")
    # Tesseract livre avec un paquet deja installe (reconstruction depuis le
    # dossier « Code source » d'un zip recu) : a cote du dossier du code.
    candidats.append(str(SOURCE.parent / "Tesseract" / "tesseract.exe"))
    candidats += [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                  os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
                  os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
                  "/usr/bin/tesseract"]
    for c in candidats:
        if c and os.path.isfile(c):
            return pathlib.Path(c), cfg
    return None, cfg


def _dossiers_tessdata(exe, cfg):
    candidats = [cfg.get("ocr", {}).get("tessdata", "")]
    if exe:
        candidats.append(str(exe.parent / "tessdata"))
    candidats += [r"C:\Program Files\Tesseract-OCR\tessdata",
                  "/usr/share/tesseract-ocr/5/tessdata", "/usr/share/tesseract-ocr/4.00/tessdata"]
    return [pathlib.Path(c) for c in candidats if c and os.path.isdir(c)]


def chercher_tessdata(exe, cfg):
    for d in _dossiers_tessdata(exe, cfg):
        if (d / "fra.traineddata").is_file():
            return d
    return None


def chercher_pack(exe, cfg, langue):
    """Premier dossier qui possede le pack demande (fra, eng, osd...)."""
    for d in _dossiers_tessdata(exe, cfg):
        if (d / f"{langue}.traineddata").is_file():
            return d / f"{langue}.traineddata"
    return None


def demander(question):
    print()
    print(question)
    reponse = input("> ").strip().strip('"')
    return reponse


def main():
    if len(sys.argv) < 2:
        print("Usage : embarquer_tesseract.py <dossier_cible>")
        return 2
    cible = pathlib.Path(sys.argv[1])

    exe, cfg = chercher_tesseract()
    while not exe:
        rep = demander("Tesseract est introuvable. Tapez le chemin complet de tesseract.exe\n"
                       "(ou laissez vide pour construire le paquet SANS lecture des PDF scannes) :")
        if not rep:
            print("Paquet construit sans Tesseract : les PDF scannes ne seront pas lus.")
            return 0
        if os.path.isfile(rep):
            exe = pathlib.Path(rep)

    tessdata = chercher_tessdata(exe, cfg)
    while not tessdata:
        rep = demander("Le pack de langue fra.traineddata est introuvable. Tapez le chemin du\n"
                       "dossier tessdata qui le contient (ou vide pour abandonner l'OCR) :")
        if not rep:
            print("Paquet construit sans Tesseract : les PDF scannes ne seront pas lus.")
            return 0
        if os.path.isfile(os.path.join(rep, "fra.traineddata")):
            tessdata = pathlib.Path(rep)

    if cible.exists():
        shutil.rmtree(cible)
    cible.mkdir(parents=True)
    (cible / "tessdata").mkdir()

    copies = 0
    shutil.copy2(exe, cible / exe.name)
    copies += 1
    for f in exe.parent.iterdir():
        if f.suffix.lower() in (".dll",) or f.name.upper().startswith("LICENSE"):
            shutil.copy2(f, cible / f.name)
            copies += 1
    packs = []
    for langue in ("fra", "osd", "eng"):
        f = chercher_pack(exe, cfg, langue)
        if f:
            shutil.copy2(f, cible / "tessdata" / f.name)
            copies += 1
            packs.append(langue)
    print("Packs de langue embarques :", ", ".join(packs))
    if "eng" not in packs:
        print("ATTENTION : eng.traineddata introuvable : la version anglaise ne lira pas les PDF scannes.")
    # Fichiers de configuration de tessdata, s'ils existent (petits)
    for sous in ("configs", "tessconfigs"):
        d = tessdata / sous
        if d.is_dir():
            shutil.copytree(d, cible / "tessdata" / sous)

    licence = ""
    for f in list(exe.parent.glob("LICENSE*")) + list(exe.parent.glob("doc/LICENSE*")):
        try:
            licence = f.read_text(encoding="utf-8", errors="replace")
            break
        except Exception:
            pass
    if not licence:
        licence = ("Texte de la licence Apache 2.0 : https://www.apache.org/licenses/LICENSE-2.0\n")
    (cible / "LICENCE_Tesseract.txt").write_text(NOTICE_APACHE + licence, encoding="utf-8")

    taille = sum(f.stat().st_size for f in cible.rglob("*") if f.is_file()) / 1e6
    print(f"Tesseract embarque : {copies} fichier(s), {taille:.0f} Mo, depuis {exe.parent}")
    print(f"Pack de langue : {tessdata}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
