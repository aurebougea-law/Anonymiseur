"""Licence : GNU Affero General Public License v3 (LICENSE_AGPL-3.0.txt).

inventaire_paquet.py
--------------------
Etape de FABRICATION uniquement (appelee par Construire_EXE.bat) : rien de ce
script n'est embarque dans l'executable, rien ne change pour l'utilisateur.

Deux usages :

  py inventaire_paquet.py paquet <dossier_du_paquet> <fr|en>
      Ecrit dans le paquet, avant le zip :
        - COMPOSANTS_TIERS.txt (THIRD_PARTY_COMPONENTS.txt) : les composants
          embarques, avec la version REELLEMENT installee au moment de la
          fabrication et leur licence ;
        - EMPREINTES_SHA256.txt (CHECKSUMS_SHA256.txt) : l'empreinte de chaque
          fichier du paquet, chemin relatif complet, pour reperer un fichier
          abime ou modifie. Le fichier ne se liste pas lui-meme.
      Les deux sont ranges dans le sous-dossier « Documentation technique »
      (« Technical documentation »), avec les deux textes de licence.

  py inventaire_paquet.py zip <chemin_du_zip> <fr|en>
      Ecrit A COTE du zip (hors du zip) <zip>.sha256.txt : l'empreinte du zip
      a publier avec le lien de telechargement. C'est elle qui permet a un
      confrere de verifier que le zip recu est bien celui qui a ete diffuse
      (une empreinte placee DANS le zip ne le prouve pas : qui modifie le zip
      peut la recalculer).
"""

import datetime
import hashlib
import os
import subprocess
import sys

TEXTES = {
    "fr": {
        "dossier_doc": "Documentation technique",
        "composants": "COMPOSANTS_TIERS.txt",
        "empreintes": "EMPREINTES_SHA256.txt",
        "titre_comp": "COMPOSANTS TIERS ET LICENCES",
        "intro_comp": ("{nom} est diffusé sous licence AGPL v3 (LICENSE_AGPL-3.0.txt). Il embarque "
                       "les composants libres suivants, dans la version installée lors de la "
                       "fabrication du {date}. Chacun reste sous sa propre licence."),
        "dans_exe": "DANS L'EXÉCUTABLE {nom}.exe",
        "dossier_tess": "DANS LE DOSSIER Tesseract",
        "tess": ("Tesseract OCR {v} : reconnaissance de caractères des PDF scannés. Licence "
                 "Apache 2.0 (texte joint : Tesseract\\LICENCE_Tesseract.txt). "
                 "https://github.com/tesseract-ocr/tesseract"),
        "packs": "Packs de langue (tessdata, licence Apache 2.0) : {l}",
        "dll": ("Bibliothèques livrées avec la version Windows de Tesseract, chacune sous sa "
                "propre licence libre (voir le site de chaque projet) :"),
        "fabrication": "OUTIL DE FABRICATION (non embarqué comme bibliothèque)",
        "pyinstaller": ("PyInstaller {v} : assemble l'exécutable. Licence : {l} ; l'exception de sa "
                        "licence autorise la diffusion des programmes qu'il fabrique. "
                        "https://pyinstaller.org"),
        "inconnue": "version non déterminée",
        "titre_emp": "EMPREINTES SHA-256 DU PAQUET",
        "intro_emp": ("Fabrication du {date}. Chaque ligne : empreinte, puis chemin du fichier dans "
                      "le dossier {nom}.\nPour vérifier un fichier sous Windows, dans une invite "
                      "de commandes ouverte dans ce dossier :\n"
                      "    certutil -hashfile \"chemin\\du\\fichier\" SHA256\n"
                      "et comparez le résultat à la ligne correspondante (majuscules et "
                      "minuscules indifférentes).\nPour vérifier que le zip reçu est bien "
                      "l'original, comparez l'empreinte du zip à celle publiée avec le lien "
                      "de téléchargement."),
        "titre_zip": "Empreinte SHA-256 du fichier à publier avec le lien de téléchargement :",
        "verif_zip": "Vérification sous Windows :  certutil -hashfile \"{z}\" SHA256",
    },
    "en": {
        "dossier_doc": "Technical documentation",
        "composants": "THIRD_PARTY_COMPONENTS.txt",
        "empreintes": "CHECKSUMS_SHA256.txt",
        "titre_comp": "THIRD-PARTY COMPONENTS AND LICENCES",
        "intro_comp": ("{nom} is distributed under the AGPL v3 licence (LICENSE_AGPL-3.0.txt). It "
                       "bundles the following free components, in the version installed when it "
                       "was built on {date}. Each remains under its own licence."),
        "dans_exe": "INSIDE THE EXECUTABLE {nom}.exe",
        "dossier_tess": "IN THE Tesseract FOLDER",
        "tess": ("Tesseract OCR {v}: text recognition for scanned PDFs. Apache 2.0 licence "
                 "(text included: Tesseract\\LICENCE_Tesseract.txt). "
                 "https://github.com/tesseract-ocr/tesseract"),
        "packs": "Language packs (tessdata, Apache 2.0 licence): {l}",
        "dll": ("Libraries shipped with the Windows build of Tesseract, each under its own free "
                "licence (see each project's website):"),
        "fabrication": "BUILD TOOL (not bundled as a library)",
        "pyinstaller": ("PyInstaller {v}: builds the executable. Licence {l}; its licence exception "
                        "allows distribution of the programs it builds. https://pyinstaller.org"),
        "inconnue": "version not determined",
        "titre_emp": "SHA-256 CHECKSUMS OF THE PACKAGE",
        "intro_emp": ("Built on {date}. Each line: checksum, then path of the file in the {nom} "
                      "folder.\nTo check a file on Windows, in a command prompt opened in this "
                      "folder:\n    certutil -hashfile \"path\\to\\file\" SHA256\n"
                      "and compare the result with the matching line (case does not matter).\n"
                      "To check that the zip you received is the original, compare the zip's "
                      "checksum with the one published with the download link."),
        "titre_zip": "SHA-256 checksum of the file, to publish with the download link:",
        "verif_zip": "Check on Windows:  certutil -hashfile \"{z}\" SHA256",
    },
}

# Composants Python embarques dans l'executable : (paquet pip, role FR, role EN,
# licence de repli si les metadonnees n'en donnent pas, site).
COMPOSANTS = [
    ("python-docx", "lecture et écriture des documents Word", "reading and writing Word documents",
     "MIT", "https://github.com/python-openxml/python-docx"),
    ("openpyxl", "lecture et écriture des classeurs Excel", "reading and writing Excel workbooks",
     "MIT", "https://openpyxl.readthedocs.io"),
    ("lxml", "traitement du XML des documents Word", "processing the XML of Word documents",
     "BSD-3-Clause", "https://lxml.de"),
    ("PyMuPDF", "lecture des PDF", "reading PDFs",
     "AGPL-3.0 (ou licence commerciale Artifex)", "https://github.com/pymupdf/pymupdf"),
]


def _metadonnees(paquet):
    """(version, licence) lues dans les metadonnees du paquet installe."""
    from importlib import metadata
    try:
        md = metadata.metadata(paquet)
        version = metadata.version(paquet)
    except Exception:
        return None, None
    licence = md.get("License-Expression") or ""
    if not licence:
        classes = [c.split("::")[-1].strip() for c in (md.get_all("Classifier") or [])
                   if c.startswith("License ::")]
        licence = ", ".join(classes)
    if not licence:
        brut = (md.get("License") or "").strip()
        licence = brut if brut and len(brut) <= 120 and "\n" not in brut else ""
    return version, licence or None


def _version_tesseract(dossier_tess):
    exe = os.path.join(dossier_tess, "tesseract.exe")
    if not os.path.isfile(exe):
        return None
    options = {}
    if sys.platform == "win32":
        options["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        r = subprocess.run([exe, "--version"], capture_output=True, text=True,
                           timeout=30, **options)
        ligne = ((r.stdout or "") + (r.stderr or "")).strip().splitlines()
        return ligne[0].replace("tesseract", "").strip() if ligne else None
    except Exception:
        return None


def ecrire_composants(paquet, langue, nom):
    T = TEXTES[langue]
    date = datetime.date.today().isoformat()
    lignes = [T["titre_comp"], "=" * len(T["titre_comp"]), "",
              T["intro_comp"].format(nom=nom, date=date), "",
              T["dans_exe"].format(nom=nom), "-" * len(T["dans_exe"].format(nom=nom))]
    v = sys.version.split()[0]
    lignes.append(f"- Python {v} : "
                  + ("interpréteur" if langue == "fr" else "interpreter")
                  + ". Licence : PSF-2.0. https://www.python.org")
    for paquet_pip, role_fr, role_en, repli, site in COMPOSANTS:
        version, licence = _metadonnees(paquet_pip)
        role = role_fr if langue == "fr" else role_en
        lignes.append(f"- {paquet_pip} {version or T['inconnue']} : {role}. "
                      f"Licence : {licence or repli}. {site}")

    dossier_tess = os.path.join(paquet, "Tesseract")
    if os.path.isdir(dossier_tess):
        lignes += ["", T["dossier_tess"], "-" * len(T["dossier_tess"])]
        lignes.append("- " + T["tess"].format(v=_version_tesseract(dossier_tess) or T["inconnue"]))
        tessdata = os.path.join(dossier_tess, "tessdata")
        if os.path.isdir(tessdata):
            packs = sorted(f for f in os.listdir(tessdata) if f.endswith(".traineddata"))
            if packs:
                lignes.append("- " + T["packs"].format(l=", ".join(packs)))
        dll = sorted(f for f in os.listdir(dossier_tess) if f.lower().endswith(".dll"))
        if dll:
            lignes.append("- " + T["dll"])
            ligne = "    "
            for f in dll:
                if len(ligne) + len(f) + 2 > 78:
                    lignes.append(ligne.rstrip(", "))
                    ligne = "    "
                ligne += f + ", "
            lignes.append(ligne.rstrip(", "))

    # Licence de PyInstaller : GPL v2 ou ulterieure AVEC exception (ses
    # metadonnees ne la resument pas en une ligne exploitable).
    version, _ = _metadonnees("pyinstaller")
    lignes += ["", T["fabrication"], "-" * len(T["fabrication"]),
               "- " + T["pyinstaller"].format(v=version or T["inconnue"],
                                              l="GPL-2.0-or-later")]
    if langue == "en":                          # typographie anglaise
        lignes = [l.replace(" : ", ": ") for l in lignes]
    os.makedirs(os.path.join(paquet, T["dossier_doc"]), exist_ok=True)
    chemin = os.path.join(paquet, T["dossier_doc"], T["composants"])
    with open(chemin, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(lignes) + "\n")
    return chemin


def empreinte(chemin):
    h = hashlib.sha256()
    with open(chemin, "rb") as f:
        for bloc in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloc)
    return h.hexdigest().upper()


def ecrire_empreintes(paquet, langue, nom):
    T = TEXTES[langue]
    os.makedirs(os.path.join(paquet, T["dossier_doc"]), exist_ok=True)
    sortie = os.path.join(paquet, T["dossier_doc"], T["empreintes"])
    fichiers = []
    for racine, _, noms in os.walk(paquet):
        for n in noms:
            complet = os.path.join(racine, n)
            if os.path.abspath(complet) == os.path.abspath(sortie):
                continue
            fichiers.append(complet)
    fichiers.sort(key=lambda c: os.path.relpath(c, paquet).lower())
    lignes = [T["titre_emp"], "=" * len(T["titre_emp"]), "",
              T["intro_emp"].format(date=datetime.date.today().isoformat(), nom=nom), ""]
    for c in fichiers:
        relatif = os.path.relpath(c, paquet).replace("/", "\\")
        lignes.append(f"{empreinte(c)}  {relatif}")
    with open(sortie, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(lignes) + "\n")
    return sortie, len(fichiers)


def ecrire_empreinte_zip(chemin_zip, langue):
    T = TEXTES[langue]
    nom = os.path.basename(chemin_zip)
    sortie = chemin_zip + ".sha256.txt"
    with open(sortie, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(f"{T['titre_zip']}\n\n{empreinte(chemin_zip)}  {nom}\n\n"
                f"{T['verif_zip'].format(z=nom)}\n")
    return sortie


def main(argv):
    if len(argv) != 4 or argv[1] not in ("paquet", "zip") or argv[3] not in TEXTES:
        print(__doc__)
        return 2
    mode, cible, langue = argv[1], argv[2], argv[3]
    if mode == "paquet":
        if not os.path.isdir(cible):
            print(f"ERREUR : dossier introuvable : {cible}")
            return 1
        nom = os.path.basename(os.path.normpath(cible))
        print("        " + ecrire_composants(cible, langue, nom))
        chemin, n = ecrire_empreintes(cible, langue, nom)
        print(f"        {chemin} ({n} fichiers)")
    else:
        if not os.path.isfile(cible):
            print(f"ERREUR : zip introuvable : {cible}")
            return 1
        print("        " + ecrire_empreinte_zip(cible, langue))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
