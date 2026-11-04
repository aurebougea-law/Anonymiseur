"""Licence : GNU Affero General Public License v3 (LICENSE_AGPL-3.0.txt).

test_paquet.py
--------------
Recette du paquet distribue : simule un premier demarrage dans un dossier vide,
puis rejoue le cycle complet d'un utilisateur sur le dossier fictif « Exemple ».

Ne touche a aucun fichier reel : tout se passe dans un dossier temporaire.
Lancement :  python test_paquet.py
"""

import os
import sys
import json
import shutil
import pathlib
import tempfile

# Dans l'executable, les fichiers livres (detection_config.json) sont dans le
# dossier temporaire de PyInstaller ; en script, a cote de ce fichier.
SCRIPT_DIR = pathlib.Path(getattr(sys, "_MEIPASS", pathlib.Path(__file__).parent))
sys.path.insert(0, str(SCRIPT_DIR))

import anonymisation_core as coeur
import detection_core as det
import detection_classeur as classeur
import detection_roles as roles
import modele_onglet as modele
import classeur_modele

# Langue du paquet (langue.txt livre avec l'executable ; francais par defaut)
LANGUE = "fr"
for _c in (pathlib.Path(getattr(sys, "executable", "")).parent / "langue.txt", SCRIPT_DIR / "langue.txt"):
    try:
        LANGUE = _c.read_text(encoding="utf-8").strip().lower()[:2] or "fr"
        break
    except OSError:
        pass
for _m in (modele, classeur, classeur_modele, coeur):
    _m.LANGUE = LANGUE
EX = classeur_modele.nom_onglet_exemple()
GUIDE = classeur_modele.nom_onglet_mode_emploi()
MOT_CONFIDENTIEL = "confidential" if LANGUE == "en" else "confidentiel"
MOT_NOUVEAU = "New matter" if LANGUE == "en" else "Nouveau dossier"

OK, KO = [], []


def verifier(condition, libelle):
    (OK if condition else KO).append(libelle)
    print(f"  [{'OK ' if condition else 'KO '}] {libelle}")


COMPTE_RENDU = None          # fichier ou recopier la sortie (mode --verifier)
_print_console = print


def print(*a, **k):          # recopie chaque ligne dans le compte rendu
    _print_console(*a, **k)
    if COMPTE_RENDU is not None:
        _print_console(*a, **k, file=COMPTE_RENDU)
        COMPTE_RENDU.flush()


def main(fichier_compte_rendu=None):
    """fichier_compte_rendu : fichier texte ou recopier le compte rendu."""
    global COMPTE_RENDU
    if fichier_compte_rendu:
        COMPTE_RENDU = open(fichier_compte_rendu, "w", encoding="utf-8")
    racine = pathlib.Path(tempfile.mkdtemp(prefix="recette_paquet_"))
    try:
        print("\n1. Premier demarrage dans un dossier vide")
        config_det = det.charger_config_detection(SCRIPT_DIR / "detection_config.json")
        modele.CONFIG = config_det
        table = racine / "Correspondances.xlsx"
        classeur_modele.creer_classeur_modele(str(table), config_det)
        verifier(table.exists(), "le classeur modele est cree")
        try:
            classeur_modele.creer_classeur_modele(str(table), config_det)
            verifier(False, "refus d'ecraser un classeur existant")
        except FileExistsError:
            verifier(True, "refus d'ecraser un classeur existant")

        onglets = coeur.lister_onglets(str(table))
        verifier(onglets == [EX], f"le programme ne voit que l'onglet {EX} ({GUIDE} masque)")
        modele_ex, entrees_ex = modele.lire_onglet(str(table), EX)
        verifier(modele_ex == "par_partie", "l'onglet Exemple suit le modele par partie")
        verifier(len(entrees_ex) == len(classeur_modele.ENTREES_EXEMPLE),
                 f"{len(entrees_ex)} lignes d'exemple lues")

        import openpyxl
        wb = openpyxl.load_workbook(str(table), read_only=True)
        me = wb[GUIDE]
        textes = [str(c) for row in me.iter_rows(values_only=True) for c in row if c]
        wb.close()
        verifier(any(MOT_CONFIDENTIEL in t.lower() for t in textes),
                 "le mode d'emploi rappelle la confidentialite du classeur")
        verifier(any(MOT_NOUVEAU in t for t in textes),
                 "le mode d'emploi renvoie au bouton Nouveau dossier")

        # Parties du menu (reglage parties_menu : demandeur/defendeur personne
        # physique ou morale, societe, personne physique, personne morale)
        menu = modele.parties_menu(config_det)
        verifier(len(menu) == 3, f"{len(menu)} parties du menu lues dans les reglages")
        codes_menu = modele.codes_parties_menu(config_det)
        plus_grand = max(menu, key=lambda r: int(r.get("numeros") or 1))
        verifier(f"{plus_grand['code']}{plus_grand['numeros']}" in codes_menu
                 and f"{plus_grand['code']}{plus_grand['numeros'] + 1}" not in codes_menu,
                 f"les numeros proposes vont de 1 a {plus_grand['numeros']} pour {plus_grand['code']}")
        _, entrees_mod = modele.lire_onglet(str(table), classeur_modele.nom_onglet_modele())
        balises_mod = {e["balise"] for e in entrees_mod}
        verifier(all(f"[{r['code']}1]" in balises_mod for r in menu),
                 "le modele complet contient le n° 1 de chaque partie du menu")
        physiques = [r["code"] for r in menu if r.get("personne_physique")]
        morales = [r["code"] for r in menu if not r.get("personne_physique")]
        sfx_siret = modele.suffixe_de("SIRET", config_det)
        verifier(physiques and morales and sfx_siret
                 and all(f"[{c}1_{sfx_siret}]" not in balises_mod for c in physiques)
                 and all(f"[{c}1_{sfx_siret}]" in balises_mod for c in morales),
                 "une personne physique n'a pas de SIRET dans le modele, une personne morale en a un")
        wb = openpyxl.load_workbook(str(table), read_only=True)
        feuille_roles = modele._nom_feuille_roles(classeur_modele.nom_onglet_modele())
        verifier(feuille_roles in wb.sheetnames, "chaque onglet a sa feuille Roles masquee")
        lignes_roles = list(wb[feuille_roles].iter_rows(min_col=4, max_col=6, values_only=True))
        wb.close()
        parties_menu_excel = {str(l[0]) for l in lignes_roles if l and l[0]}
        balises_menu_excel = {str(l[2]) for l in lignes_roles if l and len(l) > 2 and l[2]}
        verifier(all(c in parties_menu_excel for c in codes_menu),
                 f"le menu Partie du classeur propose les {len(codes_menu)} parties numerotees")
        verifier(all(f"[{c}]" in balises_menu_excel for c in codes_menu),
                 "le menu Balise du classeur propose chacune de ces parties")

        config = {
            "tables": [{"nom": "Mes correspondances", "fichier": str(table)}],
            "dossier_journaux": str(racine / "Journaux"),
            "masquer_proprietes": True,
            "onglets_masques": coeur.ONGLETS_MASQUES,
            "suffixes_onglets_masques": coeur.SUFFIXES_ONGLETS_MASQUES,
        }
        os.makedirs(config["dossier_journaux"])

        print("\n2. Document d'exemple : anonymiser, controler, restituer")
        exemple = racine / "Exemples" / "Conclusions d'exemple.docx"
        exemple.parent.mkdir()
        classeur_modele.creer_document_exemple(str(exemple))
        entrees = coeur.charger_entrees(str(table), EX)

        classement = coeur.analyser_pertinence(str(exemple), config["tables"], config)
        verifier(classement and classement[0]["onglet"] == EX and classement[0]["trouvees"] > 5,
                 "l'onglet Exemple est reconnu comme pertinent pour le document")

        res = coeur.anonymiser_document(str(exemple), entrees, config, "Mes correspondances", EX)
        texte = coeur.texte_du_document(res["chemin_sortie"]).lower()
        verifier(res["total"] >= 15, f"{res['total']} remplacements")
        for mot in classeur_modele.MOTS_TEST[LANGUE]:
            verifier(mot not in texte, f"« {mot} » n'apparait plus dans le document anonymise")
        verifier("[DATE]" in coeur.texte_du_document(res["chemin_sortie"]),
                 "le marqueur de redaction [DATE] est laisse intact")
        verifier(res.get("total_proprietes", 0) >= 1, "l'auteur du fichier a ete masque")

        verifier(os.path.basename(res["chemin_sortie"]).endswith(coeur.suffixe("anonymise") + ".docx"),
                 "le fichier anonymise porte le suffixe de la langue (" + coeur.suffixe("anonymise") + ")")
        # Un document nomme a l'autre langue est reconnu (echange FR <-> EN)
        autre = "_ANONYMISE" if LANGUE == "en" else "_ANONYMISED"
        verifier(coeur._base_propre("x" + autre + ".docx") == "x" and coeur._base_propre("x_RESTORED.docx") == "x"
                 and coeur._base_propre("x_RESTITUE.docx") == "x",
                 "les suffixes des deux langues sont retires du nom de base (journal retrouve)")
        ctrl = coeur.controler(res["chemin_sortie"], entrees, config)
        verifier(not ctrl["alerte"] and ctrl["etat"] == coeur.ETAT_ANONYMISE,
                 "le controle juge le document anonymise bon pour une IA")
        ctrl_orig = coeur.controler(str(exemple), entrees, config)
        verifier(ctrl_orig["alerte"], "le controle alerte sur le document d'origine (noms en clair)")

        rest = coeur.desanonymiser_avec_journal(res["chemin_sortie"], res["chemin_journal"])
        verifier(coeur.texte_du_document(rest["chemin_sortie"]) == coeur.texte_du_document(str(exemple)),
                 "la restitution avec journal redonne exactement le texte d'origine")
        ctrl_rest = coeur.controler(rest["chemin_sortie"], entrees, config)
        verifier(ctrl_rest["etat"] == coeur.ETAT_RESTITUE and not ctrl_rest["alerte"],
                 "le controle juge le document restitue complet")

        rest2 = coeur.desanonymiser_sans_journal(res["chemin_sortie"], entrees, confirme=True,
                                                 config=config)
        t2 = coeur.texte_du_document(rest2["chemin_sortie"])
        verifier("[" not in t2.replace("[DATE]", ""), "la restitution sans journal ne laisse aucune balise")
        verifier(classeur_modele.PREMIERE_FORME[LANGUE] in t2, "la restitution sans journal utilise la premiere forme")

        print("\n3. Pre-detection sur le document d'exemple, injection dans un nouveau dossier")
        nouveau = modele.creer_onglet(str(table), "Nouveau dossier", config_det)
        verifier("Nouveau dossier" in coeur.lister_onglets(str(table)), "un nouveau dossier (onglet) se cree")
        entites, infos, _, source = det.analyser_document(
            str(exemple), str(table), nouveau, str(SCRIPT_DIR / "detection_config.json"),
            type_doc=roles.AUTOMATIQUE)
        verifier(source == "texte" and len(entites) >= (5 if LANGUE == "fr" else 3), f"{len(entites)} termes detectes")
        verifier(infos.get("type") == "contentieux", "le document est reconnu comme contentieux")
        termes = " | ".join(e.canonique for e in entites).lower()
        attendus = ("horizon", "bleu marine", "lumivrac", "01234") if LANGUE == "fr" else ("horizon", "lumivrac", "01234")
        for mot in attendus:
            verifier(mot in termes, f"la detection trouve « {mot} »")

        sortie = racine / "Exemples" / "_Detection.xlsx"
        _, existantes = modele.lire_onglet(str(table), nouveau)
        classeur.ecrire_classeur_detection(str(sortie), entites, infos, config_det, str(exemple),
                                           "Mes correspondances", nouveau, existantes)
        verifier(sortie.exists(), "le classeur de detection a relire est ecrit")

        # L'utilisateur « relit » : on retient tout.
        wb = openpyxl.load_workbook(str(sortie))
        ws = wb.active
        col = entete_r = None
        for r in range(1, 10):
            vals = [c.value for c in ws[r]]
            c2 = next((i for i, v in enumerate(vals, 1)
                       if v and str(v).strip().lower() in ("garder ?", "à retenir ?", "keep?")), None)
            if c2:
                col, entete_r = c2, r
                break
        verifier(col is not None, "le classeur de detection a une colonne « A retenir ? »")
        for r in range(entete_r + 1, ws.max_row + 1):
            if ws.cell(row=r, column=1).value and ws.cell(row=r, column=2).value:
                ws.cell(row=r, column=col, value="YES" if LANGUE == "en" else "OUI")
        wb.save(str(sortie))
        wb.close()

        retenues = classeur.lire_classeur_detection(str(sortie), config_det)
        verifier(len(retenues) >= 3, f"{len(retenues)} lignes retenues apres relecture")
        ecrits, refuses = classeur.injecter(str(table), nouveau, retenues, config_det)
        verifier(len(ecrits) >= 3 and (not refuses or LANGUE == "en"),
                 f"{len(ecrits)} lignes injectees, {len(refuses)} refusee(s)")
        entrees_nouv = coeur.charger_entrees(str(table), nouveau)
        res2 = coeur.anonymiser_document(str(exemple), entrees_nouv, config, "Mes correspondances", nouveau)
        verifier(res2["total"] >= (8 if LANGUE == "fr" else 4), f"le dossier injecte anonymise le document ({res2['total']} remplacements)")
        verifier(all(b.startswith("[") and "1" in b for b, _ in entrees_nouv),
                 "les balises injectees sont par partie ([DEM1], ...)")

        print("\n4. Lecture d'un PDF natif")
        try:
            fitz = det._pymupdf()
            pdf = racine / "Exemples" / "exemple.pdf"
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((72, 72), "La Societe Horizon Numerique, 12 rue des Lilas, 59000 Lille,"
                             " SIRET 812 345 678, contre la Societe Atelier Bleu Marine.",
                             fontsize=9)
            doc.save(str(pdf))
            doc.close()
            texte, source = det.extraire_texte_ou_ocr(str(pdf))
            verifier(source == "texte" and "Horizon" in texte, "un PDF natif est lu sans OCR")
        except ImportError:
            verifier(False, "PyMuPDF est installe (lecture des PDF)")

    finally:
        shutil.rmtree(racine, ignore_errors=True)

    print("\n" + "=" * 68)
    print(f"RESULTAT : {len(OK)} controle(s) reussi(s), {len(KO)} echec(s)")
    for k in KO:
        print("   ECHEC :", k)
    print("=" * 68)
    if COMPTE_RENDU is not None:
        COMPTE_RENDU.close()
        COMPTE_RENDU = None
    return 0 if not KO else 1


if __name__ == "__main__":
    sys.exit(main())
