"""Licence : GNU Affero General Public License v3 (LICENSE_AGPL-3.0.txt).

test_detection.py
-----------------
Recette automatique de la pre-detection et de l'injection.

Tout est fictif : documents, societes, personnes, tableurs. Le test cree son
propre materiel dans un dossier temporaire et le supprime a la fin. Aucun
dossier reel n'est touche.

Lancement : double-clic sur Recette_detection.bat
"""

import os
import sys
import json
import shutil
import pathlib
import tempfile

SCRIPT_DIR = pathlib.Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

import openpyxl
from docx import Document

import detection_core as det
import detection_classeur as outil
import detection_roles as roles
import modele_onglet as modele

OK, KO = [], []


def _leve(fonction):
    """Vrai si l'appel est refusé par une exception (refus explicite)."""
    try:
        fonction()
    except Exception:
        return True
    return False


def verifier(libelle, condition, detail=""):
    (OK if condition else KO).append(libelle if condition else (libelle, detail))
    print(f"  [{'OK ' if condition else 'KO '}] {libelle}"
          + ("" if condition or not detail else f"\n         {detail}"))


# ------------------------------------------------------------------
# MATERIEL FICTIF
# ------------------------------------------------------------------

ACTE = """TRIBUNAL JUDICIAIRE DE LILLE
N° RG : 24/01234

ASSIGNATION

POUR :

La société ALDEBARAN SYSTEMES, SAS au capital de 50 000 euros, immatriculée au
RCS de Lille 812 345 678, dont le siège social est situé 14 rue des Trois Moulins,
59800 LILLE, prise en la personne de son président, Monsieur Jean DURAND,
demeurant 8 avenue du General Leclerc - 59000 LILLE, joignable au 03 20 11 22 33,

Exploitant sous l'enseigne ALDEBARAN STORE,

Titulaire de la marque française « ALDEBARAN » déposée à l'INPI sous le n° 4567890,
et du brevet EP 3 456 789,

Ayant pour avocat Maître Sophie MARCHAND, avocat au barreau de Lille,
SELARL MARCHAND ASSOCIES, 3 place du Theatre, 59000 LILLE,
courriel s.marchand@exemple-avocats.fr

DEMANDERESSE

CONTRE :

La société VEGA INDUSTRIES SARL, immatriculée sous le SIRET 498 765 432 00025,
dont le siège est 27 boulevard Victor Hugo, 75008 PARIS, représentée par son
gérant Monsieur Come VANTERRE,

Ayant pour avocat Maître Paul BERTIN, du barreau de Paris, joignable au
p.bertin@exemple2.fr et sur www.bertin-avocats.fr

DEFENDERESSE

FAITS

ALDEBARAN SYSTEMES commercialise depuis 2019 son logiciel de gestion.
Monsieur DURAND a constate que VEGA INDUSTRIES diffusait un produit similaire.
ALDEBARAN a mis en demeure VEGA le 12 mars 2025. Monsieur VANTERRE n'a pas repondu.
Un constat a ete dresse par Maître Luc THEVENIN, commissaire de justice a Lille,
5 rue Nationale, 59000 LILLE.
Jean DURAND rappelle que la societe ALDEBARAN est titulaire des droits.
"""


CONTRAT_CLIENT_ANONYME = """CONTRAT DE PRESTATION DE SERVICES INFORMATIQUES

ENTRE LES SOUSSIGNÉS :

La société NEBULA CLOUD, SAS au capital de 10 000 euros, dont le siège social est situé 12 rue de la Paix, 59000 LILLE, représentée par son Président, Monsieur Hugo LAMBERT,

Ci-après dénommée « le Prestataire »,

D'UNE PART,

ET :

Le Client, tel qu'identifié dans le Bon de commande,

Ci-après dénommé « le Client »,

D'AUTRE PART,

IL A ÉTÉ CONVENU CE QUI SUIT :

ARTICLE 1 - OBJET
Le Prestataire fournit au Client une solution d'hébergement.

Fait à Lille, en deux exemplaires.

Pour le Prestataire : Monsieur Hugo LAMBERT
Pour le Client : la société ORION DISTRIBUTION, représentée par Madame Julie MOREAU
"""

CONTRAT_LICENCE = """CONTRAT DE LICENCE

ENTRE :

La société BOREAL SOFT, SAS dont le siège est 4 rue Pasteur, 59000 LILLE (ci-après « le Concédant »),

ET :

La société ARCTURUS, SARL dont le siège est 9 rue Gambetta, 62000 ARRAS (ci-après « le Licencié »),

ARTICLE 1 - OBJET
Le Concédant concède au Licencié une licence non exclusive.
"""

MED_EMISE = """Maître Aurore TESTEUR
Avocat au Barreau de [Barreau]

Société DELTA FABRICATION
Monsieur Marc ROUSSEL
45 avenue de la Marne, 59100 ROUBAIX

Lettre recommandée avec accusé de réception

Objet : Mise en demeure

Monsieur,

J'interviens en qualité de conseil de la société GAMMA LOGICIELS, dont le siège est 3 rue Faidherbe, 59000 LILLE.

Ma cliente constate que la société DELTA FABRICATION exploite sans licence son logiciel. Je vous mets en demeure de cesser.

Je vous prie d'agréer, Monsieur, l'expression de mes salutations distinguées.

Maître Aurore TESTEUR
"""

MED_RECUE = """Maître Bernard VOLANT
SELARL VOLANT AVOCATS
8 place Rihour, 59000 LILLE

Société GAMMA LOGICIELS
3 rue Faidherbe, 59000 LILLE

LRAR

Objet : Mise en demeure

Madame, Monsieur,

Je suis le conseil de la société DELTA FABRICATION, dont le siège est 45 avenue de la Marne, 59100 ROUBAIX.

Ma cliente vous met en demeure de lui régler la somme due.

Veuillez agréer, Madame, Monsieur, mes salutations distinguées.

Maître Bernard VOLANT
"""

AUDIT = """RAPPORT D'AUDIT RGPD

Audit de la société HELIOS MEDICAL, SAS dont le siège est 10 boulevard Carnot, 59800 LILLE (ci-après « la Société »).

Interlocuteur : Monsieur Paul GIRARD, DSI.

1. Registre des traitements
La Société utilise le logiciel de paie Silae et la solution CRM Salesforce. L'hébergement est confié au sous-traitant la société OVH CLOUD.
Madame Lucie PETIT, assistante RH, tient le registre. Le responsable de traitement est la Société. Aucun DPO n'est désigné.
Les données à caractère personnel sont conservées sans durée de conservation définie.
"""

CONCLUSIONS_APPEL = """COUR D'APPEL DE DOUAI
N° RG : 25/00123

CONCLUSIONS D'INTIMÉE

POUR :

La société LYRA TEXTILES, SARL, dont le siège est 5 rue Nationale, 59000 LILLE,
Ayant pour avocat Maître Anne LEBLANC, avocat au barreau de Lille

INTIMÉE

CONTRE :

La société CASTOR INDUSTRIE, SAS, dont le siège est 2 quai de la Loire, 75019 PARIS,
Ayant pour avocat Maître Pierre NOIR

APPELANTE

PLAISE À LA COUR

La société CASTOR INDUSTRIE a interjeté appel. Madame Sylvie BLANC, expert judiciaire, a déposé son rapport. La société ZETA TRANSPORTS, qui n'est pas partie, a livré les produits.
"""

CONCLUSIONS_SUITE = """TRIBUNAL JUDICIAIRE DE LILLE

CONCLUSIONS EN DÉFENSE

POUR :

La société VEGA INDUSTRIES SARL, dont le siège est 11 rue Solferino, 59000 LILLE,

DÉFENDERESSE

CONTRE :

La société ALDEBARAN SYSTEMES, SAS,

DEMANDERESSE

EN PRÉSENCE DE :

La société SIRIUS NEGOCE, SAS,

INTERVENANTE VOLONTAIRE

DISCUSSION

VEGA INDUSTRIES conteste.
"""

# Gabarit imitant la structure reelle : titres, en-tetes ligne 3, sections,
# emplacements numerotes vides.
GABARIT = [
    ("titre", "TABLEAU DE CORRESPONDANCES — FICTIF"),
    ("titre", "Jeu d'essai"),
    ("entetes", None),
    ("section", "CÔTÉ DEMANDEUR"),
    ("ligne", ("[PP_DEM_1]", "Personne physique", "Demande")),
    ("ligne", ("[PP_DEM_2]", "Personne physique", "Demande")),
    ("ligne", ("[PM_DEM_1]", "Personne morale", "Demande")),
    ("ligne", ("[PM_DEM_2]", "Personne morale", "Demande")),
    ("ligne", ("[ADR_DEM_1]", "Adresse", "Demande")),
    ("ligne", ("[ADR_DEM_2]", "Adresse", "Demande")),
    ("ligne", ("[SIRET_DEM_1]", "SIRET", "Demande")),
    ("ligne", ("[RL_DEM_1]", "Représentant légal", "Demande")),
    ("ligne", ("[AV_DEM_1]", "Avocat", "Demande")),
    ("ligne", ("[AV_DEM_2]", "Avocat", "Demande")),
    ("ligne", ("[MAIL_DEM_1]", "E-mail", "Demande")),
    ("section", "CÔTÉ DÉFENDEUR"),
    ("ligne", ("[PP_DEF_1]", "Personne physique", "Défense")),
    ("ligne", ("[PM_DEF_1]", "Personne morale", "Défense")),
    ("ligne", ("[ADR_DEF_1]", "Adresse", "Défense")),
    ("ligne", ("[SIRET_DEF_1]", "SIRET", "Défense")),
    ("ligne", ("[RL_DEF_1]", "Représentant légal", "Défense")),
    ("ligne", ("[AV_DEF_1]", "Avocat", "Défense")),
    ("ligne", ("[MAIL_DEF_1]", "E-mail", "Défense")),
    ("ligne", ("[SIT_DEF_1]", "Site web", "Défense")),
    ("section", "ÉLÉMENTS COMMUNS"),
    ("ligne", ("[ADR_SUPP_1]", "Adresse", "")),
    ("ligne", ("[CJ_1]", "Commissaire", "")),
    ("ligne", ("[RG_1]", "Référence RG", "")),
    ("ligne", ("[MRQ_1]", "Marque", "")),
    ("ligne", ("[BRV_1]", "Brevet", "")),
    ("ligne", ("[DEPOT_1]", "Dépôt INPI", "")),
]


def ecrire_gabarit(chemin, lignes_sup=()):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    wb.create_sheet("Mode d'emploi")["A1"] = "onglet technique"
    ws = wb.create_sheet("Dossier fictif")
    r = 1
    for genre, contenu in list(GABARIT) + list(lignes_sup):
        if genre == "titre":
            ws.cell(row=r, column=1, value=contenu)
        elif genre == "entetes":
            for i, t in enumerate(["Terme original", "Balise", "Catégorie", "Camp",
                                   "Dossier", "Actif ?", "Notes"], start=1):
                ws.cell(row=r, column=i, value=t)
        elif genre == "section":
            ws.cell(row=r, column=1, value=contenu)
        else:
            balise, categorie, camp = contenu
            ws.cell(row=r, column=2, value=balise)
            ws.cell(row=r, column=3, value=categorie)
            ws.cell(row=r, column=4, value=camp)
        r += 1
    wb.save(chemin)


def docx_depuis_texte(chemin, texte):
    doc = Document()
    for ligne in texte.split("\n"):
        doc.add_paragraph(ligne)
    doc.save(chemin)


def trouver(entites, fragment):
    return next((e for e in entites
                 if any(fragment.lower() in v.lower() for v in e.variantes)), None)



# ------------------------------------------------------------------
# Un PDF « scanne » : une image de texte, sans aucun texte extractible.
# C'est exactement ce qu'est un acte recu par courrier puis numerise.
# ------------------------------------------------------------------

def pdf_scanne(chemin, lignes, police_taille=34):
    from PIL import Image, ImageDraw, ImageFont
    import glob

    chemin_police = None
    for motif in ("**/LiberationSerif-Regular.ttf", "**/DejaVuSerif.ttf",
                  "**/Carlito-Regular.ttf", "**/*.ttf"):
        trouves = glob.glob("/usr/share/fonts/" + motif, recursive=True)
        if trouves:
            chemin_police = trouves[0]
            break
    police = (ImageFont.truetype(chemin_police, police_taille)
              if chemin_police else ImageFont.load_default())

    image = Image.new("RGB", (2100, 180 + 60 * len(lignes)), "white")
    dessin = ImageDraw.Draw(image)
    y = 60
    for ligne in lignes:
        dessin.text((80, y), ligne, fill="black", font=police)
        y += 60
    image.save(chemin, "PDF", resolution=200.0)


def pdf_texte(chemin, lignes):
    """Un PDF normal, dont le texte est extractible sans OCR."""
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    c = canvas.Canvas(str(chemin), pagesize=A4)
    y = 780
    for ligne in lignes:
        c.drawString(60, y, ligne)
        y -= 20
    c.save()


# ------------------------------------------------------------------

def main():
    racine = pathlib.Path(tempfile.mkdtemp(prefix="recette_detection_"))
    print(f"\nDossier de recette (fictif, supprimé à la fin) :\n  {racine}\n")

    try:
        config_det = det.charger_config_detection(SCRIPT_DIR / "detection_config.json")
        doc = racine / "Assignation fictive.docx"
        docx_depuis_texte(doc, ACTE)

        print("1. Lecture du document")
        texte = det.extraire_texte(str(doc))
        verifier("texte extrait d'un .docx", "ALDEBARAN SYSTEMES" in texte)

        entites = det.detecter(texte)
        roles.attribuer_roles(texte, entites, config_det)
        print(f"\n2. Détection ({len(entites)} entités)")

        attendus = [
            ("personne morale demanderesse", "ALDEBARAN SYSTEMES", "personne_morale"),
            ("personne morale défenderesse", "VEGA INDUSTRIES", "personne_morale"),
            ("représentant légal demandeur", "Jean DURAND", "representant_legal"),
            ("représentant légal défendeur", "Come VANTERRE", "representant_legal"),
            ("adresse avec séparateur virgule", "14 rue des Trois Moulins", "adresse"),
            ("adresse avec séparateur tiret", "General Leclerc", "adresse"),
            ("SIRET", "498 765 432 00025", "siret"),
            ("RCS", "812 345 678", "rcs"),
            ("adresse électronique", "s.marchand@exemple-avocats.fr", "email"),
            ("site web", "bertin-avocats.fr", "site_web"),
            ("avocat (titre Maître)", "Sophie MARCHAND", "avocat"),
            ("structure d'exercice", "MARCHAND ASSOCIES", "avocat"),
            ("commissaire de justice", "Luc THEVENIN", "commissaire_justice"),
            ("enseigne", "ALDEBARAN STORE", "enseigne"),
            ("marque", "ALDEBARAN", "marque"),
            ("brevet", "EP 3 456 789", "brevet"),
            ("référence RG", "24/01234", "rg"),
            ("dépôt INPI", "4567890", "depot_inpi"),
            ("téléphone", "03 20 11 22 33", "telephone"),
        ]
        for libelle, fragment, type_attendu in attendus:
            e = next((x for x in entites
                      if x.type == type_attendu
                      and any(fragment.lower() in v.lower() for v in x.variantes)), None)
            verifier(f"détecté : {libelle}", e is not None,
                     f"« {fragment} » attendu en type {type_attendu}")

        print("\n3. Faux positifs")
        verifier("un verbe n'est pas avalé dans un patronyme",
                 not any("repondu" in v.lower() for e in entites for v in e.variantes),
                 str([v for e in entites for v in e.variantes if "repondu" in v.lower()]))
        verifier("aucun terme vide ou d'un seul caractère",
                 all(len(v.strip()) > 1 for e in entites for v in e.variantes))

        print("\n4. Rattachement des occurrences à la première apparition")
        durand = trouver(entites, "Jean DURAND")
        verifier("les formes d'une même personne tiennent sur une seule ligne",
                 durand is not None
                 and {"Monsieur Jean DURAND", "Monsieur DURAND", "Jean DURAND", "DURAND"}
                 <= set(durand.variantes),
                 str(durand.variantes) if durand else "entité introuvable")
        verifier("une occurrence tardive ne crée pas une seconde entité",
                 sum(1 for e in entites
                     if e.type in ("personne_physique", "representant_legal")
                     and any("DURAND" in v for v in e.variantes)) == 1)
        aldebaran = trouver(entites, "ALDEBARAN SYSTEMES")
        verifier("la forme abrégée d'une société rejoint sa forme complète",
                 aldebaran is not None and "ALDEBARAN" in aldebaran.variantes,
                 str(aldebaran.variantes) if aldebaran else "entité introuvable")
        verifier("le rattachement conserve la partie de la première apparition",
                 durand is not None and durand.partie == "DEM1",
                 f"partie = {durand.partie!r}" if durand else "")

        print("\n5. Ordre d'apparition")
        positions = [e.position for e in entites]
        verifier("les entités sortent dans l'ordre du document",
                 positions == sorted(positions))
        verifier("la partie demanderesse apparaît avant la défenderesse",
                 aldebaran.position < trouver(entites, "VEGA INDUSTRIES").position)

        print("\n6. Ambiguïtés signalées plutôt que tranchées")
        marque = next((e for e in entites if e.type == "marque"), None)
        verifier("une forme partagée entre deux entités est signalée",
                 marque is not None and marque.ambigu and bool(marque.note),
                 marque.note if marque else "")
        siret = next((e for e in entites if e.type == "siret"), None)
        verifier("un numéro à clé de contrôle fausse est retenu et signalé",
                 siret is not None and siret.confiance == "à vérifier"
                 and "clé" in siret.note,
                 f"{siret.confiance} / {siret.note}" if siret else "")

        print("\n7. Parties et rôles (assignation)")
        verifier("la société demanderesse est la partie DEM1",
                 aldebaran.partie == "DEM1" and aldebaran.role == "Identité",
                 f"{aldebaran.partie} / {aldebaran.role}")
        vega = trouver(entites, "VEGA INDUSTRIES")
        verifier("la société défenderesse est la partie DEF1",
                 vega.partie == "DEF1", f"{vega.partie}")
        verifier("le président de la demanderesse est son dirigeant",
                 durand.partie == "DEM1" and durand.role == "Dirigeant",
                 f"{durand.partie} / {durand.role}")
        verifier("le gérant de la défenderesse est son dirigeant, pas celui de la demanderesse",
                 trouver(entites, "VANTERRE").partie == "DEF1")
        verifier("le siège est rattaché à sa société",
                 trouver(entites, "14 rue des Trois Moulins").partie == "DEM1"
                 and trouver(entites, "27 boulevard Victor Hugo").partie == "DEF1")
        verifier("l'adresse personnelle du dirigeant est distinguée du siège",
                 trouver(entites, "General Leclerc").role == "Adresse dirigeant",
                 trouver(entites, "General Leclerc").role)
        verifier("chaque avocat est rattaché à la partie qu'il représente",
                 trouver(entites, "Sophie MARCHAND").partie == "DEM1"
                 and trouver(entites, "Paul BERTIN").partie == "DEF1")
        verifier("les coordonnées de l'avocat sont celles de l'avocat, pas de la partie",
                 trouver(entites, "s.marchand@").role == "E-mail avocat"
                 and trouver(entites, "3 place du Theatre").role == "Adresse avocat")
        verifier("le commissaire de justice est un acteur à part (COMMISSAIRE1)",
                 trouver(entites, "Luc THEVENIN").partie == "COMMISSAIRE1")
        verifier("la référence RG est un élément commun",
                 trouver(entites, "24/01234").partie == "COMMUN")
        verifier("les balises se lisent seules",
                 aldebaran.balise == "[DEM1]"
                 and durand.balise == "[DEM1_DIRIGEANT]"
                 and trouver(entites, "Paul BERTIN").balise == "[DEF1_AVOCAT]",
                 f"{aldebaran.balise} {durand.balise} {trouver(entites, 'Paul BERTIN').balise}")

        print("\n8. Onglet par partie : création")
        gabarit = racine / "Table_fictive.xlsx"
        ecrire_gabarit(gabarit)
        modele.creer_onglet(str(gabarit), "Dossier neuf", config_det)
        verifier("un onglet par partie est créé", modele.modele_onglet(
            str(gabarit), "Dossier neuf") == "par_partie")
        verifier("un onglet existant n'est jamais écrasé",
                 _leve(lambda: modele.creer_onglet(str(gabarit), "Dossier neuf", config_det)))
        verifier("un nom d'onglet interdit par Excel est refusé",
                 _leve(lambda: modele.creer_onglet(str(gabarit), "A/B", config_det)))
        verifier("l'ancien gabarit est reconnu comme tel",
                 modele.modele_onglet(str(gabarit), "Dossier fictif") == "ancien")

        entites2, infos, _, source2 = det.analyser_document(
            str(doc), str(gabarit), "Dossier neuf",
            str(SCRIPT_DIR / "detection_config.json"))
        verifier("un .docx est lu comme texte, pas par OCR", source2 == "texte", source2)
        verifier("le type de document est reconnu", infos["type"] == "contentieux", infos["type"])
        verifier("deux entités ne reçoivent jamais la même balise",
                 len({e.balise for e in entites2 if e.balise})
                 == len([e for e in entites2 if e.balise]))

        print("\n9. Classeur de proposition")
        sortie = racine / "_Detection_Assignation fictive.xlsx"
        outil.ecrire_classeur_detection(str(sortie), entites2, infos, config_det,
                                        str(doc), "Table fictive", "Dossier neuf")
        verifier("classeur de proposition créé", sortie.is_file())
        wb = openpyxl.load_workbook(sortie)
        verifier("onglets Détection et Parties présents",
                 "Détection" in wb.sheetnames and "Parties" in wb.sheetnames)
        ws = wb["Détection"]
        verifier("une ligne par entité détectée",
                 sum(1 for r in ws.iter_rows(min_row=outil.LIGNE_ENTETES + 1, values_only=True)
                     if r and r[outil.COL_NATURE - 1]) == len(entites2))
        colonnes = [ws.cell(row=outil.LIGNE_ENTETES, column=i).value
                    for i in range(1, len(outil.TEXTES["fr"]["en_tetes"]) + 1)]
        verifier("colonnes attendues",
                 [c.split("  (")[0] for c in colonnes] == outil.TEXTES["fr"]["en_tetes"], str(colonnes))
        verifier("Garder ?, la partie et le rôle se choisissent dans des listes déroulantes",
                 len(ws.data_validations.dataValidation) >= 3)
        verifier("plus de menu de toutes les balises sur la colonne Balise",
                 not any("B" in str(dv.sqref).split(":")[0] for dv in ws.data_validations.dataValidation))
        f_balise = str(ws.cell(row=outil.LIGNE_ENTETES + 2, column=outil.COL_BALISE).value or "")
        verifier("la balise se recalcule depuis Partie + Rôle (formule)",
                 f_balise.startswith("=IF(") and "AC" in f_balise, f_balise[:60])
        menu = [c.value for c in wb["Listes"]["A"][1:] if c.value and not str(c.value).startswith("=")]
        nouvelles = [c.value for c in wb["Listes"]["H"][1:] if c.value]
        verifier("le menu des parties nomme chaque partie connue",
                 any(m.startswith("DEM1  =  ") and "ALDEBARAN" in m for m in menu)
                 and any(m.startswith("DEF1  =  ") for m in menu), str(menu))
        verifier("« nouvelle partie » sans numéro, pour chaque sorte de partie",
                 any(m.startswith("DEF+  (") for m in nouvelles) and any(m.startswith("SOCIETE+  (") for m in nouvelles),
                 str(nouvelles))
        verifier("le menu grandit avec les parties créées dans le tableau (formules)",
                 str(wb["Listes"].cell(row=len(menu) + 2, column=1).value).startswith("=IF(")
                 and "OFFSET" in str(ws.data_validations.dataValidation[0].formula1))
        verifier("la partie choisie se lit dans le libellé (premier mot)",
                 outil.code_de_libelle("DEF1  =  VEGA INDUSTRIES") == "DEF1"
                 and outil.code_de_libelle("COMMUN  (aucune partie)") == "COMMUN")
        verifier("une identité déjà prise se signale en rouge (mise en forme)",
                 any("COUNTIF" in str(r.formula) for rg in ws.conditional_formatting for r in rg.rules))
        verifier("le type de document est indiqué en tête",
                 "Contentieux" in str(ws["A3"].value), str(ws["A3"].value))
        wb.close()

        retenues = outil.lire_classeur_detection(str(sortie), config_det)
        verifier("seules les lignes marquées OUI sont relues",
                 0 < len(retenues) < len(entites2),
                 f"{len(retenues)} retenue(s) sur {len(entites2)}")
        verifier("une société dont une forme est partagée avec une marque attend votre décision",
                 not any("ALDEBARAN SYSTEMES" in r["terme"] for r in retenues))
        # Decision de l'utilisateur : il retient la societe (la marque reste a NON)
        ald = trouver(entites2, "ALDEBARAN SYSTEMES")
        retenues.insert(0, {"terme": ald.terme_pour_tableur(), "partie": ald.partie,
                            "role": ald.role})

        print("\n10. Injection dans l'onglet par partie")
        ecrits, refuses = outil.injecter(str(gabarit), "Dossier neuf", retenues, config_det)
        verifier("des termes ont été écrits", len(ecrits) > 0)
        _, entrees_onglet = modele.lire_onglet(str(gabarit), "Dossier neuf")
        place = {x["balise"]: x["terme"] for x in entrees_onglet}
        verifier("la société demanderesse est sous [DEM1]",
                 "ALDEBARAN" in place.get("[DEM1]", ""), str(place.get("[DEM1]")))
        verifier("la société défenderesse est sous [DEF1]",
                 "VEGA" in place.get("[DEF1]", ""), str(place.get("[DEF1]")))
        verifier("la colonne Actif est renseignée",
                 all(x["actif"] == "OUI" for x in entrees_onglet))
        verifier("les lignes sont groupées par partie",
                 [x["partie"] for x in entrees_onglet]
                 == sorted([x["partie"] for x in entrees_onglet],
                           key=[x["partie"] for x in entrees_onglet].index))
        wb = openpyxl.load_workbook(gabarit)
        verifier("l'ancien onglet n'a pas été touché",
                 wb["Dossier fictif"].cell(row=4, column=1).value == "CÔTÉ DEMANDEUR")
        wb.close()
        verifier("l'injection est refusée dans un onglet à l'ancien modèle",
                 _leve(lambda: outil.injecter(str(gabarit), "Dossier fictif", retenues, config_det)))

        ecrits2, refuses2 = outil.injecter(str(gabarit), "Dossier neuf", retenues, config_det)
        verifier("une seconde injection ne crée aucun doublon",
                 len(ecrits2) == 0 and len(refuses2) == len(retenues),
                 f"{len(ecrits2)} écrit(s), {len(refuses2)} refusé(s)")
        verifier("chaque refus porte un motif explicite",
                 all("déjà présent" in r for _, r in refuses2), str(refuses2[:3]))
        _, apres2 = modele.lire_onglet(str(gabarit), "Dossier neuf")
        verifier("les balises existantes sont inchangées après une nouvelle injection",
                 {x["balise"]: x["terme"] for x in apres2} == place)

        print("\n11. Correction de la partie par l'utilisateur")
        modele.creer_onglet(str(gabarit), "Dossier corrige", config_det)
        modifiees = [dict(r) for r in retenues]
        cible = next(r for r in modifiees if "VEGA" in r["terme"])
        cible["partie"] = "DEF2"
        outil.injecter(str(gabarit), "Dossier corrige", modifiees, config_det)
        _, e3 = modele.lire_onglet(str(gabarit), "Dossier corrige")
        verifier("la partie corrigée à la main commande la balise",
                 any(x["balise"] == "[DEF2]" and "VEGA" in x["terme"] for x in e3),
                 str([(x["balise"], x["terme"][:20]) for x in e3 if "VEGA" in x["terme"]]))
        verifier("une ligne sans partie est refusée, jamais placée d'office",
                 any("partie à choisir" in r for _, r in outil.injecter(
                     str(gabarit), "Dossier corrige",
                     [{"terme": "SOCIETE ORPHELINE", "partie": "", "role": "Identité"}],
                     config_det)[1]))
        ecr, _ = outil.injecter(str(gabarit), "Dossier corrige",
                                [{"terme": "ALDEBARAN SYST.", "partie": "DEM1",
                                  "role": "Identité"}], config_det)
        verifier("une nouvelle forme d'une partie connue devient une variante",
                 ecr and ecr[0][1] == "[DEM1]" and ecr[0][2] == "variantes ajoutées",
                 str(ecr))
        ecr, ref = outil.injecter(str(gabarit), "Dossier corrige",
                                  [{"terme": "VEGA INDUSTRIES SARL", "partie": "DEM3",
                                    "role": "Identité"}], config_det)
        verifier("un terme déjà placé n'est jamais déplacé vers une autre partie",
                 not ecr and ref and "jamais déplacée" in ref[0][1], str(ref))

        print("\n12. Évolutivité : un rôle ajouté à la configuration")
        config3 = json.loads(json.dumps(config_det))
        config3["attributs"].append({"libelle": "Numéro de client", "suffixe": "NUM_CLIENT"})
        ecr, _ = outil.injecter(str(gabarit), "Dossier corrige",
                                [{"terme": "CLI-000123", "partie": "DEF1",
                                  "role": "Numéro de client"}], config3)
        verifier("un rôle inconnu du code fonctionne dès qu'il est déclaré",
                 ecr and ecr[0][1] == "[DEF1_NUM_CLIENT]", str(ecr))
        verifier("aucune balise n'est codée en dur dans le programme",
                 all("DEMANDEUR1" not in open(SCRIPT_DIR / f, encoding="utf-8").read()
                     for f in ("detection_core.py", "detection_roles.py", "modele_onglet.py",
                               "detection_classeur.py")))

        print("\n13. Termes produits utilisables par l'anonymiseur")
        import anonymisation_core as coeur
        wb_tmp = racine / "verif.xlsx"
        shutil.copy(gabarit, wb_tmp)
        entrees = coeur.charger_entrees(str(wb_tmp), "Dossier neuf")
        verifier("l'anonymiseur relit les lignes injectées", len(entrees) > 0,
                 f"{len(entrees)} correspondance(s)")
        toutes = {v for _, variantes in entrees for v in variantes}
        verifier("les variantes séparées par / sont bien éclatées",
                 any(len(variantes) > 1 for _, variantes in entrees))
        verifier("les formes détectées se retrouvent dans la table",
                 "Monsieur Jean DURAND" in toutes or "Jean DURAND" in toutes,
                 str(sorted(toutes)[:6]))

        verifier("un numéro RG « 24/01234 » reste d'un seul tenant (le « / » collé ne sépare pas)",
                 coeur.decouper_variantes("N° RG : 24/01234 / 24/01234")
                 == ["N° RG : 24/01234", "24/01234"]
                 and not any(v in ("24", "01234") for v in toutes),
                 str(coeur.decouper_variantes("N° RG : 24/01234 / 24/01234")))

        print("\n14. Lecture des PDF")
        lignes_pdf = ["TRIBUNAL JUDICIAIRE DE LILLE",
                      "La societe ALDEBARAN SYSTEMES, representee par",
                      "Monsieur Jean DURAND, demeurant a LILLE."]

        pdf_normal = racine / "Acte natif.pdf"
        pdf_texte(pdf_normal, lignes_pdf)
        verifier("un PDF normal n'est pas pris pour un scan",
                 not det.pdf_est_scanne(str(pdf_normal)))
        texte_n, source_n = det.extraire_texte_ou_ocr(str(pdf_normal))
        verifier("le texte d'un PDF natif est lu sans OCR",
                 source_n == "texte" and "ALDEBARAN" in texte_n.upper(), source_n)

        scan = racine / "Acte scanne.pdf"
        pdf_scanne(scan, lignes_pdf)
        verifier("un PDF image est reconnu comme scanné",
                 det.pdf_est_scanne(str(scan)))

        verifier("pages lues par OCR résumées en plages (« 3, 5-7 »)",
                 det.plages_pages([7, 3, 5, 6]) == "3, 5-7", det.plages_pages([7, 3, 5, 6]))

        print("\n15. OCR")
        # Le pack de langue est cherche la ou il se trouve reellement : dans le
        # dossier declare en configuration, puis aux emplacements habituels.
        # Le test s'execute donc aussi bien ici que sur le poste.
        try:
            det.trouver_tesseract({})
            tesseract_present = True
        except det.OcrIndisponible:
            tesseract_present = False
        verifier("Tesseract est localisé", tesseract_present)

        tessdata = None
        for candidat in ({"tessdata": str(SCRIPT_DIR / "_tessdata_test")},
                         config_det.get("ocr") or {}, {}):
            try:
                tessdata = det.trouver_tessdata(candidat, "fra")
                break
            except det.OcrIndisponible:
                continue
        config_ocr = {"tessdata": tessdata, "langue": "fra", "resolution": 300}

        if tesseract_present and tessdata:
            verifier("le pack de langue français est localisé",
                     os.path.isfile(os.path.join(tessdata, "fra.traineddata")), tessdata)

            texte_ocr = det.ocr_pdf(str(scan), config_ocr)
            verifier("l'OCR restitue le texte de l'image",
                     "ALDEBARAN" in texte_ocr.upper(), repr(texte_ocr[:120]))
            verifier("l'OCR lit aussi les noms de personnes",
                     "DURAND" in texte_ocr.upper(), repr(texte_ocr[:120]))

            texte_a, source_a = det.extraire_texte_ou_ocr(str(scan), config_ocr)
            verifier("un PDF scanné bascule automatiquement sur l'OCR",
                     source_a == "ocr", source_a)

            entites_ocr = det.detecter(texte_ocr, "ocr")
            verifier("l'OCR a permis de détecter des entités",
                     len(entites_ocr) > 0, f"{len(entites_ocr)} entité(s)")
            verifier("TOUTE entité issue d'un OCR est marquée à vérifier",
                     all(e.confiance == "à vérifier" for e in entites_ocr),
                     str([(e.canonique, e.confiance) for e in entites_ocr[:3]]))
            verifier("et la raison est écrite dans la remarque",
                     all("OCR" in e.note for e in entites_ocr),
                     str([e.note for e in entites_ocr[:2]]))

            gabarit_ocr = racine / "Table_ocr.xlsx"
            ecrire_gabarit(gabarit_ocr)
            modele.creer_onglet(str(gabarit_ocr), "Dossier neuf", config_det)
            # Config de test pointant sur le pack de langue du bac a sable.
            config_fichier = racine / "detection_config_test.json"
            base = json.loads(open(SCRIPT_DIR / "detection_config.json",
                                   encoding="utf-8").read())
            base["ocr"] = config_ocr
            config_fichier.write_text(json.dumps(base, ensure_ascii=False),
                                      encoding="utf-8")
            ents, infos_ocr, _, src = det.analyser_document(
                str(scan), str(gabarit_ocr), "Dossier neuf", str(config_fichier))
            verifier("la chaîne complète signale que la source est un OCR", src == "ocr")

            sortie_ocr = racine / "_Detection_scan.xlsx"
            outil.ecrire_classeur_detection(str(sortie_ocr), ents, infos_ocr, config_det,
                                            str(scan), "Table", "Dossier neuf")
            verifier("aucune ligne issue d'un OCR n'est proposée à OUI",
                     outil.lire_classeur_detection(str(sortie_ocr), config_det) == [],
                     "une ligne OCR proposée à OUI serait injectée sans relecture")

            print("\n15 bis. PDF mixte : OCR page par page (7/10/2026)")
            import fitz as _fitz
            p1, p2, p3 = racine / "m1.pdf", racine / "m2.pdf", racine / "m3.pdf"
            pdf_texte(p1, ["TRIBUNAL JUDICIAIRE DE LILLE",
                           "La societe ALDEBARAN SYSTEMES, representee par",
                           "Monsieur Jean DURAND, demeurant a LILLE."])
            pdf_scanne(p2, ["Pièce 2 : courrier de Monsieur Paul VERNIER,",
                            "demeurant 12 rue des Arts, 59100 ROUBAIX."])
            pdf_texte(p3, [""])                      # page blanche, sans image
            mixte = racine / "Acte mixte.pdf"
            with _fitz.open() as m:
                for part in (p1, p2, p3):
                    with _fitz.open(str(part)) as d:
                        m.insert_pdf(d)
                m.save(str(mixte))
            texte_m, source_m, pages_m = det.extraire_texte_detail(str(mixte), config_ocr)
            verifier("un PDF mixte est reconnu comme tel", source_m == "mixte", source_m)
            verifier("seule la page scannée passe à l'OCR (pas la page blanche sans image)",
                     pages_m == [2], str(pages_m))
            verifier("le texte de la page scannée est lu",
                     "VERNIER" in texte_m.upper() and "ALDEBARAN" in texte_m.upper(),
                     repr(texte_m[:300]))
            verifier("la pagination est conservée (3 pages)",
                     texte_m.count(det.SAUT_PAGE) == 2, str(texte_m.count(det.SAUT_PAGE)))
            ent_m = det.detecter(texte_m, source_m, pages_m)
            p1_ent = [e for e in ent_m if det.numero_page(texte_m, e.position) == 1]
            p2_ent = [e for e in ent_m if det.numero_page(texte_m, e.position) == 2]
            verifier("des entités sont trouvées sur la page scannée", bool(p2_ent),
                     str([e.canonique for e in ent_m]))
            verifier("les entités de la page scannée sont marquées OCR, à vérifier",
                     all(e.confiance == "à vérifier" and "OCR" in e.note for e in p2_ent),
                     str([(e.canonique, e.confiance) for e in p2_ent]))
            verifier("les entités de la page en texte ne portent pas la mention OCR",
                     bool(p1_ent) and not any("OCR" in e.note for e in p1_ent),
                     str([(e.canonique, e.note) for e in p1_ent]))
            _, src_mix, = det.extraire_texte_ou_ocr(str(mixte), config_ocr)
            verifier("l'ancienne fonction reste compatible (texte, source)", src_mix == "mixte")
            texte_n2, source_n2, pages_n2 = det.extraire_texte_detail(str(pdf_normal), config_ocr)
            verifier("un PDF entièrement en texte ne déclenche aucun OCR",
                     source_n2 == "texte" and pages_n2 == [], f"{source_n2} {pages_n2}")
        else:
            print("  [-- ] tests OCR non exécutés : "
                  + ("Tesseract introuvable" if not tesseract_present
                     else "pack de langue français introuvable"))
            print("         Ce n'est pas un échec : l'OCR est simplement "
                  "indisponible sur ce poste.")

        print("\n16. Messages quand l'OCR ne peut pas fonctionner")
        try:
            det.trouver_tessdata({"tessdata": str(racine / "nexistepas")}, "klingon")
            verifier("une langue absente est signalée", False, "aucune erreur levée")
        except det.OcrIndisponible as e:
            verifier("une langue absente est signalée clairement",
                     "klingon.traineddata" in str(e) and "detection_config.json" in str(e),
                     str(e)[:160])
        chemin_trouve = det.trouver_tesseract({"tesseract": "/aucun/chemin/tesseract"})
        verifier("un chemin Tesseract erroné retombe sur la recherche automatique",
                 os.path.isfile(chemin_trouve), str(chemin_trouve))

        # --------------------------------------------------------------
        print("\n17. Actes réels, formes qui posaient problème")
        # Materiel entierement fictif, ecrit pour reproduire les defauts
        # constates a l'usage : adresses ramassant le mot suivant, avocat pris
        # pour un commissaire de justice, denomination commencant par un
        # article, cabinet tronque a son premier associe.

        assignation = """ASSIGNATION DEVANT LE TRIBUNAL DE COMMERCE

A LA REQUÊTE DE :

La société FROMAGERIE DU PETIT NORD, société par actions simplifiée,
immatriculée au RCS de Lille Métropole sous le numéro 812 345 678, dont le
siège social est situé 14 rue des Trois Meuniers, 59800 LILLE, agissant
poursuites et diligences de son Président, Monsieur Éloi Vanderquiche,

exploitant sous l'enseigne « LE PETIT NORD »,

Ayant pour avocat Maître Solange Pérrichon, avocate au Barreau de Lille

CONTRE :

La société LES VERGERS DE L'ESCAUT, société à responsabilité limitée, SIRET
902 456 789 00014, dont le siège est 8 avenue des Peupliers Gris, 93100
MONTREUIL, prise en la personne de sa gérante, Madame Ambroisine Ketelbuters
Gérante

Ayant pour avocat Maître Gontran Delaflaque, avocat au Barreau de Paris

J'ai, Maître Basile Trompette, Commissaire de Justice associé au sein de la
SELARL TROMPETTE ET CHÉRUBIN, Commissaires de Justice associés, demeurant 27
boulevard Carnot, 59000 LILLE, signifié le présent acte.
"""
        ents = det.detecter(assignation, source="assignation")
        par_type = {}
        for e in ents:
            par_type.setdefault(e.type, []).append(e.canonique)

        adresses = par_type.get("adresse", [])
        verifier("les trois adresses de l'acte sont détectées",
                 len(adresses) == 3, str(adresses))
        verifier("une adresse ne ramasse pas le mot qui suit la ville",
                 all(not a.rstrip().upper().endswith(("CONTRE", "OBJET", "AYANT"))
                     for a in adresses), str(adresses))
        verifier("l'adresse s'arrête à la ville",
                 "14 rue des Trois Meuniers, 59800 LILLE" in adresses, str(adresses))
        verifier("une adresse dont le code postal et la ville sont sur deux lignes "
                 "est détectée",
                 any("MONTREUIL" in a for a in adresses), str(adresses))

        verifier("l'avocat du défendeur n'est pas pris pour un commissaire de justice",
                 "Maître Gontran Delaflaque" in par_type.get("avocat", []),
                 str(par_type.get("avocat")) + " / "
                 + str(par_type.get("commissaire_justice")))
        verifier("le commissaire de justice est bien reconnu comme tel",
                 "Maître Basile Trompette" in par_type.get("commissaire_justice", []),
                 str(par_type.get("commissaire_justice")))
        verifier("le cabinet du commissaire n'est pas tronqué à son premier associé",
                 any("CHÉRUBIN" in x for x in par_type.get("commissaire_justice", [])),
                 str(par_type.get("commissaire_justice")))
        verifier("le cabinet n'est pas compté une seconde fois en personne morale",
                 not any("TROMPETTE" in x for x in par_type.get("personne_morale", [])),
                 str(par_type.get("personne_morale")))

        verifier("une dénomination commençant par un article est détectée",
                 any("VERGERS DE L'ESCAUT" in x
                     for x in par_type.get("personne_morale", [])),
                 str(par_type.get("personne_morale")))
        verifier("l'enseigne entre guillemets est détectée",
                 "LE PETIT NORD" in par_type.get("enseigne", []),
                 str(par_type.get("enseigne")))

        representants = par_type.get("representant_legal", [])
        verifier("un représentant légal ne ramasse pas sa qualité",
                 all(not r.upper().endswith(("GÉRANTE", "GERANTE", "PRÉSIDENT"))
                     for r in representants), str(representants))
        verifier("les deux représentants légaux sont détectés",
                 len(representants) == 2, str(representants))

        contrat = """CONTRAT DE DÉVELOPPEMENT

ENTRE LES SOUSSIGNÉS :

La société LES ATELIERS DU CARILLON, dont le siège social est sis 22 quai de la
Batellerie, 59000
LILLE, immatriculée au RCS de Lille Métropole sous le numéro 888 111 222,
représentée par Madame Pélagie Vandenbroucke, en sa qualité de Présidente,

ET

La société CONFISERIES HAUTE-BORNE, dont le siège est 9 allée des Berlingots,
59650 VILLENEUVE-D'ASCQ, SIRET 777 333 444 00027, représentée par Monsieur
Hyacinthe Delbecque, Directeur Général.
"""
        ents = det.detecter(contrat, source="contrat")
        types = {}
        for e in ents:
            types.setdefault(e.type, []).append(e.canonique)
        verifier("dans un contrat, les deux sociétés sont détectées",
                 len(types.get("personne_morale", [])) == 2,
                 str(types.get("personne_morale")))
        verifier("l'article de tête est conservé dans la dénomination",
                 any("ATELIERS DU CARILLON" in x
                     for x in types.get("personne_morale", [])),
                 str(types.get("personne_morale")))
        verifier("une ville renvoyée à la ligne suivante est retrouvée",
                 any("LILLE" in a for a in types.get("adresse", [])),
                 str(types.get("adresse")))
        verifier("une ville à trait d'union et apostrophe est retrouvée",
                 any("VILLENEUVE-D'ASCQ" in a for a in types.get("adresse", [])),
                 str(types.get("adresse")))
        roles.attribuer_roles(contrat, ents, config_det)
        verifier("un contrat n'envoie personne du côté demandeur ou défendeur",
                 not any(e.partie.startswith(("DEM", "DEF")) for e in ents),
                 str([(e.canonique[:20], e.partie) for e in ents]))
        verifier("sans définition ni mot-clé, les parties d'un contrat sont PARTIE1 et PARTIE2, "
                 "à vérifier",
                 sorted(e.partie for e in ents if e.role == "Identité") == ["PARTIE1", "PARTIE2"],
                 str([(e.canonique[:20], e.partie) for e in ents if e.role == "Identité"]))

        verifier("deux sociétés citées côte à côte sont signalées à vérifier "
                 "plutôt que tranchées",
                 all(e.confiance != "haute" for e in det.detecter(
                     "les sociétés ALPHA ET BÉTA, immatriculées au RCS")
                     if e.type == "personne_morale" and "ET" in e.canonique.upper()),
                 str([(e.canonique, e.confiance) for e in det.detecter(
                     "les sociétés ALPHA ET BÉTA, immatriculées au RCS")]))

        verifier("un nom de personne ne traverse pas une fin de ligne",
                 all(len(e.canonique.split()) <= 4 for e in det.detecter(
                     "Monsieur Anselme Beaupertuis\n        Demandeur\n"
                     "        Ayant pour avocat")
                     if e.type in ("personne_physique", "representant_legal")),
                 str([e.canonique for e in det.detecter(
                     "Monsieur Anselme Beaupertuis\n        Demandeur")]))


        # ==============================================================
        print("\n18. Contrat : le rôle vient des définitions du contrat")
        texte_c = CONTRAT_CLIENT_ANONYME
        ents = det.detecter(texte_c)
        info = roles.attribuer_roles(texte_c, ents, config_det)
        verifier("le document est reconnu comme un contrat", info["type"] == "contrat",
                 str(info["scores"]))
        verifier("« ci-après le Prestataire » fait de la société le PRESTATAIRE",
                 trouver(ents, "NEBULA CLOUD").partie == "PRESTATAIRE")
        verifier("son siège et son dirigeant la suivent",
                 trouver(ents, "12 rue de la Paix").balise == "[PRESTATAIRE_ADRESSE]"
                 and trouver(ents, "Hugo LAMBERT").balise == "[PRESTATAIRE_DIRIGEANT]",
                 f"{trouver(ents, '12 rue de la Paix').balise} {trouver(ents, 'Hugo LAMBERT').balise}")
        verifier("un Client présenté sans nom ne crée aucune ligne en tête",
                 not any(e.partie == "CLIENT" and e.position < texte_c.find("ARTICLE 1")
                         for e in ents))
        verifier("le Client nommé au bloc de signature devient [CLIENT]",
                 trouver(ents, "ORION DISTRIBUTION").balise == "[CLIENT]")
        verifier("son signataire devient [CLIENT_SIGNATAIRE]",
                 trouver(ents, "Julie MOREAU").balise == "[CLIENT_SIGNATAIRE]",
                 trouver(ents, "Julie MOREAU").balise)
        verifier("le signataire du Prestataire n'est pas pris pour celui du Client",
                 trouver(ents, "Hugo LAMBERT").partie == "PRESTATAIRE")

        texte_l = CONTRAT_LICENCE
        ents = det.detecter(texte_l)
        roles.attribuer_roles(texte_l, ents, config_det)
        verifier("tout terme défini devient un rôle (« le Licencié » -> LICENCIE)",
                 trouver(ents, "ARCTURUS").partie == "LICENCIE"
                 and trouver(ents, "BOREAL SOFT").partie == "CONCEDANT",
                 f"{trouver(ents, 'ARCTURUS').partie} {trouver(ents, 'BOREAL SOFT').partie}")

        # ==============================================================
        print("\n19. Mise en demeure : les rôles sont vus depuis le cabinet")
        config_med = json.loads(json.dumps(config_det))
        config_med["cabinet"]["avocats"] = ["Aurore TESTEUR"]
        ents = det.detecter(MED_EMISE)
        info = roles.attribuer_roles(MED_EMISE, ents, config_med)
        verifier("le courrier est reconnu comme une mise en demeure",
                 info["type"] == "mise_en_demeure", str(info["scores"]))
        verifier("l'avocat du cabinet reconnu : le courrier est émis",
                 info["sens"] == roles.SENS_EMIS, info["sens"])
        verifier("émis : le client du cabinet est CLI",
                 trouver(ents, "GAMMA LOGICIELS").partie == "CLI")
        verifier("émis : le destinataire est la ADV",
                 trouver(ents, "DELTA FABRICATION").partie == "ADV")
        verifier("« ma cliente constate que la société X » ne fait pas de X notre client",
                 trouver(ents, "DELTA FABRICATION").partie != "CLI")
        verifier("l'avocat du cabinet est [CLI_AVOCAT]",
                 trouver(ents, "Aurore TESTEUR").balise == "[CLI_AVOCAT]")

        ents = det.detecter(MED_RECUE)
        info = roles.attribuer_roles(MED_RECUE, ents, config_med)
        verifier("reçu : le courrier est reconnu comme reçu", info["sens"] == roles.SENS_RECU,
                 info["sens"])
        verifier("reçu : notre client, destinataire, reste CLI",
                 trouver(ents, "GAMMA LOGICIELS").partie == "CLI")
        verifier("reçu : l'expéditeur est la ADV",
                 trouver(ents, "DELTA FABRICATION").partie == "ADV")
        verifier("reçu : l'avocat qui écrit est l'avocat adverse",
                 trouver(ents, "Bernard VOLANT").balise == "[ADV_AVOCAT]",
                 trouver(ents, "Bernard VOLANT").balise)

        config_vide = json.loads(json.dumps(config_det))
        config_vide["cabinet"]["avocats"] = []
        questions = []
        ents = det.detecter(MED_RECUE)
        info = roles.attribuer_roles(MED_RECUE, ents, config_vide,
                                     demander_sens=lambda: questions.append(1) or roles.SENS_RECU)
        verifier("sans liste du cabinet, le programme pose la question au lieu de deviner",
                 questions == [1] and info["sens"] == roles.SENS_RECU)
        ents = det.detecter(MED_RECUE)
        info = roles.attribuer_roles(MED_RECUE, ents, config_vide, sens=roles.SENS_RECU)
        verifier("le sens choisi dans la fenêtre fait foi",
                 trouver(ents, "GAMMA LOGICIELS").partie == "CLI")

        # ==============================================================
        print("\n20. Audit RGPD")
        ents = det.detecter(AUDIT)
        ents = sorted(ents + roles.detecter_outils(AUDIT, config_det, ents), key=lambda e: e.position)
        info = roles.attribuer_roles(AUDIT, ents, config_det)
        verifier("le document est reconnu comme un audit RGPD", info["type"] == "audit_rgpd",
                 str(info["scores"]))
        verifier("la société auditée est [ORGANISME_AUDITE]",
                 trouver(ents, "HELIOS MEDICAL").balise == "[ORGANISME_AUDITE]")
        verifier("le sous-traitant est [SOUS_TRAITANT1]",
                 trouver(ents, "OVH CLOUD").balise == "[SOUS_TRAITANT1]")
        verifier("le DSI est un référent", trouver(ents, "Paul GIRARD").partie == "REFERENT1")
        verifier("les logiciels nommés sont proposés comme outils",
                 trouver(ents, "Silae") is not None and trouver(ents, "Silae").partie.startswith("OUTIL"))
        verifier("un outil n'emporte pas la ponctuation ou le mot suivant",
                 all(not e.canonique.endswith((".", " L")) for e in ents if e.type == "outil"),
                 str([e.canonique for e in ents if e.type == "outil"]))

        # ==============================================================
        print("\n21. Appel, tiers et « à classer »")
        ents = det.detecter(CONCLUSIONS_APPEL)
        roles.attribuer_roles(CONCLUSIONS_APPEL, ents, config_det)
        verifier("la qualité écrite sous le bloc donne le rôle (INTIME1, APP1)",
                 trouver(ents, "LYRA TEXTILES").partie == "INTIME1"
                 and trouver(ents, "CASTOR INDUSTRIE").partie == "APP1")
        verifier("un expert nommé avec sa qualité après le nom est EXPERT1",
                 trouver(ents, "Sylvie BLANC").partie == "EXPERT1")
        tiers = trouver(ents, "ZETA TRANSPORTS")
        verifier("une société citée seulement dans les motifs reçoit une balise neutre, gardée",
                 tiers.partie.startswith("SOCIETE") and tiers.balise == f"[{tiers.partie}]"
                 and roles.proposer_a_retenir(tiers) == "OUI", tiers.partie)
        ents = det.detecter(ACTE + "\nLa société OMEGA LOGISTIQUE SAS a livré le matériel.\n")
        roles.attribuer_roles(ACTE, ents, config_det)
        verifier("AUCUN terme sans partie n'est versé du côté demandeur",
                 all(e.partie.startswith("DEM") is False
                     for e in ents if "OMEGA" in e.canonique))

        # ==============================================================
        print("\n22. Deuxième document du même dossier : les parties connues sont reprises")
        modele.creer_onglet(str(gabarit), "Dossier suite", config_det)
        outil.injecter(str(gabarit), "Dossier suite",
                       [{"terme": "VEGA INDUSTRIES SARL / VEGA INDUSTRIES", "partie": "DEF1",
                         "role": "Identité"},
                        {"terme": "ALDEBARAN SYSTEMES", "partie": "DEM1",
                         "role": "Identité"}], config_det)
        _, existantes = modele.lire_onglet(str(gabarit), "Dossier suite")
        texte_2 = CONCLUSIONS_SUITE
        ents = det.detecter(texte_2)
        roles.attribuer_roles(texte_2, ents, config_det, existantes=existantes)
        vega2 = trouver(ents, "VEGA INDUSTRIES")
        verifier("une partie déjà dans l'onglet garde sa balise, même avec un autre rôle "
                 "procédural", vega2.balise == "[DEF1]", vega2.balise)
        verifier("ses attributs suivent le code déjà connu",
                 trouver(ents, "11 rue Solferino").partie == "DEF1",
                 trouver(ents, "11 rue Solferino").partie)
        nouvelle = trouver(ents, "SIRIUS NEGOCE")
        verifier("une nouvelle partie prend un numéro libre",
                 nouvelle.partie not in ("DEF1", "DEM1") and nouvelle.partie,
                 nouvelle.partie)
        ald2 = trouver(ents, "ALDEBARAN SYSTEMES")
        verifier("un terme déjà présent n'est proposé que pour ses formes nouvelles",
                 ald2.balise == "[DEM1]"
                 and "ALDEBARAN SYSTEMES" not in ald2.variantes_nouvelles,
                 str(ald2.variantes_nouvelles))
        adr2 = trouver(ents, "11 rue Solferino")
        verifier("un attribut nouveau d'une partie connue est proposé à OUI",
                 roles.proposer_a_retenir(adr2) == "OUI")
        ents_bis = det.detecter(texte_2)
        ex_bis = existantes + [{"terme": "11 rue Solferino, 59000 LILLE", "balise": "[DEF1_ADRESSE]",
                                "partie": "DEF1", "role": "Adresse"}]
        roles.attribuer_roles(texte_2, ents_bis, config_det, existantes=ex_bis)
        verifier("un terme déjà présent à l'identique est proposé à NON",
                 roles.proposer_a_retenir(trouver(ents_bis, "11 rue Solferino")) == "NON")

        # ==============================================================
        print("\n23. Conversion d'un ancien onglet, avec archive")
        import convertir_onglets as conv
        import anonymisation_core as coeur
        table_c = racine / "Table_conversion.xlsx"
        ecrire_gabarit(table_c)
        wb = openpyxl.load_workbook(table_c)
        ws = wb["Dossier fictif"]
        remplis = {"[PM_DEM_1]": "ALDEBARAN SYSTEMES / ALDEBARAN",
                   "[ADR_DEM_1]": "14 rue des Trois Moulins, 59800 LILLE",
                   "[RL_DEM_1]": "Monsieur Jean DURAND / Jean DURAND",
                   "[PM_DEF_1]": "VEGA INDUSTRIES SARL / VEGA INDUSTRIES",
                   "[AV_DEF_1]": "Maître Paul BERTIN",
                   "[SIT_DEF_1]": "www.bertin-avocats.fr",
                   "[CJ_1]": "Maître Luc THEVENIN",
                   "[RG_1]": "24/01234"}
        for row in ws.iter_rows(min_row=4):
            b = row[1].value
            if b in remplis:
                row[0].value = remplis[b]
                row[5].value = "OUI"
        wb.save(table_c)
        config_outil = {"tables": [{"nom": "Conv", "fichier": str(table_c)}],
                        "dossier_journaux": str(racine / "Journaux")}
        chemin_cfg = racine / "config_test.json"
        chemin_cfg.write_text(json.dumps(config_outil), encoding="utf-8")

        # un document anonymise AVANT conversion, avec les anciennes balises
        doc_ancien = racine / "Courrier ancien.docx"
        docx_depuis_texte(doc_ancien, "La société ALDEBARAN SYSTEMES assigne VEGA INDUSTRIES "
                          "devant le tribunal (RG 24/01234), constat de Maître Luc THEVENIN.")
        cfg_c = coeur.charger_config(str(chemin_cfg))
        res_a = coeur.anonymiser_document(str(doc_ancien),
                                          coeur.charger_entrees(str(table_c), "Dossier fictif"),
                                          cfg_c, "Conv", "Dossier fictif")
        verifier("témoin : un document est anonymisé avec les anciennes balises",
                 "[PM_DEM_1]" in coeur.texte_du_document(res_a["chemin_sortie"]))

        verifier("l'onglet à l'ancien modèle est proposé à la conversion",
                 conv.onglets_ancien_modele(str(table_c)) == ["Dossier fictif"])
        chemin_classeur, lignes_c = conv.preparer(str(table_c), "Dossier fictif",
                                                  "DEM", "DEF", config_det,
                                                  dossier_sortie=str(racine))
        verifier("chaque ligne remplie est reprise, les cases vides ne le sont pas",
                 len(lignes_c) == len(remplis), f"{len(lignes_c)} / {len(remplis)}")
        prop = {l["Ancienne balise"]: l for l in lignes_c}
        verifier("[PM_DEM_1] devient [DEM1]", prop["[PM_DEM_1]"]["balise"] == "[DEM1]")
        verifier("[RL_DEM_1] devient le dirigeant de DEM1",
                 prop["[RL_DEM_1]"]["balise"] == "[DEM1_DIRIGEANT]")
        verifier("[SIT_DEF_1] rejoint le camp écrit dans la balise, pas la colonne Camp erronée",
                 prop["[SIT_DEF_1]"]["partie"] == "DEF1", prop["[SIT_DEF_1]"]["partie"])
        verifier("[CJ_1] devient [COMMISSAIRE1] et [RG_1] devient [RG1]",
                 prop["[CJ_1]"]["balise"] == "[COMMISSAIRE1]" and prop["[RG_1]"]["balise"] == "[RG1]")

        # l'onglet change entre preparation et application : refus
        wb = openpyxl.load_workbook(table_c)
        wb["Dossier fictif"]["A30"] = "MODIF"
        wb["Dossier fictif"]["B30"] = "[XX_1]"
        wb.save(table_c)
        verifier("un onglet modifié depuis la préparation n'est pas converti",
                 _leve(lambda: conv.appliquer(str(table_c), "Dossier fictif", chemin_classeur,
                                              config_det, str(chemin_cfg), "Conv")))
        wb = openpyxl.load_workbook(table_c)
        wb["Dossier fictif"]["A30"] = None
        wb["Dossier fictif"]["B30"] = None
        wb.save(table_c)

        res_c = conv.appliquer(str(table_c), "Dossier fictif", chemin_classeur, config_det,
                               str(chemin_cfg), "Conv")
        verifier("l'onglet est converti au modèle par partie",
                 modele.modele_onglet(str(table_c), "Dossier fictif") == "par_partie")
        archive = res_c["archive"]
        verifier("une archive est créée à côté de la table",
                 os.path.isfile(archive) and "_ARCHIVE_anciennes_balises" in archive)
        verifier("l'archive conserve l'onglet à l'ancien modèle",
                 modele.modele_onglet(archive, "Dossier fictif") == "ancien")
        cfg_apres = json.loads(chemin_cfg.read_text(encoding="utf-8"))
        verifier("l'archive est ajoutée à la liste des tables",
                 any(t["fichier"] == archive for t in cfg_apres["tables"]))
        _, conv_entrees = modele.lire_onglet(str(table_c), "Dossier fictif")
        verifier("l'ancienne balise est notée dans la colonne Notes",
                 all("ancienne balise" in x["notes"] for x in conv_entrees))

        # restitution sans journal de l'ANCIEN document : par l'archive
        res_r = coeur.desanonymiser_sans_journal(
            res_a["chemin_sortie"], coeur.charger_entrees(archive, "Dossier fictif"),
            confirme=True, config=coeur.charger_config(str(chemin_cfg)))
        verifier("un document aux anciennes balises se restitue sans journal par l'archive",
                 "ALDEBARAN SYSTEMES" in coeur.texte_du_document(res_r["chemin_sortie"])
                 and "[PM_DEM_1]" not in coeur.texte_du_document(res_r["chemin_sortie"]))
        verifier("la table convertie refuse ce document (balises inconnues) : pas de mélange",
                 _leve(lambda: coeur.desanonymiser_sans_journal(
                     res_a["chemin_sortie"], coeur.charger_entrees(str(table_c), "Dossier fictif"),
                     confirme=True, config=coeur.charger_config(str(chemin_cfg)))))

        # un nouveau document, anonymise avec l'onglet converti
        doc_neuf = racine / "Courrier neuf.docx"
        docx_depuis_texte(doc_neuf, "VEGA INDUSTRIES répond à ALDEBARAN SYSTEMES.")
        os.remove(res_r["chemin_sortie"])
        res_n = coeur.anonymiser_document(str(doc_neuf),
                                          coeur.charger_entrees(str(table_c), "Dossier fictif"),
                                          coeur.charger_config(str(chemin_cfg)), "Conv",
                                          "Dossier fictif")
        texte_n = coeur.texte_du_document(res_n["chemin_sortie"])
        verifier("un nouveau document porte les balises par partie",
                 "[DEF1]" in texte_n and "[DEM1]" in texte_n, texte_n)
        res_n2 = coeur.desanonymiser_sans_journal(
            res_n["chemin_sortie"], coeur.charger_entrees(str(table_c), "Dossier fictif"),
            confirme=True, config=coeur.charger_config(str(chemin_cfg)))
        verifier("et se restitue sans journal avec la table convertie",
                 "ALDEBARAN SYSTEMES" in coeur.texte_du_document(res_n2["chemin_sortie"])
                 and "[" not in coeur.texte_du_document(res_n2["chemin_sortie"]),
                 coeur.texte_du_document(res_n2["chemin_sortie"]))

        print("\n24. IBAN, numéro de sécurité sociale, date de naissance")
        acte_id = """TRIBUNAL JUDICIAIRE DE LILLE
ASSIGNATION

POUR :

Madame Claire FONTAINE, née le 14 février 1979 à Lille, demeurant 3 rue Solferino, 59000 LILLE,
immatriculée à la sécurité sociale sous le n° 2 55 08 14 168 025 38,

DEMANDERESSE

CONTRE :

La société VEGA INDUSTRIES SARL, dont le siège est 27 boulevard Victor Hugo, 75008 PARIS,
titulaire du compte IBAN FR76 3000 6000 0112 3456 7890 189 PARIS,

DEFENDERESSE

FAITS

Madame FONTAINE a versé la somme sur le compte FR76 3000 6000 0112 3456 7890 189.
Un second compte BE68 5390 0754 7034 a été cité. Le relevé porte un IBAN mal recopié :
FR76 3000 6000 0112 3456 7890 188.
Facture 1 85 05 78 006 084 12 du 3 mars 2025, référence ABCD EFGH IJKL MNOP.
"""
        ents = det.detecter(acte_id)
        roles.attribuer_roles(acte_id, ents, config_det)
        iban = trouver(ents, "FR76 3000 6000 0112 3456 7890 189")
        verifier("IBAN français détecté, clé juste, mot suivant écarté",
                 iban is not None and iban.type == "iban" and iban.confiance == "haute"
                 and iban.canonique.endswith("189"), repr(iban))
        verifier("l'IBAN garde aussi sa forme sans espaces",
                 iban is not None and "FR7630006000011234567890189" in iban.variantes)
        verifier("IBAN belge (16 caractères) détecté",
                 any(e.type == "iban" and "BE68" in e.canonique for e in ents))
        faux = trouver(ents, "7890 188")
        verifier("IBAN à clé fausse : gardé, mais « à vérifier »",
                 faux is not None and faux.type == "iban" and faux.confiance == "à vérifier",
                 repr(faux))
        verifier("une référence en capitales n'est pas un IBAN",
                 not any("ABCD" in v for e in ents for v in e.variantes))
        nir = trouver(ents, "2 55 08 14 168 025 38")
        verifier("NIR détecté, clé juste", nir is not None and nir.type == "nir"
                 and nir.confiance == "haute", repr(nir))
        verifier("un numéro de facture sans clé ni contexte n'est pas un NIR",
                 not any(e.type == "nir" and "084 12" in e.canonique for e in ents))
        naissance = trouver(ents, "14 février 1979")
        verifier("date de naissance détectée après « née le »",
                 naissance is not None and naissance.type == "date_naissance", repr(naissance))
        verifier("une date ordinaire n'est pas une date de naissance",
                 not any(e.type == "date_naissance" and "2025" in e.canonique for e in ents))
        verifier("rôles proposés : [DEM1_NIR], [DEM1_DATE_NAISSANCE], [DEF1_IBAN]",
                 nir is not None and nir.balise == "[DEM1_NIR]"
                 and naissance is not None and naissance.balise == "[DEM1_DATE_NAISSANCE]"
                 and iban is not None and iban.balise == "[DEF1_IBAN]",
                 f"{getattr(nir, 'balise', '')} {getattr(naissance, 'balise', '')} "
                 f"{getattr(iban, 'balise', '')}")
        corse = det._detecter_nir("assuré n° 1 80 01 2A 123 456 20")
        verifier("NIR corse (2A) : clé calculée juste",
                 len(corse) == 1 and corse[0].confiance == "haute")
        anglais = det._detecter_date_naissance("Date of birth: March 5, 1990. Born on 1st May 1988.")
        verifier("dates de naissance en anglais détectées", len(anglais) == 2,
                 str([e.canonique for e in anglais]))

        # Un role ajoute au seul fichier de reglages, sans toucher au code
        cfg_plus = json.loads(json.dumps(config_det))
        cfg_plus["attributs"].append({"libelle": "Passeport", "suffixe": "PASSEPORT"})
        verifier("un nouveau rôle s'ajoute par le seul fichier de réglages",
                 "Passeport" in modele.libelles_attributs(cfg_plus)
                 and modele.calculer_balise("DEM1", "Passeport", set(), cfg_plus)[0]
                 == "[DEM1_PASSEPORT]")
        for nom_cfg in ("fr", "en"):
            c = json.loads((SCRIPT_DIR / "_build" / "langue" / nom_cfg / "detection_config.json")
                           .read_text(encoding="utf-8"))
            libs = {a["libelle"] for a in c["attributs"]}
            verifier(f"réglages {nom_cfg.upper()} du paquet : IBAN, NIR, date de naissance déclarés",
                     all(t in c["types_detection_vers_attribut"] for t in ("iban", "nir", "date_naissance"))
                     and all(c["types_detection_vers_attribut"][t] in libs
                             for t in ("iban", "nir", "date_naissance")))

        print("\n24. Revue du 8 octobre 2026")
        t = ("Monsieur Jean DURAND, demeurant à Lille, et Madame Marie DURAND, son épouse. "
             "Monsieur DURAND a signé. Mme DURAND aussi.")
        pers = [e for e in det.detecter(t) if e.type == "personne_physique"]
        verifier("deux personnes du même patronyme ne sont pas fusionnées",
                 {e.canonique for e in pers} == {"Monsieur Jean DURAND", "Madame Marie DURAND"},
                 str([(e.canonique, sorted(e.variantes)) for e in pers]))
        verifier("et chacune est marquée à vérifier (patronyme partagé)",
                 all(e.confiance == "à vérifier" for e in pers))
        verifier("« Mme DURAND » n'est pas une variante de Monsieur",
                 all("Mme DURAND" not in e.variantes for e in pers if e.canonique.startswith("Monsieur")))
        tels = [e.canonique for e in det.detecter("Tél. +33 6 12 34 56 78 ou tel:+33612345678.") if e.type == "telephone"]
        verifier("un numéro en +33 est détecté", tels == ["+33 6 12 34 56 78", "+33612345678"], str(tels))
        adr = [e.canonique for e in det.detecter("Siège : 1 rue des Lilas 59000 Lille et Madame Y, née à 2 rue Haute, 59000 Lille né le 3 mai.")
               if e.type == "adresse"]
        verifier("la ville n'avale pas les mots suivants",
                 adr == ["1 rue des Lilas 59000 Lille", "2 rue Haute, 59000 Lille"], str(adr))
        pm = [e.canonique for e in det.detecter("Il a remis sa Proposition Commerciale. La société MIMOSA et la NASA. BETA SARL puis la SARL BETA.")
              if e.type == "personne_morale"]
        verifier("« sa Proposition » et « MIMOSA » ne sont pas des sociétés", pm == ["MIMOSA", "BETA SARL"], str(pm))
        ents = det.detecter("RCS LILLE 123 456 789\n\n2. Attendu. TVA FR 12 345 678 901. Brevet EP 1 234 567.")
        verifier("un numéro ne traverse pas le saut de paragraphe et la TVA n'est pas un brevet",
                 [e.canonique for e in ents if e.type == "rcs"] == ["RCS LILLE 123 456 789"]
                 and [e.canonique for e in ents if e.type == "brevet"] == ["EP 1 234 567"],
                 str([(e.type, e.canonique) for e in ents]))
        contrat = ("Entre la société ALPHA SAS, dont le siège est 12 rue des Lilas, 59000 Lille, "
                   "représentée par Monsieur Jean DURAND, ci-après le « Prestataire »,\n"
                   "Et la société BETA SARL, ci-après le « Client ». Le Prestataire s'engage.")
        ents = det.detecter(contrat)
        infos = roles.attribuer_roles(contrat, ents, config_det)
        sortie_c = racine / "contrat_revue.xlsx"
        outil.ecrire_classeur_detection(str(sortie_c), ents, infos, config_det, "c.docx", "T", "Dossier")
        verifier("le classeur de détection s'écrit pour un contrat (parties non numérotées)",
                 sortie_c.is_file())
        verifier("le mot « Prestataire » n'est pas rattaché comme identité de la partie",
                 all(not e.partie for e in ents if e.type == "candidat"))

        print("\n25. Notes, commentaires et zones de texte lus par la pré-détection (8/10/2026)")
        import zipfile
        base = racine / "annexes_base.docx"
        docx_depuis_texte(base, "Assignation de la société ALPHA SAS contre Monsieur Jean DURAND.")
        annexes = racine / "annexes.docx"
        ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

        def para(t):
            return f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>"
        parties = {
            "word/footnotes.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {ns}>'
                                  f'<w:footnote w:id="2">{para("Voir Maître Paul LEGRAND, avocat au barreau de Lille.")}</w:footnote></w:footnotes>',
            "word/endnotes.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes {ns}>'
                                 f'<w:endnote w:id="2">{para("Contact : paul.legrand@exemple.invalid")}</w:endnote></w:endnotes>',
            "word/comments.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {ns}>'
                                 f'<w:comment w:id="1">{para("Vérifier la société GAMMA SARL.")}</w:comment></w:comments>',
        }
        types = {"word/footnotes.xml": "footnotes", "word/endnotes.xml": "endnotes", "word/comments.xml": "comments"}
        zone = ('<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                '<mc:Choice Requires="wps"><w:drawing><wps:txbx xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
                f'<w:txbxContent>{para("Madame Sophie MARTIN, tél. 03 20 11 22 33")}</w:txbxContent></wps:txbx></w:drawing></mc:Choice>'
                '<mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox>'
                f'<w:txbxContent>{para("Madame Sophie MARTIN, tél. 03 20 11 22 33")}</w:txbxContent>'
                '</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>')
        with zipfile.ZipFile(base) as z1, zipfile.ZipFile(annexes, "w", zipfile.ZIP_DEFLATED) as z2:
            for it in z1.infolist():
                data = z1.read(it.filename)
                if it.filename == "[Content_Types].xml":
                    ajout = "".join(f'<Override PartName="/{p}" ContentType="application/vnd.openxmlformats-'
                                    f'officedocument.wordprocessingml.{t}+xml"/>' for p, t in types.items())
                    data = data.replace(b"</Types>", ajout.encode("utf-8") + b"</Types>")
                elif it.filename == "word/document.xml":
                    data = data.replace(b"</w:body>", zone.encode("utf-8") + b"<w:sectPr/></w:body>", 1) \
                        if b"<w:sectPr" not in data else data.replace(b"<w:sectPr", zone.encode("utf-8") + b"<w:sectPr", 1)
                z2.writestr(it, data)
            for nom, xml in parties.items():
                z2.writestr(nom, xml)
        texte = det.texte_docx(str(annexes))
        verifier("la note de bas de page est lue", "Paul LEGRAND" in texte)
        verifier("la note de fin est lue", "paul.legrand@exemple.invalid" in texte)
        verifier("le commentaire est lu", "GAMMA SARL" in texte)
        verifier("la zone de texte est lue, une seule fois (forme moderne et forme de secours)",
                 texte.count("Sophie MARTIN") == 1, str(texte.count("Sophie MARTIN")))
        verifier("le corps reste en tête du texte",
                 texte.index("ALPHA SAS") < texte.index("Paul LEGRAND"))
        ents = det.detecter(texte)
        verifier("la détection y trouve l'avocat, le courriel, la société et la personne de la zone de texte",
                 any(e.type == "avocat" and "LEGRAND" in e.canonique for e in ents)
                 and any(e.type == "email" for e in ents)
                 and any("GAMMA" in e.canonique for e in ents)
                 and any("Sophie MARTIN" in e.canonique for e in ents),
                 str([(e.type, e.canonique) for e in ents]))
        verifier("un document sans annexes se lit comme avant",
                 det.texte_docx(str(base)).strip() == "Assignation de la société ALPHA SAS contre Monsieur Jean DURAND.")

        print("\n26. Refonte du tableau de détection (15/10/2026)")
        acte_d = """TRIBUNAL JUDICIAIRE DE LILLE
CONCLUSIONS

POUR :

La société ALDEBARAN SYSTEMES SAS, dont le siège est 14 rue des Trois Moulins, 59000 LILLE,
représentée par son président, Monsieur Jean DURAND,

DEMANDERESSE

CONTRE :

La société VEGA INDUSTRIES SARL, dont le siège est 27 boulevard Victor Hugo, 75008 PARIS,

DEFENDERESSE

La société ORION LOGISTIQUE SAS, dont le siège est 5 rue de la Gare, 69000 LYON,

DEFENDERESSE

FAITS

Monsieur Paul TESTARD, président directeur général de la société ORION LOGISTIQUE, a reconnu les faits. Plus loin, Monsieur TESTARD a confirmé. La société VEGA INDUSTRIES, représentée par son gérant, Monsieur Luc PAVON, a contesté. M. PAVON n'a rien ajouté.

Monsieur Marc DOUBLE, gérant de la société VEGA INDUSTRIES, a signé. Puis Monsieur DOUBLE, président de la société ORION LOGISTIQUE, a refusé.

L'intérêt à agir a été mis en évidence par MM. Cornu et Foyer. « Une association a le droit de se défendre » selon la cour d'appel de Versailles. La SCP LE STANC, société Civile Professionnelle d'avocats. La cour d'appel de Versailles a jugé de même. Les parties visent le Bâtiment A puis le Bâtiment B.
"""
        ents = det.detecter(acte_d)
        infos = roles.attribuer_roles(acte_d, ents, config_det)
        testard = trouver(ents, "TESTARD")
        verifier("dirigeant nommé dans le corps : rattaché sur formule explicite",
                 testard is not None and testard.partie == "DEF2" and testard.role == "Dirigeant",
                 repr(testard and (testard.partie, testard.role)))
        verifier("le rattachement vaut pour toutes ses écritures (une seule entité)",
                 testard is not None and "Monsieur TESTARD" in testard.variantes, str(testard and testard.variantes))
        pavon = trouver(ents, "PAVON")
        verifier("« représentée par son gérant, Monsieur X » : dirigeant de la société citée",
                 pavon is not None and pavon.partie == "DEF1" and pavon.role == "Dirigeant",
                 repr(pavon and (pavon.partie, pavon.role)))
        double = trouver(ents, "DOUBLE")
        verifier("dit dirigeant de deux sociétés : aucun rattachement, balise neutre et les deux signalées",
                 double is not None and double.partie.startswith("PERSONNE_PHYSIQUE")
                 and "plusieurs sociétés" in double.note,
                 repr(double and (double.partie, double.note)))
        verifier("auteurs cités (« MM. Cornu ») écartés", trouver(ents, "Cornu") is None)
        verifier("« Une association » et « société Civile Professionnelle » ne sont pas des sociétés",
                 trouver(ents, "Une association") is None and trouver(ents, "Civile Professionnelle") is None)
        verifier("ville de juridiction et mot d'adresse écartés des candidats",
                 not any(e.type == "candidat" and e.canonique in ("Versailles", "Bâtiment") for e in ents),
                 str([e.canonique for e in ents if e.type == "candidat"]))

        # Onglet déjà rempli : DEF1 y désigne une personne ; une identité prise
        # se signale en rouge, et la feuille Parties montre l'onglet d'abord.
        existantes_d = [{"terme": "Monsieur Pierre AUTRE / AUTRE", "balise": "[DEF1]",
                         "partie": "DEF1", "role": "Identité"}]
        ents = det.detecter(acte_d)
        infos = roles.attribuer_roles(acte_d, ents, config_det, existantes=existantes_d)
        vega = trouver(ents, "VEGA INDUSTRIES")
        verifier("une partie détectée ne reprend jamais un code déjà pris dans l'onglet",
                 vega is not None and vega.partie not in ("DEF1", "") and vega.balise != "[DEF1]",
                 repr(vega and vega.partie))
        sortie_d = racine / "_Detection_refonte.xlsx"
        outil.ecrire_classeur_detection(str(sortie_d), ents, infos, config_det, "d.docx", "T", "Dossier",
                                        existantes_d)
        wb = openpyxl.load_workbook(sortie_d)
        menu = [c.value for c in wb["Listes"]["A"][1:] if c.value and not str(c.value).startswith("=")]
        verifier("le menu Partie nomme la partie de l'onglet",
                 any(m.startswith("DEF1  =  Monsieur Pierre AUTRE") for m in menu), str(menu))
        verifier("les identités de l'onglet alimentent le contrôle « rouge »",
                 "[DEF1]" in [c.value for c in wb["Listes"]["F"]])
        verifier("la feuille Parties commence par ce qui est déjà dans l'onglet",
                 wb["Parties"]["C4"].value == "[DEF1]" and "onglet" in str(wb["Parties"]["E4"].value))
        # L'utilisateur range un terme : Partie et Rôle choisis dans les menus,
        # sans que la formule ait été calculée (fichier jamais ouvert dans Excel).
        ws = wb["Détection"]
        ligne_d = next(r for r in range(outil.LIGNE_ENTETES + 1, ws.max_row + 1)
                       if "DOUBLE" in str(ws.cell(row=r, column=1).value or "")
                       and ws.cell(row=r, column=outil.COL_NATURE).value)
        libelle = next(m for m in menu if m.startswith("DEF2  ="))
        ws.cell(row=ligne_d, column=outil.COL_PARTIE, value=libelle)
        ws.cell(row=ligne_d, column=outil.COL_ROLE, value="Dirigeant")
        ws.cell(row=ligne_d, column=outil.COL_RETENIR, value="OUI")
        wb.save(sortie_d)
        wb.close()
        relues = outil.lire_classeur_detection(str(sortie_d), config_det)
        choisie = next((l for l in relues if "DOUBLE" in l["terme"]), None)
        verifier("la partie choisie dans le menu donne la balise ([DEF2_DIRIGEANT])",
                 choisie is not None and choisie["balise"] == "[DEF2_DIRIGEANT]"
                 and choisie["partie"] == "DEF2", repr(choisie))


        print("\n27. Rattacher après coup une balise neutre (15/10/2026)")
        modele.CONFIG = config_det
        table_r = racine / "Correspondances_rattachement.xlsx"
        openpyxl.Workbook().save(table_r)
        modele.creer_onglet(str(table_r), "Dossier R", config_det)
        ecr, ref = modele.injecter(str(table_r), "Dossier R",
                                   [{"terme": "la société ALPHA SAS / ALPHA SAS", "partie": "DEF1", "role": "Identité"},
                                    {"terme": "3 rue Haute, 59000 Lille", "partie": "COMMUN", "role": "Adresse"}],
                                   config_det)
        _, ent_r = modele.lire_onglet(str(table_r), "Dossier R")
        verifier("une adresse sans partie certaine reçoit une balise neutre",
                 any(x["balise"] == "[ADRESSE1]" for x in ent_r), str([x["balise"] for x in ent_r]))
        cfg_r = {"tables": [{"nom": "R", "fichier": str(table_r)}], "dossier_journaux": str(racine / "Journaux")}
        chemin_cfg_r = racine / "config_r.json"
        chemin_cfg_r.write_text(json.dumps(cfg_r), encoding="utf-8")
        doc_r1 = racine / "Piece R1.docx"
        docx_depuis_texte(doc_r1, "La société ALPHA SAS a son siège au 3 rue Haute, 59000 Lille.")
        res_r1 = coeur.anonymiser_document(str(doc_r1), coeur.charger_entrees(str(table_r), "Dossier R"),
                                           coeur.charger_config(str(chemin_cfg_r)), "R", "Dossier R")
        verifier("témoin : la pièce 1 porte [ADRESSE1]",
                 "[ADRESSE1]" in coeur.texte_du_document(res_r1["chemin_sortie"]))
        # L'utilisateur rattache l'adresse a DEF1 dans l'onglet (colonne Partie)
        wb = openpyxl.load_workbook(table_r)
        ws = wb["Dossier R"]
        for row in ws.iter_rows():
            if row[1].value == "[ADRESSE1]":
                row[3].value = "DEF1"
        wb.save(table_r)
        wb.close()
        neuves = coeur.charger_entrees(str(table_r), "Dossier R")
        verifier("après rattachement, la ligne prend sa nouvelle balise ([DEF1_ADRESSE])",
                 any(b == "[DEF1_ADRESSE]" for b, _ in neuves) and not any(b == "[ADRESSE1]" for b, _ in neuves),
                 str([b for b, _ in neuves]))
        doc_r2 = racine / "Piece R2.docx"
        docx_depuis_texte(doc_r2, "Courrier reçu au 3 rue Haute, 59000 Lille.")
        res_r2 = coeur.anonymiser_document(str(doc_r2), neuves, coeur.charger_config(str(chemin_cfg_r)),
                                           "R", "Dossier R")
        verifier("une nouvelle pièce porte [DEF1_ADRESSE]",
                 "[DEF1_ADRESSE]" in coeur.texte_du_document(res_r2["chemin_sortie"]))
        os.remove(doc_r1)
        res_rr = coeur.desanonymiser_sans_journal(
            res_r1["chemin_sortie"], coeur.charger_entrees(str(table_r), "Dossier R", anciennes=True),
            confirme=True, config=coeur.charger_config(str(chemin_cfg_r)))
        verifier("la pièce anonymisée AVANT le rattachement se restitue sans journal",
                 "3 rue Haute" in coeur.texte_du_document(res_rr["chemin_sortie"])
                 and "[" not in coeur.texte_du_document(res_rr["chemin_sortie"]),
                 coeur.texte_du_document(res_rr["chemin_sortie"]))
        # Une injection suivante reecrit l'onglet : l'ancienne balise est gardee
        modele.injecter(str(table_r), "Dossier R",
                        [{"terme": "Maître Jean UN", "partie": "DEF1", "role": "Avocat"}], config_det)
        wb = openpyxl.load_workbook(table_r)
        ws = wb["Dossier R"]
        ligne = next(row for row in ws.iter_rows() if row[0].value == "3 rue Haute, 59000 Lille")
        verifier("après réécriture : nouvelle balise en B, ancienne gardée (colonne masquée)",
                 ligne[1].value == "[DEF1_ADRESSE]" and "[ADRESSE1]" in str(ligne[8].value)
                 and ws.column_dimensions["I"].hidden, f"{ligne[1].value} / {ligne[8].value}")
        wb.close()
        verifier("et la restitution sans journal reconnaît toujours les deux",
                 {"[ADRESSE1]", "[DEF1_ADRESSE]"} <= {b for b, _ in coeur.charger_entrees(str(table_r), "Dossier R", anciennes=True)})

        print("\n28. Revue du 15/10/2026 (soir)")
        e_ocr = det.EntiteDetectee("Monsieur Jean DURAND", "personne_physique", 0)
        e_ocr.existante = {"balise": "[DEF1]", "partie": "DEF1", "role": "Identité"}
        e_ocr.variantes_nouvelles = ["DURAND"]; e_ocr.ocr = True; e_ocr.partie, e_ocr.role = "DEF1", "Identité"
        verifier("une variante nouvelle lue par OCR reste à NON", roles.proposer_a_retenir(e_ocr) == "NON")
        e_ocr.ocr = False; e_ocr.ambigu = True
        verifier("une variante nouvelle d'écriture partagée reste à NON", roles.proposer_a_retenir(e_ocr) == "NON")
        # Note tapee a la main en colonne H : jamais prise pour un etat du programme
        table_h = racine / "colonne_h.xlsx"
        openpyxl.Workbook().save(table_h)
        modele.creer_onglet(str(table_h), "H", config_det)
        modele.injecter(str(table_h), "H", [{"terme": "ALPHA SAS", "partie": "DEF1", "role": "Identité"},
                                            {"terme": "3 rue Haute", "partie": "DEF1", "role": "Adresse"}], config_det)
        wb = openpyxl.load_workbook(table_h)
        for row in wb["H"].iter_rows():
            if row[0].value == "3 rue Haute":
                row[7].value = "voir | pièce 3"
        wb.save(table_h)
        wb.close()
        _, ent_h = modele.lire_onglet(str(table_h), "H")
        verifier("une note en colonne H ne change jamais une balise",
                 [x["balise"] for x in ent_h if x["terme"]] == ["[DEF1]", "[DEF1_ADRESSE]"]
                 and [b for b, _ in coeur.charger_entrees(str(table_h), "H")] == ["[DEF1]", "[DEF1_ADRESSE]"],
                 str([x["balise"] for x in ent_h]))
        # Trou dans la numerotation : jamais rebouche
        ex_trou = [{"terme": "UN SAS", "balise": "[SOCIETE1]", "partie": "SOCIETE1", "role": "Identité"},
                   {"terme": "TROIS SAS", "balise": "[SOCIETE3]", "partie": "SOCIETE3", "role": "Identité"}]
        ents = det.detecter("La société QUATRE SAS a signé.")
        roles.attribuer_roles("La société QUATRE SAS a signé.", ents, config_det, existantes=ex_trou)
        verifier("une partie nouvelle prend le numéro suivant le plus grand (SOCIETE4), jamais un trou",
                 trouver(ents, "QUATRE").partie == "SOCIETE4", trouver(ents, "QUATRE").partie)
        # Meme adresse a la ponctuation pres : variante, pas nouvelle ligne
        ecr_v, _ = modele.injecter(str(table_h), "H", [{"terme": "3 rue Haute / 3, rue Haute", "partie": "DEF1", "role": "Adresse"}], config_det)
        verifier("une adresse qui ne diffère que par la ponctuation rejoint sa ligne (variante)",
                 ecr_v and ecr_v[0][2] == "variantes ajoutées" and ecr_v[0][1] == "[DEF1_ADRESSE]", str(ecr_v))
        ex_adr = [{"terme": "3 rue Haute, 59000 Lille", "balise": "[DEF1_ADRESSE]", "partie": "DEF1", "role": "Adresse"}]
        t_adr = "Siège : 3 rue Haute 59000 Lille."
        ents = det.detecter(t_adr)
        roles.attribuer_roles(t_adr, ents, config_det, existantes=ex_adr)
        verifier("la détection rattache aussi cette adresse à sa ligne existante",
                 trouver(ents, "3 rue Haute") is not None and trouver(ents, "3 rue Haute").balise == "[DEF1_ADRESSE]",
                 repr(trouver(ents, "3 rue Haute")))
        # Societe neutre : son dirigeant et son mail la rejoignent
        t_n = ("La société ALPHA SAS, demanderesse, assigne la société BETA SARL.\n\nFAITS\n\n"
               "Monsieur Paul NEUF, gérant de la société GAMMA SAS, a écrit (contact : paul@gamma.example).")
        ents = det.detecter(t_n)
        roles.attribuer_roles(t_n, ents, config_det)
        gamma = trouver(ents, "GAMMA SAS")
        verifier("une société sans partie certaine devient SOCIETE1 et son gérant son dirigeant",
                 gamma is not None and gamma.partie.startswith("SOCIETE")
                 and trouver(ents, "NEUF").partie == gamma.partie and trouver(ents, "NEUF").role == "Dirigeant",
                 repr((gamma and gamma.partie, trouver(ents, "NEUF").partie)))
        verifier("et le mail cité à côté d'elle la rejoint",
                 trouver(ents, "paul@gamma").partie == gamma.partie, trouver(ents, "paul@gamma").partie)
        # Tableau : une ligne deja dans l'onglet ne compte pas dans le rang des suivantes
        ents = det.detecter(t_adr + " Autre siège : 9 rue Basse, 59000 Lille.")
        infos_adr = roles.attribuer_roles(t_adr, ents, config_det, existantes=ex_adr)
        sortie_adr = racine / "_Detection_adr.xlsx"
        outil.ecrire_classeur_detection(str(sortie_adr), ents, infos_adr, config_det, "a.docx", "T", "H", ex_adr)
        ws = openpyxl.load_workbook(sortie_adr)["Détection"]
        r_ex = next(r for r in range(outil.LIGNE_ENTETES + 1, ws.max_row + 1)
                    if "3 rue Haute" in str(ws.cell(row=r, column=1).value or "") and ws.cell(row=r, column=outil.COL_NATURE).value)
        verifier("la clé de rang d'une ligne existante est vide tant que Partie et Rôle n'ont pas bougé",
                 str(ws.cell(row=r_ex, column=outil.C_CLE).value).startswith('=IF(AND(D'))

        print("\n29. Une balise ne révèle jamais le nom qu'elle masque (19/10/2026)")
        cg = ("CONDITIONS GENERALES\n\nENTRE\nLa société AGICE EDITIQUE, SAS immatriculée au RCS de Nantes sous le n° 448 111 013, "
              "dont le siège est 7 rue Edouard Branly à SAINTE-LUCE-SUR-LOIRE (44980), représentée par Monsieur Eric LAFONTAINE, "
              "ci-après « AGICE EDITIQUE » ou « AGICE »,\nET\nLe Client, ci-après le « Client ».\n\nARTICLE 1 - OBJET\n"
              "AGICE fournit au Client des prestations. La société Yousign est sous-traitante.\n")
        ents = det.detecter(cg)
        infos_cg = roles.attribuer_roles(cg, ents, config_det)
        agice = trouver(ents, "AGICE EDITIQUE")
        verifier("un terme défini qui est le nom de la société ne devient jamais un code de partie",
                 agice is not None and "AGICE" not in agice.partie and "AGICE" not in agice.balise
                 and all("AGICE" not in e.balise for e in ents), repr([(e.canonique, e.balise) for e in ents]))
        verifier("la société définie par son nom reçoit l'autre rôle du contrat (PRESTATAIRE, l'autre partie étant le Client)",
                 agice.partie == "PRESTATAIRE" and "à vérifier" in agice.note, repr((agice.partie, agice.note)))
        verifier("ses SIRET, adresse et dirigeant la suivent",
                 trouver(ents, "448 111 013").balise == "[PRESTATAIRE_SIRET]"
                 and trouver(ents, "LAFONTAINE").balise == "[PRESTATAIRE_DIRIGEANT]")
        verifier("un rôle défini (« ci-après le Client ») reste reconnu",
                 roles._definitions(roles._Contexte(cg, ents, config_det, "contrat"))[1][1] == "CLIENT")
        ctx_f = roles._Contexte("La société AGICE EDITIQUE a signé.", det.detecter("La société AGICE EDITIQUE a signé."), config_det, "contrat")
        ef = ctx_f.entites[0]
        ctx_f.parties["AGICE"] = ef; ctx_f._attribuer(ef, "AGICE", "Identité")
        roles._sans_fuite(ctx_f)
        verifier("le filet renomme une partie dont le code reprend le nom (AGICE -> SOCIETE1)",
                 ef.partie.startswith("SOCIETE") and "reprenait le nom" in ef.note, repr((ef.partie, ef.note)))
        table_f = racine / "fuite.xlsx"
        openpyxl.Workbook().save(table_f)
        modele.creer_onglet(str(table_f), "F", config_det)
        ecr_f, ref_f = modele.injecter(str(table_f), "F",
                                       [{"terme": "AGICE EDITIQUE / AGICE", "partie": "AGICE_EDITIQUE", "role": "Identité"},
                                        {"terme": "7 rue Branly", "partie": "AGICE_EDITIQUE", "role": "Adresse"},
                                        {"terme": "Le Fournisseur Pro SAS", "partie": "FOURNISSEUR", "role": "Identité"},
                                        {"terme": "Monsieur Jean SOCIETE", "partie": "SOCIETE1", "role": "Identité"}], config_det)
        verifier("l'injection refuse une partie qui reprend le nom (identité ET adresse), pas les mots génériques",
                 len(ref_f) == 2 and all("reprend le nom" in r for _, r in ref_f)
                 and [e[1] for e in ecr_f] == ["[FOURNISSEUR]", "[SOCIETE1]"], str((ecr_f, ref_f)))
        wb = openpyxl.load_workbook(table_f)
        for row in wb["F"].iter_rows():
            if row[0].value == "Le Fournisseur Pro SAS":
                row[1].value = "[FOURNISSEUR_PRO]"
        wb.save(table_f)
        wb.close()
        cfg_f = {"tables": [{"nom": "F", "fichier": str(table_f)}], "dossier_journaux": str(racine / "Journaux")}
        chemin_cfg_f = racine / "config_f.json"
        chemin_cfg_f.write_text(json.dumps(cfg_f), encoding="utf-8")
        doc_f = racine / "Piece F.docx"
        docx_depuis_texte(doc_f, "Le Fournisseur Pro SAS a livré.")
        res_f = coeur.anonymiser_document(str(doc_f), coeur.charger_entrees(str(table_f), "F"),
                                          coeur.charger_config(str(chemin_cfg_f)), "F", "F")
        ctl_f = coeur.controler(res_f["chemin_sortie"], coeur.charger_entrees(str(table_f), "F"),
                                coeur.charger_config(str(chemin_cfg_f)))
        verifier("le contrôle met en ALERTE une balise qui révèle le nom ([FOURNISSEUR_PRO] pour « Fournisseur Pro »)",
                 ctl_f["alerte"] and [t["balise"] for t in ctl_f["balises_revelatrices"]] == ["[FOURNISSEUR_PRO]"]
                 and "REPRENNENT LE NOM" in ctl_f["conseil"], str(ctl_f["balises_revelatrices"]))
        verifier("et le rapport écrit le dit",
                 "RÉVÈLENT" in pathlib.Path(coeur.ecrire_rapport(res_f["chemin_sortie"], ctl_f, "F", "F")).read_text(encoding="utf-8"))
        # Version anglaise : « hereinafter », « referred to as », premier article numerote
        cg_en = ("AGREEMENT\nBETWEEN\nAlpha Software Ltd (hereinafter referred to as the \"Provider\")\nAND\n"
                 "Beta Retail Ltd, hereinafter the \"Customer\".\n1. Definitions\nGamma Ltd is a subprocessor.\n")
        cfg_en = json.loads((SCRIPT_DIR / "_build" / "langue" / "en" / "detection_config.json").read_text(encoding="utf-8"))
        ents = det.detecter(cg_en)
        roles.attribuer_roles(cg_en, ents, cfg_en)
        verifier("version anglaise : « hereinafter » et « referred to as » donnent les rôles, le 1er article ferme l'en-tête",
                 trouver(ents, "Alpha Software").partie == "PROVIDER" and trouver(ents, "Beta Retail").partie == "CUSTOMER"
                 and trouver(ents, "Gamma Ltd").partie.startswith("COMPANY"),
                 str([(e.canonique, e.partie) for e in ents if e.type == "personne_morale"]))
    finally:
        print("\n" + "=" * 70)
        print(f"RESULTAT : {len(OK)} contrôle(s) réussi(s), {len(KO)} échec(s)")
        if KO:
            print("\nÉchecs à corriger :")
            for libelle, detail in KO:
                print(f"  - {libelle}" + (f"\n      {detail}" if detail else ""))
        else:
            print("\nTous les contrôles sont passés.")
        print("=" * 70 + "\n")
        shutil.rmtree(racine, ignore_errors=True)

    return 1 if KO else 0


if __name__ == "__main__":
    sys.exit(main())
