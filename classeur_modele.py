"""
classeur_modele.py
------------------
Cree le classeur de correspondances remis aux utilisateurs, et le document
d'exemple qui va avec.

Le classeur contient :
  - un onglet « Mode d'emploi » (ignore par le programme, lu par l'humain) ;
  - un onglet « Exemple » au modele par partie, rempli avec un dossier fictif,
    pour essayer le programme tout de suite.

Toutes les donnees sont inventees. Aucun nom reel.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import os

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

import modele_onglet as modele

LANGUE = "fr"                      # fixe par le programme principal

NOMS_ONGLETS = {"fr": ("Mode d'emploi", "Exemple", "Modèle complet"),
                "en": ("User guide", "Example", "Full template")}


def nom_onglet_modele():
    return NOMS_ONGLETS.get(LANGUE, NOMS_ONGLETS["fr"])[2]


def nom_onglet_mode_emploi():
    return NOMS_ONGLETS.get(LANGUE, NOMS_ONGLETS["fr"])[0]


def nom_onglet_exemple():
    return NOMS_ONGLETS.get(LANGUE, NOMS_ONGLETS["fr"])[1]


# Couleurs sobres (legal design) : une seule teinte, des gris.
BLEU_NUIT = "1B3A5C"
BLEU_PALE = "E8EEF4"
GRIS_TEXTE = "4A5561"
BLANC = "FFFFFF"

# ------------------------------------------------------------------
# DOSSIER FICTIF
# ------------------------------------------------------------------
# Partie DEM1 : une societe, son dirigeant, son avocat.
# Partie DEF1 : une societe et son adresse.
# Elements communs : la marque en litige et le numero de RG.

ENTREES_EXEMPLE = [
    # terme, balise, role, partie, notes
    ("Société Horizon Numérique", "[DEM1]", "Identité", "DEM1",
     "Dénomination complète ; les variantes se séparent par « / »"),
    ("Horizon Numérique / HORIZON NUMERIQUE", "[DEM1]", "Identité", "DEM1",
     "Formes abrégées du même nom : même balise"),
    ("12 rue des Lilas, 59000 Lille", "[DEM1_ADRESSE]", "Adresse", "DEM1", ""),
    ("812 345 678", "[DEM1_SIRET]", "SIRET", "DEM1", "SIREN ou SIRET"),
    ("Madame Claire Vasseur", "[DEM1_DIRIGEANT]", "Dirigeant", "DEM1",
     "Civilité comprise : « Madame Claire Vasseur » est remplacé en entier"),
    ("Claire Vasseur / VASSEUR", "[DEM1_DIRIGEANT]", "Dirigeant", "DEM1",
     "Formes courtes"),
    ("contact@horizon-numerique.example", "[DEM1_EMAIL]", "E-mail", "DEM1",
     "Les adresses e-mail et sites web se cachent aussi"),
    ("Maître Paul Lemaire", "[DEM1_AVOCAT]", "Avocat", "DEM1", ""),
    ("Société Atelier Bleu Marine", "[DEF1]", "Identité", "DEF1", ""),
    ("Atelier Bleu Marine / ATELIER BLEU MARINE", "[DEF1]", "Identité", "DEF1", ""),
    ("8 quai de la Citadelle, 59140 Dunkerque", "[DEF1_ADRESSE]", "Adresse", "DEF1", ""),
    ("LUMIVRAC", "[MARQUE1]", "Marque", "COMMUN", "Marque en litige"),
    ("24/01234", "[RG1]", "Référence RG", "COMMUN", "Numéro de rôle"),
]


ENTREES_EXEMPLE_EN = [
    ("Horizon Digital Ltd", "[CLM1]", "Identité", "CLM1",
     "Full name; spellings are separated by \" / \""),
    ("Horizon Digital / HORIZON DIGITAL", "[CLM1]", "Identité", "CLM1",
     "Short forms of the same name: same tag"),
    ("12 Lilac Street, Lille 59000", "[CLM1_ADDRESS]", "Adresse", "CLM1", ""),
    ("812 345 678", "[CLM1_COMPANY_NO]", "SIRET", "CLM1", "Registration number"),
    ("Ms Claire Vasseur", "[CLM1_DIRECTOR]", "Dirigeant", "CLM1",
     "Title included: \"Ms Claire Vasseur\" is replaced as a whole"),
    ("Claire Vasseur / VASSEUR", "[CLM1_DIRECTOR]", "Dirigeant", "CLM1", "Short forms"),
    ("contact@horizon-digital.example", "[CLM1_EMAIL]", "E-mail", "CLM1",
     "E-mail addresses and websites are hidden too"),
    ("Paul Lemaire", "[CLM1_LAWYER]", "Avocat", "CLM1", "Counsel"),
    ("Blue Marine Workshop Ltd", "[DEF1]", "Identité", "DEF1", ""),
    ("Blue Marine Workshop / BLUE MARINE WORKSHOP", "[DEF1]", "Identité", "DEF1", ""),
    ("8 Citadel Quay, Dunkirk 59140", "[DEF1_ADDRESS]", "Adresse", "DEF1", ""),
    ("LUMIVRAC", "[TRADEMARK1]", "Marque", "COMMUN", "Trademark in dispute"),
    ("24/01234", "[CASE_NO1]", "Référence RG", "COMMUN", "Case number"),
]

# Mots qui ne doivent plus apparaitre dans le document anonymise (recette)
MOTS_TEST = {"fr": ("horizon", "vasseur", "lemaire", "bleu marine", "lumivrac", "01234", "lilas"),
             "en": ("horizon", "vasseur", "lemaire", "blue marine", "lumivrac", "01234", "lilac")}
PREMIERE_FORME = {"fr": "Société Horizon Numérique", "en": "Horizon Digital Ltd"}


# Lignes du « Modele complet » : chaque partie avec tous ses roles, les
# acteurs communs, puis les elements communs. Terme vide : a remplir.
ROLES_SANS_PERSONNE = {"Étude", "Référence RG", "Montant", "Pièce", "Donnée", "Empreinte fichier",
                       "URL dépôt", "Clé de licence", "Version logiciel", "Dépôt INPI", "Dépôt européen"}
ROLES_OBJETS = {"Marque", "Brevet", "Dessin et modèle", "Dépôt INPI", "Dépôt européen",
                "Logiciel", "Produit", "Version logiciel", "Empreinte fichier", "URL dépôt",
                "Clé de licence", "Référence RG", "Montant", "Pièce", "Donnée"}
ACTEURS_COMMUNS = {"fr": [("COMMISSAIRE1", ["Identité", "Étude", "Adresse"]),
                          ("EXPERT1", ["Identité"]), ("MAGISTRAT1", ["Identité"]),
                          ("TIERS1", ["Identité", "Adresse"])],
                   "en": [("BAILIFF1", ["Identité", "Étude", "Adresse"]),
                          ("EXPERT1", ["Identité"]), ("JUDGE1", ["Identité"]),
                          ("THIRD_PARTY1", ["Identité", "Adresse"])]}
PARTIES_MODELE = {"fr": ["DEM1", "DEF1"], "en": ["CLM1", "DEF1"]}
INDICATIONS = {
    "fr": {"Identité": "Nom complet (société ou personne) ; les autres écritures séparées par « / »",
           "Adresse": "Adresse du siège ou du domicile", "SIRET": "SIREN ou SIRET",
           "Dirigeant": "Représentant légal (Madame / Monsieur + nom)",
           "Adresse dirigeant": "Adresse personnelle du dirigeant", "Signataire": "Signataire d'un acte",
           "Personne liée": "Associé, conjoint, contact…", "Salarié": "Salarié cité dans les pièces",
           "Enseigne": "Nom commercial ou enseigne", "E-mail": "Adresse e-mail", "Téléphone": "Numéro",
           "Site web": "Nom de domaine", "Avocat": "Maître + nom", "Cabinet d'avocat": "Structure d'exercice",
           "Adresse avocat": "Adresse du cabinet", "E-mail avocat": "", "Téléphone avocat": "",
           "Site web avocat": "", "Toque avocat": "Toque ou case du Palais", "Avocat postulant": "Maître + nom", "Adresse avocat postulant": "",
           "Étude": "Étude du commissaire de justice", "Marque": "Signe tel que déposé",
           "Brevet": "Numéro ou titre du brevet", "Dessin et modèle": "", "Dépôt INPI": "Numéro de dépôt",
           "Dépôt européen": "Numéro EUIPO / OEB", "Logiciel": "Nom du logiciel", "Produit": "Nom du produit ou de la gamme",
           "Version logiciel": "", "Empreinte fichier": "Hash d'une pièce", "URL dépôt": "Adresse d'un dépôt en ligne",
           "Clé de licence": "", "Référence RG": "Numéro RG ou de rôle", "Montant": "Montant à masquer",
           "Pièce": "Référence d'une pièce", "Donnée": "Toute autre donnée à masquer"},
    "en": {"Identité": "Full name (company or person); other spellings separated by \" / \"",
           "Adresse": "Registered or home address", "SIRET": "Company number",
           "Dirigeant": "Legal representative", "Adresse dirigeant": "Director's personal address",
           "Signataire": "Signatory of a deed", "Personne liée": "Partner, spouse, contact…",
           "Salarié": "Employee named in the documents", "Enseigne": "Trade name", "E-mail": "E-mail address",
           "Téléphone": "Phone number", "Site web": "Domain name", "Avocat": "Lawyer's name",
           "Cabinet d'avocat": "Law firm", "Adresse avocat": "Firm's address", "E-mail avocat": "",
           "Téléphone avocat": "", "Site web avocat": "", "Toque avocat": "Court box number", "Avocat postulant": "", "Adresse avocat postulant": "",
           "Étude": "Bailiff's office", "Marque": "Trademark as filed", "Brevet": "Patent number or title",
           "Dessin et modèle": "", "Dépôt INPI": "Filing number", "Dépôt européen": "EUIPO / EPO number",
           "Logiciel": "Software name", "Produit": "Product or range name", "Version logiciel": "",
           "Empreinte fichier": "Hash of an exhibit", "URL dépôt": "Online repository address",
           "Clé de licence": "", "Référence RG": "Case number", "Montant": "Amount to mask",
           "Pièce": "Exhibit reference", "Donnée": "Any other data to mask"},
}


def _entrees_modele(config):
    """Lignes du modele complet, balises calculees comme a l'injection."""
    lang = "en" if LANGUE == "en" else "fr"
    attributs = modele._attributs_complets(config)          # (libelle, suffixe, affichage)
    indications = INDICATIONS[lang]
    entrees, prises = [], set()

    def ligne(partie, libelle):
        balise, _ = modele.calculer_balise(partie, libelle, prises, config)
        prises.add(balise)
        entrees.append({"terme": "", "balise": balise, "role": libelle, "partie": partie,
                        "dossier": "", "actif": "YES" if lang == "en" else "OUI",
                        "notes": indications.get(libelle, "")})

    for partie in PARTIES_MODELE[lang]:
        for lib, sfx, _ in attributs:
            if lib in ROLES_SANS_PERSONNE:
                continue
            ligne(partie, lib)
    # Parties du menu (reglage « parties_menu » : SOCIETE,
    # PERSONNE_PHYSIQUE...) : le n° 1 de chacune, avec ses roles ; les
    # numeros suivants restent dans les menus deroulants.
    for r in modele.parties_menu(config):
        partie = f"{r['code']}1"
        for lib, sfx, _ in attributs:
            if lib in ROLES_SANS_PERSONNE or lib in modele.ROLES_DOSSIER:
                continue
            if r.get("personne_physique") and lib in modele.ROLES_PERSONNE_MORALE_SEULEMENT:
                continue
            ligne(partie, lib)
    for partie, libelles in ACTEURS_COMMUNS[lang]:
        for lib in libelles:
            ligne(partie, lib)
    for lib, sfx, _ in attributs:
        if lib in ROLES_OBJETS:
            ligne(modele.COMMUN, lib)
    return entrees


def _entrees():
    source = ENTREES_EXEMPLE_EN if LANGUE == "en" else ENTREES_EXEMPLE
    return [{"terme": t, "balise": b, "role": r, "partie": p, "dossier": "",
             "actif": "YES" if LANGUE == "en" else "OUI", "notes": n} for t, b, r, p, n in source]


# ------------------------------------------------------------------
# MODE D'EMPLOI (onglet)
# ------------------------------------------------------------------

LIGNES_MODE_EMPLOI = [
    ("TITRE", "Tableau de correspondances : mode d'emploi"),
    ("", "Ce classeur est votre clé. Il dit au programme quel vrai nom devient quelle balise. "
         "Il est confidentiel : ne le transmettez jamais avec un document anonymisé, "
         "et ne le déposez dans aucun outil d'intelligence artificielle."),
    ("", ""),
    ("SECTION", "Un onglet par dossier"),
    ("", "Chaque onglet correspond à un dossier (une affaire, un contrat, un audit). "
         "Créez-le depuis le programme (bouton « Nouveau dossier… ») : l'onglet arrive "
         "déjà au bon format. L'onglet « Exemple » est un dossier fictif pour essayer."),
    ("L'essentiel", "L'anonymisation ne se prépare pas document par document, mais UNE FOIS pour tout "
                    "le dossier. L'onglet est la liste des noms du dossier et de leurs balises : une "
                    "fois rempli et bien complété, il sert à toutes les pièces, passées et à venir. "
                    "Pour anonymiser une nouvelle pièce : choisir le document, choisir l'onglet, "
                    "cliquer. Inutile de refaire une détection. La détection n'est qu'une aide pour "
                    "remplir l'onglet au départ, puis l'enrichir : un nom nouveau = une ligne de plus, "
                    "et toutes les pièces en profitent."),
    ("", "Le programme ignore cet onglet « Mode d'emploi »."),
    ("", ""),
    ("SECTION", "Les sept colonnes"),
    ("Terme original", "Le vrai nom, écrit exactement comme dans vos documents. Plusieurs façons "
                       "d'écrire le même nom se séparent par «  /  » (espace, barre, espace)."),
    ("Balise", "Ce qui remplace le terme : entre crochets, en majuscules. Elle se construit "
               "toute seule : PARTIE + rôle, par exemple [DEF1_ADRESSE]. "
               "Ne modifiez jamais une balise déjà utilisée : la restitution en dépend."),
    ("Rôle", "Ce que le terme est pour la partie : Identité, Adresse, SIRET, Dirigeant, Avocat…"),
    ("Partie", "Qui est concerné, en code court numéroté : DEM1 (demandeur), DEF1 (défendeur), APP1, "
               "INTIME1, INTERV1, CLI (notre client), ADV (partie adverse), SOCIETE1, "
               "PERSONNE_PHYSIQUE1… ou COMMUN pour ce qui n'appartient à personne. Ce qui appartient "
               "à une partie s'y rattache : [DEF1_ADRESSE], [DEF1_AVOCAT]. Pour rattacher une ligne à "
               "une autre partie, changez simplement sa Partie : la balise suivra au prochain usage, et "
               "l'ancienne reste reconnue pour restituer les documents déjà anonymisés."),
    ("Dossier", "Facultatif : votre référence interne."),
    ("Actif ?", "OUI (ou vide) : le terme est remplacé. NON : la ligne est gardée pour mémoire "
                "mais ignorée."),
    ("Notes", "Libre."),
    ("", ""),
    ("SECTION", "Remplir la table à la main"),
    ("Étape 1", "Allez dans l'onglet du dossier (créé par le bouton « Nouveau dossier… » du programme)."),
    ("Étape 2", "Sur une des lignes préparées sous le tableau : le vrai nom en colonne A, puis "
                "le rôle et la partie dans les menus déroulants. La balise s'écrit toute seule, "
                "numéro compris, et la ligne prend la couleur de sa partie. Un rôle absent du "
                "menu se tape librement (la balise se forme avec)."),
    ("Chercher vite", "Dans Microsoft 365, tapez les premières lettres dans la case (sans accent : "
                      "« defe », « soc ») : le menu ne garde que les choix qui les contiennent ; "
                      "flèche bas puis Entrée pour valider. Les versions Excel 2016 à 2021 n'ont "
                      "pas ce filtre : déroulez la liste, ou passez par Rôle + Partie."),
    ("Étape 3", "Enregistrez (Ctrl+S). C'est tout : le programme relit ce classeur à chaque clic, "
                "inutile de le relancer. Si un nouvel onglet n'apparaît pas dans sa liste "
                "« Dossier », cliquez sur « Actualiser »."),
    ("Excel ouvert", "Vous pouvez laisser Excel ouvert pour anonymiser, contrôler ou restituer. "
                     "Fermez-le seulement quand le programme doit écrire ici (« Nouveau "
                     "dossier… », « Injecter »), car Excel bloque le fichier."),
    ("Modèle complet", "L'onglet « Modèle complet » contient toutes les balises prêtes (demandeur, "
                       "défendeur, personnes physiques et morales, sociétés, commissaire de justice, "
                       "expert, éléments communs), colonne Terme "
                       "vide. Pour un nouveau dossier à remplir à la main : clic droit sur l'onglet > "
                       "Déplacer ou copier > Créer une copie, renommez-la, remplissez les termes. Les "
                       "lignes laissées vides sont ignorées ; la détection remplit aussi ces lignes."),
    ("Astuce", "Le programme peut aussi remplir l'onglet pour vous : onglet « 1. Préparer la "
               "table », bouton « Détecter ». Vous relisez le classeur de détection dans Excel "
               "(Garder ?, partie, balise), vous l'enregistrez et le fermez, puis vous injectez. "
               "Les deux méthodes se mélangent sans problème."),
    ("", ""),
    ("SECTION", "Trois règles d'or"),
    ("1", "Un vrai nom = une seule balise, pour tout le dossier."),
    ("2", "Écrivez le terme le plus long d'abord (« Société Horizon Numérique »), puis les formes "
          "courtes sur une autre ligne avec la même balise."),
    ("3", "Relisez toujours le document produit. Le programme vérifie beaucoup, il ne lit pas "
          "à votre place."),
]


LIGNES_MODE_EMPLOI_EN = [
    ("TITRE", "Correspondence table: user guide"),
    ("", "This workbook is your key. It tells the program which real name becomes which tag. "
         "It is confidential: never send it with an anonymised document, and never put it "
         "into an AI tool."),
    ("", ""),
    ("SECTION", "One sheet per matter"),
    ("", "Each sheet is one matter (a case, a contract, an audit). Create it from the program "
         "(\"New matter…\" button): the sheet comes in the right format. The \"Example\" sheet "
         "is a fictional matter for trying the program out."),
    ("Key point", "Anonymisation is not prepared document by document, but ONCE for the whole matter. "
                  "The sheet is the list of the matter's names and their tags: once filled in and "
                  "completed, it serves every document, past and future. To anonymise a new document: "
                  "pick the document, pick the sheet, click. No need to run detection again. Detection "
                  "is only an aid to fill the sheet at the start, then to enrich it: a new name = one "
                  "more row, and every document benefits."),
    ("", "The program ignores this \"User guide\" sheet."),
    ("", ""),
    ("SECTION", "The seven columns"),
    ("Original term", "The real name, written exactly as in your documents. Several spellings of "
                      "the same name are separated by \"  /  \" (space, slash, space)."),
    ("Tag", "What replaces the term: in square brackets, in capitals. It is built automatically: "
            "PARTY + role, for instance [DEF1_ADDRESS]. Never change a tag already in "
            "use: restoration depends on it."),
    ("Role", "What the term is to the party: Identity, Address, Company number, Director, "
             "Lawyer…"),
    ("Party", "Who the term relates to, as a short numbered code: CLM1 (claimant), DEF1 (defendant), "
              "APP1, RESP1, INTERV1, CLI (our client), OPP (opposing party), COMPANY1, INDIVIDUAL1… or "
              "SHARED for items that belong to no party. What belongs to a party hangs off it: "
              "[DEF1_ADDRESS], [DEF1_LAWYER]. To attach a row to another party, simply change its "
              "Party: the tag follows at the next use, and the former one is still recognised to "
              "restore documents already anonymised."),
    ("Matter", "Optional: your internal reference."),
    ("Active?", "YES (or empty): the term is replaced. NO: the row is kept for the record but "
                "ignored."),
    ("Notes", "Free text."),
    ("", ""),
    ("SECTION", "Filling the table by hand"),
    ("Step 1", "Go to the matter sheet (created by the program's \"New matter…\" button)."),
    ("Step 2", "On one of the rows prepared below the table: the real name in column A, then "
               "the role and the party from the drop-down lists. The tag fills itself in, number "
               "included, and the row takes its party's colour. A role missing from the menu "
               "can be typed freely (the tag is built from it)."),
    ("Quick search", "In Microsoft 365, type the first letters in the cell (\"defe\", \"comp\"): the "
                     "menu keeps only the entries containing them; arrow down then Enter to confirm. "
                     "Excel 2016 to 2021 have no such filter: scroll the list, or use Role + Party."),
    ("Step 3", "Save (Ctrl+S). Nothing else is needed: the program re-reads this workbook at "
               "every click, so there is no need to restart it. If a new sheet does not show in its \"Matter\" list, "
               "click \"Refresh\"."),
    ("Excel open", "You may leave Excel open to anonymise, check or restore. Close it only when "
                   "the program must write here (\"New matter…\", \"Inject\"), because Excel "
                   "locks the file."),
    ("Full template", "The \"Full template\" sheet contains every tag, ready to use (claimant, defendant, "
                      "individuals and legal entities, companies, bailiff, expert, shared items) with an "
                      "empty Term column. For a new matter "
                      "filled by hand: right-click the sheet tab > Move or Copy > Create a copy, "
                      "rename it and fill in the terms. Rows left empty are ignored; detection also fills "
                      "these rows."),
    ("Tip", "The program can also fill the sheet for you: \"1. Prepare the table\" tab, "
            "\"Detect\" button. You review the detection workbook in Excel (Keep?, party, tag), "
            "save and close it, then inject. The two methods can be combined freely."),
    ("", ""),
    ("SECTION", "Three golden rules"),
    ("1", "One real name = one tag, for the whole matter."),
    ("2", "Write the longest form first (\"Horizon Digital Ltd\"), then the short forms on "
          "another row with the same tag."),
    ("3", "Always proofread the produced document. The program checks a lot; it does not read "
          "for you."),
]


def _ecrire_mode_emploi(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 95
    r = 2
    for genre, texte in (LIGNES_MODE_EMPLOI_EN if LANGUE == "en" else LIGNES_MODE_EMPLOI):
        if genre == "TITRE":
            c = ws.cell(row=r, column=2, value=texte)
            c.font = Font(bold=True, size=16, color=BLEU_NUIT)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            ws.row_dimensions[r].height = 30
        elif genre == "SECTION":
            c = ws.cell(row=r, column=2, value=texte.upper())
            c.font = Font(bold=True, size=11, color=BLANC)
            c.fill = PatternFill("solid", fgColor=BLEU_NUIT)
            ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor=BLEU_NUIT)
            ws.row_dimensions[r].height = 20
        elif genre == "" and texte == "":
            ws.row_dimensions[r].height = 8
        elif genre == "":
            c = ws.cell(row=r, column=2, value=texte)
            c.font = Font(size=10, color=GRIS_TEXTE)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            ws.row_dimensions[r].height = 15 * max(1, (len(texte) // 110) + 1)
        else:
            c1 = ws.cell(row=r, column=2, value=genre)
            c1.font = Font(bold=True, size=10, color=BLEU_NUIT)
            c1.alignment = Alignment(vertical="top")
            c1.fill = PatternFill("solid", fgColor=BLEU_PALE)
            c2 = ws.cell(row=r, column=3, value=texte)
            c2.font = Font(size=10, color=GRIS_TEXTE)
            c2.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * max(1, (len(texte) // 95) + 1)
        r += 1


def creer_classeur_modele(chemin, config_detection):
    """Cree le classeur complet. Refuse d'ecraser un fichier existant."""
    if os.path.exists(chemin):
        raise FileExistsError(chemin)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = nom_onglet_mode_emploi()
    _ecrire_mode_emploi(ws)
    wb.save(chemin)
    wb.close()

    modele.creer_onglet(chemin, nom_onglet_exemple(), config_detection)
    # La mise en forme (couleurs par partie, menus, lignes preformulees) est
    # celle de modele_onglet : rien a ajouter.
    modele.reecrire_onglet(chemin, nom_onglet_exemple(), _entrees(), config_detection)
    # Modele complet : toutes les balises pretes, colonne Terme a remplir
    modele.creer_onglet(chemin, nom_onglet_modele(), config_detection)
    titres = [[T_MODELE[LANGUE if LANGUE in T_MODELE else "fr"][0]] + [""] * 6,
              [T_MODELE[LANGUE if LANGUE in T_MODELE else "fr"][1]] + [""] * 6]
    modele.reecrire_onglet(chemin, nom_onglet_modele(), _entrees_modele(config_detection),
                           config_detection, titres=titres)
    return chemin


T_MODELE = {
    "fr": ("MODÈLE COMPLET — toutes les balises, à copier pour un nouveau dossier",
           "Clic droit sur l'onglet > Déplacer ou copier > Créer une copie ; renommez-la du nom du dossier, "
           "remplissez la colonne Terme original (les lignes vides sont ignorées). La détection remplit "
           "aussi ces lignes."),
    "en": ("FULL TEMPLATE — every tag, to copy for a new matter",
           "Right-click the sheet tab > Move or Copy > Create a copy; rename it after the matter, fill the "
           "Original term column (empty rows are ignored). Detection also fills these rows."),
}


# ------------------------------------------------------------------
# DOCUMENT D'EXEMPLE
# ------------------------------------------------------------------

PARAGRAPHES_EXEMPLE = [
    ("CONCLUSIONS EN DEMANDE", True),
    ("Tribunal judiciaire de Lille", False),
    ("N° RG : 24/01234", False),
    ("POUR :", True),
    ("La Société Horizon Numérique, SAS immatriculée au RCS de Lille sous le numéro "
     "812 345 678, dont le siège social est 12 rue des Lilas, 59000 Lille, prise en la "
     "personne de sa présidente, Madame Claire Vasseur.", False),
    ("Ayant pour avocat Maître Paul Lemaire, avocat au barreau de Lille.", False),
    ("DEMANDERESSE", True),
    ("CONTRE :", True),
    ("La Société Atelier Bleu Marine, SARL dont le siège social est 8 quai de la Citadelle, "
     "59140 Dunkerque.", False),
    ("DÉFENDERESSE", True),
    ("PLAISE AU TRIBUNAL", True),
    ("1. Les faits", True),
    ("HORIZON NUMERIQUE est titulaire de la marque française « LUMIVRAC », déposée pour "
     "des logiciels. Elle a constaté que ATELIER BLEU MARINE commercialise une application "
     "sous un signe identique.", False),
    ("Par courrier du 3 mars, Madame VASSEUR a mis en demeure la défenderesse de cesser "
     "tout usage du signe LUMIVRAC (contact@horizon-numerique.example). Ce courrier est "
     "resté sans réponse.", False),
    ("2. Discussion", True),
    ("La reproduction à l'identique de la marque LUMIVRAC pour des produits identiques "
     "constitue une contrefaçon au sens de l'article L. 713-2 du code de la propriété "
     "intellectuelle. Horizon Numérique est donc fondée à solliciter l'interdiction de "
     "l'usage du signe et la réparation de son préjudice.", False),
    ("PAR CES MOTIFS", True),
    ("Il est demandé au tribunal de faire droit aux demandes de la Société Horizon "
     "Numérique. Fait à Lille, le [DATE], pour Maître Paul Lemaire.", False),
]


PARAGRAPHES_EXEMPLE_EN = [
    ("PARTICULARS OF CLAIM", True),
    ("Case No: 24/01234", False),
    ("BETWEEN:", True),
    ("Horizon Digital Ltd, a company registered under number 812 345 678, whose registered "
     "office is at 12 Lilac Street, Lille 59000, represented by its director, Ms Claire "
     "Vasseur, and by its counsel, Paul Lemaire.", False),
    ("CLAIMANT", True),
    ("AND:", True),
    ("Blue Marine Workshop Ltd, whose registered office is at 8 Citadel Quay, Dunkirk 59140.",
     False),
    ("DEFENDANT", True),
    ("1. Facts", True),
    ("HORIZON DIGITAL is the owner of the registered trademark \"LUMIVRAC\", registered for "
     "software. It has come to its attention that BLUE MARINE WORKSHOP is marketing an "
     "application under an identical sign.", False),
    ("By letter dated 3 March, Ms VASSEUR demanded that the Defendant cease all use of the "
     "sign LUMIVRAC (contact@horizon-digital.example). The letter went unanswered.",
     False),
    ("2. Infringement", True),
    ("The use of a sign identical to the trademark LUMIVRAC in relation to identical goods "
     "constitutes infringement. Horizon Digital is therefore entitled to an injunction and "
     "to damages.", False),
    ("AND THE CLAIMANT CLAIMS:", True),
    ("(1) An injunction restraining the Defendant from infringing the trademark LUMIVRAC; "
     "(2) damages. Dated [DATE]. Paul Lemaire, counsel for Horizon Digital Ltd.", False),
]


def creer_document_exemple(chemin):
    """Document Word fictif qui correspond a l'onglet « Exemple »."""
    from docx import Document
    if os.path.exists(chemin):
        raise FileExistsError(chemin)
    doc = Document()
    for texte, titre in (PARAGRAPHES_EXEMPLE_EN if LANGUE == "en" else PARAGRAPHES_EXEMPLE):
        p = doc.add_paragraph()
        run = p.add_run(texte)
        if titre:
            run.bold = True
    doc.core_properties.author = "Paul Lemaire"
    doc.save(chemin)
    return chemin
