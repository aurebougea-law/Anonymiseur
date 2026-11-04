"""
anonymisation_core.py
---------------------
Coeur metier de l'anonymisation reversible de documents Word (.docx).

Ce module ne contient AUCUN code d'interface : il est importable et testable
sans tkinter et sans ecran. L'interface graphique vit dans
anonymiseur.py et n'appelle que les fonctions publiques ci-dessous.

Principes de securite tenus par ce module :
  1. Une restitution n'est jamais ecrite si la table choisie ne couvre pas
     TOUTES les balises presentes dans le document.
  2. Lorsque plusieurs tables couvrent le document, l'ambiguite est signalee
     avec les divergences terme par terme : l'outil refuse de trancher seul.
  3. Le journal porte son propre en-tete (table, onglet, horodatage)
     et son nom integre l'onglet, pour qu'un document homonyme d'un autre
     dossier ne puisse pas ecraser le journal d'un premier dossier.
  4. Un document deja anonymise ne peut pas etre anonymise une seconde fois.
  5. Le module n'AJOUTE jamais de metadonnee au document : l'identification de
     la table se fait par les balises presentes dans le texte.
     En revanche, depuis la version 4, l'anonymisation MASQUE les proprietes
     d'identite du fichier (auteur, societe) et neutralise les cibles de liens
     qui portent un nom. Les valeurs d'origine sont dans le journal et la
     restitution les remet en place. Ce comportement se desactive par la cle
     "masquer_proprietes" de config.json.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import os
import re
import json
import hashlib
import pathlib
import datetime
import unicodedata

import openpyxl
from docx import Document


# ==================================================================
# CONSTANTES
# ==================================================================

AUTEUR = "Aurélien Bougeard"
LANGUE = "fr"            # langue des messages destines a l'utilisateur ("fr" / "en")


def _t(fr, en):
    """Message dans la langue du programme. Les textes techniques (journaux,
    noms de fichiers) ne changent jamais : seul ce que lit l'utilisateur."""
    return en if LANGUE == "en" else fr
# Format 5 (octobre 2026) : les paragraphes sont lus directement dans le XML
# (hyperliens, controles de contenu, tableaux imbriques et suivi des
# modifications compris). Les journaux de format 4 et anterieurs restent
# restitues avec l'ancienne lecture, paragraphe par paragraphe.
FORMAT_JOURNAL = 5

# Une balise est un bloc en majuscules entre crochets : [PP_DEM_1], [SOCIETE_3]
MOTIF_BALISE = re.compile(r"\[[A-Z0-9ÉÈÀÇ_\-]{2,}\]")


# ==================================================================
# RECHERCHE D'UN TERME : UNE SEULE REGLE POUR TOUT L'OUTIL
# ==================================================================
#
# L'anonymisation, le controle et la verification de pertinence cherchent les
# termes de la table avec EXACTEMENT la meme regle, construite ici :
#
#   - frontieres de mot : « Aldebaran » ne touche pas « Aldebaranesque », et
#     une variante courte (« SA », « X ») ne remplace plus l'interieur des mots.
#     La frontiere n'est posee que du cote ou la variante commence ou finit
#     par une lettre ou un chiffre (une variante « (SA) » garde sa parenthese) ;
#     \w reconnait les lettres accentuees : « Élodie » est bien un mot ;
#   - espaces : un espace de la table accepte tout blanc du document, espace
#     insecable compris (« 59000 Lille » saisi au clavier retrouve
#     « 59000[insecable]Lille » mis en forme par Word) ;
#   - apostrophes : droite ou typographique, indifferemment ;
#   - casse indifferente, comme avant ;
#   - accents et ligatures indifferents (depuis le 7 octobre 2026, voir plus bas).
#
# Auparavant le remplacement comparait litteralement et sans frontiere, alors
# que le controle exigeait des frontieres et ignorait les variantes courtes :
# le controle ne voyait donc pas les degats de l'anonymisation.

APOSTROPHES = "'’‘ʼ"
LONGUEUR_MIN_CONSEILLEE = 3

# Accents et ligatures (7 octobre 2026). Une table « Societe Dupre » doit
# trouver « Société Dupré », et une table « Société Dupré » doit trouver
# « SOCIETE DUPRE » : les titres en capitales perdent souvent leurs accents.
# Sans cette regle, le nom passait en clair et le controle ne le voyait pas.
LIGATURES = {"œ": "oe", "æ": "ae"}
LIGATURE_DE = {"oe": "œ", "ae": "æ"}


def sans_accents(texte):
    """Texte sans accents ni ligatures (« Lœuvre Société » -> « Loeuvre Societe »)."""
    t = unicodedata.normalize("NFC", str(texte or ""))
    for lig, rempl in LIGATURES.items():
        t = t.replace(lig, rempl).replace(lig.upper(), rempl.upper())
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    return unicodedata.normalize("NFC", t)


def _classes_accents():
    """Pour chaque lettre de base, la classe de ses formes accentuees : e -> [eéèêë...]."""
    formes = {}
    for code in range(0x00C0, 0x0250):
        c = chr(code).lower()
        base = sans_accents(c)
        if len(base) == 1 and base.isascii() and base.isalpha() and base != c:
            formes.setdefault(base, set()).add(c)
    return {b: "[" + b + "".join(sorted(f)) + "]" for b, f in formes.items()}


CLASSES_ACCENTS = _classes_accents()


def normaliser(texte):
    """Forme de comparaison : sans accents ni ligatures, blancs reduits a un
    espace, apostrophes droites, minuscules. Sert de cle commune a la table et au document."""
    t = sans_accents(str(texte or ""))
    t = re.sub(f"[{APOSTROPHES}]", "'", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip().lower()


def _motif_mot(mot):
    """Un mot de la table, lettre par lettre, chaque lettre acceptant ses
    formes accentuees ; « oe » / « ae » acceptent aussi la ligature."""
    out, i = [], 0
    while i < len(mot):
        c = mot[i]
        paire = mot[i:i + 2].lower()
        if paire in ("oe", "ae"):
            lettres = "".join(_classe_lettre(x) for x in mot[i:i + 2])
            out.append(f"(?:{lettres}|{LIGATURE_DE[paire]})")
            i += 2
            continue
        if c.lower() in LIGATURES:
            lig = c.lower()
            out.append(f"(?:{lig}|{LIGATURES[lig]})")
        else:
            out.append(_classe_lettre(c))
        i += 1
    return "".join(out)


def _classe_lettre(c):
    base = sans_accents(c).lower()
    return CLASSES_ACCENTS.get(base, re.escape(c)) if len(base) == 1 else re.escape(c)


def corps_motif(variante, blancs=r"\s+", apostrophes=None):
    """Corps de l'expression d'une variante, sans les frontieres de mot.
    blancs / apostrophes : ce qu'acceptent un espace et une apostrophe de la
    table (par defaut tout blanc, et toute apostrophe)."""
    apostrophes = apostrophes or f"[{APOSTROPHES}]"
    v = unicodedata.normalize("NFC", str(variante or "").strip())
    morceaux = []
    for p in re.split(f"(\\s+|[{APOSTROPHES}])", v):
        if not p:
            continue
        if p.isspace():
            morceaux.append(blancs)
        elif p in APOSTROPHES:
            morceaux.append(apostrophes)
        else:
            morceaux.append(_motif_mot(p))
    return "".join(morceaux)


def motif_variante(variante):
    """Expression reguliere (texte) d'une variante de la table."""
    v = unicodedata.normalize("NFC", str(variante or "").strip())
    gauche = r"(?<!\w)" if re.match(r"\w", v) else ""
    droite = r"(?!\w)" if re.search(r"\w$", v) else ""
    return gauche + corps_motif(v) + droite


def compiler_variante(variante):
    return re.compile(motif_variante(variante), re.IGNORECASE)


def avertissements_entrees(entrees):
    """Ce qui merite d'etre signale dans une table, sans bloquer :
      - une variante de moins de 3 caracteres (« SA », « X ») : elle remplace
        chaque mot isole identique du document, ce qui est rarement voulu ;
      - une meme variante declaree sous deux balises : une seule s'applique ;
      - une meme balise portee par deux lignes aux termes differents : la
        restitution sans journal remet partout le terme de la premiere ligne
        (8 octobre 2026)."""
    messages = []
    premiers = {}
    for b, vs in entrees:
        if not vs:
            continue
        if b not in premiers:
            premiers[b] = vs[0]
        elif normaliser(vs[0]) != normaliser(premiers[b]):
            messages.append(_t(
                f"La balise {b} figure sur deux lignes aux termes différents (« {premiers[b]} » "
                f"et « {vs[0]} »). Une restitution sans journal remettra « {premiers[b]} » partout : "
                "regroupez les variantes sur une seule ligne ou donnez-leur deux balises.",
                f"Tag {b} appears on two rows with different terms (\"{premiers[b]}\" and "
                f"\"{vs[0]}\"). A restoration without log will put \"{premiers[b]}\" everywhere: "
                "group the variants on one row or give them two tags."))
    vues = {}
    plat = sorted([(v, b) for b, vs in entrees for v in vs], key=lambda x: len(x[0]), reverse=True)
    for v, b in plat:
        cle = normaliser(v)
        if cle in vues and vues[cle][0] != b:
            messages.append(_t(
                f"« {v} » figure sous {vues[cle][0]} et sous {b} : seule {vues[cle][0]} est appliquée.",
                f"\"{v}\" is listed under {vues[cle][0]} and under {b}: only {vues[cle][0]} is applied."))
        else:
            vues.setdefault(cle, (b, v))
    for b, vs in entrees:
        for v in vs:
            if len(v.strip()) < LONGUEUR_MIN_CONSEILLEE:
                messages.append(_t(
                    f"Variante très courte « {v} » ({b}) : chaque mot isolé « {v} » du document "
                    "est remplacé. Vérifiez le résultat ou retirez-la de la table.",
                    f"Very short variant \"{v}\" ({b}): every standalone word \"{v}\" of the "
                    "document is replaced. Check the result or remove it from the table."))
    return messages


# ==================================================================
# BALISES ALTEREES PAR UN OUTIL D'IA
# ==================================================================
#
# Un outil d'IA recopie souvent les balises imparfaitement : [Demandeur1],
# [DEMANDEUR 1], [DEMANDEUR1 ], DEMANDEUR1 sans crochets. Le motif strict ne
# les reconnait pas : elles n'etaient ni restituees ni signalees, et une
# balise mise en minuscules passait le controle « restitution complete ».
#
# Regle retenue : ces formes ne sont JAMAIS restituees en silence. Elles sont
# reperees a partir des balises connues (table ou journal) et LISTEES, a la
# restitution comme au controle, pour que l'utilisateur les corrige a la main.

def _jetons_balise(balise):
    corps = str(balise).strip().strip("[]")
    return [j for j in re.split(r"[\s_\-]+|(?<=[^\W\d_])(?=\d)|(?<=\d)(?=[^\W\d_])", corps) if j]


def _cle_alteree(texte):
    return re.sub(r"[\s_\-\[\]]", "", texte).upper()


def _motif_balises_alterees(balises):
    alternatives, index = [], {}
    for b in sorted(set(balises), key=len, reverse=True):
        jetons = _jetons_balise(b)
        if not jetons:
            continue
        # Separateurs tolérés : espace, insécable, tiret bas, tiret ; jamais un
        # retour à la ligne (deux paragraphes ne forment pas une balise).
        alternatives.append(r"[ \t\u00a0\u202f_\-]*".join(re.escape(j) for j in jetons))
        index.setdefault(_cle_alteree("".join(jetons)), b)
    if not alternatives:
        return None, index
    motif = re.compile(
        r"(?P<o>\[[^\S\n]*)?(?(o)|(?<![\w\[]))"
        r"(?P<c>" + "|".join(alternatives) + r")"
        r"(?P<f>[^\S\n]*\])?(?(f)|(?![\w\]]))",
        re.IGNORECASE)
    return motif, index


def balises_alterees_dans(texte, balises_connues, _cache=None):
    """[(forme_trouvee, balise_canonique, debut, fin)] : formes approchantes
    d'une balise connue qui ne sont pas la balise exacte."""
    motif, index = _cache or _motif_balises_alterees(balises_connues)
    if motif is None:
        return []
    sortie = []
    for m in motif.finditer(texte or ""):
        forme = m.group(0)
        balise = index.get(_cle_alteree(m.group("c")))
        if balise is None or forme == balise:
            continue
        if not m.group("o") and not m.group("f"):
            # Sans crochets, seule une forme d'un seul tenant, en capitales et
            # composee (DEMANDEUR1, SOC_DEM_1) est retenue : un titre
            # « DEFENDEUR 1 » ou un mot ordinaire n'est pas une balise.
            c = m.group("c")
            if (c != c.upper() or re.search(r"\s", c)
                    or len(_jetons_balise(balise)) < 2):
                continue
        sortie.append((forme, balise, m.start(), m.end()))
    return sortie


def balises_alterees_document(chemin_docx, balises_connues):
    """Formes alterees dans TOUTES les parties du document.
    [{"zone", "forme", "balise", "occurrences", "extrait"}]"""
    cache = _motif_balises_alterees(balises_connues)
    if cache[0] is None:
        return []
    trouvees = {}
    for zone, partie, texte in _textes_du_docx(chemin_docx):
        for forme, balise, debut, fin in balises_alterees_dans(texte, None, cache):
            cle = (zone, forme)
            if cle not in trouvees:
                d = max(0, debut - 45)
                extrait = re.sub(r"\s+", " ", texte[d:fin + 45]).strip()
                trouvees[cle] = {"zone": zone, "partie": partie, "forme": forme,
                                 "balise": balise, "occurrences": 0,
                                 "extrait": ("..." if d else "") + extrait + "..."}
            trouvees[cle]["occurrences"] += 1
    return sorted(trouvees.values(), key=lambda t: (t["zone"], t["balise"], t["forme"]))

VALEURS_ACTIF_VRAI = ("OUI", "YES", "1", "TRUE", "")

# « _Rôles » : feuille masquee propre a chaque onglet (menus deroulants et
# formule de balise), une par onglet depuis le 8 octobre 2026.
SUFFIXES_ONGLETS_MASQUES = ("_DataSrc", "_Rôles", "_Roles")
ONGLETS_MASQUES = ("Mode d'emploi", "Mode d’emploi", "User guide", "Rôles", "Roles",
                   "Modèle complet", "Full template")


class ErreurAnonymisation(Exception):
    """Erreur metier destinee a etre affichee telle quelle a l'utilisateur."""


class ErreurConcordance(ErreurAnonymisation):
    """Le journal ne correspond pas (ou plus) au document soumis.

    Porte le taux de concordance, pour que l'interface puisse distinguer un
    journal manifestement etranger au document (refus sec) d'un document
    seulement retouche depuis son anonymisation (confirmation possible).
    """

    def __init__(self, message, taux=0.0, details=""):
        super().__init__(message)
        self.taux = taux
        self.details = details


# ==================================================================
# CONFIGURATION
# ==================================================================

def _resoudre(chemin, dossier_base):
    """Un chemin relatif est resolu depuis le dossier du script, jamais depuis
    le repertoire courant : c'est ce qui rendait l'outil dependant de la facon
    dont il etait lance."""
    if not chemin:
        return ""
    p = pathlib.Path(str(chemin))
    if not p.is_absolute():
        p = pathlib.Path(dossier_base) / p
    return str(p)


def charger_config(chemin_config):
    """Lit config.json et retourne un dictionnaire normalise :
        {"tables": [{"nom":..., "fichier":...}], "dossier_journaux": ...}
    Accepte l'ancien format a cle unique "fichier_excel" / "dossier_json".
    """
    chemin_config = str(chemin_config)
    if not os.path.isfile(chemin_config):
        raise ErreurAnonymisation(f"config.json introuvable :\n{chemin_config}")

    with open(chemin_config, encoding="utf-8") as f:
        brut = json.load(f)

    dossier_base = os.path.dirname(os.path.abspath(chemin_config))

    tables = []
    for entree in brut.get("tables", []) or []:
        fichier = _resoudre(entree.get("fichier", ""), dossier_base)
        nom = str(entree.get("nom") or os.path.basename(fichier) or "Table")
        if fichier:
            tables.append({"nom": nom, "fichier": fichier})

    # Compatibilite ancien format
    if not tables and brut.get("fichier_excel"):
        tables.append({
            "nom": "Table principale",
            "fichier": _resoudre(brut["fichier_excel"], dossier_base),
        })

    if not tables:
        raise ErreurAnonymisation(
            "Aucune table de correspondances declaree dans config.json.\n\n"
            "Attendu : une liste \"tables\", chaque entree ayant un \"nom\" et un \"fichier\"."
        )

    dossier_journaux = brut.get("dossier_journaux") or brut.get("dossier_json") or "Journaux"
    dossier_journaux = _resoudre(dossier_journaux, dossier_base)

    return {
        "tables": tables,
        "dossier_journaux": dossier_journaux,
        "masquer_proprietes": bool(brut.get("masquer_proprietes", True)),
        "onglets_masques": tuple(brut.get("onglets_masques", ONGLETS_MASQUES)),
        "suffixes_onglets_masques": tuple(
            brut.get("suffixes_onglets_masques", SUFFIXES_ONGLETS_MASQUES)
        ),
    }


def verifier_tables(config):
    """Retourne (tables_ok, problemes). Permet a l'interface de signaler un
    fichier manquant AVANT que l'utilisateur ne lance un traitement."""
    ok, problemes = [], []
    for t in config["tables"]:
        if os.path.isfile(t["fichier"]):
            ok.append(t)
        else:
            problemes.append(f"« {t['nom']} » : fichier introuvable\n    {t['fichier']}")
    return ok, problemes


# ==================================================================
# LECTURE DES TABLES DE CORRESPONDANCES
# ==================================================================

def lister_onglets(chemin_excel, config=None):
    """Onglets exploitables d'un classeur, onglets techniques exclus."""
    if not os.path.isfile(chemin_excel):
        raise ErreurAnonymisation(f"Table de correspondances introuvable :\n{chemin_excel}")

    # Les onglets techniques du programme sont toujours exclus, meme si le
    # config.json de l'utilisateur ne les liste pas.
    masques = tuple((config or {}).get("onglets_masques", ())) + ONGLETS_MASQUES
    suffixes = tuple((config or {}).get("suffixes_onglets_masques", ())) + SUFFIXES_ONGLETS_MASQUES

    wb = openpyxl.load_workbook(chemin_excel, read_only=True)
    try:
        noms = [
            ws.title for ws in wb.worksheets
            if ws.title not in masques and not ws.title.endswith(suffixes)
        ]
    finally:
        wb.close()
    return noms


def _terme_en_texte(valeur):
    """Un SIRET stocke comme nombre doit etre relu comme texte."""
    if isinstance(valeur, float) and valeur == int(valeur):
        return str(int(valeur))
    return str(valeur).strip()


def decouper_variantes(terme):
    """« A / B » -> [A, B]. Le « / » ne separe deux formes que s'il touche un
    espace : un numero RG « 24/01234 » ou une date « 12/03/2025 » reste d'un
    seul tenant. Auparavant « 24/01234 » etait coupe en « 24 » et « 01234 »,
    et tout « 24 » du document etait remplace par la balise du RG."""
    return [v.strip() for v in re.split(r"\s+/\s*|\s*/\s+", str(terme or "")) if v.strip()]


def charger_entrees(chemin_excel, nom_onglet, anciennes=False):
    """Retourne [(balise, [variante, ...]), ...] pour les lignes actives.

    Colonnes attendues : A terme original, B balise, F actif.
    Les lignes de titre de section (terme ou balise vide) sont ignorees.

    Rattachement apres coup (15 octobre 2026) : si la Partie ou le Role d'une
    ligne a change depuis son ecriture (colonne H masquee), la ligne prend sa
    nouvelle balise. anciennes=True (restitution, controle) ajoute aussi les
    balises qu'elle portait avant (colonne I et ancienne colonne B), pour que
    les documents deja anonymises restent restituables.
    """
    if not os.path.isfile(chemin_excel):
        raise ErreurAnonymisation(f"Table de correspondances introuvable :\n{chemin_excel}")

    wb = openpyxl.load_workbook(chemin_excel, read_only=True, data_only=True)
    try:
        if nom_onglet not in wb.sheetnames:
            dispo = ", ".join(wb.sheetnames)
            raise ErreurAnonymisation(
                f"Onglet « {nom_onglet} » introuvable.\n\nOnglets du classeur : {dispo}"
            )

        ws = wb[nom_onglet]
        entrees = []
        etats = []                            # (indice, partie, role, origine, anciennes)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            terme, balise = row[0], row[1]
            actif = row[5] if len(row) > 5 else None
            if terme and not balise and len(row) > 3 and row[2] and row[3]:
                # Ligne tapee a la main dont la balise (formule Excel) n'a pas
                # encore de valeur enregistree : calculee ici, comme Excel le ferait.
                try:
                    import modele_onglet as _modele
                    balise, _ = _modele.calculer_balise(
                        str(row[3]).strip(), str(row[2]).strip(),
                        {b for b, _ in entrees}, _modele.config_courante())
                except Exception:
                    balise = None
            if not terme or not balise:
                continue

            terme_str = _terme_en_texte(terme)
            balise_str = str(balise).strip()
            if not terme_str or not balise_str or terme_str == balise_str:
                continue
            if balise_str.startswith("="):
                continue                      # formule sans valeur : traitee ci-dessous
            if actif is not None and str(actif).strip().upper() not in VALEURS_ACTIF_VRAI:
                continue

            variantes = decouper_variantes(terme_str)
            if variantes:
                entrees.append((balise_str, variantes))
                h = _terme_en_texte(row[7]) if len(row) > 7 and row[7] else ""
                if "|" in h:
                    i_col = _terme_en_texte(row[8]) if len(row) > 8 and row[8] else ""
                    j_col = _terme_en_texte(row[9]) if len(row) > 9 and row[9] else ""
                    if j_col.startswith("[") and j_col != balise_str:
                        i_col = f"{i_col}; {j_col}"      # balise changee a la main en B
                    etats.append((len(entrees) - 1, row[3], row[2], h, i_col))
        if etats:
            entrees = _rattachements(entrees, etats, anciennes)
        return entrees
    finally:
        wb.close()


def _rattachements(entrees, etats, anciennes):
    """Voir charger_entrees : nouvelle balise des lignes rattachees apres
    coup, et anciennes balises en plus si demande."""
    import modele_onglet as _modele
    config = _modele.config_courante()
    lignes = [{"balise": b, "terme": " / ".join(v), "partie": "", "role": "", "origine": "",
               "anciennes": []} for b, v in entrees]
    for i, partie, role, origine, vieilles in etats:
        lignes[i].update(partie=str(partie or "").strip(), role=str(role or "").strip(),
                         origine=origine,
                         anciennes=[b for b in re.split(r"[;\s]+", vieilles) if b.startswith("[")])
    _modele.appliquer_rattachements(lignes, config)
    sortie = [(l["balise"], v) for l, (_, v) in zip(lignes, entrees)]
    if anciennes:
        actuelles = {b for b, _ in sortie}
        for l, (_, v) in zip(lignes, entrees):
            for b in l["anciennes"]:
                if b not in actuelles:
                    sortie.append((b, v))
                    actuelles.add(b)
    return sortie


def balises_de(entrees):
    return {b for b, _ in entrees}


# ==================================================================
# LECTURE DU TEXTE D'UN DOCUMENT
# ==================================================================

def _paragraphes_du_document(doc):
    """LECTURE HISTORIQUE (journaux de format 4 et anterieurs uniquement).

    Paragraphes parcourus : corps, tableaux, en-tetes et pieds, dans l'ordre
    ou python-docx les expose, pour retrouver les para_id de ces anciens
    journaux. Les notes, commentaires et zones de texte sont traites a part
    (voir ZONES INVISIBLES, plus bas) : ce n'est plus une limite.
    """
    for para in doc.paragraphs:
        yield para
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            if hf is not None:
                for para in hf.paragraphs:
                    yield para


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P, W_T, W_DELTEXT, W_TXBX = _W + "p", _W + "t", _W + "delText", _W + "txbxContent"
W_TAB, W_BR, W_CR, W_DEL = _W + "tab", _W + "br", _W + "cr", _W + "del"

# Sauts de ligne (Maj+Entree), retours chariot et tabulations : des noeuds
# SANS texte dans le XML de Word. Ignores, les mots voisins se collent
# (« Dupre12 rue ») et le terme n'est plus reconnu (8 octobre 2026). Chaque
# separateur compte donc pour un caractere dans le texte compare.
SEPARATEURS = {W_TAB: "\t", W_BR: "\n", W_CR: "\n"}


class _Separateur:
    """Enveloppe d'un w:tab / w:br / w:cr, qui se presente comme un porteur
    de texte d'un caractere. Si un remplacement le recouvre (terme a cheval
    sur un saut de ligne), le noeud est retire ou remplace par un w:t."""

    def __init__(self, noeud):
        self.noeud = noeud
        self._texte = SEPARATEURS[noeud.tag]

    @property
    def text(self):
        return self._texte

    @text.setter
    def text(self, valeur):
        parent = self.noeud.getparent()
        if valeur == self._texte or parent is None:
            return
        if valeur:
            nouveau = etree.Element(W_T)
            nouveau.text = valeur
            nouveau.set(XML_SPACE, "preserve")
            parent.replace(self.noeud, nouveau)
            self.noeud = nouveau
        else:
            parent.remove(self.noeud)
        self._texte = valeur

    def set(self, cle, valeur):
        if self.noeud.tag == W_T:
            self.noeud.set(cle, valeur)


def _envelopper(noeud):
    return _Separateur(noeud) if noeud.tag in SEPARATEURS else noeud


def _est_supprime(noeud, p):
    """Le noeud est-il dans une suppression suivie (w:del) du paragraphe p ?"""
    for a in noeud.iterancestors():
        if a is p:
            return False
        if a.tag == W_DEL:
            return True
    return False


def _hors_zone_de_texte(p):
    return next(p.iterancestors(W_TXBX), None) is None


def _parties_entete_pied(doc):
    """En-tetes et pieds de page, chacun UNE fois (python-docx renvoie la meme
    partie pour chaque section qui la partage), dans un ordre stable."""
    vues = {}
    for rel in doc.part.rels.values():
        if rel.is_external:
            continue
        if rel.reltype.endswith("/header") or rel.reltype.endswith("/footer"):
            part = rel.target_part
            if hasattr(part, "element"):
                vues[str(part.partname)] = part
    return [vues[k] for k in sorted(vues)]


def _paragraphes_xml(doc):
    """Paragraphes parcourus depuis le format 5 : TOUS les w:p du corps (y
    compris dans les controles de contenu et les tableaux imbriques), puis
    ceux des en-tetes et pieds de page, tableaux compris.

    Les paragraphes des zones de texte sont exclus : la seconde passe
    (anonymiser_zones_invisibles) les traite deja."""
    racines = [doc.element.body] + [part.element for part in _parties_entete_pied(doc)]
    for racine in racines:
        for p in racine.iter(W_P):
            if _hors_zone_de_texte(p):
                yield p


def _porteurs(p, genre="t"):
    """Noeuds de texte propres au paragraphe p : w:t (texte affiche, liens,
    insertions suivies compris) ou w:delText (suppressions suivies).
    Le texte d'une zone de texte ancree dans p n'en fait pas partie."""
    balise = W_DELTEXT if genre == "suppr" else W_T
    suppr = genre == "suppr"
    sortie = []
    for n in p.iter(balise, *SEPARATEURS):
        if next(n.iterancestors(W_P), None) is not p:
            continue
        if n.tag in SEPARATEURS and _est_supprime(n, p) != suppr:
            continue
        sortie.append(_envelopper(n))
    return sortie


def _porteurs_xml(para):
    """Porteurs de texte d'un paragraphe des zones invisibles (passe 2) :
    w:t et separateurs, sans distinction de suivi des modifications."""
    return [_envelopper(n) for n in para.iter(W_T, *SEPARATEURS)]


def _texte_porteurs(noeuds):
    return "".join(n.text or "" for n in noeuds)


# ==================================================================
# FICHIERS TEXTE (.txt, .md) - 7 octobre 2026
# ==================================================================
#
# Cas d'usage : la reponse d'un outil d'IA copiee en texte, a restituer sans
# passer par Word. Un fichier texte n'a ni notes, ni commentaires, ni
# proprietes : chaque LIGNE y tient le role d'un paragraphe, et la regle de
# recherche est exactement celle du Word. Le fichier produit garde son
# encodage et ses fins de ligne.

EXTENSIONS_TEXTE = (".txt", ".md")


def est_texte(chemin):
    return os.path.splitext(str(chemin))[1].lower() in EXTENSIONS_TEXTE


def _lire_texte_brut(chemin):
    """(texte, encodage). Encodages reconnus : UTF-8 (avec ou sans BOM),
    UTF-16 avec BOM, puis Windows-1252 (ancien Bloc-notes)."""
    with open(chemin, "rb") as f:
        brut = f.read()
    if brut.startswith(b"\xef\xbb\xbf"):
        return brut[3:].decode("utf-8"), "utf-8-sig"
    if brut.startswith((b"\xff\xfe", b"\xfe\xff")):
        return brut.decode("utf-16"), "utf-16"
    try:
        return brut.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    try:
        return brut.decode("cp1252"), "cp1252"
    except UnicodeDecodeError:
        return brut.decode("latin-1"), "latin-1"


def _ecrire_texte_brut(chemin, texte, encodage):
    """Ecrit dans l'encodage d'origine ; si un nom restitue n'y existe pas,
    bascule en UTF-8 avec BOM (lisible partout) plutot que de l'abimer."""
    try:
        donnees = texte.encode(encodage)
        if encodage == "utf-8-sig" and not donnees.startswith(b"\xef\xbb\xbf"):
            donnees = b"\xef\xbb\xbf" + donnees
    except UnicodeEncodeError:
        donnees = b"\xef\xbb\xbf" + texte.encode("utf-8")
    with open(chemin, "wb") as f:
        f.write(donnees)


def _decouper_lignes(texte):
    """[contenu, fin, contenu, fin, ..., contenu] : les contenus (indices
    pairs) sont les paragraphes, les fins de ligne sont conservees telles
    quelles pour reconstituer le fichier a l'identique."""
    return re.split(r"(\r\n|\n|\r)", texte)


def _appliquer_ops_chaine(texte, ops):
    for debut, fin, remplacement in sorted(ops, key=lambda o: o[0], reverse=True):
        texte = texte[:debut] + remplacement + texte[fin:]
    return texte


def _paragraphes_texte(chemin):
    return _decouper_lignes(_lire_texte_brut(chemin)[0])[0::2]


def texte_du_document(chemin_docx):
    """Texte affiche du corps, des tableaux et des en-tetes / pieds."""
    if est_texte(chemin_docx):
        return _lire_texte_brut(chemin_docx)[0]
    doc = Document(chemin_docx)
    return "\n".join(_texte_porteurs(_porteurs(p)) for p in _paragraphes_xml(doc))


def balises_presentes(chemin_docx):
    """Balises du corps, des tableaux et des en-tetes (ce que parcourt
    python-docx). Pour les zones invisibles, voir balises_partout."""
    return set(MOTIF_BALISE.findall(texte_du_document(chemin_docx)))


def balises_partout(chemin_docx):
    """Balises presentes dans TOUTE l'archive, zones invisibles comprises."""
    trouvees = set(balises_presentes(chemin_docx))
    try:
        for _, _, texte in _textes_du_docx(chemin_docx):
            trouvees |= set(MOTIF_BALISE.findall(texte))
    except Exception:
        pass
    return trouvees


# ==================================================================
# NOMMAGE DES FICHIERS
# ==================================================================

# Suffixes des fichiers produits, selon la langue de la version. La lecture
# reconnait les deux langues : un document anonymise par la version francaise
# se restitue avec la version anglaise, et inversement.
SUFFIXES = {"fr": {"anonymise": "_ANONYMISE", "restitue": "_RESTITUE", "controle": "_Controle_"},
            "en": {"anonymise": "_ANONYMISED", "restitue": "_RESTORED", "controle": "_Check_"}}
# Le suffixe n'est reconnu qu'en fin de jeton : « Lettre_anonymisee.docx » ou
# « Note_restituee » sont des noms ordinaires, pas des fichiers produits par
# l'outil (8 octobre 2026).
MOTIF_SUFFIXES = re.compile(r"_(?:ANONYMISED|ANONYMISE|RESTORED|RESTITUE)(?![A-Za-z0-9])", re.IGNORECASE)
MOTIF_RESTITUE = re.compile(r"_(?:RESTORED|RESTITUE)(?![A-Za-z0-9])", re.IGNORECASE)
MOTIF_ANONYMISE = re.compile(r"_ANONYMISED?(?![A-Za-z0-9])", re.IGNORECASE)


DOSSIER_PRIVE = {"fr": "_Privé", "en": "_Private"}


def dossier_prive(chemin_docx):
    """Sous-dossier, a cote du document, qui recoit les fichiers de travail EN
    CLAIR (classeur de detection, rapport de controle). Ainsi un dossier zippe
    ou depose tel quel n'embarque pas, par megarde, la liste des vrais noms
    a cote du document anonymise : il suffit de ne pas joindre « _Privé »."""
    dossier = os.path.join(os.path.dirname(os.path.abspath(chemin_docx)),
                           DOSSIER_PRIVE.get(LANGUE, DOSSIER_PRIVE["fr"]))
    os.makedirs(dossier, exist_ok=True)
    return dossier


def suffixe(genre):
    """« anonymise », « restitue » ou « controle » : suffixe de la version courante."""
    return SUFFIXES.get(LANGUE, SUFFIXES["fr"])[genre]


def _base_propre(chemin_docx):
    base = os.path.splitext(os.path.basename(chemin_docx))[0]
    return MOTIF_SUFFIXES.sub("", base)


def _slug(texte):
    return re.sub(r"[^A-Za-z0-9À-ÿ_\-]+", "_", str(texte)).strip("_") or "sans_onglet"


def _empreinte_document(chemin_docx):
    """Empreinte courte et stable du nom du document (sans son contenu).
    Le nom du journal ne reprend ainsi jamais en clair le nom des parties
    qui figure souvent dans le nom du fichier source."""
    base = _base_propre(chemin_docx)
    return "DOC-" + hashlib.sha256(base.encode("utf-8")).hexdigest()[:12].upper()


def _empreinte_onglet(nom_onglet):
    """Empreinte courte du nom de l'onglet : l'onglet porte en general le nom
    du dossier ou du client, il ne doit pas non plus figurer en clair."""
    return "ONG-" + hashlib.sha256(_slug(nom_onglet).encode("utf-8")).hexdigest()[:8].upper()


def chemin_journal(config, chemin_docx, nom_onglet):
    """Nom du journal incluant l'onglet, pour qu'un document homonyme d'un
    autre dossier ne puisse pas ecraser un journal existant.
    Le document est designe par son empreinte, jamais par son nom."""
    dossier = config["dossier_journaux"]
    os.makedirs(dossier, exist_ok=True)
    return os.path.join(dossier, f"{_empreinte_document(chemin_docx)}__{_empreinte_onglet(nom_onglet)}_JOURNAL.json")


def _conserver_journal_existant(chemin_j):
    """Deux documents de meme nom de base (dans deux sous-dossiers, ou un
    document repris) produisent le meme nom de journal. Le journal precedent
    n'est plus ecrase : il est conserve sous un nom horodate, que la
    restitution avec journal continue de proposer. Retourne son nouveau nom."""
    if not os.path.isfile(chemin_j):
        return None
    horo = datetime.datetime.fromtimestamp(os.path.getmtime(chemin_j)).strftime("%Y%m%d-%H%M%S")
    racine = chemin_j[: -len("_JOURNAL.json")]
    cible = f"{racine}__{horo}_JOURNAL.json"
    n = 2
    while os.path.exists(cible):
        cible = f"{racine}__{horo}-{n}_JOURNAL.json"
        n += 1
    os.replace(chemin_j, cible)
    return cible


def journaux_candidats(config, chemin_docx):
    """Journaux utilisables pour ce document, le format historique inclus.
    Retourne [(chemin, meta_ou_None)], du plus recent au plus ancien."""
    dossier = config["dossier_journaux"]
    if not os.path.isdir(dossier):
        return []

    base = _base_propre(chemin_docx)
    empreinte = _empreinte_document(chemin_docx)
    trouves = []
    for nom in os.listdir(dossier):
        if not nom.endswith("_JOURNAL.json"):
            continue
        racine = nom[: -len("_JOURNAL.json")]
        # Nouveau nommage (empreinte), puis les deux formes historiques (nom en clair).
        if (racine.startswith(empreinte + "__")
                or racine == base or racine.startswith(base + "__")):
            chemin = os.path.join(dossier, nom)
            try:
                meta, _ = lire_journal(chemin)
            except Exception:
                meta = None
            trouves.append((chemin, meta))

    trouves.sort(key=lambda c: os.path.getmtime(c[0]), reverse=True)
    return trouves


# ==================================================================
# JOURNAL
# ==================================================================

def ecrire_journal(chemin, meta, remplacements, remplacements_xml=None,
                   proprietes=None, liens=None):
    with open(chemin, "w", encoding="utf-8") as f:
        json.dump(
            {"format": FORMAT_JOURNAL, "meta": meta,
             "remplacements": remplacements,
             "remplacements_xml": remplacements_xml or [],
             "proprietes": proprietes or [],
             "liens": liens or []},
            f, ensure_ascii=False, indent=2,
        )


def _charger_journal(chemin):
    """Contenu brut du journal (liste pour le format historique, dict sinon).
    Lu une seule fois : les lecteurs ci-dessous travaillent sur ce contenu."""
    if not os.path.isfile(chemin):
        raise ErreurAnonymisation(f"Journal introuvable :\n{chemin}")
    with open(chemin, encoding="utf-8") as f:
        brut = json.load(f)
    if not isinstance(brut, (list, dict)):
        raise ErreurAnonymisation(f"Journal illisible (format inattendu) :\n{chemin}")
    return brut


def _journal_brut(chemin_ou_brut):
    """Accepte un chemin ou un journal deja charge ; retourne le contenu brut,
    ou None si le journal est absent ou illisible."""
    if isinstance(chemin_ou_brut, (list, dict)):
        return chemin_ou_brut
    try:
        return _charger_journal(chemin_ou_brut)
    except Exception:
        return None


def lire_journal(chemin):
    """Retourne (meta, remplacements). Accepte le format historique, qui est
    une simple liste sans en-tete : meta vaut alors {} ."""
    brut = _charger_journal(chemin) if not isinstance(chemin, (list, dict)) else chemin
    if isinstance(brut, list):
        return {}, brut
    return brut.get("meta", {}) or {}, brut.get("remplacements", []) or []


# ==================================================================
# IDENTIFICATION DE LA TABLE D'ORIGINE
# ==================================================================

def analyser_couverture(chemin_docx, tables, config=None):
    """Pour chaque couple (table, onglet), mesure la couverture des balises du
    document. Retourne (balises_du_document, liste_analyses).

    Une analyse : {"table", "fichier", "onglet", "entrees", "couvertes",
                   "manquantes", "complet"}
    Un onglet "complet" couvre TOUTES les balises du document : lui seul est
    un candidat acceptable pour une restitution sans journal.

    Les balises sont relevees dans TOUT le document, zones invisibles comprises
    (notes, commentaires, zones de texte), et les marqueurs de redaction sont
    ecartes : memes regles que desanonymiser_sans_journal, sinon un onglet
    juge « complet » ici pouvait etre refuse a l'ecriture (8 octobre 2026).
    """
    cibles, _ = trier_balises(balises_partout(chemin_docx), config)
    analyses = []

    for t in tables:
        if not os.path.isfile(t["fichier"]):
            continue
        for onglet in lister_onglets(t["fichier"], config):
            try:
                entrees = charger_entrees(t["fichier"], onglet, anciennes=True)
            except ErreurAnonymisation:
                continue
            dispo = balises_de(entrees)
            if not dispo:
                continue
            couvertes = cibles & dispo
            manquantes = cibles - dispo
            analyses.append({
                "table": t["nom"],
                "fichier": t["fichier"],
                "onglet": onglet,
                "entrees": entrees,
                "couvertes": couvertes,
                "manquantes": manquantes,
                "complet": bool(cibles) and not manquantes,
            })

    analyses.sort(key=lambda a: (not a["complet"], len(a["manquantes"]), -len(a["couvertes"])))
    return cibles, analyses


def analyser_pertinence(chemin_docx, tables, config=None):
    """AVANT anonymisation : quel onglet correspond reellement a ce document ?

    Contrairement a la restitution, ce controle est fiable. Un document non
    encore anonymise contient les VRAIS termes, et ceux-ci sont propres a chaque
    dossier : un onglet etranger ne trouve rien. C'est donc a l'entree que la
    verification protege le mieux contre une erreur d'onglet.

    Retourne [{"table", "fichier", "onglet", "trouvees", "total"}], le meilleur
    en tete. « trouvees » compte les balises dont au moins une variante apparait
    dans le document.
    """
    texte = texte_du_document(chemin_docx)
    resultats = []

    for t in tables:
        if not os.path.isfile(t["fichier"]):
            continue
        for onglet in lister_onglets(t["fichier"], config):
            try:
                entrees = charger_entrees(t["fichier"], onglet)
            except ErreurAnonymisation:
                continue
            if not entrees:
                continue
            trouvees = 0
            for balise, variantes in entrees:
                for v in variantes:
                    if compiler_variante(v).search(texte):
                        trouvees += 1
                        break
            resultats.append({
                "table": t["nom"], "fichier": t["fichier"], "onglet": onglet,
                "trouvees": trouvees, "total": len(entrees),
            })

    resultats.sort(key=lambda r: r["trouvees"], reverse=True)
    return resultats


def divergences(analyses_completes, balises):
    """Pour des candidats concurrents, liste les balises dont la restitution
    differe d'un candidat a l'autre : c'est ce que l'utilisateur doit voir pour
    trancher sans se tromper de dossier.

    Retourne [(balise, {"Table / Onglet": terme_restitue})].
    """
    sortie = []
    for balise in sorted(balises):
        propositions = {}
        for a in analyses_completes:
            table_inv = {b: v[0] for b, v in a["entrees"] if v}
            if balise in table_inv:
                propositions[f"{a['table']} / {a['onglet']}"] = table_inv[balise]
        if len(set(propositions.values())) > 1:
            sortie.append((balise, propositions))
    return sortie


# ==================================================================
# ANONYMISATION
# ==================================================================

def _motif_variantes(plat_trie):
    """Un seul motif alternatif, variantes de la plus longue a la plus courte,
    plus un index variante -> balise.

    Le balayage se fait ainsi en UNE passe de gauche a droite. C'est essentiel :
    le journal enregistre alors les occurrences dans leur ordre d'apparition
    reelle, seule facon de restituer exactement un paragraphe ou plusieurs
    variantes d'une meme balise coexistent (« Societe X », puis « X »).
    Une passe unique empeche aussi qu'un remplacement deja pose soit re-analyse.

    Chaque variante suit la regle commune (motif_variante) : frontieres de
    mot, blancs et apostrophes tolerants. L'index est cle par forme normalisee.
    """
    index = {}
    for variante, balise in plat_trie:          # deja trie par longueur decroissante
        index.setdefault(normaliser(variante), balise)
    motif = re.compile("|".join(motif_variante(v) for v, _ in plat_trie), re.IGNORECASE)
    return motif, index


def anonymiser_document(chemin_docx, entrees, config, nom_table, nom_onglet):
    """Ecrit <nom>_ANONYMISE.docx et son journal. Retourne un compte rendu."""
    if not os.path.isfile(chemin_docx):
        raise ErreurAnonymisation(_t("Document introuvable :\n", "Document not found:\n") + chemin_docx)
    if not entrees:
        raise ErreurAnonymisation(_t(
            f"Aucune correspondance active dans l'onglet « {nom_onglet} ».\n\n"
            "Verifiez la colonne A (terme original) et la colonne F (actif).",
            f"No active correspondence in sheet \"{nom_onglet}\".\n\n"
            "Check column A (original term) and column F (active)."))

    # Garde-fou : ne jamais anonymiser deux fois le meme document.
    nom = os.path.basename(chemin_docx)
    deja = balises_partout(chemin_docx) & balises_de(entrees)
    if MOTIF_ANONYMISE.search(nom) or deja:
        raise ErreurAnonymisation(
            _t("Ce document semble DEJA anonymise.\n\n", "This document seems ALREADY anonymised.\n\n")
            + (_t("Balises deja presentes : ", "Tags already present: ") + ", ".join(sorted(deja)[:6])
               + (" ..." if len(deja) > 6 else "") + "\n\n" if deja else "")
            + _t("Une seconde anonymisation imbriquerait les balises et rendrait\n"
                 "la restitution impossible. Reprenez le document d'origine.",
                 "A second anonymisation would nest the tags and make restoration\n"
                 "impossible. Start again from the original document.")
        )

    plat_trie = sorted(
        [(v, balise) for balise, variantes in entrees for v in variantes],
        key=lambda x: len(x[0]), reverse=True,
    )
    motif, index = _motif_variantes(plat_trie)

    if est_texte(chemin_docx):
        return _anonymiser_texte(chemin_docx, entrees, config, nom_table, nom_onglet,
                                 motif, index)

    doc = Document(chemin_docx)
    remplacements = []
    total = 0

    # Chaque paragraphe est traite sur ses noeuds de texte (w:t), et non sur
    # ses « runs » : le texte des hyperliens, des insertions suivies et des
    # controles de contenu est ainsi atteint, et un appel de note ou une image
    # qui partage le fragment d'un terme remplace n'est plus efface.
    # Le texte supprime en suivi des modifications (w:delText) est traite a
    # part : il reste lisible dans le fichier et dans le volet des revisions.
    for para_id, p in enumerate(_paragraphes_xml(doc)):
        for genre in ("t", "suppr"):
            noeuds = _porteurs(p, genre)
            texte = _texte_porteurs(noeuds)
            if not texte.strip():
                continue
            locaux = []
            ops = _ops_anonymisation(texte, motif, index, para_id, locaux)
            if genre == "suppr":
                for e in locaux:
                    e["suppr"] = True
            remplacements.extend(locaux)
            total += _appliquer_ops(noeuds, ops, xml=True)

    dossier = os.path.dirname(os.path.abspath(chemin_docx))
    _, ext = os.path.splitext(chemin_docx)
    chemin_sortie = os.path.join(dossier, f"{_base_propre(chemin_docx)}{suffixe('anonymise')}{ext}")

    meta = {
        "document_source": os.path.basename(chemin_docx),
        "table": nom_table,
        "onglet": nom_onglet,
        "horodatage": datetime.datetime.now().isoformat(timespec="seconds"),
        "nb_remplacements": total,
        "ordre": "apparition",
        "outil": f"anonymisation_core (auteur : {AUTEUR})",
    }

    chemin_j = chemin_journal(config, chemin_docx, nom_onglet)
    doc.save(chemin_sortie)

    # Les passes suivantes reecrivent le fichier produit. Si l'une d'elles
    # echoue (antivirus, OneDrive, fichier verrouille), un _ANONYMISE a moitie
    # traite ne doit pas rester sur le disque sous un nom rassurant
    # (8 octobre 2026).
    try:
        # Seconde passe : notes de bas de page, notes de fin, commentaires et
        # zones de texte, que python-docx n'a pas vus.
        remplacements_xml = anonymiser_zones_invisibles(chemin_sortie, motif, index)
        total_invisible = len(remplacements_xml)
        total += total_invisible

        # Proprietes du fichier et cibles des liens : invisibles a la lecture,
        # et pourtant transmises avec le document.
        masquer = True if config is None else config.get("masquer_proprietes", True)
        proprietes = traiter_proprietes(chemin_sortie, motif, index,
                                        masquer_identites=masquer)
        if masquer:
            proprietes += traiter_auteurs(chemin_sortie)
        liens = traiter_liens(chemin_sortie, entrees)
    except Exception:
        _supprimer_sortie_partielle(chemin_sortie)
        raise

    meta["nb_remplacements"] = total
    meta["nb_zones_invisibles"] = total_invisible
    meta["nb_proprietes"] = len(proprietes)
    meta["nb_liens"] = len(liens)

    # Aucun remplacement nulle part : le fichier produit serait la copie
    # conforme de l'original. Le livrer sous le nom « _ANONYMISE » serait le
    # piege le plus grave de tout l'outil.
    # Seuls les remplacements issus de la TABLE prouvent que l'onglet est le bon.
    issus_table = [c for c in proprietes if c.get("origine") == "terme"]
    if total == 0 and not issus_table and not liens:
        try:
            os.remove(chemin_sortie)
        except OSError:
            pass
        raise _erreur_aucun_terme(nom_onglet)

    ancien_journal = _conserver_journal_existant(chemin_j)
    ecrire_journal(chemin_j, meta, remplacements, remplacements_xml,
                   proprietes, liens)
    return _compte_rendu_anonymisation(chemin_sortie, chemin_j, total, total_invisible,
                                       len(proprietes), len(liens), meta, entrees,
                                       ancien_journal)


def _verifier_sortie_distincte(chemin_entree, chemin_sortie):
    """Restituer « X_RESTITUE.docx » produirait... « X_RESTITUE.docx » : le
    fichier soumis serait ecrase par son propre resultat (8 octobre 2026)."""
    if os.path.abspath(chemin_entree) == os.path.abspath(chemin_sortie):
        raise ErreurAnonymisation(_t(
            "Le fichier produit porterait le meme nom que le fichier soumis :\n"
            f"{os.path.basename(chemin_sortie)}\n\nRenommez le document avant de recommencer.",
            "The output file would have the same name as the submitted file:\n"
            f"{os.path.basename(chemin_sortie)}\n\nRename the document and try again."))


def _supprimer_sortie_partielle(chemin_sortie):
    for c in (chemin_sortie, chemin_sortie + ".en_cours"):
        try:
            if os.path.isfile(c):
                os.remove(c)
        except OSError:
            pass


def _erreur_aucun_terme(nom_onglet):
    return ErreurAnonymisation(_t(
            f"AUCUN terme de l'onglet « {nom_onglet} » n'a ete trouve dans ce document,\n"
            "y compris dans les notes de bas de page et les zones de texte.\n\n"
            "Aucun fichier n'a ete cree : un document _ANONYMISE identique a l'original\n"
            "serait plus dangereux qu'utile.\n\n"
            "Verifiez que l'onglet selectionne est bien celui de ce dossier.",
            f"NO term of sheet \"{nom_onglet}\" was found in this document,\n"
            "footnotes and text boxes included.\n\n"
            "No file was created: an _ANONYMISED document identical to the original\n"
            "would be more dangerous than useful.\n\n"
            "Check that the selected sheet is the one of this matter."))


def _compte_rendu_anonymisation(chemin_sortie, chemin_j, total, total_invisible,
                                nb_proprietes, nb_liens, meta, entrees, ancien_journal):
    avertissements = avertissements_entrees(entrees)
    if ancien_journal:
        avertissements.append(_t(
            "Un journal existait déjà pour un document de même nom et le même onglet : "
            f"il a été conservé sous {os.path.basename(ancien_journal)}.",
            "A log already existed for a document with the same name and sheet: "
            f"it was kept as {os.path.basename(ancien_journal)}."))

    return {
        "chemin_sortie": chemin_sortie,
        "chemin_journal": chemin_j,
        "total": total,
        "total_zones_invisibles": total_invisible,
        "total_proprietes": nb_proprietes,
        "total_liens": nb_liens,
        "meta": meta,
        "avertissements": avertissements,
    }


def _anonymiser_texte(chemin, entrees, config, nom_table, nom_onglet, motif, index):
    """Anonymisation d'un .txt / .md : une ligne = un paragraphe du journal."""
    texte, encodage = _lire_texte_brut(chemin)
    morceaux = _decouper_lignes(texte)
    remplacements = []
    total = 0
    for i in range(0, len(morceaux), 2):
        ops = _ops_anonymisation(morceaux[i], motif, index, i // 2, remplacements)
        morceaux[i] = _appliquer_ops_chaine(morceaux[i], ops)
        total += len(ops)
    if total == 0:
        raise _erreur_aucun_terme(nom_onglet)

    dossier = os.path.dirname(os.path.abspath(chemin))
    _, ext = os.path.splitext(chemin)
    chemin_sortie = os.path.join(dossier, f"{_base_propre(chemin)}{suffixe('anonymise')}{ext}")
    _ecrire_texte_brut(chemin_sortie, "".join(morceaux), encodage)

    meta = {
        "document_source": os.path.basename(chemin),
        "table": nom_table,
        "onglet": nom_onglet,
        "horodatage": datetime.datetime.now().isoformat(timespec="seconds"),
        "nb_remplacements": total,
        "ordre": "apparition",
        "nature": "texte",
        "outil": f"anonymisation_core (auteur : {AUTEUR})",
    }
    chemin_j = chemin_journal(config, chemin, nom_onglet)
    ancien_journal = _conserver_journal_existant(chemin_j)
    ecrire_journal(chemin_j, meta, remplacements)
    return _compte_rendu_anonymisation(chemin_sortie, chemin_j, total, 0, 0, 0, meta,
                                       entrees, ancien_journal)


# ==================================================================
# RESTITUTION AVEC JOURNAL (exacte)
# ==================================================================

def _chemin_restitue(chemin):
    dossier = os.path.dirname(os.path.abspath(chemin))
    _, ext = os.path.splitext(chemin)
    return os.path.join(dossier, f"{_base_propre(chemin)}{suffixe('restitue')}{ext}")


def _par_paragraphe(remplacements):
    par_para = {}
    for e in remplacements:
        par_para.setdefault(e["para_id"], []).append(e)
    return par_para


def _format_journal(chemin):
    """Numero de format du journal (chemin ou contenu deja charge)."""
    brut = _journal_brut(chemin)
    if brut is None:
        return FORMAT_JOURNAL
    if isinstance(brut, dict):
        return int(brut.get("format", 1) or 1)
    return 1


def _textes_par_genre(p, format_journal):
    """[(genre, porteurs, texte)] d'un paragraphe, selon la lecture du journal."""
    if format_journal >= 5:
        sortie = []
        for genre in ("t", "suppr"):
            noeuds = _porteurs(p, genre)
            sortie.append((genre, noeuds, _texte_porteurs(noeuds), True))
        return sortie
    runs = p.runs
    return [("t", runs, "".join(r.text or "" for r in runs), False)]


def _paragraphes_selon_format(doc, format_journal):
    if format_journal >= 5:
        return list(_paragraphes_xml(doc))
    return list(_paragraphes_du_document(doc))


def mesurer_concordance(chemin_docx, remplacements, format_journal=FORMAT_JOURNAL):
    """Le journal decrit-il bien CE document ?

    Controle structurel : pour chaque paragraphe reference par le journal, les
    balises attendues doivent etre effectivement presentes, au bon nombre.

    Ce controle est indispensable : les onglets partagent le meme gabarit de
    balises, donc la seule presence de « [PP_DEM_1] » dans un document ne dit
    rien du dossier dont il vient. Seule la correspondance position par position
    distingue le bon journal d'un journal etranger.

    Retourne (taux, attendues, satisfaites, exemples_de_manques).
    """
    par_para = _par_paragraphe(remplacements)

    attendues = len(remplacements)
    satisfaites = 0
    manques = []

    if est_texte(chemin_docx):
        paragraphes = _paragraphes_texte(chemin_docx)
    else:
        paragraphes = _paragraphes_selon_format(Document(chemin_docx), format_journal)
    for para_id, entrees_para in par_para.items():
        if para_id >= len(paragraphes):
            manques.append(f"paragraphe {para_id} absent du document")
            continue
        if est_texte(chemin_docx):
            textes = {"t": paragraphes[para_id]}
        else:
            textes = {g: t for g, _, t, _ in _textes_par_genre(paragraphes[para_id],
                                                               format_journal)}
        besoins = {}
        for e in entrees_para:
            cle = ("suppr" if e.get("suppr") else "t", e["balise"])
            besoins[cle] = besoins.get(cle, 0) + 1
        for (genre, balise), attendu in besoins.items():
            trouve = textes.get(genre, "").count(balise)
            satisfaites += min(trouve, attendu)
            if trouve < attendu:
                manques.append(
                    f"paragraphe {para_id} : {balise} attendue {attendu} fois, trouvee {trouve}")

    # Un document dont seules les notes de bas de page portaient des termes n'a
    # aucun remplacement de corps : il n'y a alors rien a verifier ici.
    taux = (satisfaites / attendues) if attendues else 1.0
    return taux, attendues, satisfaites, manques[:6]


def desanonymiser_avec_journal(chemin_docx, chemin_j, tolerer_partiel=False):
    """Restitution exacte a partir du journal. Le journal fait foi : la table
    n'est pas relue, donc aucune confusion de dossier n'est possible.

    Deux garde-fous avant toute ecriture :
      - concordance structurelle du journal avec le document (voir ci-dessus) ;
      - restitution dans l'ordre d'apparition, et non dans l'ordre du fichier
        journal, pour ne pas permuter deux variantes d'une meme balise.
    """
    journal = _charger_journal(chemin_j)          # lu une seule fois
    meta, remplacements = lire_journal(journal)
    remplacements_xml = lire_journal_xml(journal)
    proprietes, liens = lire_journal_annexes(journal)
    if not remplacements and not remplacements_xml and not proprietes and not liens:
        raise ErreurAnonymisation(
            f"Le journal ne contient aucun remplacement :\n{chemin_j}\n\nRien a restituer.")

    format_j = _format_journal(journal)
    taux, attendues, satisfaites, manques = mesurer_concordance(chemin_docx, remplacements,
                                                                format_j)

    if taux < 1.0 and not tolerer_partiel:
        attendu_nom = meta.get("document_source", "(non renseigne)")
        detail = "\n".join("    - " + m for m in manques)
        message = _t(
            "REFUS : ce journal ne decrit pas ce document.\n\n"
            f"Concordance : {satisfaites} / {attendues} remplacement(s) retrouve(s) "
            f"({taux * 100:.0f} %).\n"
            f"Journal : {os.path.basename(chemin_j)}\n"
            f"Document d'origine attendu : {attendu_nom}\n"
            f"Document soumis : {os.path.basename(chemin_docx)}\n\n"
            f"{detail}\n\n"
            "Restituer avec un journal etranger melangerait deux dossiers.\n"
            "Aucune ecriture effectuee.",
            "REFUSED: this log does not describe this document.\n\n"
            f"Match: {satisfaites} / {attendues} replacement(s) found ({taux * 100:.0f}%).\n"
            f"Log: {os.path.basename(chemin_j)}\n"
            f"Expected original document: {attendu_nom}\n"
            f"Submitted document: {os.path.basename(chemin_docx)}\n\n"
            f"{detail}\n\n"
            "Restoring with a foreign log would mix two matters.\nNothing was written.")
        raise ErreurConcordance(message, taux=taux, details=detail)

    par_para = _par_paragraphe(remplacements)
    if est_texte(chemin_docx):
        texte, encodage = _lire_texte_brut(chemin_docx)
        morceaux = _decouper_lignes(texte)
        total = 0
        for i in range(0, len(morceaux), 2):
            du_para = [e for e in par_para.get(i // 2, []) if not e.get("suppr")]
            if du_para:
                ops = _ops_restitution(morceaux[i], du_para)
                morceaux[i] = _appliquer_ops_chaine(morceaux[i], ops)
                total += len(ops)
        chemin_sortie = _chemin_restitue(chemin_docx)
        _ecrire_texte_brut(chemin_sortie, "".join(morceaux), encodage)
        connues = {e["balise"] for e in remplacements}
        return {"chemin_sortie": chemin_sortie, "total": total, "meta": meta,
                "total_zones_invisibles": 0, "total_proprietes": 0, "total_liens": 0,
                "concordance": taux,
                "balises_alterees": balises_alterees_document(chemin_sortie, connues)}

    doc = Document(chemin_docx)
    total = 0

    for para_id, para in enumerate(_paragraphes_selon_format(doc, format_j)):
        entrees_para = par_para.get(para_id)
        if not entrees_para:
            continue
        for genre, porteurs, texte, xml in _textes_par_genre(para, format_j):
            du_genre = [e for e in entrees_para
                        if ("suppr" if e.get("suppr") else "t") == genre]
            if not du_genre:
                continue
            # File consommee de gauche a droite : la n-ieme occurrence d'une
            # balise dans le texte recoit la n-ieme variante enregistree.
            total += _appliquer_ops(porteurs, _ops_restitution(texte, du_genre), xml=xml)

    dossier = os.path.dirname(os.path.abspath(chemin_docx))
    _, ext = os.path.splitext(chemin_docx)
    chemin_sortie = os.path.join(dossier, f"{_base_propre(chemin_docx)}{suffixe('restitue')}{ext}")
    _verifier_sortie_distincte(chemin_docx, chemin_sortie)
    doc.save(chemin_sortie)

    try:
        total_invisible = restituer_zones_invisibles(chemin_sortie, remplacements_xml)
        total += total_invisible
        total_prop = restituer_proprietes(chemin_sortie, proprietes)
        total_liens = restituer_liens(chemin_sortie, liens)
    except Exception:
        _supprimer_sortie_partielle(chemin_sortie)
        raise

    connues = {e["balise"] for e in list(remplacements) + list(remplacements_xml)}
    return {"chemin_sortie": chemin_sortie, "total": total, "meta": meta,
            "total_zones_invisibles": total_invisible,
            "total_proprietes": total_prop, "total_liens": total_liens,
            "concordance": taux,
            # Formes alterees par l'IA : laissees en place et listees, jamais
            # restituees en silence.
            "balises_alterees": balises_alterees_document(chemin_sortie, connues)}


# ==================================================================
# RESTITUTION SANS JOURNAL (voie degradee, sous conditions)
# ==================================================================

def _premiere_variante_par_balise(entrees):
    """Balise -> premiere variante de la PREMIERE ligne qui la porte.

    Une balise peut occuper plusieurs lignes de la table (denomination
    complete, puis formes courtes) : la premiere ligne, la plus complete par
    convention, fait foi. Un dictionnaire construit en une passe retiendrait
    la derniere ligne, donc la forme courte."""
    table = {}
    for balise, variantes in entrees:
        if variantes and balise not in table:
            table[balise] = variantes[0]
    return table


def preparer_restitution_sans_journal(chemin_docx, tables, config=None):
    """Prepare la decision SANS rien ecrire. Retourne :
        {"balises": set, "candidats": [...], "autres": [...],
         "divergences": [...], "verdict": "aucune_balise" | "aucun_candidat"
                                          | "candidat_unique" | "ambigu"}
    """
    balises, analyses = analyser_couverture(chemin_docx, tables, config)
    candidats = [a for a in analyses if a["complet"]]
    autres = [a for a in analyses if not a["complet"]]

    if not balises:
        verdict = "aucune_balise"
    elif not candidats:
        verdict = "aucun_candidat"
    elif len(candidats) == 1:
        verdict = "candidat_unique"
    else:
        verdict = "ambigu"

    return {
        "balises": balises,
        "candidats": candidats,
        "autres": autres,
        "divergences": divergences(candidats, balises) if len(candidats) > 1 else [],
        "verdict": verdict,
    }


def desanonymiser_sans_journal(chemin_docx, entrees, confirme=False, config=None):
    """Restitue en relisant la table (premiere variante de chaque balise).

    Deux verrous, non contournables depuis l'interface :
      - refus si la table ne couvre pas toutes les balises du document ;
      - refus si l'utilisateur n'a pas confirme explicitement (confirme=True).
    """
    if not confirme:
        raise ErreurAnonymisation(
            "Restitution sans journal non confirmee : aucune ecriture effectuee.\n\n"
            "Cette voie relit la table de correspondances et non le journal :\n"
            "elle doit etre validee a l'ecran avant execution."
        )
    if not entrees:
        raise ErreurAnonymisation("La table selectionnee ne contient aucune correspondance active.")

    dispo = balises_de(entrees)
    presentes = balises_partout(chemin_docx)   # zones invisibles comprises

    # Un marqueur de redaction ([DATE], [MONTANT]...) n'est pas une balise :
    # aucune table ne pourrait le restituer, bloquer dessus n'a aucun sens.
    reelles, marqueurs = trier_balises(presentes, config)
    manquantes = reelles - dispo

    if manquantes:
        apercu = ", ".join(sorted(manquantes)[:8]) + (" ..." if len(manquantes) > 8 else "")
        raise ErreurAnonymisation(_t(
            "REFUS : la table selectionnee ne couvre pas tout le document.\n\n"
            f"{len(manquantes)} balise(s) du document sont absentes de cette table :\n"
            f"{apercu}\n\n"
            "Restituer malgre tout laisserait des balises non resolues, ou pire,\n"
            "melangerait deux dossiers. Aucune ecriture effectuee.",
            "REFUSED: the selected table does not cover the whole document.\n\n"
            f"{len(manquantes)} tag(s) of the document are missing from this table:\n"
            f"{apercu}\n\n"
            "Restoring anyway would leave unresolved tags or, worse, mix two\n"
            "matters. Nothing was written."))

    table_inverse = _premiere_variante_par_balise(entrees)
    balises_triees = sorted(table_inverse.items(), key=lambda x: len(x[0]), reverse=True)

    if est_texte(chemin_docx):
        texte, encodage = _lire_texte_brut(chemin_docx)
        morceaux = _decouper_lignes(texte)
        total = 0
        for i in range(0, len(morceaux), 2):
            ops = _ops_restitution_table(morceaux[i], balises_triees)
            morceaux[i] = _appliquer_ops_chaine(morceaux[i], ops)
            total += len(ops)
        chemin_sortie = _chemin_restitue(chemin_docx)
        _ecrire_texte_brut(chemin_sortie, "".join(morceaux), encodage)
        return {"chemin_sortie": chemin_sortie, "total": total,
                "total_zones_invisibles": 0,
                "marqueurs_laisses": sorted(marqueurs),
                "balises_alterees": balises_alterees_document(chemin_sortie, dispo)}

    doc = Document(chemin_docx)
    total = 0

    for p in _paragraphes_xml(doc):
        for genre in ("t", "suppr"):
            noeuds = _porteurs(p, genre)
            texte = _texte_porteurs(noeuds)
            if not texte.strip():
                continue
            total += _appliquer_ops(noeuds, _ops_restitution_table(texte, balises_triees),
                                    xml=True)

    dossier = os.path.dirname(os.path.abspath(chemin_docx))
    _, ext = os.path.splitext(chemin_docx)
    chemin_sortie = os.path.join(dossier, f"{_base_propre(chemin_docx)}{suffixe('restitue')}{ext}")
    _verifier_sortie_distincte(chemin_docx, chemin_sortie)
    doc.save(chemin_sortie)

    try:
        total_invisible = restituer_zones_invisibles_par_table(chemin_sortie, entrees)
    except Exception:
        _supprimer_sortie_partielle(chemin_sortie)
        raise
    total += total_invisible

    return {"chemin_sortie": chemin_sortie, "total": total,
            "total_zones_invisibles": total_invisible,
            "marqueurs_laisses": sorted(marqueurs),
            "balises_alterees": balises_alterees_document(chemin_sortie, dispo)}


# ==================================================================
# CONTROLE AVANT ENVOI
# ==================================================================
#
# L'anonymisation traite le corps, les tableaux, les en-tetes, puis les zones
# invisibles (notes, commentaires, zones de texte), les proprietes et les
# liens. Ce controle, lui, n'en sait rien : il ouvre le .docx comme l'archive
# qu'il est et relit TOUTES ses parties, y compris celles que python-docx
# n'expose pas, sans se fier a ce que l'anonymisation croit avoir fait.
#
# Il ne corrige rien : il constate. C'est volontaire. Un controle independant
# de l'outil qu'il verifie detecte aussi ce que cet outil n'a pas prevu.

import zipfile
import xml.etree.ElementTree as ET

ZONES_LISIBLES = {
    "word/document.xml": "corps du document",
    "word/footnotes.xml": "NOTES DE BAS DE PAGE",
    "word/endnotes.xml": "NOTES DE FIN",
    "word/comments.xml": "COMMENTAIRES",
    "word/commentsExtended.xml": "commentaires (complément)",
    "docProps/core.xml": "PROPRIÉTÉS DU DOCUMENT (auteur, titre)",
    "docProps/app.xml": "PROPRIÉTÉS DU DOCUMENT (société, modèle)",
    "word/_rels/document.xml.rels": "LIENS HYPERTEXTE",
    "word/settings.xml": "réglages du document",
    "docProps/custom.xml": "PROPRIÉTÉS PERSONNALISÉES",
}


ZONES_LISIBLES_EN = {
    "corps du document": "document body", "NOTES DE BAS DE PAGE": "FOOTNOTES",
    "NOTES DE FIN": "ENDNOTES", "COMMENTAIRES": "COMMENTS",
    "commentaires (complément)": "comments (extended)",
    "PROPRIÉTÉS DU DOCUMENT (auteur, titre)": "DOCUMENT PROPERTIES (author, title)",
    "PROPRIÉTÉS DU DOCUMENT (société, modèle)": "DOCUMENT PROPERTIES (company, template)",
    "LIENS HYPERTEXTE": "HYPERLINKS", "réglages du document": "document settings",
    "en-tête de page": "page header", "pied de page": "page footer",
    "auteurs des commentaires": "comment authors",
    "PROPRIÉTÉS PERSONNALISÉES": "CUSTOM PROPERTIES",
    "MODIFICATIONS SUIVIES (texte supprimé)": "TRACKED CHANGES (deleted text)",
    "AUTEURS (commentaires, révisions)": "AUTHORS (comments, revisions)",
    "CODES DE CHAMPS": "FIELD CODES",
    "ZONES DE TEXTE ET FORMES": "TEXT BOXES AND SHAPES",
}


def _zone_nommee(nom):
    return ZONES_LISIBLES_EN.get(nom, nom) if LANGUE == "en" else nom


# Attributs qui portent le nom d'une personne : auteur et initiales des
# commentaires, des insertions et des suppressions, identifiants de people.xml.
ATTRIBUTS_PERSONNES = ("author", "initials", "userId")


def _zone_lisible(partie):
    if partie in ZONES_LISIBLES:
        nom = ZONES_LISIBLES[partie]
    else:
        nom = os.path.basename(partie)
        if nom.startswith("header"):
            nom = "en-tête de page"
        elif nom.startswith("footer"):
            nom = "pied de page"
        elif nom.startswith("people"):
            nom = "auteurs des commentaires"
        else:
            return partie
    return ZONES_LISIBLES_EN.get(nom, nom) if LANGUE == "en" else nom


def _texte_et(elem, balise, dans_del=False):
    """Texte d'un element lu avec xml.etree : noeuds `balise` plus les
    separateurs (tabulation, saut de ligne), ceux-ci comptes seulement dans
    le suivi des suppressions (dans_del) ou hors de celui-ci."""
    morceaux = []

    def parcourir(el, supprime):
        for enfant in el:
            tag = enfant.tag
            if tag == balise:
                morceaux.append(enfant.text or "")
            elif tag in SEPARATEURS:
                if supprime == dans_del:
                    morceaux.append(SEPARATEURS[tag])
            parcourir(enfant, supprime or tag == W_DEL)

    parcourir(elem, False)
    return "".join(morceaux)


def _textes_du_docx(chemin_docx):
    """Retourne [(zone_lisible, partie, texte)] pour toute l'archive.

    Trois lectures complementaires :
      - le texte redige (noeuds w:t), partie par partie ;
      - le texte des zones de texte (w:txbxContent), isole pour etre localisable ;
      - les cibles des liens hypertexte et les proprietes du document.
    """
    if est_texte(chemin_docx):
        return [(_zone_lisible("word/document.xml"), "texte", _lire_texte_brut(chemin_docx)[0])]
    morceaux = []
    with zipfile.ZipFile(chemin_docx) as zf:
        for partie in zf.namelist():
            if not partie.endswith((".xml", ".rels")):
                continue
            try:
                brut = zf.read(partie)
            except Exception:
                continue

            # Liens hypertexte : la cible vit dans les relations, jamais dans le
            # texte. Un lien vers le site du client survit donc a l'anonymisation.
            if partie.endswith(".rels"):
                try:
                    racine = ET.fromstring(brut)
                except ET.ParseError:
                    continue
                cibles = [el.get("Target", "") for el in racine
                          if el.get("TargetMode") == "External"]
                if cibles:
                    morceaux.append((_zone_lisible(partie), partie, "\n".join(cibles)))
                continue

            try:
                racine = ET.fromstring(brut)
            except ET.ParseError:
                continue

            if partie.startswith("docProps/"):
                textes = [(el.text or "") for el in racine.iter() if el.text]
                if textes:
                    morceaux.append((_zone_lisible(partie), partie, "\n".join(textes)))
                continue

            # Texte redige : un paragraphe par ligne, pour un reperage utile.
            lignes, supprimees, champs = [], [], []
            for para in racine.iter(f"{_W}p"):
                texte = _texte_et(para, f"{_W}t")
                if texte.strip():
                    lignes.append(texte)
                texte = _texte_et(para, f"{_W}delText", dans_del=True)
                if texte.strip():
                    supprimees.append(texte)
                texte = "".join(t.text or "" for t in para.iter(f"{_W}instrText"))
                if texte.strip():
                    champs.append(texte)
            if lignes:
                morceaux.append((_zone_lisible(partie), partie, "\n".join(lignes)))
            # Suivi des modifications : le texte supprime reste dans le fichier.
            if supprimees:
                morceaux.append((_zone_nommee("MODIFICATIONS SUIVIES (texte supprimé)"),
                                 partie, "\n".join(supprimees)))
            # Codes de champs (lien « mailto: », reference) : invisibles a l'ecran.
            if champs:
                morceaux.append((_zone_nommee("CODES DE CHAMPS"), partie, "\n".join(champs)))
            # Auteurs des commentaires et des revisions, people.xml compris.
            personnes = []
            for el in racine.iter():
                for cle, valeur in el.attrib.items():
                    if cle.rsplit("}", 1)[-1] in ATTRIBUTS_PERSONNES and valeur.strip():
                        if valeur not in personnes:
                            personnes.append(valeur)
            if personnes:
                morceaux.append((_zone_nommee("AUTEURS (commentaires, révisions)"),
                                 partie, "\n".join(personnes)))

            # Zones de texte et formes, isolees pour pouvoir les nommer.
            zones = []
            for boite in racine.iter(f"{_W}txbxContent"):
                texte = _texte_et(boite, f"{_W}t")
                if texte.strip():
                    zones.append(texte)
            if zones:
                morceaux.append((_zone_nommee("ZONES DE TEXTE ET FORMES"), partie, "\n".join(zones)))

    return morceaux


def controler_document(chemin_docx, entrees):
    """Cherche, dans tout le document, ce qui aurait du etre pseudonymise.

    Retourne (trouvailles, zones_examinees) ; une trouvaille vaut
        {"zone", "partie", "terme", "balise", "occurrences", "extrait"}
    """
    if not os.path.isfile(chemin_docx):
        raise ErreurAnonymisation(f"Document introuvable :\n{chemin_docx}")
    if not entrees:
        raise ErreurAnonymisation(
            "La table sélectionnée ne contient aucune correspondance active :\n"
            "il n'y a rien à vérifier.")

    morceaux = _textes_du_docx(chemin_docx)
    trouvailles = []
    deja = set()

    for zone, partie, texte in morceaux:
        for balise, variantes in entrees:
            for variante in sorted(variantes, key=len, reverse=True):
                # Meme regle que l'anonymisation, variantes courtes comprises :
                # le controle voit ainsi exactement ce que l'anonymisation a fait.
                motif = compiler_variante(variante)
                trouves = list(motif.finditer(texte))
                if not trouves:
                    continue
                # Deux variantes qui ne different que par la casse ou les
                # accents (« Société X », « SOCIETE X ») reconnaissent les memes
                # occurrences : comptees une seule fois (8 octobre 2026).
                cle = (zone, balise, normaliser(variante))
                if cle in deja:
                    continue
                deja.add(cle)
                debut = max(0, trouves[0].start() - 45)
                extrait = re.sub(r"\s+", " ",
                                 texte[debut:trouves[0].end() + 45]).strip()
                trouvailles.append({
                    "zone": zone,
                    "partie": partie,
                    "terme": variante,
                    "balise": balise,
                    "occurrences": len(trouves),
                    "extrait": ("..." if debut else "") + extrait + "...",
                })

    # Les zones sensibles d'abord : ce sont celles que l'anonymisation ne voit pas.
    priorite = {"NOTES DE BAS DE PAGE": 0, "NOTES DE FIN": 1, "COMMENTAIRES": 2,
                "ZONES DE TEXTE ET FORMES": 3, "LIENS HYPERTEXTE": 4}
    trouvailles.sort(key=lambda t: (priorite.get(t["zone"], 9), t["zone"], t["balise"]))

    zones_examinees = sorted({z for z, _, _ in morceaux})
    return trouvailles, zones_examinees


# ==================================================================
# CONTROLE DANS LES DEUX SENS
# ==================================================================
#
# Un document peut pecher de deux facons opposees, et les deux sont graves :
#
#   - un document ANONYMISE ou l'un des vrais noms subsiste : le deposer dans
#     un outil d'IA, ou le transmettre, expose une donnee couverte par le
#     secret professionnel ;
#   - un document RESTITUE ou une balise subsiste : il part au client ou au
#     confrere avec « [PM_DEF_1] » au milieu d'une phrase, ce qui ne se voit
#     pas a la relecture rapide et se remarque tres bien a l'arrivee.
#
# Le meme bouton doit donc poser la bonne question selon l'etat du fichier.

ETAT_ANONYMISE = "anonymise"
ETAT_RESTITUE = "restitue"
ETAT_INDETERMINE = "indetermine"


def balises_restantes(chemin_docx):
    """Balises trouvees dans TOUTES les parties du document, pas seulement
    celles que l'anonymisation parcourt."""
    trouvees = {}
    for zone, partie, texte in _textes_du_docx(chemin_docx):
        for m in MOTIF_BALISE.finditer(texte):
            balise = m.group(0)
            cle = (zone, balise)
            if cle not in trouvees:
                debut = max(0, m.start() - 45)
                extrait = re.sub(r"\s+", " ", texte[debut:m.end() + 45]).strip()
                trouvees[cle] = {
                    "zone": zone, "partie": partie, "balise": balise,
                    "occurrences": 0,
                    "extrait": ("..." if debut else "") + extrait + "...",
                }
            trouvees[cle]["occurrences"] += 1
    return sorted(trouvees.values(), key=lambda t: (t["zone"], t["balise"]))


def etat_document(chemin_docx, config=None):
    """Dans quel etat est ce fichier ? Le nom le dit souvent, le contenu tranche.

    Retourne (etat, motif) ; le motif est affichable tel quel.
    """
    nom = os.path.basename(chemin_docx).upper()
    m = MOTIF_RESTITUE.search(nom)
    if m:
        return ETAT_RESTITUE, _t(f"le nom du fichier porte la mention {m.group(0)}",
                                 f"the file name contains {m.group(0)}")
    m = MOTIF_ANONYMISE.search(nom)
    if m:
        return ETAT_ANONYMISE, _t(f"le nom du fichier porte la mention {m.group(0)}",
                                  f"the file name contains {m.group(0)}")

    presentes = balises_presentes(chemin_docx)
    if presentes:
        return (ETAT_ANONYMISE,
                _t(f"{len(presentes)} balise(s) trouvée(s) dans le document",
                   f"{len(presentes)} tag(s) found in the document"))

    # Un document sans balise peut aussi bien etre un original qu'un restitue,
    # et l'existence d'un journal ne les distingue pas : le journal porte le nom
    # de base du document, donc il correspond aux deux.
    #
    # Se tromper dans ce sens serait le pire defaut de l'outil : declarer
    # « restitue, tout va bien » un document d'origine plein de vrais noms
    # autoriserait a le deposer dans une IA. En cas de doute, l'etat reste donc
    # INDETERMINE, et le controle alerte sur les termes en clair.
    return ETAT_INDETERMINE, _t("ni balise, ni mention _ANONYMISE ou _RESTITUE dans le nom",
                                "no tag, and neither _ANONYMISED nor _RESTORED in the name")


# ==================================================================
# DONNEES HORS TABLE (relecture du document anonymise par le detecteur)
# ==================================================================
#
# Le controle par la table ne voit que ce que la table contient. Ici, le
# detecteur de la pre-detection relit le document anonymise et signale ce qui
# ressemble encore a une donnee personnelle, sur deux niveaux :
#   - ALERTE : IBAN, numero de securite sociale, courriel, telephone. Motifs
#     precis, presque jamais de fausse alerte : le document ne doit pas partir.
#   - A VERIFIER : civilite + nom, Prenom NOM en capitales. Simple signalement,
#     qui ne change pas le verdict du controle.

TYPES_HORS_TABLE_ALERTE = ("iban", "nir", "email", "telephone")
FORMES_STRUCTURE_EXERCICE = ("SELARL", "SELAS", "SCP", "AARPI", "SELAFA", "SELEURL")

# Mots en capitales qui suivent souvent un mot capitalise sans etre un nom de
# famille (« Lille CEDEX », « Code CIVIL »...).
MAJUSCULES_NON_NOMS = {
    "SAS", "SASU", "SARL", "EURL", "SA", "SCI", "SNC", "SCP", "SELARL", "SELAS",
    "GIE", "LLC", "LTD", "INC", "PLC", "GMBH", "RCS", "TVA", "RG", "CPI", "HT",
    "TTC", "EUR", "INPI", "EUIPO", "OMPI", "WIPO", "CNIL", "RGPD", "GDPR", "SIRET",
    "SIREN", "CEDEX", "BP", "CS", "TGI", "TJ", "TC", "CA", "CJUE", "CEDH", "UE",
    "PDF", "URL", "API", "SAAS", "IA", "AI", "DPA", "CGV", "CGU", "NDA", "PI",
}


def _type_lisible_hors_table(type_):
    return {
        "iban": "IBAN", "nir": _t("n° de sécurité sociale", "social security no."),
        "email": _t("courriel", "e-mail"), "telephone": _t("téléphone", "phone"),
        "personne_physique": _t("civilité + nom", "title + name"),
        "representant_legal": _t("civilité + nom", "title + name"),
        "avocat": _t("Maître + nom", "Maître + name"),
        "commissaire_justice": _t("Maître + nom", "Maître + name"),
        "prenom_nom": _t("Prénom NOM", "First name SURNAME"),
        "structure": _t("structure d'exercice", "firm"),
    }.get(type_, type_)


def donnees_hors_table(chemin_docx, entrees=()):
    """Relit le document anonymise et signale les donnees qui ne sont pas des
    balises. Retourne (alertes, a_verifier) ; chaque element vaut
        {"zone", "type", "terme", "occurrences", "extrait"}.
    Un terme deja dans la table n'est pas repris ici : le controle par la
    table le signale deja."""
    import detection_core as det

    connus = set()
    for _, variantes in entrees or ():
        for v in variantes:
            if len(v.strip()) > 2:
                connus.add(normaliser(v))

    def deja_dans_la_table(terme):
        n = normaliser(terme)
        return n in connus or any(c in n for c in connus if len(c) > 3)

    exclus = {sans_accents(m).lower() for m in det.MOTS_COURANTS_CAPITALISES}
    exclus |= {"le", "la", "les", "un", "une", "du", "de", "des", "au", "aux", "en", "et",
               "il", "elle", "son", "sa", "ses", "leur", "page", "tel", "fax", "mail",
               "rue", "avenue", "boulevard", "place", "quai", "chemin", "route"}
    motif_prenom_nom = re.compile(
        r"(?<![\w#\[])([A-ZÀ-Ý][a-zà-ÿ]+(?:-[A-ZÀ-Ý][a-zà-ÿ]+)?)[ \u00a0]+"
        r"([A-ZÀ-Ý]{2,}(?:-[A-ZÀ-Ý]{2,})*)(?![\w#\]])"
        # suivi d'un autre mot en capitales : raison sociale ou produit
        # (« Amazon WEB SERVICES »), pas un Prenom NOM
        r"(?![  ]+[A-ZÀ-Ý]{2,})")

    trouves = {}

    def noter(niveau, zone, type_, terme, texte, position):
        terme = terme.strip()
        if len(terme) < 3 or deja_dans_la_table(terme):
            return
        cle = (niveau, normaliser(terme))
        if cle in trouves:
            trouves[cle]["occurrences"] += 1
            return
        debut = max(0, position - 45)
        extrait = re.sub(r"\s+", " ", texte[debut:position + len(terme) + 45]).strip()
        trouves[cle] = {"zone": zone, "type": _type_lisible_hors_table(type_),
                        "terme": terme, "occurrences": 1,
                        "extrait": ("..." if debut else "") + extrait + "..."}

    # Auteurs des commentaires et des revisions : masques par l'anonymisation
    # (option masquer_proprietes, 7 octobre 2026). S'ils sont encore la, c'est
    # que l'option est desactivee, un choix et non un oubli. Zone ignoree.
    zone_auteurs = _zone_nommee("AUTEURS (commentaires, révisions)")

    for zone, partie, texte in _textes_du_docx(chemin_docx):
        if zone == zone_auteurs:
            continue
        # Les balises sont neutralisees (meme longueur) : « Madame [X] » n'est
        # pas un nom, et les positions restent celles du texte.
        neutre = MOTIF_BALISE.sub(lambda m: "#" * len(m.group(0)), texte)

        for e in det._detecter_identifiants(neutre):
            if e.type in TYPES_HORS_TABLE_ALERTE:
                noter("alerte", zone, e.type, e.canonique, texte, e.position)

        for e in det._detecter_professionnels(neutre) + det._detecter_personnes(neutre):
            type_ = e.type
            if sans_accents(e.canonique).upper().split(" ")[0] in FORMES_STRUCTURE_EXERCICE:
                type_ = "structure"
            noter("verifier", zone, type_, e.canonique, texte, e.position)

        for m in motif_prenom_nom.finditer(neutre):
            prenom, nom = m.group(1), m.group(2)
            if (sans_accents(prenom).lower() in exclus
                    or sans_accents(nom).upper() in MAJUSCULES_NON_NOMS):
                continue
            # deja signale avec sa civilite : pas de doublon
            if any(t["terme"].endswith(f"{prenom} {nom}") for (n, _), t in trouves.items()
                   if n == "verifier"):
                continue
            noter("verifier", zone, "prenom_nom", f"{prenom} {nom}", texte, m.start())

    alertes = [t for (n, _), t in trouves.items() if n == "alerte"]
    a_verifier = [t for (n, _), t in trouves.items() if n == "verifier"]
    return alertes, a_verifier


def balises_revelatrices(entrees):
    """[{"balise", "mots", "terme"}] : balises qui contiennent un mot du nom
    qu'elles masquent ([AGICE] pour « AGICE EDITIQUE »). Une telle balise
    livre le nom a l'IA malgre l'anonymisation (19 octobre 2026)."""
    import modele_onglet as _modele
    sortie = []
    for balise, variantes in entrees:
        mots = _modele.mots_revelateurs(balise, variantes)
        if mots:
            sortie.append({"balise": balise, "mots": sorted(mots), "terme": " / ".join(variantes)})
    return sortie


def controler(chemin_docx, entrees, config=None):
    """Contrôle complet, dans les deux sens.

    Retourne un compte rendu :
        {"etat", "motif_etat", "termes_en_clair", "balises_restantes",
         "zones", "alerte", "titre", "conseil"}
    """
    etat, motif = etat_document(chemin_docx, config)
    revelatrices = balises_revelatrices(entrees)
    termes, zones = controler_document(chemin_docx, entrees)
    termes += controler_chemin(chemin_docx, entrees)     # nom du fichier et dossiers
    balises = balises_restantes(chemin_docx)

    # Les marqueurs de redaction sont mis de cote : leur presence dans un
    # document restitue est normale et ne vaut pas restitution incomplete.
    _, marqueurs = trier_balises({b["balise"] for b in balises}, config)
    balises_reelles = [b for b in balises if b["balise"] not in marqueurs]
    balises_marqueurs = [b for b in balises if b["balise"] in marqueurs]
    balises = balises_reelles

    # Balises recopiees de travers par un outil d'IA ([Demandeur1], DEMANDEUR 1...).
    alterees = balises_alterees_document(chemin_docx,
                                         balises_de(entrees) | balises_declarees(config))

    priorite = {"nom du fichier": -2, "nom d'un dossier": -1,
                "NOTES DE BAS DE PAGE": 0, "NOTES DE FIN": 1, "COMMENTAIRES": 2,
                "ZONES DE TEXTE ET FORMES": 3, "LIENS HYPERTEXTE": 4}
    termes.sort(key=lambda t: (priorite.get(t["zone"], 9), t["zone"], t["balise"]))
    zones = sorted(set(zones) | {t["zone"] for t in termes
                                 if t["zone"].startswith("nom ")})

    hors_alertes, hors_a_verifier, hors_erreur = [], [], ""
    if etat == ETAT_RESTITUE:
        alerte = bool(balises) or bool(alterees)
        titre = (_t("RESTITUTION INCOMPLÈTE", "INCOMPLETE RESTORATION") if alerte
                 else _t("Document restitué, restitution complète",
                         "Restored document, restoration complete"))
        conseil = (
            (_t("Des balises subsistent dans ce document destiné au client ou au "
                "confrère.\nNe l'envoyez pas : reprenez la restitution.",
                "Tags remain in this document intended for the client or a colleague.\n"
                "Do not send it: run the restoration again.")
             if balises else
             _t("Des balises ALTÉRÉES (recopiées de travers par l'outil d'IA) subsistent :\n"
                "elles ne sont jamais restituées automatiquement. Corrigez-les à la main\n"
                "dans le document, en vous aidant de la liste ci-dessous.",
                "ALTERED tags (miscopied by the AI tool) remain: they are never restored\n"
                "automatically. Fix them by hand in the document, using the list below."))
            if alerte else
            _t("Aucune balise ne subsiste : le document peut partir au client ou "
               "au confrère.\n\nEn revanche il contient les vrais noms, "
               "par construction : ne le déposez\ndans AUCUN outil d'IA.",
               "No tag remains: the document can go to the client or a colleague.\n\n"
               "It does contain the real names, by construction: never put it\n"
               "into ANY AI tool."))
    elif etat == ETAT_ANONYMISE:
        # Donnees hors table : seul le niveau ALERTE change le verdict.
        # Si cette relecture echoue, le controle par la table reste valable,
        # mais le rapport le dit au lieu de conclure « ni IBAN, ni courriel »
        # sur une verification qui n'a pas eu lieu (8 octobre 2026).
        try:
            hors_alertes, hors_a_verifier = donnees_hors_table(chemin_docx, entrees)
        except Exception as e:
            hors_alertes, hors_a_verifier = [], []
            hors_erreur = f"{type(e).__name__} : {e}"
        alerte = bool(termes) or bool(hors_alertes) or bool(revelatrices)
        titre = (_t("NE PAS TRANSMETTRE NI DÉPOSER DANS UNE IA",
                    "DO NOT SEND OR PUT INTO AN AI") if alerte
                 else _t("Document anonymisé, aucun terme en clair",
                         "Anonymised document, no unmasked term"))
        if termes:
            conseil = _t("Des termes de la table apparaissent encore en clair.\n"
                         "Ne déposez ce document dans aucun outil d'IA et ne le transmettez pas.",
                         "Terms from the table still appear unmasked.\n"
                         "Do not put this document into any AI tool and do not send it.")
        elif hors_alertes:
            conseil = _t("Des données absentes de la table apparaissent en clair\n"
                         "(IBAN, n° de sécurité sociale, courriel ou téléphone).\n"
                         "Ajoutez-les à la table, puis relancez l'anonymisation.\n"
                         "D'ici là, ne déposez ce document dans aucun outil d'IA.",
                         "Data missing from the table appear unmasked\n"
                         "(IBAN, social security no., e-mail or phone).\n"
                         "Add them to the table, then run the anonymisation again.\n"
                         "Until then, do not put this document into any AI tool.")
        elif hors_erreur:
            conseil = _t("Aucun terme actif de cette table n'apparaît dans le document.\n"
                         "En revanche la recherche des données hors table (IBAN, n° de\n"
                         "sécurité sociale, courriel, téléphone) n'a pas pu être faite :\n"
                         f"{hors_erreur}\n"
                         "Relisez le document sur ce point avant de le déposer dans un outil d'IA.",
                         "No active term of this table appears in the document.\n"
                         "However the search for data not in the table (IBAN, social\n"
                         "security no., e-mail, phone) could not be carried out:\n"
                         f"{hors_erreur}\n"
                         "Check the document on that point before putting it into an AI tool.")
        else:
            conseil = _t("Aucun terme actif de cette table n'apparaît dans le document,\n"
                         "ni IBAN, n° de sécurité sociale, courriel ou téléphone.\n"
                         "Il peut être déposé dans un outil d'IA.",
                         "No active term of this table appears in the document,\n"
                         "nor any IBAN, social security no., e-mail or phone.\n"
                         "It can be put into an AI tool.")
        if hors_a_verifier and not termes:
            conseil += _t("\n\nQuelques noms restent à vérifier (liste ci-dessous) : simple\n"
                          "signalement, à vous de juger s'ils doivent être masqués.",
                          "\n\nA few names are worth checking (list below): just a flag,\n"
                          "it is up to you whether they should be masked.")
    else:
        alerte = bool(termes)
        titre = (_t("Document non traité : termes en clair", "Unprocessed document: unmasked terms")
                 if alerte else _t("État du document indéterminé", "Document status unclear"))
        conseil = (
            _t("Impossible d'établir si ce document est un original ou une version\n"
               "restituée : il ne porte ni balise, ni mention dans son nom.\n\n",
               "Cannot tell whether this document is an original or a restored\n"
               "version: it has no tag and no marker in its name.\n\n")
            + (_t("Par prudence, les termes trouvés en clair sont traités comme une\n"
                  "alerte. Ne déposez ce document dans aucun outil d'IA tant que vous\n"
                  "n'avez pas vérifié de quoi il s'agit.",
                  "As a precaution, the unmasked terms found are treated as an alert.\n"
                  "Do not put this document into any AI tool until you have checked\n"
                  "what it is.")
               if alerte else
               _t("Aucun terme de cette table n'y figure, et aucune balise non plus.",
                  "No term of this table appears in it, and no tag either.")))

    if revelatrices and etat == ETAT_ANONYMISE:
        conseil = (_t("Des balises de la table REPRENNENT LE NOM qu'elles masquent "
                      "(par exemple [AGICE] pour « AGICE EDITIQUE ») : le document les "
                      "montre en clair à l'IA.\nChangez la Partie de ces lignes dans "
                      "l'onglet (SOCIETE, PRESTATAIRE, CLIENT…), relancez l'anonymisation.\n"
                      "Ne déposez pas ce document dans une IA.",
                      "Some tags in the table REVEAL THE NAME they mask (for instance "
                      "[AGICE] for \"AGICE EDITIQUE\"): the document shows them to the AI.\n"
                      "Change the Party of those rows in the sheet (COMPANY, PROVIDER, "
                      "CUSTOMER…), run the anonymisation again.\nDo not put this document "
                      "into an AI.")
                   + "\n\n" + conseil)
    return {
        "etat": etat,
        "motif_etat": motif,
        "termes_en_clair": termes,
        "balises_restantes": balises,
        "marqueurs_redaction": balises_marqueurs,
        "balises_alterees": alterees,
        "hors_table_alertes": hors_alertes,
        "hors_table_a_verifier": hors_a_verifier,
        "hors_table_erreur": hors_erreur,
        "balises_revelatrices": revelatrices,
        "zones": zones,
        "alerte": alerte,
        "titre": titre,
        "conseil": conseil,
    }


def ecrire_rapport(chemin_docx, resultat, table, onglet):
    """Rapport texte a cote du document, dans les deux sens."""
    dossier = dossier_prive(chemin_docx)
    # Nom COMPLET du document controle (suffixe _ANONYMISE / _RESTITUE compris) :
    # le rapport du restitue n'ecrase plus celui de l'anonymise (8 octobre 2026).
    chemin = os.path.join(dossier, f"{suffixe('controle')}{os.path.splitext(os.path.basename(chemin_docx))[0]}.txt")
    horodatage = datetime.datetime.now().strftime("%d/%m/%Y à %H:%M")

    etats = ({ETAT_ANONYMISE: "ANONYMISED", ETAT_RESTITUE: "RESTORED",
              ETAT_INDETERMINE: "UNCLEAR"} if LANGUE == "en" else
             {ETAT_ANONYMISE: "ANONYMISÉ", ETAT_RESTITUE: "RESTITUÉ",
              ETAT_INDETERMINE: "INDÉTERMINÉ"})

    l = [
        _t("CONTRÔLE AVANT TRANSMISSION OU DÉPÔT DANS UN OUTIL D'IA",
           "CHECK BEFORE SENDING OR PUTTING INTO AN AI TOOL"),
        "=" * 70,
        f"Document : {os.path.basename(chemin_docx)}",
        f"Table    : {table} / {onglet}",
        _t("Contrôle : ", "Checked: ") + horodatage,
        _t("État     : ", "Status: ") + f"{etats.get(resultat['etat'], resultat['etat'])} "
        f"({resultat['motif_etat']})",
        "",
        _t("Zones examinées :", "Areas examined:"),
    ]
    l += [f"    - {z}" for z in resultat["zones"]]
    l += ["", "-" * 70, "", resultat["titre"], "", resultat["conseil"], ""]

    if resultat["termes_en_clair"]:
        total = sum(t["occurrences"] for t in resultat["termes_en_clair"])
        l += ["", "-" * 70,
              _t("TERMES EN CLAIR : ", "UNMASKED TERMS: ")
              + f"{len(resultat['termes_en_clair'])} terme(s), {total} occurrence(s)", ""]
        zone = None
        for t in resultat["termes_en_clair"]:
            if t["zone"] != zone:
                zone = t["zone"]
                l += ["", f"  {zone}", "  " + "-" * len(zone)]
            l += [f"    {t['balise']}  ({t['occurrences']} occurrence(s))",
                  _t("        terme   : ", "        term    : ") + t['terme'],
                  _t("        extrait : ", "        excerpt : ") + t['extrait']]

    for cle, titre_fr, titre_en in (
            ("hors_table_alertes", "ALERTE, DONNÉES HORS TABLE EN CLAIR : ",
             "ALERT, DATA NOT IN THE TABLE, UNMASKED: "),
            ("hors_table_a_verifier", "À VÉRIFIER (non bloquant), NOMS HORS TABLE : ",
             "TO CHECK (not blocking), NAMES NOT IN THE TABLE: ")):
        liste = resultat.get(cle) or []
        if not liste:
            continue
        l += ["", "-" * 70, _t(titre_fr, titre_en) + str(len(liste)), ""]
        for t in liste:
            l += [f"    {t['type']} : {t['terme']}  ({t['occurrences']} occurrence(s), {t['zone']})",
                  _t("        extrait : ", "        excerpt : ") + t['extrait']]

    if resultat.get("balises_revelatrices"):
        l += ["", "-" * 70, _t("ALERTE, BALISES QUI RÉVÈLENT LE NOM MASQUÉ : ",
                                "ALERT, TAGS THAT REVEAL THE MASKED NAME: ")
              + str(len(resultat["balises_revelatrices"])), ""]
        for t in resultat["balises_revelatrices"]:
            l += [f"    {t['balise']}  ({', '.join(t['mots'])})",
                  _t("        terme   : ", "        term    : ") + t["terme"]]

    if resultat["balises_restantes"]:
        total = sum(t["occurrences"] for t in resultat["balises_restantes"])
        l += ["", "-" * 70,
              _t("BALISES PRÉSENTES : ", "TAGS PRESENT: ")
              + f"{len(resultat['balises_restantes'])} tag(s), {total} occurrence(s)"]
        if resultat["etat"] == ETAT_ANONYMISE:
            l += [_t("(normal dans un document anonymisé)", "(normal in an anonymised document)")]
        l += [""]
        zone = None
        for t in resultat["balises_restantes"]:
            if t["zone"] != zone:
                zone = t["zone"]
                l += ["", f"  {zone}", "  " + "-" * len(zone)]
            l += [f"    {t['balise']}  ({t['occurrences']} occurrence(s))",
                  _t("        extrait : ", "        excerpt : ") + t['extrait']]

    if resultat.get("balises_alterees"):
        total = sum(t["occurrences"] for t in resultat["balises_alterees"])
        l += ["", "-" * 70,
              _t("BALISES ALTÉRÉES (jamais restituées automatiquement) : ",
                 "ALTERED TAGS (never restored automatically): ")
              + f"{len(resultat['balises_alterees'])} forme(s), {total} occurrence(s)", ""]
        for t in resultat["balises_alterees"]:
            l += [f"    {t['forme']}  ->  {t['balise']}  ({t['occurrences']} occurrence(s), {t['zone']})",
                  _t("        extrait : ", "        excerpt : ") + t['extrait']]

    if (not resultat["termes_en_clair"] and not resultat["balises_restantes"]
            and not resultat.get("balises_alterees")):
        l += ["", _t("Ni terme en clair, ni balise : le document ne contient aucune",
                     "No unmasked term and no tag: the document contains no trace"),
              _t("trace de la table de correspondances.", "of the correspondence table.")]

    with open(chemin, "w", encoding="utf-8") as f:
        f.write("\n".join(l) + "\n")
    return chemin


# ==================================================================
# A1 - LE NOM DU FICHIER ET DU DOSSIER
# ==================================================================
#
# On envoie un fichier, pas seulement son contenu. Un document parfaitement
# anonymise a l'interieur mais nomme « Conclusions Societe X.docx », ou range
# dans un dossier au nom du client, livre l'information dans la piece jointe
# elle-meme, avant meme d'etre ouvert.

# Dans un nom de fichier, les mots sont separes par des tirets bas, des tirets
# ou des points : la bordure de mot habituelle (\b) ne convient pas, car le
# tiret bas compte comme une lettre pour elle.
def _motif_dans_nom(terme):
    # Meme regle d'accents que dans le texte : un nom de fichier est souvent
    # ecrit sans accents (« Conclusions Societe Dupre.docx »). Un espace de la
    # table accepte aussi tiret bas, tiret ou point (« Societe_Dupre »), et
    # une apostrophe peut avoir disparu (« LOrbe », « L_Orbe »).
    corps = corps_motif(terme, blancs=r"[\s_.\-]+",
                        apostrophes=rf"[{APOSTROPHES}\s_.\-]?")
    return re.compile(rf"(?<![A-Za-zÀ-ÿ0-9]){corps}(?![A-Za-zÀ-ÿ0-9])",
                      re.IGNORECASE)


def controler_chemin(chemin_docx, entrees, profondeur_dossiers=4):
    """Cherche les termes de la table dans le nom du fichier et dans les noms
    des dossiers qui le contiennent.

    profondeur_dossiers borne la remontee : au-dela, on atteint des dossiers
    systeme (Utilisateurs, Documents) sans rapport avec le dossier traite.
    """
    chemin = os.path.abspath(chemin_docx)
    nom_fichier = os.path.basename(chemin)

    segments = []
    courant = os.path.dirname(chemin)
    for _ in range(profondeur_dossiers):
        nom = os.path.basename(courant)
        parent = os.path.dirname(courant)
        if not nom or parent == courant:
            break
        segments.append(nom)
        courant = parent

    cibles = [("nom du fichier", nom_fichier)]
    cibles += [("nom d'un dossier", s) for s in segments]

    trouvailles = []
    deja = set()
    for zone, texte in cibles:
        for balise, variantes in entrees:
            for variante in sorted(variantes, key=len, reverse=True):
                if len(variante) < 3:
                    continue
                if not _motif_dans_nom(variante).search(texte):
                    continue
                cle = (texte, balise)
                if cle in deja:
                    continue
                deja.add(cle)
                trouvailles.append({
                    "zone": zone,
                    "partie": texte,
                    "terme": variante,
                    "balise": balise,
                    "occurrences": len(_motif_dans_nom(variante).findall(texte)),
                    "extrait": texte,
                })
    return trouvailles



# ==================================================================
# B1 - REMPLACER SANS DETRUIRE LA MISE EN FORME
# ==================================================================
#
# Un paragraphe Word n'est pas un texte : c'est une suite de fragments
# (les « runs »), chacun avec sa propre mise en forme. « Le contrat de
# SOCIETE X est resilie » peut compter trois fragments, dont un seul en gras.
#
# L'ancienne methode ecrasait tout le paragraphe dans le premier fragment et
# vidait les autres : le texte etait juste, mais le gras, l'italique et le
# souligne places A L'INTERIEUR du paragraphe disparaissaient. Le document
# restitue revenait donc au client appauvri.
#
# La methode ci-dessous calcule les positions a remplacer dans le texte
# complet, puis n'ecrit que dans les fragments reellement concernes. Un terme
# entierement contenu dans un fragment garde exactement sa mise en forme.
#
# LIMITE ASSUMEE : un terme a cheval sur deux fragments de mise en forme
# differente (« Aldeba » en normal puis « ran » en gras, ce que Word produit
# apres une correction orthographique) recoit la mise en forme du premier
# fragment. Il n'y a pas de bonne reponse dans ce cas.

def _bornes_des_porteurs(porteurs):
    bornes, pos = [], 0
    for p in porteurs:
        longueur = len(p.text or "")
        bornes.append((pos, pos + longueur))
        pos += longueur
    return bornes


def _appliquer_ops(porteurs, ops, xml=False):
    """Applique des remplacements (debut, fin, texte) sur une suite de
    fragments porteurs de texte, sans toucher aux autres.

    Fonctionne aussi bien sur les runs de python-docx que sur les noeuds w:t
    de lxml : les deux exposent un attribut .text.

    Les operations sont appliquees de la fin vers le debut, pour que les
    positions des suivantes restent valides.
    """
    if not ops or not porteurs:
        return 0

    for debut, fin, remplacement in sorted(ops, key=lambda o: o[0], reverse=True):
        bornes = _bornes_des_porteurs(porteurs)
        premier = True
        for i, (rd, rf) in enumerate(bornes):
            if rf <= debut or rd >= fin:
                continue
            texte = porteurs[i].text or ""
            d = max(debut, rd) - rd
            f = min(fin, rf) - rd
            if premier:
                porteurs[i].text = texte[:d] + remplacement + texte[f:]
                premier = False
                if xml and ("\n" in remplacement or "\t" in remplacement):
                    _materialiser_separateurs(porteurs[i])
            else:
                porteurs[i].text = texte[:d] + texte[f:]
            if xml:
                porteurs[i].set(XML_SPACE, "preserve")
    return len(ops)


def _materialiser_separateurs(porteur):
    """Un texte restitue qui contient un saut de ligne ou une tabulation
    (terme qui etait a cheval sur un w:br / w:tab) est redecoupe en w:t et
    noeuds w:br / w:tab, pour que Word l'affiche comme a l'origine."""
    noeud = getattr(porteur, "noeud", porteur)
    if noeud.tag != W_T or not noeud.text:
        return
    morceaux = re.split(r"([\n\t])", noeud.text)
    if len(morceaux) == 1:
        return
    parent = noeud.getparent()
    position = parent.index(noeud)
    parent.remove(noeud)
    for m in morceaux:
        if m == "\n":
            el = etree.Element(W_BR)
        elif m == "\t":
            el = etree.Element(W_TAB)
        elif m:
            el = etree.Element(W_T)
            el.text = m
            el.set(XML_SPACE, "preserve")
        else:
            continue
        parent.insert(position, el)
        position += 1


def _ops_anonymisation(texte, motif, index, para_id, remplacements):
    """Positions a remplacer, dans l'ordre d'apparition.

    Une seule passe de gauche a droite : le journal enregistre donc les
    occurrences dans leur ordre reel, ce qui permet de restituer exactement
    plusieurs formes d'un meme nom dans un meme paragraphe.
    """
    ops = []
    if not texte.strip():
        return ops
    for m in motif.finditer(texte):
        trouve = m.group(0)
        balise = index.get(normaliser(trouve))
        if balise is None:
            continue
        remplacements.append({
            "para_id": para_id,
            "balise": balise,
            "variante_originale": trouve,
        })
        ops.append((m.start(), m.end(), balise))
    return ops


def _ops_restitution(texte, entrees_para):
    """Positions des balises a restituer, file consommee de gauche a droite."""
    file = list(entrees_para)
    ops = []
    for m in MOTIF_BALISE.finditer(texte):
        balise = m.group(0)
        for k, e in enumerate(file):
            if e["balise"] == balise:
                ops.append((m.start(), m.end(), file.pop(k)["variante_originale"]))
                break
    return ops


def _ops_restitution_table(texte, balises_triees):
    """Positions des balises, remplacees par la premiere variante de la table."""
    table = dict(balises_triees)
    ops = []
    for m in MOTIF_BALISE.finditer(texte):
        variante = table.get(m.group(0))
        if variante is not None:
            ops.append((m.start(), m.end(), variante))
    return ops


# ==================================================================
# A2 - LES ZONES QUE python-docx N'EXPOSE PAS
# ==================================================================
#
# Notes de bas de page, notes de fin, commentaires, zones de texte et formes.
# C'etait le seul endroit ou l'outil produisait un document qu'on croyait sur
# et qui ne l'etait pas : il ecrivait « _ANONYMISE » en laissant des noms en
# clair dans des parties qu'il ne parcourait pas.
#
# Methode : le .docx est une archive de fichiers XML. python-docx traite le
# corps, les tableaux et les en-tetes ; cette seconde passe ouvre l'archive et
# traite les parties restantes, paragraphe par paragraphe, avec exactement la
# meme logique de remplacement (une seule passe de gauche a droite).
#
# Les remplacements de cette passe sont journalises separement, reperes par
# (partie, index du paragraphe dans la partie), afin que les para_id du corps
# restent inchanges et que les journaux deja produits restent valides.

from lxml import etree

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Parties dont TOUS les paragraphes sont a traiter.
PARTIES_COMPLETES = ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml")

# Parties ou seules les zones de texte sont a traiter : le reste du contenu
# a deja ete vu par python-docx, le retraiter ferait un double remplacement.
MOTIF_PARTIE_MIXTE = re.compile(r"^word/(document|header\d*|footer\d*)\.xml$")

def _parties_invisibles(zf):
    noms = []
    for n in zf.namelist():
        if n in PARTIES_COMPLETES or MOTIF_PARTIE_MIXTE.match(n):
            noms.append(n)
    return sorted(noms)


def _paragraphes_cibles(racine, partie):
    """Les paragraphes a traiter dans cette partie.

    Une zone de texte est souvent ecrite DEUX fois par Word (forme moderne et
    forme de repli pour les anciennes versions). Les deux copies sont traitees :
    n'en anonymiser qu'une laisserait le nom visible dans l'autre.
    """
    if partie in PARTIES_COMPLETES:
        return list(racine.iter(f"{_W}p"))
    paras = []
    for boite in racine.iter(f"{_W}txbxContent"):
        paras.extend(boite.iter(f"{_W}p"))
    return paras


def _texte_para_xml(para):
    return _texte_porteurs(_porteurs_xml(para))


def _reecrire_archive(chemin_docx, parties_modifiees):
    """Reecrit le .docx en remplacant seulement les parties changees.
    Tout le reste est recopie octet pour octet."""
    if not parties_modifiees:
        return
    temporaire = chemin_docx + ".en_cours"
    with zipfile.ZipFile(chemin_docx) as source, \
         zipfile.ZipFile(temporaire, "w", zipfile.ZIP_DEFLATED) as cible:
        for item in source.infolist():
            donnees = parties_modifiees.get(item.filename)
            cible.writestr(item, donnees if donnees is not None
                           else source.read(item.filename))
    os.replace(temporaire, chemin_docx)


def _serialiser(racine):
    return etree.tostring(racine, xml_declaration=True, encoding="UTF-8",
                          standalone=True)


def anonymiser_zones_invisibles(chemin_docx, motif, index):
    """Anonymise les parties que python-docx ne voit pas.
    Retourne la liste des remplacements, reperes par partie et paragraphe."""
    remplacements = []
    modifiees = {}

    with zipfile.ZipFile(chemin_docx) as zf:
        parties = _parties_invisibles(zf)
        contenus = {p: zf.read(p) for p in parties}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for i, para in enumerate(_paragraphes_cibles(racine, partie)):
            texte = _texte_para_xml(para)
            if not texte.strip():
                continue
            locaux = []
            ops = _ops_anonymisation(texte, motif, index, i, locaux)
            noeuds = _porteurs_xml(para)
            if _appliquer_ops(noeuds, ops, xml=True):
                for e in locaux:
                    remplacements.append({
                        "partie": partie,
                        "para_index": e["para_id"],
                        "balise": e["balise"],
                        "variante_originale": e["variante_originale"],
                    })
                change = True
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return remplacements


def restituer_zones_invisibles(chemin_docx, remplacements_xml):
    """Restitution exacte des zones invisibles, meme logique de file consommee
    de gauche a droite que pour le corps."""
    if not remplacements_xml:
        return 0

    par_partie = {}
    for e in remplacements_xml:
        par_partie.setdefault(e["partie"], {}).setdefault(e["para_index"], []).append(e)

    total = 0
    modifiees = {}
    with zipfile.ZipFile(chemin_docx) as zf:
        contenus = {p: zf.read(p) for p in par_partie if p in zf.namelist()}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for i, para in enumerate(_paragraphes_cibles(racine, partie)):
            entrees_para = par_partie[partie].get(i)
            if not entrees_para:
                continue
            texte = _texte_para_xml(para)
            noeuds = _porteurs_xml(para)
            nb = _appliquer_ops(noeuds, _ops_restitution(texte, entrees_para), xml=True)
            if nb:
                total += nb
                change = True
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return total


def restituer_zones_invisibles_par_table(chemin_docx, entrees):
    """Restitution des zones invisibles sans journal, via la table."""
    table_inverse = _premiere_variante_par_balise(entrees)
    balises_triees = sorted(table_inverse.items(), key=lambda x: len(x[0]), reverse=True)
    if not balises_triees:
        return 0

    total = 0
    modifiees = {}
    with zipfile.ZipFile(chemin_docx) as zf:
        contenus = {p: zf.read(p) for p in _parties_invisibles(zf)}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for para in _paragraphes_cibles(racine, partie):
            texte = _texte_para_xml(para)
            if not texte.strip():
                continue
            noeuds = _porteurs_xml(para)
            nb = _appliquer_ops(noeuds,
                                _ops_restitution_table(texte, balises_triees), xml=True)
            if nb:
                total += nb
                change = True
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return total


def lire_journal_xml(chemin):
    """Remplacements des zones invisibles. Vide pour les journaux anterieurs.
    Accepte un chemin ou un journal deja charge."""
    brut = _journal_brut(chemin)
    if isinstance(brut, dict):
        return brut.get("remplacements_xml", []) or []
    return []


# ==================================================================
# A3 et A4 - PROPRIETES DU FICHIER ET CIBLES DES LIENS
# ==================================================================
#
# Deux informations qui ne figurent nulle part dans le texte et que personne ne
# regarde jamais :
#
#   - les proprietes du fichier. L'auteur d'un .docx recu est tres souvent le
#     nom du confrere ou du client, et « derniere modification par » aussi.
#   - la cible d'un lien hypertexte. Le texte affiche est anonymise, l'adresse
#     derriere ne l'est pas : un lien vers le site du client survit intact.
#
# Les deux traitements sont REVERSIBLES : les valeurs d'origine sont portees
# par le journal, et la restitution les remet en place.

NS_DC = "{http://purl.org/dc/elements/1.1/}"
NS_CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
NS_EP = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
NS_VT = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Champs qui designent une personne ou une societe : leur contenu est masque
# en entier, car le nom qui s'y trouve n'est pas necessairement dans la table.
CHAMPS_IDENTITE = {
    "docProps/core.xml": [(f"{NS_DC}creator", "Auteur masqué"),
                          (f"{NS_CP}lastModifiedBy", "Auteur masqué")],
    "docProps/app.xml": [(f"{NS_EP}Company", "Société masquée"),
                         (f"{NS_EP}Manager", "Responsable masqué")],
}

# Champs de texte libre : traites comme le reste du document, terme par terme.
CHAMPS_TEXTE = {
    "docProps/core.xml": [f"{NS_DC}title", f"{NS_DC}subject", f"{NS_DC}description",
                          f"{NS_CP}keywords", f"{NS_CP}category"],
    # TitlesOfParts : Word y recopie le titre du document (8 octobre 2026).
    "docProps/app.xml": [f"{NS_EP}Template", f"{NS_EP}Application", f"{NS_VT}lpstr"],
}

LIBELLES_CHAMPS = {
    f"{NS_DC}creator": "auteur", f"{NS_CP}lastModifiedBy": "dernière modification par",
    f"{NS_EP}Company": "société", f"{NS_EP}Manager": "responsable",
    f"{NS_DC}title": "titre", f"{NS_DC}subject": "sujet",
    f"{NS_VT}lpstr": "titre (copie TitlesOfParts)",
    f"{NS_DC}description": "commentaires", f"{NS_CP}keywords": "mots-clés",
    f"{NS_CP}category": "catégorie", f"{NS_EP}Template": "modèle",
    f"{NS_EP}Application": "application",
}


def traiter_proprietes(chemin_docx, motif, index, masquer_identites=True):
    """Anonymise les proprietes du fichier. Retourne la liste des changements."""
    changements = []
    modifiees = {}

    with zipfile.ZipFile(chemin_docx) as zf:
        presentes = [p for p in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml")
                     if p in zf.namelist()]
        contenus = {p: zf.read(p) for p in presentes}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False

        if masquer_identites:
            for balise, valeur_neutre in CHAMPS_IDENTITE.get(partie, []):
                for el in racine.iter(balise):
                    ancien = (el.text or "").strip()
                    if not ancien or ancien == valeur_neutre:
                        continue
                    changements.append({
                        "partie": partie, "champ": balise,
                        "libelle": LIBELLES_CHAMPS.get(balise, balise),
                        "valeur_originale": el.text,
                        "valeur_remplacee": valeur_neutre,
                        # Masquage systematique, sans rapport avec la table :
                        # ne prouve pas que l'onglet choisi est le bon.
                        "origine": "identite",
                    })
                    el.text = valeur_neutre
                    change = True

        champs = CHAMPS_TEXTE.get(partie, [])
        if partie == "docProps/custom.xml":
            # Proprietes personnalisees (numero de dossier, client, reference) :
            # toutes les valeurs texte sont traitees terme par terme.
            champs = sorted({el.tag for el in racine.iter()
                             if isinstance(el.tag, str) and len(el) == 0
                             and (el.text or "").strip()})
        for balise in champs:
            for el in racine.iter(balise):
                ancien = el.text or ""
                if not ancien.strip():
                    continue
                locaux = []
                ops = _ops_anonymisation(ancien, motif, index, 0, locaux)
                if not ops:
                    continue
                nouveau = ancien
                for debut, fin, remplacement in sorted(ops, key=lambda o: o[0],
                                                       reverse=True):
                    nouveau = nouveau[:debut] + remplacement + nouveau[fin:]
                changements.append({
                    "partie": partie, "champ": balise,
                    "libelle": LIBELLES_CHAMPS.get(balise, balise),
                    "valeur_originale": ancien,
                    "valeur_remplacee": nouveau,
                    "origine": "terme",
                })
                el.text = nouveau
                change = True

        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return changements


# ------------------------------------------------------------------
# AUTEURS DES COMMENTAIRES ET DES MODIFICATIONS SUIVIES
# ------------------------------------------------------------------
# Word attache a chaque commentaire et a chaque modification suivie le nom de
# son auteur, ses initiales et, dans people.xml, l'identifiant de son compte
# (souvent son adresse electronique). Ces noms ne sont pas forcement dans la
# table : ils sont remplaces systematiquement par « Auteur 1 », « Auteur 2 »...
# (meme regle et meme reglage que l'auteur des proprietes), et le journal
# garde les valeurs d'origine pour la restitution.

MOTIF_AUTEUR_NEUTRE = re.compile(r"^(?:Auteur|Author) \d+$|^A\d+$|^(?:utilisateur|user)-\d+$")


def _nom_local(cle):
    return cle.rsplit("}", 1)[-1]


def traiter_auteurs(chemin_docx):
    """Remplace les noms d'auteurs des commentaires et des revisions.
    Retourne les changements, au format des proprietes (champ « attribut »)."""
    if est_texte(chemin_docx):
        return []
    with zipfile.ZipFile(chemin_docx) as zf:
        parties = [n for n in zf.namelist()
                   if n.startswith("word/") and n.endswith(".xml")]
        contenus = {n: zf.read(n) for n in parties}

    # Numerotation commune a tout le document, dans l'ordre d'apparition :
    # le meme auteur porte le meme numero dans le corps et dans les commentaires.
    numeros = {}

    def numero(nom):
        if nom not in numeros:
            numeros[nom] = len(numeros) + 1
        return numeros[nom]

    autres = {}             # initiales ou identifiants sans auteur sur l'element

    def neutre_isole(attribut, valeur):
        cle = (attribut, valeur)
        if cle not in autres:
            autres[cle] = sum(1 for a, _ in autres if a == attribut) + 1
        n = autres[cle]
        return f"A{n}" if attribut == "initials" else _t(f"utilisateur-{n}", f"user-{n}")

    changements = {}
    modifiees = {}
    numero_element = {}
    for partie in parties:
        brut = contenus[partie]
        if b"author" not in brut and b"initials" not in brut and b"userId" not in brut:
            continue
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for el in racine.iter():
            if not isinstance(el.tag, str):
                continue
            attributs = {_nom_local(k): k for k in el.attrib}
            if not any(a in attributs for a in ATTRIBUTS_PERSONNES):
                continue
            auteur = el.get(attributs["author"], "").strip() if "author" in attributs else ""
            n = numero(auteur) if auteur and not MOTIF_AUTEUR_NEUTRE.match(auteur) else None
            # people.xml : l'identifiant est porte par l'enfant presenceInfo,
            # l'auteur par l'element parent person.
            if n is None and "userId" in attributs and el.getparent() is not None:
                n = numero_element.get(el.getparent())
            if n is not None:
                numero_element[el] = n
            for local in ATTRIBUTS_PERSONNES:
                if local not in attributs:
                    continue
                cle = attributs[local]
                ancien = el.get(cle) or ""
                if not ancien.strip() or MOTIF_AUTEUR_NEUTRE.match(ancien.strip()):
                    continue
                if local == "author":
                    nouveau = _t(f"Auteur {n}", f"Author {n}")
                elif n is not None:
                    nouveau = f"A{n}" if local == "initials" else _t(f"Auteur {n}", f"Author {n}")
                else:
                    nouveau = neutre_isole(local, ancien)
                el.set(cle, nouveau)
                change = True
                changements.setdefault((partie, cle, nouveau), ancien)
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return [{"partie": partie, "attribut": cle,
             "libelle": _t("auteur des commentaires et révisions",
                           "author of comments and revisions"),
             "valeur_originale": ancien, "valeur_remplacee": nouveau,
             "origine": "identite"}
            for (partie, cle, nouveau), ancien in changements.items()]


def traiter_liens(chemin_docx, entrees):
    """Neutralise les cibles de liens qui portent un terme de la table.

    L'adresse est remplacee par une adresse neutre et syntaxiquement valide
    plutot que par une balise : une URL contenant des crochets ferait tiquer
    Word a l'ouverture.
    """
    variantes = [v for _, vs in entrees for v in vs if len(v) >= 3]
    if not variantes:
        return []

    changements = []
    modifiees = {}
    compteur = 0

    with zipfile.ZipFile(chemin_docx) as zf:
        rels = [n for n in zf.namelist() if n.endswith(".rels")]
        contenus = {n: zf.read(n) for n in rels}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for el in racine.iter(f"{NS_REL}Relationship"):
            if el.get("TargetMode") != "External":
                continue
            cible = el.get("Target") or ""
            if not any(_motif_dans_nom(v).search(cible) for v in variantes):
                continue
            compteur += 1
            neutre = f"https://lien-anonymise.invalid/{compteur}"
            changements.append({
                "partie": partie,
                "id": el.get("Id"),
                "cible_originale": cible,
                "cible_remplacee": neutre,
            })
            el.set("Target", neutre)
            change = True
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return changements


def restituer_proprietes(chemin_docx, changements):
    """Remet les proprietes d'origine."""
    if not changements:
        return 0
    par_partie = {}
    for c in changements:
        par_partie.setdefault(c["partie"], []).append(c)

    total = 0
    modifiees = {}
    with zipfile.ZipFile(chemin_docx) as zf:
        contenus = {p: zf.read(p) for p in par_partie if p in zf.namelist()}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for c in par_partie[partie]:
            if "attribut" in c:
                # auteur d'un commentaire ou d'une revision : toutes les
                # occurrences de la valeur neutre reprennent la valeur d'origine
                for el in racine.iter():
                    if isinstance(el.tag, str) and el.get(c["attribut"]) == c["valeur_remplacee"]:
                        el.set(c["attribut"], c["valeur_originale"])
                        total += 1
                        change = True
                continue
            for el in racine.iter(c["champ"]):
                if (el.text or "") == c["valeur_remplacee"]:
                    el.text = c["valeur_originale"]
                    total += 1
                    change = True
                    break
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return total


def restituer_liens(chemin_docx, changements):
    """Remet les adresses d'origine."""
    if not changements:
        return 0
    par_partie = {}
    for c in changements:
        par_partie.setdefault(c["partie"], []).append(c)

    total = 0
    modifiees = {}
    with zipfile.ZipFile(chemin_docx) as zf:
        contenus = {p: zf.read(p) for p in par_partie if p in zf.namelist()}

    for partie, brut in contenus.items():
        try:
            racine = etree.fromstring(brut)
        except etree.XMLSyntaxError:
            continue
        change = False
        for el in racine.iter(f"{NS_REL}Relationship"):
            for c in par_partie[partie]:
                if el.get("Id") == c["id"] and el.get("Target") == c["cible_remplacee"]:
                    el.set("Target", c["cible_originale"])
                    total += 1
                    change = True
                    break
        if change:
            modifiees[partie] = _serialiser(racine)

    _reecrire_archive(chemin_docx, modifiees)
    return total


def lire_journal_annexes(chemin):
    """Proprietes et liens enregistres. Vides pour les journaux anterieurs.
    Accepte un chemin ou un journal deja charge."""
    brut = _journal_brut(chemin)
    if isinstance(brut, dict):
        return (brut.get("proprietes", []) or [], brut.get("liens", []) or [])
    return [], []


# ==================================================================
# MARQUEURS DE REDACTION ET BALISES REELLES
# ==================================================================
#
# Le motif de balise reconnait tout bloc en majuscules entre crochets. Or un
# acte en contient souvent qui n'ont RIEN a voir avec la pseudonymisation :
# [DATE], [MONTANT], [A COMPLETER], [Barreau]... Ce sont des marqueurs que le
# redacteur a laisses pour lui-meme.
#
# Les confondre avec des balises avait une consequence concrete : la
# restitution refusait d'ecrire, au motif que la table « ne couvrait pas »
# une balise que personne n'avait jamais produite et qu'aucune table ne
# pourrait jamais restituer.
#
# La regle est donc : seule compte comme balise ce qui est DECLARE en
# colonne B d'un onglet d'une des tables. Le reste est laisse en place et
# signale pour information.

_cache_balises_declarees = {}


def balises_declarees(config, forcer=False):
    """Toutes les balises declarees dans les onglets des tables configurees."""
    if not config:
        return set()
    # La cle porte la date de modification des classeurs : une balise ajoutee
    # dans Excel pendant la session est donc prise en compte des
    # l'enregistrement, sans redemarrer le programme.
    def _date(f):
        try:
            return os.path.getmtime(f)
        except OSError:
            return 0
    cle = tuple(sorted((t["fichier"], _date(t["fichier"])) for t in config.get("tables", [])))
    if not forcer and cle in _cache_balises_declarees:
        return _cache_balises_declarees[cle]

    declarees = set()
    for table in config.get("tables", []):
        fichier = table.get("fichier")
        if not fichier or not os.path.isfile(fichier):
            continue
        try:
            wb = openpyxl.load_workbook(fichier, read_only=True, data_only=True)
        except Exception:
            continue
        try:
            for ws in wb.worksheets:
                if ws.title in ONGLETS_MASQUES or ws.title.endswith(SUFFIXES_ONGLETS_MASQUES):
                    continue
                for row in ws.iter_rows(min_row=1, max_col=2, values_only=True):
                    if row and len(row) > 1 and row[1]:
                        valeur = str(row[1]).strip()
                        if MOTIF_BALISE.fullmatch(valeur):
                            declarees.add(valeur)
        except Exception:
            pass
        finally:
            wb.close()

    _cache_balises_declarees[cle] = declarees
    return declarees


def trier_balises(balises, config):
    """Separe les balises reelles des marqueurs de redaction.
    Retourne (balises_reelles, marqueurs)."""
    declarees = balises_declarees(config)
    if not declarees:
        # Aucune table lisible : par prudence, tout est considere comme balise.
        return set(balises), set()
    reelles = {b for b in balises if b in declarees}
    return reelles, set(balises) - reelles
