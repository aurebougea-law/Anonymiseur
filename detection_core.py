"""
detection_core.py
-----------------
Pre-detection des termes a pseudonymiser dans un document (.docx ou PDF natif).

Le moteur est ENTIEREMENT DETERMINISTE : uniquement des motifs, aucune IA,
aucun appel reseau. Le document ne quitte jamais le poste.

Ce qu'il produit : une liste d'entites, chacune portant
  - sa forme canonique (la forme la plus complete rencontree) ;
  - toutes ses variantes REELLEMENT presentes dans le texte ;
  - sa position de premiere apparition ;
  - son type, un extrait de contexte, un niveau de confiance.
(Le rattachement a une partie est fait ensuite par detection_roles.py.)

Deux principes de conception :
  1. Aucune donnee metier en dur. Les categories du classeur sont lues dans le
     classeur ; le lien type de detection -> categorie vit dans
     detection_config.json, pas dans ce fichier.
  2. Le moteur PROPOSE, il ne decide pas. Toute ambiguite est signalee plutot
     que tranchee.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import os
import re
import json
import unicodedata

MAJ = "A-ZÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸÆŒ"
MIN = "a-zà-öø-ÿ"
# Un composant de nom propre : commence obligatoirement par une majuscule.
MOT_PROPRE = rf"[{MAJ}][{MAJ}{MIN}'’\-]*"
PARTICULE_RE = r"(?:d[eu]s?|de|du|des|le|la|les|van|von|der|den|di|da|el|al)"
# Les composants d'un nom se suivent sur la MEME ligne. Avec \s, la capture
# franchissait la fin de ligne et ramassait la qualite ecrite en dessous :
# « Monsieur Théobald Grinchu\n Défenderesse » devenait un patronyme de trois
# mots, introuvable tel quel dans le document.
NOM_PROPRE = (rf"(?:{MOT_PROPRE}|{PARTICULE_RE})"
              rf"(?:[^\S\n'’\-]*[’'\-]?[^\S\n]*(?:{MOT_PROPRE}|{PARTICULE_RE})){{0,4}}")

# Particules qui font partie d'un nom et ne l'interrompent pas.
PARTICULES = {"de", "du", "des", "d'", "d’", "le", "la", "les", "van", "von",
              "der", "den", "el", "al", "di", "da", "dos", "ben"}

# Mots qui, eux, terminent un nom propre : ils ne peuvent pas en faire partie.
ARRET = {
    "et", "ou", "en", "au", "aux", "par", "pour", "dans", "sur", "sous", "chez",
    "societe", "société", "monsieur", "madame", "maitre", "maître", "demeurant",
    "domicilie", "domicilié", "domiciliee", "domiciliée", "ne", "né", "nee", "née",
    "agissant", "representee", "représentée", "represente", "représenté", "ayant",
    "sise", "sis", "immatriculee", "immatriculée", "immatricule", "immatriculé",
    "exercant", "exerçant", "avocat", "avocate", "commissaire", "huissier",
    "titulaire", "exploitant", "inscrit", "inscrite", "es", "ès", "qualites",
    "qualités", "prise", "pris", "dont", "siege", "siège", "social", "a", "à",
    "the", "of", "ci", "apres", "après", "denomme", "dénommé", "denommee", "dénommée",
    # Qualites : elles suivent le nom et n'en font pas partie.
    "gerant", "gérant", "gerante", "gérante", "president", "président",
    "presidente", "présidente", "directeur", "directrice", "dirigeant",
    "dirigeante", "administrateur", "liquidateur", "mandataire", "associe",
    "associé", "associee", "associée", "demandeur", "demanderesse",
    "defendeur", "défendeur", "defenderesse", "défenderesse", "requerant",
    "requérant", "requerante", "requérante", "intime", "intimé", "appelant",
    "appelante",
    # 4 novembre 2026 : coordonnees d'en-tete. Un Word issu d'un PDF (L'atelier
    # du PDF) recolle les lignes d'un bloc avec des espaces : « Maitre X Tel »,
    # « 75008 PARIS Avocat ». Ces mots ne font jamais partie d'un nom.
    # Mots en minuscules sans accents (la comparaison se fait ainsi).
    "tel", "telephone", "fax", "telecopie", "email", "e-mail", "courriel",
    "toque", "vestiaire", "barreau", "plaidant", "postulant",
}

# Mots qui ouvrent une phrase ou une ligne. Ecrits avec une seule majuscule
# initiale (« Dans », « Par », « À »), ils ne prolongent jamais une
# denomination : dans un Word issu d'un PDF, le paragraphe suivant est souvent
# recolle au nom (« HALDOR SAS Dans ces conditions »), ce qui creait une
# societe « HALDOR Dans » de plus (4 novembre 2026). Ecrits en capitales
# (« LES VERGERS DE L'ESCAUT ET FILS »), ils restent dans le nom.
OUVREURS_DE_PHRASE = {
    "dans", "par", "pour", "sur", "sous", "selon", "enfin", "ainsi", "or", "mais",
    "donc", "car", "cette", "ce", "ces", "cet", "il", "elle", "ils", "elles", "nous",
    "vous", "en", "au", "aux", "a", "et", "ou", "un", "une", "attendu", "considerant",
    "vu", "aussi", "depuis", "lors", "apres", "avant", "puis", "alors", "toutefois",
    "cependant", "neanmoins", "partant", "outre", "afin", "si", "qui", "que", "sans",
    "avec", "monsieur", "madame", "maitre", "avocat", "avocate", "demanderesse",
    "defenderesse", "demandeur", "defendeur", "ayant", "representee", "represente",
    "tel", "telephone", "fax", "telecopie", "mail", "email", "e-mail", "courriel",
    "toque", "palais", "vestiaire", "avocats", "barreau", "plaidant", "postulant",
}

FORMES_SOCIALES = [
    "SARL", "S.A.R.L.", "EURL", "E.U.R.L.", "SASU", "S.A.S.U.", "SAS", "S.A.S.",
    "SA", "S.A.", "SCI", "S.C.I.", "SNC", "S.N.C.", "SCM", "SCOP", "SCIC",
    "SELARL", "SELAS", "SELAFA", "SCP", "AARPI", "GIE", "GAEC", "SEM", "SPL",
    "association", "fondation", "syndicat",
    # Formes anglophones (documents en anglais) : reconnues telles quelles.
    "Ltd", "Ltd.", "Limited", "LLC", "L.L.C.", "Inc", "Inc.", "PLC", "LLP", "LP",
    "GmbH", "Corp", "Corp.", "Corporation",
]
# Formes qui ne se placent QUE derriere le nom (« JEREMIAS GmbH ») : les
# chercher devant collait la societe suivante au nom (« GmbH ATELIERS X »).
FORMES_POSTPOSEES_SEULES = {"Ltd", "Ltd.", "Limited", "LLC", "L.L.C.", "Inc", "Inc.",
                            "PLC", "LLP", "LP", "GmbH", "Corp", "Corp.", "Corporation",
                            "AG", "SA", "S.A.", "SAS", "S.A.S."}

# Le libelle de voie peut etre coupe par un retour a la ligne : un document
# converti depuis un PDF renvoie a la ligne n'importe ou. On admet UN retour,
# jamais une ligne blanche, qui marquerait un changement de paragraphe.
LIBELLE_VOIE = r"(?:[^,;\n]|\n(?![ \t]*\n)){2,70}?"

# Une ville : au plus quatre composants capitalises, sur UNE SEULE LIGNE.
# Pas de \s dans la classe, sinon la capture franchit la fin de ligne.
# Toujours employee sous (?-i:...) : les motifs d'adresse sont insensibles a
# la casse, et sans cela la ville avalait les mots suivants (« Lille et
# Madame », « Lille né le »), rendant le terme introuvable (8 octobre 2026).
VILLE = (rf"[{MAJ}][{MAJ}{MIN}\-'’]{{1,30}}"
         rf"(?:[^\S\n]+[{MAJ}][{MAJ}{MIN}\-'’]{{1,30}}){{0,3}}")

# Mots de procedure qui suivent parfois une ville sur la meme ligne et ne font
# pas partie de son nom.
ARRET_VILLE = {
    "contre", "pour", "demandeur", "demanderesse", "defendeur", "defendeur",
    "defenderesse", "objet", "ayant", "lettre", "pieces", "faits", "plaise",
    "tribunal", "par", "ci", "apres", "et", "representee", "represente",
    "immatriculee", "immatricule", "prise", "agissant", "dont", "siege",
    # 4 novembre 2026 : coordonnees recollees a la ville (Word issu d'un PDF)
    "tel", "telephone", "fax", "telecopie", "mail", "email", "e-mail", "courriel",
    "toque", "palais", "vestiaire", "avocat", "avocats", "avocate", "barreau",
    "dans", "a", "le", "la", "les", "en", "au", "aux", "sur", "selon", "enfin",
    "ainsi", "attendu", "vu", "maitre", "monsieur", "madame", "societe", "case",
}

TYPES_VOIE = [
    "rue", "avenue", "av\\.", "boulevard", "bd", "place", "impasse", "chemin",
    "route", "allée", "allee", "quai", "cours", "square", "villa", "passage",
    "promenade", "esplanade", "faubourg", "hameau", "lieu-dit", "lieudit",
    "résidence", "residence", "zone", "ZA", "ZI", "ZAC", "parc", "rond-point",
]


class EntiteDetectee:
    """Une personne, une societe, une adresse... et toutes ses formes."""

    def __init__(self, canonique, type_, position, contexte="", confiance="haute"):
        self.canonique = canonique
        self.type = type_
        self.position = position
        self.variantes = [canonique]
        self.contexte = contexte
        self.confiance = confiance
        self.camp = ""          # conserve pour compatibilite, plus utilise
        self.partie = ""
        self.role = ""
        self.balise = ""
        self.occurrences = 0
        self.page = ""          # numero de page de la premiere occurrence, si connu
        self.ambigu = False
        self.note = ""

    def ajouter_variante(self, v):
        if v and v not in self.variantes:
            self.variantes.append(v)

    def terme_pour_tableur(self):
        """Les variantes de la plus longue a la plus courte, separees par « / ».
        L'ordre importe : l'anonymiseur remplace la plus longue d'abord."""
        return " / ".join(sorted(self.variantes, key=len, reverse=True))

    def __repr__(self):
        return f"<{self.type} {self.canonique!r} pos={self.position}>"


# ==================================================================
# EXTRACTION DU TEXTE
# ==================================================================

SAUT_PAGE = "\f"           # marque le debut d'une nouvelle page dans le texte extrait


def texte_docx(chemin):
    from docx import Document
    doc = Document(chemin)
    morceaux = []

    def para_texte(p, pages=True):
        # SAUT_PAGE devant le texte d'un passage qui ouvre une nouvelle page
        # (saut rendu par Word, ou saut de page force) : sert au numero de page.
        parts = []
        for r in p.runs:
            if pages and r._r.xpath('.//w:lastRenderedPageBreak | .//w:br[@w:type="page"]'):
                parts.append(SAUT_PAGE + "\n")      # le texte suivant reste en debut de ligne
            parts.append(r.text)
        return "".join(parts)

    for p in doc.paragraphs:
        morceaux.append(para_texte(p))
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    morceaux.append(para_texte(p))
    for section in doc.sections:
        for hf in (section.header, section.footer):
            if hf is not None:
                for p in hf.paragraphs:
                    morceaux.append(para_texte(p, pages=False))
    lignes, commentaires = _annexes_docx(chemin, avec_commentaires=True)
    morceaux.extend(lignes)
    corps = "\n".join(morceaux)
    if not commentaires:
        return TexteDocument(corps)
    # Les commentaires viennent en dernier ; leur zone est memorisee pour que
    # la detection sache ce qui n'est qu'une note de relecture (4/11/2026).
    debut = len(corps) + 1
    texte = TexteDocument(corps + "\n" + "\n".join(commentaires))
    texte.zones_commentaires = [(debut, len(texte))]
    return texte


class TexteDocument(str):
    """Texte extrait, qui sait ou se trouvent les commentaires Word."""
    zones_commentaires = ()


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _annexes_docx(chemin, avec_commentaires=False):
    """Texte que python-docx ne parcourt pas : notes de bas de page, notes de
    fin, commentaires et zones de texte. Sans lui, un nom qui ne figure que
    dans une note echappait a la pre-detection (8 octobre 2026).
    Place apres le corps : la page attribuee est alors la derniere."""
    import zipfile
    import xml.etree.ElementTree as ET
    lignes = []
    commentaires = []

    def paragraphes(racine, cible=lignes):
        for para in racine.iter(f"{_W}p"):
            texte = "".join(t.text or "" for t in para.iter(f"{_W}t"))
            if texte.strip():
                cible.append(texte)

    try:
        with zipfile.ZipFile(chemin) as zf:
            noms = zf.namelist()
            for partie in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
                if partie in noms:
                    try:
                        paragraphes(ET.fromstring(zf.read(partie)),
                                    commentaires if partie.endswith("comments.xml")
                                    and avec_commentaires else lignes)
                    except ET.ParseError:
                        continue
            vues = set()
            for partie in noms:
                if not re.match(r"^word/(document|header\d*|footer\d*)\.xml$", partie):
                    continue
                try:
                    racine = ET.fromstring(zf.read(partie))
                except ET.ParseError:
                    continue
                # Une zone de texte est ecrite deux fois dans le fichier (forme
                # moderne et forme de secours) : une seule lecture.
                for boite in racine.iter(f"{_W}txbxContent"):
                    texte = "\n".join("".join(t.text or "" for t in para.iter(f"{_W}t"))
                                      for para in boite.iter(f"{_W}p")).strip()
                    if texte and texte not in vues:
                        vues.add(texte)
                        lignes.append(texte)
    except (OSError, zipfile.BadZipFile):
        return ([], []) if avec_commentaires else []
    return (lignes, commentaires) if avec_commentaires else lignes


def numero_page(texte, position):
    """Page d'une position dans le texte extrait : 1 + sauts de page avant."""
    return texte.count(SAUT_PAGE, 0, position) + 1


def pages_connues(texte):
    """Vrai si le texte porte au moins un saut de page (document de plus
    d'une page dont la pagination est connue)."""
    return SAUT_PAGE in texte


def _pymupdf():
    """PyMuPDF s'importe sous son nouveau nom (pymupdf), ou sous l'ancien
    (fitz) sur une installation plus vieille. Leve ImportError si absent."""
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        import fitz
        return fitz


def texte_pdf(chemin):
    """PDF natif. Essaie les bibliotheques disponibles, dans l'ordre de qualite."""
    erreurs = []
    try:
        fitz = _pymupdf()
        with fitz.open(chemin) as doc:
            return ("\n" + SAUT_PAGE + "\n").join(page.get_text() for page in doc)
    except ImportError:
        erreurs.append("PyMuPDF absent")
    except Exception as e:
        erreurs.append(f"PyMuPDF : {e}")

    try:
        import pdfplumber
        with pdfplumber.open(chemin) as pdf:
            return ("\n" + SAUT_PAGE + "\n").join((page.extract_text() or "") for page in pdf.pages)
    except ImportError:
        erreurs.append("pdfplumber absent")
    except Exception as e:
        erreurs.append(f"pdfplumber : {e}")

    try:
        from pypdf import PdfReader
        return ("\n" + SAUT_PAGE + "\n").join((p.extract_text() or "") for p in PdfReader(chemin).pages)
    except ImportError:
        erreurs.append("pypdf absent")
    except Exception as e:
        erreurs.append(f"pypdf : {e}")

    raise RuntimeError(
        "Aucune bibliotheque de lecture PDF n'a fonctionne.\n"
        + "\n".join("  - " + e for e in erreurs)
        + "\n\nInstallez-en une :  py -m pip install pymupdf"
    )


def extraire_texte(chemin):
    ext = os.path.splitext(chemin)[1].lower()
    if ext == ".docx":
        return texte_docx(chemin)
    if ext == ".pdf":
        texte = texte_pdf(chemin)
        if len(texte.strip()) < 40:
            raise RuntimeError(
                "Ce PDF ne contient quasiment aucun texte : c'est vraisemblablement\n"
                "un document scanne. Il faut d'abord le passer par l'OCR."
            )
        return texte
    raise RuntimeError(f"Format non pris en charge : {ext}\nAttendu : .docx ou .pdf")


# ==================================================================
# OUTILS
# ==================================================================

def _sans_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn").lower()


# Les mots qui terminent un NOM DE PERSONNE mais font partie d'une
# DENOMINATION. « Monsieur Durand et Madame Martin » fait deux personnes ;
# « LES VERGERS DE L'ESCAUT ET FILS » fait une societe. Les melanger faisait
# perdre toute denomination commencant par un article et tronquait
# « TROMPETTE ET CHERUBIN » a son premier mot.
ARRET_PERSONNE_SEULEMENT = {
    "et", "ou", "en", "au", "aux", "par", "pour", "dans", "sur", "sous", "chez",
    "the", "of", "a", "à", "es", "ès",
}

ARRET_DENOMINATION = ARRET - ARRET_PERSONNE_SEULEMENT


FORMES_STRUCTURE_AVOCATS = {"SELARL", "SELAS", "SCP", "AARPI", "SELAFA", "SELEURL", "CABINET"}


def _nettoyer_nom(brut, genre="personne"):
    """Ne garde que des composants de nom propre.

    Un composant est retenu s'il commence par une majuscule, ou s'il s'agit
    d'une particule suivie d'un autre composant. Tout le reste coupe le nom :
    c'est ce qui empeche « Monsieur VANTERRE n'a pas repondu » d'etre pris
    pour un patronyme.

    genre="denomination" pour une societe, une enseigne ou une structure
    d'exercice : l'article de tete est alors garde (« LE PETIT NORD »,
    « LES VERGERS ») et les conjonctions ne coupent plus le nom
    (« TROMPETTE ET CHERUBIN »). Avec genre="personne", rien ne change.
    """
    arret = ARRET_DENOMINATION if genre == "denomination" else ARRET
    mots = [m for m in re.split(r"\s+", brut) if m]
    if genre == "denomination":
        # « SOCIETE ZETA S.A.S. » : le mot societe en tete n'est pas le nom
        nus = [_sans_accents(m.strip(".,;:()«»\"'")).lower() for m in mots]
        if len(nus) > 1 and nus[0] in ("la", "le", "les", "the") and nus[1] == "societe":
            mots, nus = mots[1:], nus[1:]
        while nus and nus[0] == "societe":
            mots, nus = mots[1:], nus[1:]
    gardes = []
    for i, m in enumerate(mots):
        propre = m.strip(".,;:()«»\"'")
        nu = _sans_accents(propre)
        if not propre:
            break
        if nu in arret:
            break
        # « Dans », « Par », « Tel » apres un nom : debut de la ligne suivante
        if gardes and nu in OUVREURS_DE_PHRASE and propre[:1].isupper() \
                and not propre.isupper():
            break
        if nu in PARTICULES:
            suivant = mots[i + 1].strip(".,;:()«»\"'") if i + 1 < len(mots) else ""
            if not suivant or not suivant[:1].isupper():
                break
            # Une particule en tete n'ouvre un nom que pour une denomination :
            # « Monsieur le President Durand » ne doit pas devenir un patronyme.
            if not gardes and genre != "denomination":
                break
            gardes.append(propre)
            continue
        # Un nom de personne s'arrete a une forme sociale (« Maitre X SCP X
        # FAUCHER ») et a un mot repete (« Laure VIGNAC VIGNAC IP ») : signe
        # d'un en-tete recolle sur une seule ligne (4 novembre 2026).
        if genre != "denomination" and gardes and (
                propre in FORMES_SOCIALES or propre in FORMES_STRUCTURE_AVOCATS
                or nu == _sans_accents(gardes[-1])):
            break
        if propre == "&" and genre == "denomination" and gardes:
            gardes.append(propre)
            continue
        if not propre[:1].isupper():
            break
        gardes.append(propre)
    while gardes and (_sans_accents(gardes[-1]) in PARTICULES or gardes[-1] == "&"):
        gardes.pop()
    return " ".join(gardes).strip(" ,;:.")


def _contient_conjonction(nom):
    """Le nom retenu contient-il un « et » ou un « & » ?

    Deux societes citees cote a cote (« les societes ALPHA et BETA ») se
    presentent alors comme une seule. On ne tranche pas a sa place : on
    detecte la forme longue et on baisse la confiance pour qu'elle soit
    verifiee dans le classeur de propositions.
    """
    return bool(re.search(r"\b(et|&)\b", _sans_accents(nom), re.IGNORECASE))


def _contexte(texte, position, largeur=90):
    debut = max(0, position - largeur // 2)
    fin = min(len(texte), position + largeur)
    extrait = texte[debut:fin].replace("\n", " ").replace(SAUT_PAGE, " ")
    return ("..." if debut else "") + re.sub(r"\s+", " ", extrait) + ("..." if fin < len(texte) else "")


def _luhn(numero):
    """Validation de la cle de Luhn : ecarte l'essentiel des faux SIREN/SIRET."""
    chiffres = [int(c) for c in numero if c.isdigit()]
    if not chiffres:
        return False
    total, pair = 0, False
    for c in reversed(chiffres):
        if pair:
            c *= 2
            if c > 9:
                c -= 9
        total += c
        pair = not pair
    return total % 10 == 0


# Longueur de l'IBAN par pays (les plus courants). Sert a couper le numero au
# bon endroit quand un mot en capitales le suit (« ... 189 PARIS »).
LONGUEUR_IBAN = {
    "FR": 27, "MC": 27, "BE": 16, "LU": 20, "DE": 22, "NL": 18, "ES": 24,
    "IT": 27, "PT": 25, "CH": 21, "AT": 20, "IE": 22, "GB": 22, "DK": 18,
    "FI": 18, "SE": 24, "NO": 15, "PL": 28, "GR": 27, "CZ": 24, "SK": 24,
    "HU": 28, "RO": 24, "BG": 22, "HR": 21, "SI": 19, "EE": 20, "LT": 20,
    "LV": 21, "MT": 31, "CY": 28, "LI": 21, "IS": 26,
}


def _cle_iban(compact):
    """Cle IBAN (modulo 97) : vrai si le numero est coherent."""
    if len(compact) < 15 or not compact[:2].isalpha() or not compact[2:4].isdigit():
        return False
    tourne = compact[4:] + compact[:4]
    try:
        nombre = "".join(str(int(c, 36)) for c in tourne)
    except ValueError:
        return False
    return int(nombre) % 97 == 1


def _cle_nir(chiffres13, cle):
    """Cle du numero de securite sociale : 97 - (13 premiers chiffres modulo 97).
    Corse : 2A vaut 19, 2B vaut 18 dans le calcul."""
    base = chiffres13.upper().replace("2A", "19").replace("2B", "18")
    if not base.isdigit() or not cle or not cle.isdigit():
        return False
    return 97 - int(base) % 97 == int(cle)


# ==================================================================
# DETECTEURS
# ==================================================================

def _detecter_personnes(texte):
    """Monsieur / Madame / M. / Mme + nom, et les representants legaux."""
    entites = []

    civilites = (r"(?:Monsieur|Madame|Mademoiselle|MONSIEUR|MADAME|M\.|Mme|Mle|MM\.|Mmes"
                 r"|Mr\.?|Mrs\.?|Ms\.?|Miss|Dr\.?)")
    motif = re.compile(rf"\b({civilites})\s+({NOM_PROPRE})")

    for m in motif.finditer(texte):
        civilite, brut = m.group(1), m.group(2)
        # « MM. » / « Mmes » : des auteurs cites en doctrine (« MM. Cornu et
        # Foyer »), jamais une partie, qui est toujours nommee seule
        # (15 octobre 2026).
        if civilite in ("MM.", "Mmes"):
            continue
        nom = _nettoyer_nom(brut)
        if not nom or len(nom) < 2:
            continue

        # Representant legal : la mention le precede de peu.
        # 4 novembre 2026 : une simple fonction dans les 120 caracteres
        # precedents suffisait (« la Vice-Presidente... Monsieur X », « directeur
        # du laboratoire ») ; un inventeur devenait « Representant legal ». Seules
        # les formules de representation, juste avant le nom, comptent desormais.
        avant = _sans_accents(texte[max(0, m.start() - 80):m.start()])
        est_rl = bool(re.search(
            r"(?:represente(?:e|s|es)?\s+par|pris(?:e)?\s+en\s+la\s+personne\s+de"
            r"|agissant\s+(?:par|poursuites)|acting\s+through|represented\s+by)"
            r"(?:\s+(?:son|sa|ses|its|their)\s+[a-z\- ]{0,40}?)?\s*,?\s*$"
            r"|\b(?:son|sa|its)\s+(?:president|presidente|gerant|gerante|directeur general"
            r"|directrice generale|dirigeant|dirigeante|representant legal|representante legale"
            r"|chief executive officer|managing director)\s*,?\s*$",
            avant))

        e = EntiteDetectee(f"{civilite} {nom}",
                           "representant_legal" if est_rl else "personne_physique",
                           m.start(), _contexte(texte, m.start()))
        e.ajouter_variante(nom)

        # Variantes plausibles, retenues seulement si presentes dans le texte.
        parties = nom.split()
        if len(parties) >= 2:
            patronyme = parties[-1]
            prenom = parties[0]
            # Les formes abregees suivent la civilite relevee : « Madame X »
            # ne recoit pas la variante « Monsieur X » (8 octobre 2026).
            feminin = civilite.lower().rstrip(".") in CIVILITES_FEMININES
            masculin = civilite.lower().rstrip(".") in CIVILITES_MASCULINES
            candidates = [patronyme, f"{prenom} {patronyme}"]
            if not feminin:
                candidates += [f"M. {patronyme}", f"Monsieur {patronyme}"]
            if not masculin:
                candidates += [f"Mme {patronyme}", f"Madame {patronyme}"]
            for candidate in candidates:
                if len(candidate) > 2 and re.search(rf"\b{re.escape(candidate)}\b", texte):
                    e.ajouter_variante(candidate)
        entites.append(e)

    return entites


CIVILITES_FEMININES = ("madame", "mademoiselle", "mme", "mle", "mmes", "mrs", "ms", "miss")
CIVILITES_MASCULINES = ("monsieur", "m", "mm", "mr")


def _genre_civilite(nom):
    """« f », « m » ou "" selon la civilite en tete du nom."""
    tete = nom.split()[0].lower().rstrip(".") if nom.split() else ""
    if tete in CIVILITES_FEMININES:
        return "f"
    if tete in CIVILITES_MASCULINES:
        return "m"
    return ""


# Premier mot d'une denomination qui ne la designe pas a lui seul : « GROUPE
# ALPHA SAS » ne doit pas faire remplacer chaque « groupe » du document.
MOTS_GENERIQUES_DENOMINATION = {
    "groupe", "ateliers", "atelier", "laboratoires", "laboratoire", "etablissements",
    "etablissement", "compagnie", "banque", "caisse", "cabinet", "societe", "entreprise",
    "entreprises", "agence", "bureau", "centre", "institut", "maison", "editions", "studio",
    "clinique", "fondation", "association", "union", "mutuelle", "federation", "syndicat",
    "office", "service", "services", "transports", "industries", "industrie", "holding",
    "immobilier", "immobiliere", "distribution", "conseil", "consulting", "solutions",
    "technologies", "international", "france", "europe", "group", "company", "partners",
}


# Premier mot qui fait d'une « denomination » une simple phrase.
DETERMINANTS_NON_NOMS = {"une", "un", "cette", "ce", "cet", "toute", "tout", "chaque", "notre",
                         "votre", "son", "sa", "ses", "leur", "leurs", "aucune", "aucun", "autre",
                         "a", "an", "any", "each", "every", "our", "your", "its", "their"}
# Mots qui decrivent une forme sociale sans la nommer.
DESCRIPTIFS_FORME = {"civile", "professionnelle", "anonyme", "cooperative", "commerciale",
                     "immobiliere", "simplifiee", "unipersonnelle", "responsabilite", "limitee",
                     "actions", "par", "de", "d", "droit", "etranger", "etrangere", "exercice",
                     "liberal", "liberale", "en", "nom", "collectif", "commandite", "simple",
                     "interet", "economique", "europeenne", "nouvelle", "mere", "fille", "soeur",
                     "holding", "francaise", "anglaise", "russe", "americaine", "allemande"}


def _detecter_personnes_morales(texte):
    entites = []
    formes = "|".join(re.escape(f) for f in FORMES_SOCIALES)
    # « SA » et « SAS » devant le nom restent sensibles a la casse : avec
    # IGNORECASE, « sa Proposition Commerciale » devenait une societe (8 oct. 2026).
    formes_avant = "|".join(
        f"(?-i:{re.escape(f)})" if f in ("SA", "S.A.", "SAS", "S.A.S.") else re.escape(f)
        for f in FORMES_SOCIALES
        if f not in FORMES_POSTPOSEES_SEULES or f in ("SA", "S.A.", "SAS", "S.A.S."))

    # Forme anteposee : « la société ALPHA SAS », « SARL ALPHA »
    motif_a = re.compile(
        rf"(?<!\w)(?:(?:la|le|les)\s+)?(?:soci[ée]t[ée]\s+)?(?<!\w)({formes_avant})(?!\w)\s+"
        rf"(?-i:((?:(?:[{MAJ}][{MAJ}{MIN}0-9'’\-&]*|&)[^\S\n]?){{1,5}}))", re.IGNORECASE)
    # Forme postposee : « ALPHA SAS »
    # (?<=[^\S\n]) : la forme doit etre un mot a part ; sans cela « MIMOSA »
    # donnait « MIMO SA » et « NASA » « NA SA » (8 oct. 2026).
    motif_b = re.compile(
        rf"((?:(?:[{MAJ}][{MAJ}{MIN}0-9'’\-&]*|&)[^\S\n]?){{1,5}})[^\S\n]*(?<=[^\S\n])({formes})(?!\w)")
    # « la société ALPHA » sans forme sociale
    motif_c = re.compile(
        rf"\b[sS]oci[ée]t[ée]\s+((?:(?:[{MAJ}][{MAJ}{MIN}0-9'’\-&]*|&)[^\S\n]?){{1,4}})(?!\.\w)")

    vus = {}

    def ajouter(denomination, forme, position):
        denomination = _nettoyer_nom(denomination, genre="denomination")
        if not denomination or len(denomination) < 2:
            return
        # « Une association », « société Civile Professionnelle » : un article
        # ou une description de forme sociale, pas un nom (15 octobre 2026).
        nus = [_sans_accents(m).lower().strip(".,'’") for m in denomination.split()]
        premier = denomination.split()[0]
        if (nus[0] in DETERMINANTS_NON_NOMS and premier != premier.upper()) \
                or all(n in DESCRIPTIFS_FORME for n in nus):
            return
        if _sans_accents(denomination) in {_sans_accents(f) for f in FORMES_SOCIALES}:
            return
        # « la société HALDOR SAS » lue sans forme : la forme finale est
        # separee du nom, sinon « HALDOR SAS » et « HALDOR » faisaient deux cles.
        if not forme and len(denomination.split()) > 1:
            dernier = denomination.split()[-1]
            if dernier in FORMES_SOCIALES and dernier not in ("association", "fondation", "syndicat"):
                denomination, forme = denomination[: -len(dernier)].strip(), dernier
        cle = _sans_accents(denomination)
        # Meme nom, formes sociales differentes (« HALDOR SAS », filiale
        # francaise, et « HALDOR GmbH », maison mere allemande) : deux
        # societes distinctes, jamais fusionnees (4 novembre 2026). Le nom nu
        # (« HALDOR ») reste a la premiere rencontree.
        f_norm = re.sub(r"[.\s]", "", forme).upper() if forme else ""
        homonymes = vus.get(cle, [])
        e = next((x for x in homonymes
                  if not f_norm or not x.formes or f_norm in x.formes), None)
        if e is not None:
            if forme:
                e.ajouter_variante(f"{denomination} {forme}")
                e.formes.add(f_norm)
            if position < e.position:          # vue plus tot sous une autre forme
                e.position = position
                e.contexte = _contexte(texte, position)
            return
        canonique = f"{denomination} {forme}".strip() if forme else denomination
        e = EntiteDetectee(canonique, "personne_morale", position,
                           _contexte(texte, position),
                           confiance="moyenne" if _contient_conjonction(denomination)
                                     else "haute")
        e.formes = {f_norm} if f_norm else set()
        if homonymes:
            # seconde societe du meme nom : seulement les ecritures avec SA forme
            candidats = [f"{denomination} {forme}", f"société {denomination} {forme}",
                         f"la société {denomination} {forme}"]
        else:
            candidats = [denomination,
                         f"société {denomination}", f"la société {denomination}",
                         f"{denomination} {forme}".strip() if forme else ""]
            mots = denomination.split()
            if len(mots) > 1 and len(mots[0]) > 3 \
                    and _sans_accents(mots[0]).lower() not in MOTS_GENERIQUES_DENOMINATION:
                candidats.append(mots[0])          # forme abregee usuelle
        for candidate in candidats:
            if candidate and re.search(rf"\b{re.escape(candidate)}\b", texte, re.IGNORECASE):
                e.ajouter_variante(candidate)
        vus.setdefault(cle, []).append(e)
        entites.append(e)

    for m in motif_a.finditer(texte):
        ajouter(m.group(2), m.group(1), m.start())
    for m in motif_b.finditer(texte):
        ajouter(m.group(1), m.group(2), m.start())
    for m in motif_c.finditer(texte):
        ajouter(m.group(1), "", m.start())

    return entites


def _detecter_identifiants(texte):
    entites = []

    # SIRET (14 chiffres) et SIREN (9).
    #
    # La cle de Luhn ne sert PAS de filtre : un numero mal recopie dans un acte
    # reste une donnee a pseudonymiser, et le rejeter silencieusement serait la
    # pire des issues. Luhn releve la confiance, le contexte fait foi.
    CONTEXTE_ID = re.compile(r"siret|siren|immatricul|rcs|identifi|n°\s*unique",
                             re.IGNORECASE)

    # IBAN, NIR et date de naissance passent AVANT le SIRET : une suite de
    # chiffres d'un IBAN ou d'un NIR ne doit pas ressortir en SIRET.
    entites += _detecter_iban(texte)
    entites += _detecter_nir(texte)
    entites += _detecter_date_naissance(texte)

    for motif, type_, taille in (
        (re.compile(r"\b(\d{3}[\s.]?\d{3}[\s.]?\d{3}[\s.]?\d{5})\b"), "siret", 14),
        (re.compile(r"\b(\d{3}[\s.]?\d{3}[\s.]?\d{3})\b"), "siren", 9),
    ):
        for m in motif.finditer(texte):
            brut = m.group(1)
            chiffres = re.sub(r"\D", "", brut)
            if len(chiffres) != taille:
                continue
            if any(e.position <= m.start() < e.position + len(e.canonique) + 6
                   for e in entites):
                continue

            luhn_ok = _luhn(chiffres)
            avant = texte[max(0, m.start() - 90):m.start()]
            contexte_ok = bool(CONTEXTE_ID.search(avant))

            if not luhn_ok and not contexte_ok:
                continue        # suite de chiffres quelconque : on ne devine pas

            e = EntiteDetectee(brut.strip(), type_, m.start(),
                               _contexte(texte, m.start()),
                               confiance="haute" if luhn_ok else "à vérifier")
            if not luhn_ok:
                e.note = "clé de contrôle incorrecte (numéro erroné dans le document ?)"
            if chiffres != brut.strip():
                e.ajouter_variante(chiffres)
            entites.append(e)

    # Telephones
    for m in re.finditer(r"(?<!\w)(?:\+33\s?|0)[1-9](?:[\s.\-]?\d{2}){4}\b", texte):
        entites.append(EntiteDetectee(m.group(0).strip(), "telephone", m.start(),
                                      _contexte(texte, m.start())))

    # RCS
    for m in re.finditer(rf"RCS\s+(?:de\s+|d['’])?([{MAJ}][{MAJ}{MIN}\-]+)\s+(\d[\d \u00a0]{{8,}})", texte):
        entites.append(EntiteDetectee(m.group(0).strip(), "rcs", m.start(),
                                      _contexte(texte, m.start())))

    # References de procedure
    for m in re.finditer(r"(?:(?:n°\s*)?\bR\.?\s?G\.?|(?:Case|Claim)\s+(?:No\.?|Number))\s*:?\s*"
                         r"(\d{2}\s?/\s?\d{3,6}|\d{6,})",
                         texte, re.IGNORECASE):
        e = EntiteDetectee(m.group(1).strip(), "rg", m.start(1), _contexte(texte, m.start()))
        e.ajouter_variante(m.group(0).strip())
        entites.append(e)

    # E-mails
    for m in re.finditer(r"\b[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)*\.[a-zA-Z]{2,}\b", texte):
        entites.append(EntiteDetectee(m.group(0), "email", m.start(),
                                      _contexte(texte, m.start())))

    # Sites web
    for m in re.finditer(r"\b(?:https?://|www\.)[\w\-.]+\.[a-zA-Z]{2,}(?:/[\w\-./?%&=]*)?", texte):
        entites.append(EntiteDetectee(m.group(0).rstrip(".,;"), "site_web", m.start(),
                                      _contexte(texte, m.start())))

    return entites


# ------------------------------------------------------------------
# IBAN, NIR, DATE DE NAISSANCE
# ------------------------------------------------------------------
# Meme regle que le SIRET : la cle de controle ne sert PAS de filtre. Une cle
# fausse (numero mal recopie dans l'acte) abaisse la confiance ; le numero
# reste propose. Sans cle juste ET sans mot qui l'annonce, on ne devine pas.

MOTIF_IBAN = re.compile(r"\b([A-Z]{2}\d{2}(?:[ \u00a0]?[A-Z0-9]{4}){2,7}(?:[ \u00a0]?[A-Z0-9]{1,3})?)\b")
CONTEXTE_IBAN = re.compile(r"iban|rib\b|relev[ée] d.identit[ée] bancaire|compte|bank|account|virement|transfer",
                           re.IGNORECASE)


def _couper_iban(brut):
    """Garde le debut du texte trouve qui forme l'IBAN complet : longueur du
    pays si elle est connue, sinon la plus longue forme dont la cle est juste."""
    def prefixe(n):
        compte, i = 0, 0
        while i < len(brut) and compte < n:
            if brut[i].isalnum():
                compte += 1
            i += 1
        return brut[:i]

    compact = re.sub(r"[^A-Z0-9]", "", brut)
    attendue = LONGUEUR_IBAN.get(compact[:2])
    if attendue:
        return prefixe(attendue) if len(compact) >= attendue else brut
    for n in range(len(compact), 14, -1):
        if _cle_iban(compact[:n]):
            return prefixe(n)
    return brut


def _detecter_iban(texte):
    entites = []
    for m in MOTIF_IBAN.finditer(texte):
        brut = _couper_iban(m.group(1)).strip()
        compact = re.sub(r"[^A-Z0-9]", "", brut)
        if sum(c.isdigit() for c in compact) < 8:
            continue            # code en capitales, pas un numero de compte
        cle_ok = _cle_iban(compact)
        avant = texte[max(0, m.start() - 90):m.start()]
        if not cle_ok and not CONTEXTE_IBAN.search(avant):
            continue
        e = EntiteDetectee(brut, "iban", m.start(), _contexte(texte, m.start()),
                           confiance="haute" if cle_ok else "à vérifier")
        if not cle_ok:
            e.note = "clé de contrôle incorrecte (numéro erroné dans le document ?)"
        if compact != brut:
            e.ajouter_variante(compact)
        entites.append(e)
    return entites


MOTIF_NIR = re.compile(
    r"(?<![\w])([12][ \u00a0.]?\d{2}[ \u00a0.]?(?:0[1-9]|1[0-2])[ \u00a0.]?(?:\d{2}|2A|2B)"
    r"[ \u00a0.]?\d{3}[ \u00a0.]?\d{3})(?:[ \u00a0.]?(\d{2}))?(?![\w])", re.IGNORECASE)
CONTEXTE_NIR = re.compile(r"s[ée]curit[ée] sociale|\bnir\b|\binsee\b|n°\s*s\.?s\b|assur[ée]|"
                          r"inscription au r[ée]pertoire|social security", re.IGNORECASE)


def _detecter_nir(texte):
    entites = []
    for m in MOTIF_NIR.finditer(texte):
        brut = m.group(0).strip()
        treize = re.sub(r"[^0-9AB]", "", m.group(1).upper())
        cle = m.group(2)
        cle_ok = _cle_nir(treize, cle)
        avant = texte[max(0, m.start() - 90):m.start()]
        if not cle_ok and not CONTEXTE_NIR.search(avant):
            continue
        e = EntiteDetectee(brut, "nir", m.start(), _contexte(texte, m.start()),
                           confiance="haute" if cle_ok else "à vérifier")
        if cle and not cle_ok:
            e.note = "clé de contrôle incorrecte (numéro erroné dans le document ?)"
        elif not cle:
            e.note = "numéro sans sa clé : vérifiez qu'il s'agit bien d'un NIR"
        compact = re.sub(r"[^0-9AB]", "", brut.upper())
        if compact != brut:
            e.ajouter_variante(compact)
        entites.append(e)
    return entites


_MOIS_FR = (r"janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[ûu]t|septembre|octobre|"
            r"novembre|d[ée]cembre")
_MOIS_EN = (r"january|february|march|april|may|june|july|august|september|october|"
            r"november|december|jan\.?|feb\.?|mar\.?|apr\.?|jun\.?|jul\.?|aug\.?|sept?\.?|oct\.?|nov\.?|dec\.?")
_DATE = (r"(?:\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}"
         rf"|(?:1er|\d{{1,2}})(?:st|nd|rd|th)?\s+(?:{_MOIS_FR}|{_MOIS_EN})\s+\d{{4}}"
         rf"|(?:{_MOIS_EN})\s+\d{{1,2}}(?:st|nd|rd|th)?,?\s+\d{{4}}"
         r"|\d{4}-\d{2}-\d{2})")
# « né le », « née le », « né(e) le », « né à Lille le » ; « date de naissance : » ;
# en anglais « born on », « date of birth », « DOB ».
MOTIF_NAISSANCE = re.compile(
    r"(?:\bn[ée]e?(?:\s?\(e\))?(?:\s+(?:à|a|au)\s+[^\d\n,;]{1,40}?)?,?\s+le"
    r"|\bdate\s+de\s+naissance\s*:?(?:\s+le)?"
    r"|\bborn(?:\s+(?:in|at)\s+[^\d\n,;]{1,40}?)?,?(?:\s+on)?"
    r"|\bdate\s+of\s+birth\s*:?|\bDOB\s*:?)"
    rf"\s*({_DATE})", re.IGNORECASE)


def _detecter_date_naissance(texte):
    entites = []
    for m in MOTIF_NAISSANCE.finditer(texte):
        e = EntiteDetectee(m.group(1).strip(), "date_naissance", m.start(1),
                           _contexte(texte, m.start(1)))
        entites.append(e)
    return entites


def _elaguer_ville(adresse):
    """Retire les mots de procedure restes accroches a la fin de l'adresse.

    « 59000 LILLE CONTRE » sur une seule ligne reste possible ; ce mot ne fait
    pas partie de la ville et rendrait le terme introuvable dans le document.
    """
    mots = adresse.split()
    while len(mots) > 3 and _sans_accents(mots[-1]).lower().strip(":,.") in ARRET_VILLE:
        mots.pop()
    # Ville en capitales (« 75008 PARIS ») suivie d'un mot qui ne l'est pas
    # (« Tel », « Avocat ») : ce mot vient de la ligne suivante (4/11/2026).
    cp = next((i for i, m in enumerate(mots) if re.fullmatch(r"\d{5}", m.strip(","))), None)
    if cp is not None and cp + 1 < len(mots) and mots[cp + 1].isupper():
        fin = cp + 2
        while fin < len(mots) and mots[fin].isupper():
            fin += 1
        mots = mots[:fin]
    return " ".join(mots).strip(" ,;.:")


# « 3 rue du Rhin, CS 40012, 67000 Strasbourg » : boite postale ou course
# speciale entre la voie et le code postal (4 novembre 2026).
COMPLEMENT_POSTAL = r"(?:(?:CS|BP|TSA|B\.P\.)\s?\d{1,6}\s{0,3}[,\-–—]?\s{0,3})?"


def _entites_adresse(texte, m, confiance="haute", elaguer=True):
    """Une adresse ecrite sur une seule ligne donne une entite. Ecrite sur
    deux paragraphes (« 41 boulevard Malesherbes » puis « 75008 PARIS »), elle
    en donne une par ligne : l'anonymisation remplace paragraphe par
    paragraphe, et un terme a cheval sur deux paragraphes ne serait jamais
    trouve, l'adresse restant alors EN CLAIR (4 novembre 2026)."""
    brut = m.group(0)
    sortie = []
    decalage = 0
    # Coupure seulement AVANT le code postal (« voie » / « CP ville », la
    # presentation d'un en-tete) ; un retour juste apres le code postal
    # (« 93100 » / « MONTREUIL ») reste une seule adresse, comme avant.
    cp = re.search(r"\d{5}", brut)
    if cp is None or "\n" not in brut[:cp.start()]:
        propre = re.sub(r"\s+", " ", brut).strip(" ,;.")
        if elaguer:
            propre = _elaguer_ville(propre)
        return [EntiteDetectee(propre, "adresse", m.start(), _contexte(texte, m.start()),
                               confiance=confiance)]
    coupure = brut.rfind("\n", 0, cp.start())
    for morceau in (brut[:coupure], brut[coupure + 1:]):
        debut = m.start() + decalage
        decalage += len(morceau) + 1
        propre = re.sub(r"\s+", " ", morceau).strip(" ,;.-–—")
        if elaguer:
            propre = _elaguer_ville(propre)
        if len(propre) < 4 or not re.search(r"[A-Za-zÀ-ÿ]", propre):
            continue
        debut += max(0, morceau.find(propre[:6]))
        sortie.append(EntiteDetectee(propre, "adresse", debut, _contexte(texte, debut),
                                     confiance=confiance))
    return sortie


def _detecter_adresses(texte):
    entites = []
    voies = "|".join(TYPES_VOIE)

    # La ville se lit sur la MEME ligne que le code postal. Avec \s, la capture
    # sautait la fin de ligne et ramassait ce qui suivait : « 59000 LILLE
    # CONTRE », « 62000 ARRAS Demandeur Ayant po ». Le terme ainsi ecrit dans
    # la table ne correspondait plus au texte du document, donc l'anonymisation
    # ne pouvait pas le remplacer.
    motif = re.compile(
        rf"\b(\d{{1,4}}\s*(?:bis|ter|quater)?)\s*,?\s+({voies})\s+"
        rf"({LIBELLE_VOIE})\s{{0,3}}[,\-–—]?\s{{0,3}}{COMPLEMENT_POSTAL}"
        rf"(\d{{5}})[^\S\n]*\n?[^\S\n]*((?-i:{VILLE}))",
        re.IGNORECASE)

    for m in motif.finditer(texte):
        entites += _entites_adresse(texte, m)

    # Ville puis code postal entre parentheses : « 2 Allée des artisans, Lentilly (69210) »
    motif3 = re.compile(
        rf"\b(\d{{1,4}}\s*(?:bis|ter|quater)?)\s*,?\s+({voies})\s+({LIBELLE_VOIE})\s{{0,3}}[,\-–—]\s{{0,3}}"
        rf"((?-i:{VILLE}))\s*\((\d{{5}})\)", re.IGNORECASE)
    for m in motif3.finditer(texte):
        if any(e.position <= m.start() < e.position + len(e.canonique) for e in entites):
            continue
        entites += _entites_adresse(texte, m, elaguer=False)

    # Adresse sans numero de voie : « rue X, 59000 LILLE »
    motif2 = re.compile(
        rf"\b({voies})\s+({LIBELLE_VOIE})\s{{0,3}}[,\-–—]\s{{0,3}}"
        rf"(\d{{5}})[^\S\n]*\n?[^\S\n]*((?-i:{VILLE}))",
        re.IGNORECASE)
    for m in motif2.finditer(texte):
        if any(e.position <= m.start() < e.position + len(e.canonique) for e in entites):
            continue
        entites += _entites_adresse(texte, m, confiance="moyenne")

    return entites


def _qualite_la_plus_proche(apres, avant):
    """« avocat » ou « commissaire_justice » d'apres la qualite ecrite le plus
    pres du nom. Avant le 4 novembre 2026, la qualite de commissaire de
    justice l'emportait des qu'elle figurait dans la fenetre, meme loin
    derriere « Avocat au Barreau de ... » : un avocat plaidant etait alors
    classe commissaire de justice."""
    QUALITE_CJ = r"commissaires? de justice|huissiers?"
    QUALITE_AV = r"avocats?|barreau|conseil de|toque"
    cj = re.search(QUALITE_CJ, apres)
    av = re.search(QUALITE_AV, apres)
    if cj or av:
        if cj and (not av or cj.start() < av.start()):
            return "commissaire_justice"
        return "avocat"
    # Rien apres le nom : la qualite la plus proche AVANT lui
    cj = [m.end() for m in re.finditer(QUALITE_CJ, avant)]
    av = [m.end() for m in re.finditer(QUALITE_AV, avant)]
    if cj and (not av or cj[-1] > av[-1]):
        return "commissaire_justice"
    if av:
        return "avocat"
    return ""


def _detecter_professionnels(texte):
    """Avocats et commissaires de justice, reperes par leur titre."""
    entites = []
    # Un meme professionnel garde la meme qualite dans tout le document : la
    # premiere qualite lue pour un nom fait foi pour ses citations suivantes
    # (« Maitre DORVAL a dresse proces-verbal » n'en fait pas un avocat).
    qualite_de = {}

    for m in re.finditer(rf"\b(?:Ma[iî]tre|Me)\s+({NOM_PROPRE})", texte):
        nom = _nettoyer_nom(m.group(1))
        if not nom:
            continue

        # La fenetre s'arrete au confrere suivant et a la fin du paragraphe :
        # au-dela, la qualite lue est celle d'un autre.
        suivant = re.search(r"\b(?:Ma[iî]tre|Me)\s", texte[m.end():])
        fin = m.end() + (suivant.start() if suivant else 240)
        fin = min(fin, m.end() + 240)
        apres = _sans_accents(texte[m.end():fin])
        avant = _sans_accents(texte[max(0, m.start() - 160):m.start()])
        precedent = list(re.finditer(r"\b(?:ma[iî]tre|me)\s", avant))
        if precedent:
            avant = avant[precedent[-1].end():]

        patronyme = _sans_accents(nom.split()[-1]).lower()
        type_ = qualite_de.get(patronyme) or _qualite_la_plus_proche(apres, avant) or "avocat"
        qualite_de.setdefault(patronyme, type_)

        e = EntiteDetectee(f"Maître {nom}", type_, m.start(),
                           _contexte(texte, m.start()))
        e.ajouter_variante(nom)
        parties = nom.split()
        if len(parties) >= 2:
            for candidate in (f"Maître {parties[-1]}", f"Me {parties[-1]}"):
                if re.search(rf"\b{re.escape(candidate)}\b", texte):
                    e.ajouter_variante(candidate)
        entites.append(e)

    # Structures d'exercice nommees a cote du titre
    for m in re.finditer(
            rf"\b(SELARL|SELAS|SCP|AARPI|SELAFA)\s+"
            rf"((?:(?:[{MAJ}][{MAJ}{MIN}'’\-\.]*|&|ET)[^\S\n]?){{1,6}})",
            texte):
        avant = _sans_accents(texte[max(0, m.start() - 200):m.start()])
        apres = _sans_accents(texte[m.end():m.end() + 200])
        type_ = _qualite_la_plus_proche(apres, avant)
        if not type_:
            continue
        nom = _nettoyer_nom(m.group(2), genre="denomination")
        if not nom:
            continue
        e = EntiteDetectee(f"{m.group(1)} {nom}", type_, m.start(),
                           _contexte(texte, m.start()))
        e.ajouter_variante(nom)
        entites.append(e)

    # Cabinets sans forme sociale : « CABINET LESAGE AVOCATS », « VIGNAC
    # AVOCATS » (4 novembre 2026), en capitales, dans un contexte d'avocat.
    for m in re.finditer(
            rf"\b(?:CABINET[^\S\n]+((?:[{MAJ}][{MAJ}'’\-&]+[^\S\n]?){{1,4}})"
            rf"|((?:[{MAJ}][{MAJ}'’\-&]+[^\S\n]){{1,3}})AVOCATS\b)", texte):
        brut = (m.group(1) or m.group(2) or "").strip()
        mots = [x for x in brut.split() if _sans_accents(x).lower() not in ("avocats", "avocat")]
        if not mots or any(_sans_accents(x).lower() in ("societe", "le", "la", "les") for x in mots[:1]):
            continue
        entour = _sans_accents(texte[max(0, m.start() - 200):m.end() + 200])
        if not re.search(r"avocat|barreau|toque", entour):
            continue
        canon = re.sub(r"\s+", " ", m.group(0)).strip()
        e = EntiteDetectee(canon, "avocat", m.start(), _contexte(texte, m.start()),
                           confiance="moyenne")
        entites.append(e)

    # Toque et case du Palais : identifient l'avocat (4 novembre 2026)
    for m in re.finditer(
            r"\b((?:[Tt]oque|TOQUE|[Pp]alais|PALAIS|[Vv]estiaire|VESTIAIRE)\s*(?:n°|N°|no)?\s*:?\s*"
            r"[A-Z]?\s?\d{1,5})\b", texte):
        entites.append(EntiteDetectee(re.sub(r"\s+", " ", m.group(1)).strip(), "toque",
                                      m.start(), _contexte(texte, m.start())))

    return entites


def _detecter_titres_et_signes(texte):
    """Enseignes, marques, brevets, dessins et modeles, depots."""
    entites = []

    for m in re.finditer(
            rf"(?:sous\s+l['’]enseigne|[àa]\s+l['’]enseigne|enseigne)\s*«?\s*"
            rf"([{MAJ}][{MAJ}{MIN}0-9'’\-\s&]{{1,40}})", texte, re.IGNORECASE):
        nom = _nettoyer_nom(m.group(1), genre="denomination")
        if nom:
            entites.append(EntiteDetectee(nom, "enseigne", m.start(),
                                          _contexte(texte, m.start())))

    for m in re.finditer(r"marques?\s+(?:[\w'’\s]{0,30}?)«\s*([^»]{1,50})\s*»",
                         texte, re.IGNORECASE):
        entites.append(EntiteDetectee(m.group(1).strip(), "marque", m.start(),
                                      _contexte(texte, m.start())))
    # Documents en anglais : trademark "X" / “X”
    for m in re.finditer(r"trade\s?marks?\s+(?:[\w'’\s]{0,30}?)[\"“]\s*([^\"”]{1,50})\s*[\"”]",
                         texte, re.IGNORECASE):
        entites.append(EntiteDetectee(m.group(1).strip(), "marque", m.start(),
                                      _contexte(texte, m.start())))

    for m in re.finditer(r"\b((?:EP|FR|WO|US)\s?\d[\d \u00a0]{5,})", texte):
        # « TVA FR 12 345 678 901 » n'est pas un brevet.
        if re.search(r"(?:TVA|VAT)\b[^\n\d]{0,25}$", texte[max(0, m.start() - 40):m.start()], re.IGNORECASE):
            continue
        entites.append(EntiteDetectee(m.group(1).strip(), "brevet", m.start(),
                                      _contexte(texte, m.start())))

    for m in re.finditer(r"(?:dessins?\s+et\s+mod[èe]les?|mod[èe]le)\s+n°\s*([\w\s\-]{3,20})",
                         texte, re.IGNORECASE):
        entites.append(EntiteDetectee(m.group(0).strip(), "dessin_modele", m.start(),
                                      _contexte(texte, m.start()), confiance="moyenne"))

    for m in re.finditer(r"(?:INPI|dépôt|depot)[^.\n]{0,40}?n°\s*(\d{6,10})",
                         texte, re.IGNORECASE):
        entites.append(EntiteDetectee(m.group(1).strip(), "depot_inpi", m.start(1),
                                      _contexte(texte, m.start()), confiance="moyenne"))

    return entites


# ==================================================================
# ACTEURS NOMMES PAR LEUR FONCTION, PRODUITS, CANDIDATS
# ==================================================================

# Mots capitalises qui ne sont jamais des noms propres a anonymiser.
MOTS_COURANTS_CAPITALISES = {
    # Termes juridiques anglais courants (version anglaise)
    "Claimant", "Claimants", "Defendant", "Defendants", "Applicant", "Respondent", "Appellant",
    "Court", "Judge", "Tribunal", "Counsel", "Particulars", "Claim", "Defence", "Reply",
    "Agreement", "Clause", "Schedule", "Exhibit", "Witness", "Statement", "Order", "Section",
    "Provider", "Customer", "Supplier", "Licensor", "Licensee", "Party", "Parties", "Company",
    "January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
    "The", "This", "That", "These", "Those", "Dated", "Signed",
    "Tribunal", "Cour", "Conseil", "Président", "Présidente", "Juge", "Monsieur", "Madame",
    "Maître", "Pièce", "Pièces", "Article", "Articles", "Code", "Chambre", "Ordonnance",
    "Jugement", "Arrêt", "Requête", "Assignation", "Conclusions", "Audience", "Greffe",
    "Procès", "Procès-verbal", "Brevet", "Marque", "Partie", "Parties", "Société", "Groupe",
    "France", "Paris", "Europe", "Européen", "Européenne", "Union", "Office", "Institut",
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre",
    "Octobre", "Novembre", "Décembre", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi",
    "Samedi", "Dimanche", "Dans", "Pour", "Contre", "Par", "Sur", "Selon", "Enfin", "Ainsi",
    "Cette", "Cela", "Ces", "Cet", "Elle", "Elles", "Ils", "Nous", "Vous", "Les", "Des",
    "Une", "Aucun", "Aucune", "Toute", "Toutes", "Tout", "Tous", "Lorsque", "Puis", "Or",
    "Mais", "Donc", "Car", "Attendu", "Considérant", "Vu", "Plaise", "Dire", "Juger",
    "Condamner", "Débouter", "Ordonner", "Constater", "Déclarer", "Fixer", "Rejeter",
    "Rappel", "Discussion", "Faits", "Procédure", "Moyens", "Dispositif", "Motifs",
    "Monsieur", "Mesdames", "Messieurs", "Avocat", "Avocats", "Barreau", "Cabinet",
    "Direction", "Service", "Réf", "Ref", "Fait", "Signifiées", "Objet", "Annexe", "Kbis",
    "Extrait", "Catalogue", "Courrier", "Protocole", "Protocole", "Certificat", "Copie",
    "Relevé", "Justificatif", "Fascicule", "Notice", "Avis", "Rapport", "Expert", "Experts",
    "Internet", "Web", "Google", "Microsoft", "Word", "Excel", "PDF", "Windows",
    "Commission", "Règlement", "Directive", "Décret", "Loi", "Groupe", "Groupes", "État",
    "États", "Technique", "Techniques", "Spécialisé", "Spécialisés", "Spécialisée",
    "Signification", "Extraits", "Extrait", "Cass", "Civ", "Com", "Soc", "Crim", "Bull",
    "Juris", "Dalloz", "Lexis", "Jurisprudence", "Doctrine", "Convention", "Accord",
    "Contrat", "Licence", "Cahier", "Charges", "Annexes", "Préambule", "Définitions",
    "Général", "Générale", "Générales", "Particulières", "Directeur", "Directrice",
    "Gérant", "Gérante", "Responsable", "Délégué", "Déléguée", "Nationale", "National",
    "Régional", "Régionale", "Public", "Publique", "Privé", "Privée", "Industrie",
    "Industriel", "Industrielle", "Commerce", "Commercial", "Commerciale", "Judiciaire",
    "Judiciaires", "Civil", "Civile", "Pénal", "Pénale", "Administratif", "Administrative",
    "Ayant", "Ci-après", "Requérante", "Requérant", "Demanderesse", "Défenderesse",
    "Demandeur", "Défendeur", "Appelante", "Appelant", "Intimée", "Intimé",
    # 15 octobre 2026 : pluriels de procedure, qualites, mots d'adresse
    "Demandeurs", "Demanderesses", "Défendeurs", "Défenderesses", "Requérants", "Requérantes",
    "Appelants", "Appelantes", "Intimés", "Intimées", "Intervenant", "Intervenante",
    "Intervenants", "Intervenantes", "Titulaire", "Titulaires", "Concluante", "Concluantes",
    "Concluant", "Concluants", "Vice-président", "Vice-présidente", "Bâtiment", "Immeuble",
    "Résidence", "Lieudit", "Lieu-dit", "Cedex", "Étage", "Président-directeur",
    # 4 novembre 2026 : bruit releve sur des conclusions reelles
    "Registre", "Registres", "Sociétés", "Société", "Document", "Documents", "Sans",
    "Vice-Présidente", "Vice-Président", "Vice-présidente", "Commerce", "Sociétés.",
    "Inventeur", "Laboratoire", "Saisie", "Saisie-contrefaçon", "Contrefaçon",
    "Rétractation", "Référé", "Requérante", "Toque", "Palais", "Tel", "Tél", "Fax",
}
# Ville citee seulement comme siege d'une juridiction (« cour d'appel de
# Versailles ») : de la jurisprudence, pas une donnee du dossier.
JURIDICTION_AVANT = re.compile(
    r"(?:COUR\s+D'APPEL|TRIBUNAL(?:\s+(?:JUDICIAIRE|DE\s+COMMERCE|DE\s+GRANDE\s+INSTANCE|"
    r"ADMINISTRATIF|D'INSTANCE|CORRECTIONNEL))?|CONSEIL\s+DE\s+PRUD'HOMMES|\bCA|\bTJ|\bTGI|"
    r"\bT\.?\s?COM\.?|\bTC|\bCAA|\bTA)\s*(?:DE\s+|D'|,\s*)?$")


def _candidats_sans_marqueur(texte, deja):
    """Mots capitalises au milieu d'une phrase, repetes, inconnus des autres
    detecteurs : noms de produits, de gammes, de personnes sans civilite...
    Proposes « a classer », a NON, pour que rien ne passe inapercu."""
    entites = []
    pris = set()
    for e in deja:
        for v in e.variantes:
            for mot in re.findall(rf"[{MAJ}][{MAJ}{MIN}0-9'’\-]{{2,}}", v):
                pris.add(_sans_accents(mot).lower())
    compte = {}
    for m in re.finditer(rf"(?<=[{MIN}] )([{MAJ}][{MIN}][{MAJ}{MIN}0-9\-]{{2,}})\b", texte):
        mot = m.group(1)
        if mot in MOTS_COURANTS_CAPITALISES or _sans_accents(mot).lower() in pris:
            continue
        if re.match(r"^[A-Z][a-z]+$", mot) and len(mot) < 4:
            continue
        compte.setdefault(mot, []).append(m.start(1))
    for mot, positions in compte.items():
        if len(positions) < 2:
            continue
        if all(JURIDICTION_AVANT.search(_sans_accents(texte[max(0, p - 45):p]).upper())
               for p in positions):
            continue
        e = EntiteDetectee(mot, "candidat", positions[0], _contexte(texte, positions[0]),
                           confiance="basse")
        e.occurrences = len(positions)
        entites.append(e)
    return entites


# Mots en capitales qui ne sont pas des patronymes (« Lyon CEDEX »)
CAPITALES_NON_PATRONYMES = {
    "CEDEX", "SAS", "SA", "SARL", "SASU", "EURL", "SCI", "SNC", "SELARL", "SELAS", "SCP",
    "RCS", "RG", "TJ", "TGI", "CA", "INPI", "EUIPO", "OEB", "PDF", "TVA", "SIREN", "SIRET",
    "FRANCE", "PARIS", "LYON", "LILLE", "TEL", "FAX", "GMBH", "LTD", "INC", "LLC", "AG",
    "CONTRE", "POUR", "ENTRE", "ET", "DE", "DU", "DES", "LE", "LA", "LES", "PAR", "SUR",
}


def _detecter_prenom_nom(texte):
    """« Inès FAURIEL » sans civilite : prenom puis patronyme en capitales,
    l'usage des ecritures. Propose en confiance moyenne (4 novembre 2026) ;
    auparavant, un tel nom restait un « candidat » a NON."""
    entites = []
    for m in re.finditer(
            rf"(?<![\w’'])([{MAJ}][{MIN}]{{2,}}(?:-[{MAJ}][{MIN}]{{2,}})?)[^\S\n]+"
            rf"((?:[{MAJ}]{{2,}}['’\-]?)+(?:[^\S\n]+[{MAJ}]{{2,}})?)(?![\w])", texte):
        prenom, nom = m.group(1), m.group(2)
        if prenom in MOTS_COURANTS_CAPITALISES or _sans_accents(prenom).lower() in ARRET \
                or _sans_accents(prenom).lower() in OUVREURS_DE_PHRASE:
            continue
        mots_nom = nom.split()
        if any(_sans_accents(x).upper().strip("'’-") in CAPITALES_NON_PATRONYMES for x in mots_nom):
            continue
        if any(_sans_accents(x).upper() in {_sans_accents(c).upper() for c in MOTS_COURANTS_CAPITALISES}
               for x in mots_nom):
            continue
        e = EntiteDetectee(f"{prenom} {nom}", "personne_physique", m.start(),
                           _contexte(texte, m.start()), confiance="moyenne")
        if re.search(rf"(?<!\w){re.escape(nom)}(?!\w)", texte[m.end():]):
            e.ajouter_variante(nom)
        entites.append(e)
    return entites


def _detecter_fonctions_et_produits(texte):
    """Commissaires de justice et experts nommes sans « Maitre » ; produits,
    gammes, systemes et logiciels nommes apres leur mot generique."""
    entites = []

    # « les commissaires de justice Bédon et Grain », « l'huissier Durand »
    for m in re.finditer(
            rf"(?:commissaires?\s+de\s+justice|huissiers?(?:\s+de\s+justice)?)\s+"
            rf"((?:[{MAJ}][{MAJ}{MIN}'’\-]{{2,}})(?:\s*(?:,|et)\s*[{MAJ}][{MAJ}{MIN}'’\-]{{2,}})*)",
            texte):
        for nom in re.split(r"\s*(?:,|\bet\b)\s*", m.group(1)):
            nom = nom.strip()
            if nom and nom not in MOTS_COURANTS_CAPITALISES:
                pos = texte.find(nom, m.start(1))
                entites.append(EntiteDetectee(nom, "commissaire_justice", pos if pos >= 0 else m.start(1),
                                              _contexte(texte, m.start()), confiance="moyenne"))

    # « l'expert Clayton », « l'expert judiciaire M. Durand »
    for m in re.finditer(
            rf"\bexperts?(?:\s+judiciaire)?(?:\s+en\s+[{MIN}\-]+)?\s+(?:M\.|Mme|Monsieur|Madame)?\s*"
            rf"([{MAJ}][{MAJ}{MIN}'’\-]{{2,}})\b", texte):
        nom = m.group(1)
        if nom not in MOTS_COURANTS_CAPITALISES and not re.match(r"^(Le|La|Les|Du|De|Des|En|Et)$", nom):
            entites.append(EntiteDetectee(nom, "expert", m.start(1), _contexte(texte, m.start()),
                                          confiance="moyenne"))

    # « le système Coqisol », « les gammes Sileo, Dixio, Bialto et Axxio », « le logiciel Alpha »
    generiques = (r"produits?|syst[èe]mes?|gammes?|[ée]quipements?|logiciels?|solutions?|"
                  r"applications?|plateformes?|mod[èe]les?|r[ée]f[ée]rences?|d[ée]nominations?")
    for m in re.finditer(
            rf"\b(?:{generiques})\s+(?:«\s*)?(?:syst[èe]me\s+)?"
            rf"([{MAJ}][{MAJ}{MIN}0-9\-]{{2,}}(?:\s*,\s*[{MAJ}][{MAJ}{MIN}0-9\-]{{2,}})*(?:\s+et\s+[{MAJ}][{MAJ}{MIN}0-9\-]{{2,}})?)",
            texte):
        for nom in re.split(r"\s*(?:,|\bet\b)\s*", m.group(1)):
            nom = nom.strip(" »")
            if nom and nom not in MOTS_COURANTS_CAPITALISES and not nom.isupper():
                pos = texte.find(nom, m.start(1))
                entites.append(EntiteDetectee(nom, "produit", pos if pos >= 0 else m.start(1),
                                              _contexte(texte, m.start()), confiance="moyenne"))
    return entites


# ==================================================================
# ORCHESTRATION
# ==================================================================

# ==================================================================
# COMMENTAIRES WORD (4 novembre 2026)
# ==================================================================
# Les commentaires sont lus et anonymises comme le reste. Mais ce sont des
# notes de relecture : mots tronques (« HALDOR TECH… »), coquilles
# (« Fauriell », « Kalbch »). Une ecriture approchee d'une identite deja
# detectee dans le corps devient une VARIANTE de cette identite (elle sera
# donc masquee, sous la meme balise) ; elle ne cree jamais une societe ou une
# personne de plus.

SEUIL_COQUILLE = 0.8
_CIVILITE_TETE = re.compile(
    r"^(?:Monsieur|Madame|Mademoiselle|M\.|Mme|Mle|Me|Ma[iî]tre|Mr|Mrs|Ms|Dr)\s+", re.IGNORECASE)


def _dans_zones(position, zones):
    return any(a <= position < b for a, b in zones)


def _mots_distinctifs(entite):
    """Mots d'une identite qui la designent (patronyme, nom de societe)."""
    mots = set()
    for v in entite.variantes:
        for mot in re.findall(rf"[{MAJ}{MIN}][{MAJ}{MIN}'’\-]{{3,}}", _CIVILITE_TETE.sub("", v)):
            nu = _sans_accents(mot).lower()
            if mot in MOTS_COURANTS_CAPITALISES or nu in MOTS_GENERIQUES_DENOMINATION \
                    or nu in ARRET or nu in {_sans_accents(f).lower() for f in FORMES_SOCIALES}:
                continue
            if mot[:1].isupper():
                mots.add(mot)
    return mots


def _ressemble(a, b):
    """Coquille ou troncature de b ? (comparaison sans casse ni accents)"""
    import difflib
    a, b = _sans_accents(a).lower(), _sans_accents(b).lower()
    if a == b or len(a) < 4:
        return a == b
    if b.startswith(a) and len(a) >= 4:                 # tronque
        return True
    return difflib.SequenceMatcher(None, a, b).ratio() >= SEUIL_COQUILLE


def _proche_connue(mots, connues):
    """Identite connue dont un mot distinctif ressemble a l'un des mots."""
    for x in connues:
        for cible in _mots_distinctifs(x):
            if any(_ressemble(m, cible) for m in mots):
                return x
    return None


def _rattacher_commentaires(texte, finales, zones):
    hors_zone = lambda e: any(not _dans_zones(m.start(), zones)
                              for v in e.variantes
                              for m in re.finditer(rf"(?<!\w){re.escape(v)}(?!\w)", texte))
    connues = [e for e in finales if hors_zone(e)]
    sortie = []
    for e in finales:
        if e in connues or e.type not in ("personne_physique", "representant_legal",
                                          "personne_morale", "avocat", "commissaire_justice",
                                          "enseigne", "expert", "produit"):
            sortie.append(e)
            continue
        mots = re.findall(rf"[{MAJ}{MIN}][{MAJ}{MIN}'’\-]{{3,}}", _CIVILITE_TETE.sub("", e.canonique))
        cible = _proche_connue(mots, connues)
        if cible is not None:
            for v in e.variantes:
                if len(_CIVILITE_TETE.sub("", v)) >= 4:
                    cible.ajouter_variante(v)
            cible.note = (cible.note + " ; écriture approchée relevée dans un commentaire : "
                          + e.canonique).strip(" ;")
            continue
        fin = e.position + len(e.canonique)
        if re.match(r"\s*(?:…|\.\.\.)", texte[fin:fin + 4]):
            continue                                    # mot tronque, sans identite connue
        sortie.append(e)

    # Mots isoles des commentaires (« Kalbch », sans civilite) : rattaches
    # s'ils ressemblent a un nom connu, sans jamais creer d'entite.
    deja = {_sans_accents(v).lower() for e in sortie for v in e.variantes}
    for a, b in zones:
        for m in re.finditer(rf"(?<![\w])([{MAJ}][{MAJ}{MIN}'’\-]{{3,}})(?![\w])", texte[a:b]):
            mot = m.group(1)
            if mot in MOTS_COURANTS_CAPITALISES or _sans_accents(mot).lower() in deja:
                continue
            suite = texte[a + m.end():a + m.end() + 3]
            cible = _proche_connue([mot], connues)
            if cible is None:
                continue
            # Mot tronque (« TECH… ») : le reste du nom porte deja la balise ;
            # un fragment ajoute seul masquerait ce mot partout ailleurs.
            if suite.startswith(("…", "...")):
                continue
            cible.ajouter_variante(mot)
            deja.add(_sans_accents(mot).lower())
            cible.note = (cible.note + f" ; écriture approchée relevée dans un commentaire : {mot}").strip(" ;")
    return sortie


def detecter(texte, source="texte", pages_ocr=None):
    """Retourne les entites, dedoublonnees, ordonnees par premiere apparition.

    source="ocr" degrade la confiance de TOUTES les entites : le texte lu par
    un OCR est approche, et un terme mal reconnu recopie dans la table ne
    correspondra pas au texte reel du document a anonymiser.
    """
    entites = []
    entites += _detecter_professionnels(texte)     # avant les personnes : « Maître X »
    entites += _detecter_personnes(texte)
    entites += _detecter_personnes_morales(texte)
    entites += _detecter_identifiants(texte)
    entites += _detecter_adresses(texte)
    entites += _detecter_titres_et_signes(texte)
    entites += _detecter_fonctions_et_produits(texte)
    sans_civilite = _detecter_prenom_nom(texte)

    # Un terme deja capte par un detecteur prioritaire n'est pas repris.
    entites.sort(key=lambda e: (e.position, -len(e.canonique)))
    retenues, pris = [], []
    for e in entites:
        fin = e.position + len(e.canonique)
        if any(d <= e.position and fin <= f for d, f in pris):
            continue
        pris.append((e.position, fin))
        retenues.append(e)

    # « Prenom NOM » sans civilite : seulement si ce nom n'est pas deja connu
    # sous une autre forme (« Maitre Gontran LESAGE » signe « Gontran LESAGE »).
    connues = {_sans_accents(v) for e in retenues for v in e.variantes}
    for e in sans_civilite:
        fin = e.position + len(e.canonique)
        if _sans_accents(e.canonique) in connues \
                or any(d <= e.position < f or d < fin <= f for d, f in pris):
            continue
        # « Gontran LESAGE » en signature, « Maitre LESAGE » plus haut : meme avocat
        patronyme = _sans_accents(e.canonique.split()[-1])
        confrere = next((x for x in retenues if x.type in ("avocat", "commissaire_justice")
                         and patronyme in {_sans_accents(v.split()[-1]) for v in x.variantes}), None)
        if confrere is not None:
            confrere.ajouter_variante(e.canonique)
            continue
        pris.append((e.position, fin))
        retenues.append(e)
    retenues.sort(key=lambda e: e.position)

    # Fusion des doublons : meme forme canonique detectee deux fois.
    par_cle = {}
    finales = []
    for e in retenues:
        cle = (_sans_accents(e.canonique), e.type)
        if cle in par_cle:
            for v in e.variantes:
                par_cle[cle].ajouter_variante(v)
            continue
        par_cle[cle] = e
        finales.append(e)

    # --- Une structure d'exercice n'est pas deux entites ---------------------
    # « SELARL TROMPETTE ET CHERUBIN » sortait deux fois : une fois comme
    # structure d'avocats ou de commissaires, une fois comme personne morale,
    # avec l'ordre des mots inverse. C'est le meme cabinet : on garde la
    # qualification professionnelle, qui est la plus informative.
    def _mots_significatifs(nom):
        formes = {_sans_accents(f).lower().strip(".") for f in FORMES_SOCIALES}
        mots = {_sans_accents(x).lower().strip(".,;") for x in nom.split()}
        return {m for m in mots if m and m not in formes and m not in ("&", "et")}

    professionnels = [e for e in finales
                      if e.type in ("avocat", "commissaire_justice")]
    sans_doublon = []
    for e in finales:
        if e.type == "personne_morale":
            mots = _mots_significatifs(e.canonique)
            jumeau = next((p for p in professionnels
                           if mots and _mots_significatifs(p.canonique) == mots), None)
            if jumeau is not None:
                for v in e.variantes:
                    jumeau.ajouter_variante(v)
                continue
        sans_doublon.append(e)
    finales = sans_doublon

    # --- Rattachement des occurrences ulterieures a la premiere apparition ---
    # « Monsieur DURAND » plus loin dans l'acte n'est pas une nouvelle personne :
    # c'est la meme, deja rencontree sous une forme plus complete.
    GROUPES = [
        {"personne_physique", "representant_legal"},
        {"personne_morale", "enseigne"},
    ]

    def _compatibles(t1, t2):
        if t1 == t2:
            return True
        return any(t1 in g and t2 in g for g in GROUPES)

    PERSONNES = {"personne_physique", "representant_legal"}
    MOTIF_CIVILITE_TETE = re.compile(
        r"^(?:Monsieur|Madame|Mademoiselle|MONSIEUR|MADAME|M\.|Mme|Mle|MM\.|Mmes"
        r"|Me|Maître|Maitre|Dr|Pr|Mr|Mrs|Ms|Miss)\s+", re.IGNORECASE)

    def _personnes_distinctes(a, b):
        """« Monsieur Jean DURAND » et « Madame Marie DURAND » partagent le
        patronyme mais pas le prenom : deux personnes, jamais fusionnees
        (8 octobre 2026). Compare les prenoms quand les deux noms en ont un."""
        if a.type not in PERSONNES or b.type not in PERSONNES:
            return False
        if _genre_civilite(a.canonique) and _genre_civilite(b.canonique) \
                and _genre_civilite(a.canonique) != _genre_civilite(b.canonique):
            return True
        mots_a = MOTIF_CIVILITE_TETE.sub("", a.canonique).split()
        mots_b = MOTIF_CIVILITE_TETE.sub("", b.canonique).split()
        if len(mots_a) < 2 or len(mots_b) < 2:
            return False
        return _sans_accents(mots_a[0]).lower() != _sans_accents(mots_b[0]).lower()

    fusionnees = []
    for e in finales:
        absorbe = None
        for deja in fusionnees:
            if deja.position >= e.position or not _compatibles(deja.type, e.type):
                continue
            if _personnes_distinctes(deja, e):
                continue
            formes_deja = {_sans_accents(v) for v in deja.variantes}
            if _sans_accents(e.canonique) in formes_deja or any(
                    _sans_accents(v) in formes_deja for v in e.variantes):
                absorbe = deja
                break
            # « LES ATELIERS X » et « ATELIERS X » : meme societe, article en plus
            if deja.type in ("personne_morale", "enseigne") and e.type in ("personne_morale", "enseigne"):
                noyau = lambda t: re.sub(r"^(?:LES|LE|LA|L')\s*", "", _sans_accents(t).upper()).strip()
                if noyau(e.canonique) and noyau(e.canonique) in {noyau(v) for v in deja.variantes}:
                    absorbe = deja
                    break
        if absorbe is not None:
            for v in e.variantes:
                absorbe.ajouter_variante(v)
            continue
        fusionnees.append(e)
    finales = fusionnees

    # Commentaires Word : un mot tronque ou une coquille ne cree jamais
    # d'identite nouvelle, il rejoint l'identite deja connue (4/11/2026).
    zones = getattr(texte, "zones_commentaires", ())
    if zones:
        finales = _rattacher_commentaires(texte, finales, zones)

    # Occurrences reelles et detection d'homonymie.
    proprietaire = {}
    for e in finales:
        e.occurrences = sum(len(re.findall(rf"(?<!\w){re.escape(v)}(?!\w)", texte))
                            for v in e.variantes)
        for v in e.variantes:
            cle = _sans_accents(v)
            if cle in proprietaire and proprietaire[cle] is not e:
                autre = proprietaire[cle]
                for x in (e, autre):
                    x.ambigu = True
                    x.confiance = "à vérifier"
                    note = f"forme « {v} » partagée avec « {autre.canonique if x is e else e.canonique} »"
                    x.note = (x.note + " ; " + note).strip(" ;")
            else:
                proprietaire[cle] = e

    # Candidats sans marqueur (produits, personnes sans civilite...), ajoutes en
    # dernier : tout ce qui est deja reconnu les exclut.
    for c in _candidats_sans_marqueur(texte, finales):
        c.note = "mot répété, non reconnu : produit, personne, lieu ? à vous de dire"
        finales.append(c)

    for e in finales:
        if _sur_page_ocr(texte, e.position, source, pages_ocr):
            e.confiance = "à vérifier"
            e.ocr = True
            note = "issu d'un OCR : vérifiez l'orthographe exacte avant injection"
            e.note = (e.note + " ; " + note).strip(" ;")

    finales.sort(key=lambda e: e.position)
    return finales


# ==================================================================
# ATTRIBUTION DES BALISES, D'APRES LE CLASSEUR
# ==================================================================

def plages_pages(pages):
    """[3, 5, 6, 7] -> « 3, 5-7 » (pour les messages)."""
    pages = sorted(set(pages or []))
    plages, debut = [], None
    for i, n in enumerate(pages):
        if debut is None:
            debut = n
        if i + 1 == len(pages) or pages[i + 1] != n + 1:
            plages.append(str(debut) if debut == n else f"{debut}-{n}")
            debut = None
    return ", ".join(plages)


def _sur_page_ocr(texte, position, source, pages_ocr):
    """Vrai si la position tombe sur une page lue par OCR. Un document « ocr »
    l'est en entier ; un document « mixte » seulement sur ses pages_ocr."""
    if source == "ocr":
        return True
    if source == "mixte" and pages_ocr:
        return numero_page(texte, position) in set(pages_ocr)
    return False


def charger_config_detection(chemin):
    """Lien type de detection -> libelle de categorie du classeur.
    Ce fichier est modifiable : ajouter une categorie dans le tableur ne
    demande jamais de toucher au code."""
    if not os.path.isfile(chemin):
        raise RuntimeError(f"detection_config.json introuvable :\n{chemin}")
    with open(chemin, encoding="utf-8") as f:
        return json.load(f)


def analyser_document(chemin_document, chemin_excel=None, nom_onglet=None,
                      chemin_config=None, autoriser_ocr=True, progression=None,
                      type_doc=None, sens=None, demander_sens=None):
    """Chaine complete : texte (ou OCR) -> entites -> parties et roles -> balises.

    type_doc      : type de document force (code de detection_config.json),
                    ou None / « automatique » pour le deduire du texte ;
    sens          : pour un courrier, « emis » ou « recu » (None = automatique) ;
    demander_sens : fonction appelee si le sens d'un courrier reste inconnu.

    Si l'onglet vise suit le modele par partie, ses termes servent a garder
    les parties deja connues et a ne pas proposer deux fois le meme terme.

    Retourne (entites, infos, texte, source).
    """
    import detection_roles as roles
    import modele_onglet as modele

    config = charger_config_detection(chemin_config) if chemin_config else {}
    texte, source, pages_ocr = extraire_texte_detail(
        chemin_document, config.get("ocr"), autoriser_ocr, progression)
    entites = detecter(texte, source, pages_ocr)

    type_eff = type_doc
    if not type_doc or type_doc == roles.AUTOMATIQUE:
        type_eff, _ = roles.detecter_type(texte, config)
    if type_eff == "audit_rgpd":
        entites = sorted(entites + roles.detecter_outils(texte, config, entites),
                         key=lambda e: e.position)
        for e in entites:
            if _sur_page_ocr(texte, e.position, source, pages_ocr):
                e.confiance = "à vérifier"
                e.ocr = True

    existantes, modele_cible = [], ""
    if chemin_excel and nom_onglet:
        modele_cible, existantes = modele.lire_onglet(chemin_excel, nom_onglet)
        if modele_cible != "par_partie":
            existantes = []

    infos = roles.attribuer_roles(texte, entites, config, type_doc, sens,
                                  existantes, demander_sens)
    infos["modele_onglet"] = modele_cible
    infos["pages_ocr"] = pages_ocr
    numeroter_pages(texte, entites)
    return entites, infos, texte, source


def numeroter_pages(texte, entites):
    """Renseigne e.page quand la pagination du texte est connue."""
    if not pages_connues(texte):
        return
    for e in entites:
        e.page = numero_page(texte, e.position)


# ==================================================================
# OCR - LES PDF SCANNES
# ==================================================================
#
# Un PDF scanne ne contient pas de texte, seulement des images de pages.
# Tesseract sait en extraire du texte, a condition de disposer du pack de
# langue francais (fra.traineddata), qui n'est pas installe par defaut.
#
# Choix d'implementation : tesseract est appele EN LIGNE DE COMMANDE, pas via
# une bibliotheque Python. C'est le plus sur, puisque le programme est deja
# installe sur le poste et qu'aucune dependance nouvelle n'est requise.
#
# AVERTISSEMENT CENTRAL : le texte issu d'un OCR est APPROCHE. Un « I » lu pour
# un « l », un accent perdu, et le terme recopie dans la table ne correspondra
# plus au texte reel du document a anonymiser. C'est pourquoi toute entite
# issue d'un OCR est marquee « a verifier », sans exception.

import sys
import shutil
import tempfile
import subprocess

# Emplacements ajoutes par le programme principal (Tesseract embarque a cote
# de l'executable) : consultes avant tous les autres.
EMPLACEMENTS_SUPPLEMENTAIRES = []
TESSDATA_SUPPLEMENTAIRES = []

EMPLACEMENTS_TESSERACT = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    "/usr/bin/tesseract",
    "/usr/local/bin/tesseract",
]


class OcrIndisponible(RuntimeError):
    """L'OCR ne peut pas fonctionner, avec la raison exacte et le geste a faire."""


def trouver_tesseract(config_ocr=None):
    config_ocr = config_ocr or {}
    declare = config_ocr.get("tesseract")
    if declare and os.path.isfile(declare):
        return declare
    for candidat in EMPLACEMENTS_SUPPLEMENTAIRES:
        if candidat and os.path.isfile(candidat):
            return candidat
    trouve = shutil.which("tesseract")
    if trouve:
        return trouve
    for candidat in EMPLACEMENTS_TESSERACT:
        if os.path.isfile(candidat):
            return candidat
    raise OcrIndisponible(
        "Tesseract est introuvable sur ce poste.\n\n"
        "Si vous l'avez installe ailleurs, indiquez son chemin dans\n"
        "detection_config.json, rubrique « ocr », cle « tesseract »."
    )


def trouver_tessdata(config_ocr=None, langue="fra"):
    """Dossier contenant le pack de langue. Le francais n'etant pas installe
    par defaut, il vit souvent a cote de l'outil qui s'en sert."""
    config_ocr = config_ocr or {}
    fichier = f"{langue}.traineddata"

    candidats = []
    if config_ocr.get("tessdata"):
        candidats.append(config_ocr["tessdata"])
    candidats += list(TESSDATA_SUPPLEMENTAIRES)
    if os.environ.get("TESSDATA_PREFIX"):
        candidats.append(os.environ["TESSDATA_PREFIX"])
    candidats += [
        # A cote du programme (paquet distribue : dossier Tesseract livre).
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tesseract", "tessdata"),
        r"C:\Program Files\Tesseract-OCR\tessdata",
        "/usr/share/tesseract-ocr/5/tessdata",
        "/usr/share/tesseract-ocr/4.00/tessdata",
    ]

    for dossier in candidats:
        if dossier and os.path.isfile(os.path.join(dossier, fichier)):
            return dossier

    raise OcrIndisponible(
        f"Le pack de langue « {fichier} » est introuvable.\n\n"
        "Tesseract n'installe pas le francais par defaut. Deposez le fichier\n"
        f"{fichier} dans un dossier, puis indiquez ce dossier dans\n"
        "detection_config.json, rubrique « ocr », cle « tessdata ».\n\n"
        "Dossiers cherches :\n" + "\n".join(f"    {c}" for c in candidats if c)
    )


def _pdf_en_images(chemin_pdf, dossier, resolution=300, pages=None):
    """Une image par page (ou par page de la liste « pages », numerotees a
    partir de 1). Essaie PyMuPDF puis pdf2image (poppler)."""
    erreurs = []
    try:
        fitz = _pymupdf()
        images = []
        with fitz.open(chemin_pdf) as doc:
            zoom = resolution / 72.0
            for i, page in enumerate(doc):
                if pages is not None and (i + 1) not in pages:
                    continue
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                cible = os.path.join(dossier, f"page_{i + 1:04d}.png")
                pix.save(cible)
                images.append(cible)
        if images:
            return images
        erreurs.append("PyMuPDF : document sans page")
    except ImportError:
        erreurs.append("PyMuPDF absent")
    except Exception as e:
        erreurs.append(f"PyMuPDF : {e}")

    try:
        from pdf2image import convert_from_path
        images = []
        for i, image in enumerate(convert_from_path(chemin_pdf, dpi=resolution)):
            if pages is not None and (i + 1) not in pages:
                continue
            cible = os.path.join(dossier, f"page_{i + 1:04d}.png")
            image.save(cible)
            images.append(cible)
        if images:
            return images
        erreurs.append("pdf2image : document sans page")
    except ImportError:
        erreurs.append("pdf2image absent")
    except Exception as e:
        erreurs.append(f"pdf2image : {e}")

    raise OcrIndisponible(
        "Impossible de convertir ce PDF en images pour l'OCR.\n\n"
        + "\n".join("    - " + e for e in erreurs)
        + "\n\nInstallez PyMuPDF :  py -m pip install pymupdf"
    )


def ocr_pdf(chemin_pdf, config_ocr=None, progression=None, pages=None):
    """Texte d'un PDF scanne. Leve OcrIndisponible avec un message utilisable.

    pages : numeros de page (a partir de 1) a lire ; toutes si None. Le texte
    rendu contient alors les seules pages lues, separees par un saut de page.
    """
    config_ocr = config_ocr or {}
    langue = config_ocr.get("langue", "fra")
    resolution = int(config_ocr.get("resolution", 300))

    binaire = trouver_tesseract(config_ocr)
    tessdata = trouver_tessdata(config_ocr, langue)

    dossier = tempfile.mkdtemp(prefix="ocr_anonymisation_")
    try:
        images = _pdf_en_images(chemin_pdf, dossier, resolution, pages)
        morceaux = []
        for i, image in enumerate(images, start=1):
            if progression:
                progression(i, len(images))
            sortie = os.path.join(dossier, f"texte_{i:04d}")
            commande = [binaire, image, sortie, "-l", langue,
                        "--tessdata-dir", tessdata]
            options = {}
            if sys.platform == "win32":       # pas de fenetre noire sous Windows
                options["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            resultat = subprocess.run(commande, capture_output=True, text=True,
                                      **options)
            if resultat.returncode != 0:
                raise OcrIndisponible(
                    f"Tesseract a echoue sur la page {i}.\n\n"
                    f"{(resultat.stderr or '').strip()[:400]}"
                )
            fichier_texte = sortie + ".txt"
            lu = ""
            if os.path.isfile(fichier_texte):
                with open(fichier_texte, encoding="utf-8", errors="replace") as f:
                    lu = f.read()
            # Tesseract termine chaque page par un saut de page : le garder
            # decalait la numerotation des pages du texte (une page en trop
            # par page lue). Le seul separateur de pages est celui ajoute ici.
            lu = lu.replace(SAUT_PAGE, "").rstrip()
            morceaux.append(lu)         # une entree par page, meme vide
        return ("\n" + SAUT_PAGE + "\n").join(morceaux)
    finally:
        shutil.rmtree(dossier, ignore_errors=True)


def pdf_est_scanne(chemin_pdf, seuil=40):
    """Un PDF dont on n'extrait presque aucun texte est un document scanne."""
    try:
        return len(texte_pdf(chemin_pdf).strip()) < seuil
    except Exception:
        return True


# OCR page par page (7 octobre 2026). Un PDF peut etre MIXTE : premiere page
# en texte (bordereau, page de garde), pages suivantes scannees. Avant, le
# document entier etait juge sur son texte total : des qu'il depassait 40
# caracteres, l'OCR n'etait pas lance et les pages scannees n'etaient pas lues.
# Desormais une page passe a l'OCR si elle porte moins de SEUIL_PAGE_OCR
# lettres ou chiffres ET contient une image (une page blanche volontaire n'a
# pas d'image : elle n'est pas envoyee a Tesseract).
SEUIL_PAGE_OCR = 20


def _lisibles(texte):
    return len(re.findall(r"[^\W_]", texte or ""))


def _pages_pdf(chemin):
    """[(texte, porte_une_image)] par page, ou None si PyMuPDF manque."""
    try:
        fitz = _pymupdf()
    except ImportError:
        return None
    try:
        with fitz.open(chemin) as doc:
            return [(page.get_text(), bool(page.get_images(full=True))) for page in doc]
    except Exception:
        return None


def extraire_texte_ou_ocr(chemin, config_ocr=None, autoriser_ocr=True,
                          progression=None):
    """Retourne (texte, source) ou source vaut "texte", "ocr" ou "mixte"."""
    texte, source, _ = extraire_texte_detail(chemin, config_ocr, autoriser_ocr, progression)
    return texte, source


def extraire_texte_detail(chemin, config_ocr=None, autoriser_ocr=True,
                          progression=None):
    """Retourne (texte, source, pages_ocr).

    source : "texte", "ocr" (tout le document lu par OCR) ou "mixte" (seules
    certaines pages) ; pages_ocr : numeros des pages lues par OCR.
    """
    config_ocr = config_ocr or {}
    ext = os.path.splitext(chemin)[1].lower()
    if ext == ".docx":
        return texte_docx(chemin), "texte", []
    if ext in (".txt", ".md"):
        import anonymisation_core
        return anonymisation_core._lire_texte_brut(chemin)[0], "texte", []
    if ext != ".pdf":
        raise RuntimeError(f"Format non pris en charge : {ext}\nAttendu : .docx ou .pdf")

    texte = ""
    try:
        texte = texte_pdf(chemin)
    except Exception:
        texte = ""

    # Document quasi vide : c'est un scan, tout passe a l'OCR (comme avant).
    if len(texte.strip()) < 40:
        if not autoriser_ocr:
            raise RuntimeError(
                "Ce PDF ne contient quasiment aucun texte : c'est un document scanne.\n"
                "Activez l'OCR pour le lire."
            )
        texte_ocr = ocr_pdf(chemin, config_ocr, progression)
        nb = texte_ocr.count(SAUT_PAGE) + 1
        return texte_ocr, "ocr", list(range(1, nb + 1))

    # Document avec du texte : seules les pages scannees passent a l'OCR.
    pages = _pages_pdf(chemin)
    if not pages or not autoriser_ocr:
        return texte, "texte", []
    seuil = int(config_ocr.get("seuil_page", SEUIL_PAGE_OCR))
    a_lire = [i + 1 for i, (t, image) in enumerate(pages)
              if image and _lisibles(t) < seuil]
    if not a_lire:
        return texte, "texte", []

    lus = ocr_pdf(chemin, config_ocr, progression, pages=a_lire).split("\n" + SAUT_PAGE + "\n")
    textes = [t for t, _ in pages]
    for numero, lu in zip(a_lire, lus):
        textes[numero - 1] = lu
    complet = ("\n" + SAUT_PAGE + "\n").join(textes)
    source = "ocr" if len(a_lire) == len(pages) else "mixte"
    return complet, source, a_lire
