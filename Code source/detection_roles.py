"""
detection_roles.py
------------------
Qui est qui : rattache chaque terme detecte a une PARTIE et lui donne un ROLE.

Le detecteur (detection_core) trouve les noms, societes, adresses, SIRET...
Ce module decide a qui ils appartiennent. Il ne lit plus le document de facon
lineaire : il procede en trois passes.

    Passe 0  Type de document : contentieux, contrat, mise en demeure,
             audit RGPD (ou force par l'utilisateur). Le type fixe les roles.
    Passe 1  Presentation des parties : les blocs « POUR : ... DEMANDERESSE »,
             « ENTRE ... ci-apres « le Prestataire » », la lettre et son
             destinataire... Chaque societe ou personne en tete de bloc devient
             une partie ; ce qui suit dans le bloc (siege, SIRET, dirigeant,
             avocat) lui est rattache.
    Passe 2  Le reste du document : les occurrences rejoignent l'entite deja
             connue ; un attribut isole rejoint la partie citee juste avant lui ;
             un nom inconnu reste « a classer ». Il n'est JAMAIS verse d'office
             chez le demandeur.

Tous les mots reperes (qualites, ouvreurs, definitions...) sont lus dans
detection_config.json : ajouter un type de document ou un role ne demande
jamais de toucher a ce fichier.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import re
import collections

import modele_onglet
from modele_onglet import (plier, code_propre, cle_variante, calculer_balise, variantes_de, mots_revelateurs,
                           est_identite, COMMUN, ErreurModele)

TYPES_TETE = {"personne_morale", "personne_physique"}
TYPES_TITRES = {"marque", "brevet", "dessin_modele", "depot_inpi", "depot_europeen",
                "produit", "logiciel"}
FORMES_STRUCTURE = ("SELARL", "SELAS", "SCP", "AARPI", "SELAFA", "SELEURL")

AUTOMATIQUE = "automatique"
SENS_EMIS, SENS_RECU = "emis", "recu"

# Fenetre de rattachement d'un attribut isole a la partie citee avant lui.
FENETRE_PROXIMITE = 260


# ==================================================================
# MOTIFS CONSTRUITS DEPUIS LA CONFIGURATION
# ==================================================================

def _motif(mots, debut_ligne=False):
    """Liste de mots de la configuration -> expression reguliere sur texte plie.
    Les espaces deviennent « un ou plusieurs blancs » ; les accents sont
    ignores puisque le texte et les mots sont plies de la meme facon."""
    alternatives = []
    for m in sorted({plier(x).strip() for x in mots if x and x.strip()}, key=len, reverse=True):
        corps = r"\s+".join(re.escape(p) for p in m.split())
        debut = r"(?<![A-Z0-9])" if m[:1].isalnum() else ""
        fin = r"(?![A-Z0-9])" if m[-1:].isalnum() else ""
        alternatives.append(debut + corps + fin)
    if not alternatives:
        return None
    corps = "(?:" + "|".join(alternatives) + ")"
    if debut_ligne:
        return re.compile(r"(?m)^[ \t]*" + corps)
    return re.compile(corps)


# ==================================================================
# CONTEXTE DE TRAVAIL
# ==================================================================

class _Contexte:
    def __init__(self, texte, entites, config, type_doc):
        self.texte = texte
        self.T = plier(texte)
        self.entites = entites
        self.config = config
        self.type_doc = type_doc
        self.def_type = (config.get("types_documents") or {}).get(type_doc, {})
        self.compteurs = collections.Counter()
        self.parties = {}                 # code -> entite de tete
        self.pluriels = {}                # code de role -> pluriel ?
        for r in self.def_type.get("roles", []) + config.get("roles_communs", []):
            self.pluriels[r["code"]] = bool(r.get("pluriel"))
        self.occurrences = {id(e): self._occurrences(e) for e in entites}

    def _occurrences(self, e):
        spans = []
        for v in sorted(e.variantes, key=len, reverse=True):
            if len(v) < 2:
                continue
            for m in re.finditer(rf"(?<!\w){re.escape(v)}(?!\w)", self.texte):
                if not any(a <= m.start() < b for a, b in spans):
                    spans.append((m.start(), m.end()))
        if not spans:
            spans.append((e.position, e.position + len(e.canonique)))
        return sorted(spans)

    def nouvelle_partie(self, code_role, tete=None):
        code_role = code_propre(code_role)
        pluriel = self.pluriels.get(code_role, False)
        self.compteurs[code_role] += 1
        n = self.compteurs[code_role]
        code = f"{code_role}{n}" if pluriel else (code_role if n == 1 else f"{code_role}{n}")
        if tete is not None:
            self.parties[code] = tete
            self._attribuer(tete, code, "Identité")
        return code

    def _attribuer(self, e, partie, role, confiance=None, note=""):
        e.partie, e.role = partie, role
        if confiance and e.confiance == "haute":
            e.confiance = confiance
        if note:
            e.note = (e.note + " ; " + note).strip(" ;")

    def attribuer(self, e, partie, role, confiance=None, note=""):
        if not e.partie:
            self._attribuer(e, partie, role, confiance, note)

    def dans(self, e, debut, fin):
        return any(debut <= a < fin for a, _ in self.occurrences[id(e)])

    def premiere_dans(self, e, debut, fin):
        return next((a for a, _ in self.occurrences[id(e)] if debut <= a < fin), None)


def _role_attribut(e, config):
    """Role propose pour un attribut d'apres son type de detection."""
    carte = config.get("types_detection_vers_attribut", {})
    if e.type == "avocat":
        if plier(e.canonique).split(" ")[0] in FORMES_STRUCTURE:
            return "Cabinet d'avocat"
        return "Avocat"
    return carte.get(e.type, "")


# ==================================================================
# PASSE 0 : TYPE DE DOCUMENT
# ==================================================================

def detecter_type(texte, config):
    """Retourne (type, scores). Le debut du document compte double : c'est
    la que se lisent l'intitule et la presentation des parties."""
    T = plier(texte)
    tete = T[:4000]
    scores = {}
    for code, d in (config.get("types_documents") or {}).items():
        motif = _motif(d.get("signaux", []))
        if motif is None:
            scores[code] = 0
            continue
        total = 0
        for m in (plier(x).strip() for x in d.get("signaux", [])):
            mo = _motif([m])
            total += min(3, len(mo.findall(T))) + min(3, len(mo.findall(tete)))
        scores[code] = total
    if not scores or max(scores.values()) == 0:
        return "", scores
    return max(scores, key=lambda c: scores[c]), scores


# ==================================================================
# PASSE 1 : PRESENTATION DES PARTIES (BLOCS)
# ==================================================================

QUALITE_INCIDENTE = re.compile(
    r"\s*(?:AU|A\s+L'|A\s+LA|EN|SUR)\s+(?:RECOURS|INCIDENT|REFERE|RETRACTATION|REQUETE|OPPOSITION"
    r"|APPEL\s+INCIDENT|DEMANDE\s+RECONVENTIONNELLE|INTERVENTION)", re.IGNORECASE)


_PREMIER_ARTICLE = re.compile(r"(?m)^[ \t]*(?:(?:ARTICLE|CLAUSE|SECTION|ART\.?)\s*)?1\s*(?:[.)\-:]|\s+[A-Z])")


def _fin_entete(ctx, apres):
    """Fin de la presentation des parties : un titre de la liste « fin_entete »
    ou, quelle que soit la langue, le premier article numerote (« 1. OBJET »,
    « ARTICLE 1 », « 1 - SCOPE ») en debut de ligne (19/10/2026)."""
    motif = _motif(ctx.def_type.get("fin_entete", []), debut_ligne=True)
    fins = [m.start() for m in _PREMIER_ARTICLE.finditer(ctx.T) if m.start() > apres]
    if motif is not None:
        fins += [m.start() for m in motif.finditer(ctx.T) if m.start() > apres]
    return min(fins) if fins else min(len(ctx.T), apres + 6000)


def _segments(ctx):
    """Decoupe la presentation des parties en blocs.

    Un bloc s'ouvre sur un mot en debut de ligne (POUR, CONTRE, ENTRE, ET,
    A LA REQUETE DE...) et se ferme sur une qualite (DEMANDERESSE, INTIMEE...)
    ou sur l'ouverture du bloc suivant. Retourne [(debut, fin, role_ou_None)].
    """
    evenements = []
    for r in ctx.def_type.get("roles", []):
        q = _motif(r.get("qualites", []))
        if q:
            for m in q.finditer(ctx.T):
                # « defenderesse AU RECOURS », « demanderesse A L'INCIDENT » :
                # qualite dans l'incident, pas dans le dossier. Ignoree.
                if QUALITE_INCIDENTE.match(ctx.T, m.end()):
                    continue
                evenements.append((m.end(), "qualite", r["code"]))
        o = _motif(r.get("ouvreurs", []), debut_ligne=True)
        if o:
            for m in o.finditer(ctx.T):
                evenements.append((m.start(), "ouvreur", r["code"]))
    s = _motif(ctx.def_type.get("separateurs", []), debut_ligne=True)
    if s:
        for m in s.finditer(ctx.T):
            # « ET » n'ouvre un bloc que s'il est seul sur sa ligne (« ET : »)
            reste = ctx.T[m.end():ctx.T.find("\n", m.end()) if "\n" in ctx.T[m.end():] else len(ctx.T)]
            if plier(m.group(0)).strip() == "ET" and reste.strip(" :\t,") != "":
                continue
            evenements.append((m.start(), "separateur", None))
    for m in re.finditer(r"D'(?:UNE|AUTRE)\s+PART", ctx.T):
        evenements.append((m.end(), "fermeture", None))
    if not evenements:
        return []

    evenements.sort(key=lambda x: x[0])
    premier = next((p for p, genre, _ in evenements if genre in ("ouvreur", "separateur")),
                   evenements[0][0])
    fin_entete = _fin_entete(ctx, premier)
    evenements = [ev for ev in evenements if ev[0] <= fin_entete]

    # Un bloc va d'un ouvreur (ou separateur) au suivant, ou a la fin de la
    # presentation. La qualite (DEMANDERESSE, INTIMEE...) ne ferme PAS le bloc :
    # les avocats et leurs adresses sont souvent ecrits apres elle et doivent
    # rester rattaches a la partie. Sans ouvreur (bloc ouvert par une qualite
    # seule, ou par « D'UNE PART »), le bloc court sur les 1500 caracteres qui
    # precedent la qualite.
    segments = []
    debut, role_ouvreur, explicite, qualites = None, None, False, []
    for pos, genre, role in evenements:
        if genre in ("ouvreur", "separateur"):
            if debut is not None and pos > debut:
                segments.append([debut, pos, role_ouvreur, explicite, qualites])
            debut, role_ouvreur, explicite, qualites = pos, role, True, []
        elif genre == "qualite":
            if debut is None:
                segments.append([max(0, pos - 1500), pos, None, True, [role] if role else []])
                debut, role_ouvreur, explicite, qualites = pos, None, False, []
            elif role:
                qualites.append(role)
        elif genre == "fermeture":
            d = debut if debut is not None else max(0, pos - 1500)
            segments.append([d, pos, role_ouvreur, True, qualites])
            debut, role_ouvreur, explicite, qualites = pos, None, False, []
    if debut is not None and explicite and fin_entete > debut:
        segments.append([debut, fin_entete, role_ouvreur, explicite, qualites])

    sortie = []
    for d, f, role_ouvreur, explicite, qualites in segments:
        if not explicite:
            continue
        # une qualite ecrite dans le bloc (« ci-apres la demanderesse ») compte aussi
        if not qualites:
            trouvees = []
            for r in ctx.def_type.get("roles", []):
                q = _motif(r.get("qualites", []))
                if not q:
                    continue
                for m in q.finditer(ctx.T, d, f):
                    if not QUALITE_INCIDENTE.match(ctx.T, m.end()):
                        trouvees.append((m.start(), r["code"]))
                        break
            qualites = [code for _, code in sorted(trouvees)]
        # L'ouvreur (CONTRE, A LA REQUETE DE) fait foi : dans un refere-
        # retractation, la qualite ecrite (« defenderesse au recours ») est
        # celle de l'incident, pas celle du dossier. Sans ouvreur, la PREMIERE
        # qualite du bloc est celle de la partie (les suivantes qualifient
        # l'incident : « Requerante, defenderesse au recours »).
        role = qualites[0] if qualites else role_ouvreur
        sortie.append((d, f, role))
    return sortie


def _rattacher_dans_bloc(ctx, debut, fin, tetes):
    """Ce qui figure dans le bloc d'une partie lui appartient : siege, SIRET,
    dirigeant, avocat et son cabinet, et les coordonnees qui suivent l'avocat."""
    if not tetes:
        return
    avocats = []
    # « candidat » exclu : « ci-apres le Prestataire » faisait du mot
    # « Prestataire » l'identite de la partie (8 octobre 2026).
    membres = sorted((e for e in ctx.entites
                      if not e.partie and e.type not in TYPES_TETE | {"commissaire_justice", "rg", "candidat"}
                      and ctx.dans(e, debut, fin)),
                     key=lambda e: ctx.premiere_dans(e, debut, fin))
    for e in membres:
        pos = ctx.premiere_dans(e, debut, fin)
        precedentes = [t for t in tetes if ctx.premiere_dans(t, debut, fin) is not None
                       and ctx.premiere_dans(t, debut, fin) <= pos]
        tete = precedentes[-1] if precedentes else tetes[0]
        partie = tete.partie
        role = _role_attribut(e, ctx.config)
        if e.type == "avocat":
            ctx.attribuer(e, partie, role)
            avocats.append((pos, partie))
            continue
        if role == "Adresse" and re.search(r"(DEMEURANT|DOMICILIE)E?S?\s*:?\s*(A\s+|AU\s+)?$",
                                           ctx.T[max(0, pos - 40):pos]):
            dirigeants = [x for x in ctx.entites if x.partie == partie and x.role == "Dirigeant"
                          and ctx.premiere_dans(x, debut, fin) is not None
                          and ctx.premiere_dans(x, debut, fin) < pos]
            if dirigeants:
                ctx.attribuer(e, partie, "Adresse dirigeant")
                continue
        avocat_avant = [p for p, pa in avocats if p < pos and pa == partie]
        tete_apres_avocat = avocat_avant and ctx.premiere_dans(tete, debut, fin) < avocat_avant[-1]
        if avocat_avant and tete_apres_avocat:
            role = {"Adresse": "Adresse avocat", "E-mail": "E-mail avocat",
                    "Téléphone": "Téléphone avocat", "Site web": "Site web avocat"}.get(role, role)
        if role:
            ctx.attribuer(e, partie, role)


def _tetes_du_bloc(ctx, debut, fin):
    return sorted((e for e in ctx.entites
                   if e.type in TYPES_TETE and not e.partie and ctx.dans(e, debut, fin)),
                  key=lambda e: ctx.premiere_dans(e, debut, fin))


# ==================================================================
# DEFINITIONS CONTRACTUELLES : « ci-apres « le Prestataire » »
# ==================================================================

_ART = r"(?:(?:LE|LA|LES|THE)\s+|L')?"
_GUIL_O, _GUIL_F = r"[«\"“]", r"[»\"”]"
_TERME = r"([A-Z][A-Z0-9'\- ]{1,40}?)"
MOTIF_DEF_APRES = re.compile(
    rf"(?:CI-?\s?APRES|CI-?\s?DESSOUS|HEREIN\s?AFTER|HEREAFTER|REFERRED\s+TO\s+AS)\s*,?\s*"
    rf"(?:(?:[A-Z]+(?:EE?S?)|REFERRED\s+TO|KNOWN|CALLED|DEFINED)\s+)?(?:PAR\s+|AS\s+|TO\s+AS\s+)?:?\s*{_ART}"
    rf"{_GUIL_O}\s*{_ART}{_TERME}\s*{_GUIL_F}")
MOTIF_DEF_PARENTHESE = re.compile(rf"\(\s*{_ART}{_GUIL_O}\s*{_ART}{_TERME}\s*{_GUIL_F}\s*\)")
MOTIF_DEF_DESIGNE = re.compile(
    rf"{_GUIL_O}\s*{_ART}{_TERME}\s*{_GUIL_F}\s*:?\s*(?:DESIGNE|SIGNIFIE|S'ENTEND|VISE|MEANS|REFERS\s+TO|SHALL\s+MEAN)")


# Termes definis admis comme code de partie. Un contrat definit souvent une
# partie par son propre nom (« ci-apres « AGICE » ») : ce terme-la ne doit
# JAMAIS devenir une balise, il revelerait le nom qu'il masque (19/10/2026).
ROLES_DEFINITION_ADMIS = {
    "PRESTATAIRE", "CLIENT", "CLIENTE", "FOURNISSEUR", "EDITEUR", "VENDEUR", "ACHETEUR",
    "CONCEDANT", "CEDANT", "CESSIONNAIRE", "LICENCIE", "ABONNE", "UTILISATEUR", "SOUS_TRAITANT",
    "DISTRIBUTEUR", "REVENDEUR", "AGENT", "MANDATAIRE", "PARTENAIRE", "HEBERGEUR", "INTEGRATEUR",
    "BAILLEUR", "PRENEUR", "LOCATAIRE", "EMPRUNTEUR", "PRETEUR", "GARANT", "CAUTION", "BENEFICIAIRE",
    "DONNEUR_D_ORDRE", "MAITRE_D_OUVRAGE", "MAITRE_D_OEUVRE", "ENTREPRENEUR", "CONSULTANT",
    "AUTEUR", "PRODUCTEUR", "ARTISTE", "INTERPRETE", "EMPLOYEUR", "SALARIE", "STAGIAIRE",
    "RESPONSABLE_DE_TRAITEMENT", "RESPONSABLE", "DESTINATAIRE", "EMETTEUR", "PARTIE_DIVULGATRICE",
    "PARTIE_RECEPTRICE", "DIVULGATRICE", "RECEPTRICE", "ASSUREUR", "ASSURE", "FRANCHISEUR",
    "FRANCHISE", "APPORTEUR", "INVESTISSEUR", "ASSOCIE", "ACQUEREUR", "REPRENEUR",
    # version anglaise
    "PROVIDER", "SUPPLIER", "VENDOR", "CUSTOMER", "BUYER", "PURCHASER", "LICENSOR", "LICENSEE",
    "SUBSCRIBER", "USER", "PROCESSOR", "CONTROLLER", "SUBPROCESSOR", "DISTRIBUTOR", "RESELLER",
    "PARTNER", "CONTRACTOR", "SUBCONTRACTOR", "CONSULTANT", "LESSOR", "LESSEE", "TENANT",
    "LANDLORD", "BORROWER", "LENDER", "GUARANTOR", "DISCLOSING_PARTY", "RECEIVING_PARTY",
    "RECIPIENT", "EMPLOYER", "EMPLOYEE", "INSURER", "INSURED", "FRANCHISOR", "FRANCHISEE",
    "ASSIGNOR", "ASSIGNEE", "AUTHOR", "PUBLISHER", "PRODUCER", "SERVICE_PROVIDER",
}


def _roles_admis(ctx):
    admis = set(ROLES_DEFINITION_ADMIS)
    for r in ctx.def_type.get("roles", []):
        admis.add(code_propre(r.get("code", "")))
        admis |= {code_propre(m) for m in r.get("mots", [])}
    admis |= {code_propre(x) for x in ctx.config.get("definitions_roles_admis", [])}
    admis.discard("")
    return admis


def _definitions(ctx):
    """[(position, code_du_role, genre)] ; genre = « apres » (la partie est
    nommee avant la definition) ou « designe » (nommee apres).

    Seul un terme defini qui est un ROLE (Prestataire, Client, Bailleur...)
    devient un code de partie. Un terme defini qui est un nom, un sigle ou
    une marque (« ci-apres « AGICE » ») est ecarte : la societe recevra un
    role de repli ou une balise neutre, jamais son propre nom en balise."""
    ignorees = {plier(x).strip() for x in ctx.def_type.get("definitions_ignorees", [])}
    admis = _roles_admis(ctx)
    sortie = []
    ctx.zones_definitions = []
    for motif, genre in ((MOTIF_DEF_APRES, "apres"), (MOTIF_DEF_PARENTHESE, "apres"),
                         (MOTIF_DEF_DESIGNE, "designe")):
        for m in motif.finditer(ctx.T):
            terme = re.sub(r"\s+", " ", m.group(1)).strip(" '-")
            if not terme or terme in ignorees or len(terme) < 3:
                continue
            if any(abs(p - m.start()) < 5 for p, _, _ in sortie):
                continue
            code = code_propre(terme)
            # Terme non admis : garde sa position (code vide) pour qu'une
            # definition suivante ne remonte pas jusqu'a la societe qu'il nomme.
            sortie.append((m.start(), code if code in admis else "", genre))
            ctx.zones_definitions.append((m.start(), m.end()))
    return sorted(sortie)


def _entre_guillemets(ctx, a, b):
    """« AGICE » : un nom cite entre guillemets est un terme defini, pas une
    tete de partie."""
    avant = ctx.T[max(0, a - 3):a]
    apres = ctx.T[b:b + 3]
    return bool(re.search(r'[«"“]\s*$', avant) and re.match(r'\s*[»"”]', apres))


def _roles_par_definitions(ctx, definitions, correspondance=None):
    """Une definition donne son role a la societe ou la personne nommee juste
    avant elle (ou juste apres, pour « « X » designe la societe Y »)."""
    correspondance = correspondance or {}
    for pos, code, genre in definitions:
        if not code:
            continue
        code = correspondance.get(code, code)
        zones = getattr(ctx, "zones_definitions", [])
        if genre == "apres":
            # Un nom cite DANS une definition (« ci-apres « AGICE » ») n'est
            # pas une tete de partie pour la definition suivante.
            candidats = [(a, e) for e in ctx.entites if e.type in TYPES_TETE
                         for a, b in ctx.occurrences[id(e)] if pos - 700 <= a < pos
                         and not any(d <= a < f for d, f in zones)
                         and not _entre_guillemets(ctx, a, b)]
            if not candidats:
                continue
            a, e = max(candidats, key=lambda x: x[0])
            # une autre definition entre la tete et celle-ci : ce n'est pas la sienne
            if any(a < p < pos for p, _, _ in definitions):
                continue
        else:
            candidats = [(a, e) for e in ctx.entites if e.type in TYPES_TETE
                         for a, _ in ctx.occurrences[id(e)] if pos < a <= pos + 250]
            if not candidats:
                continue
            a, e = min(candidats, key=lambda x: x[0])
        if e.partie:
            continue
        ctx.pluriels.setdefault(code, False)
        if code in ctx.parties:
            code = ctx.nouvelle_partie(code)
            ctx.parties[code] = e
            ctx._attribuer(e, code, "Identité")
        else:
            ctx.parties[code] = e
            ctx.compteurs[code] += 1
            ctx._attribuer(e, code, "Identité")


def _signatures(ctx, codes):
    """« Pour le Client : Madame X » : le signataire rejoint la partie ; une
    societe nommee la rejoint aussi quand la partie n'etait pas identifiee
    (« le Client » presente sans nom en tete du contrat)."""
    for code in codes:
        mots = code.replace("_", " ")
        motif = re.compile(rf"(?<![A-Z])POUR\s+{_ART}{re.escape(mots)}(?![A-Z])")
        for m in motif.finditer(ctx.T):
            fin_ligne = ctx.T.find("\n", m.end())
            fin_ligne = len(ctx.T) if fin_ligne < 0 else fin_ligne
            if not ctx.T[m.end():fin_ligne].strip(" :\t,"):
                suivante = ctx.T.find("\n", fin_ligne + 1)
                fin_ligne = len(ctx.T) if suivante < 0 else suivante
            fenetre = (m.end(), fin_ligne)
            for e in sorted(ctx.entites, key=lambda x: ctx.premiere_dans(x, *fenetre) or 10**9):
                if e.partie or not ctx.dans(e, *fenetre):
                    continue
                if e.type in ("personne_physique", "representant_legal"):
                    ctx.attribuer(e, code, "Signataire")
                elif e.type == "personne_morale" and code not in ctx.parties:
                    ctx.parties[code] = e
                    ctx.compteurs[code] += 1
                    ctx._attribuer(e, code, "Identité")


# ==================================================================
# STRATEGIES PAR TYPE DE DOCUMENT
# ==================================================================

def _par_blocs(ctx, role_par_defaut=None):
    """Contentieux et presentation generique : un role par bloc."""
    for debut, fin, role in _segments(ctx):
        tetes = _tetes_du_bloc(ctx, debut, fin)
        if not tetes:
            continue
        role = role or role_par_defaut
        if not role:
            continue
        for t in tetes:
            ctx.nouvelle_partie(role, t)
        _rattacher_dans_bloc(ctx, debut, fin, tetes)


SEUIL_PARTIE_PROBABLE = 8


def _parties_probables(ctx):
    """Contentieux : une societe tres citee dans le corps mais absente de
    l'en-tete (co-defendeur, societe du meme groupe assignee a part...) est
    proposee comme partie du second role (defendeur, intime...), a verifier.
    Sans cette passe, elle restait « a classer » avec ses adresse et SIRET."""
    codes = [r["code"] for r in ctx.def_type.get("roles", [])]
    if len(codes) < 2 or not ctx.parties:
        return
    utilises = [c for c in codes if ctx.compteurs.get(c)]
    if len(utilises) >= 2:
        cible = utilises[1]
    elif utilises:
        i = codes.index(utilises[0])
        cible = codes[i + 1] if i + 1 < len(codes) else codes[0]
    else:
        return
    maxi = max(len(ctx.occurrences[id(t)]) for t in ctx.parties.values())
    seuil = max(SEUIL_PARTIE_PROBABLE, maxi // 20)
    for e in sorted(ctx.entites, key=lambda x: -len(ctx.occurrences[id(x)])):
        if e.partie or e.type != "personne_morale" or e.ambigu:
            continue
        if len(ctx.occurrences[id(e)]) < seuil:
            continue
        code = ctx.nouvelle_partie(cible, e)
        ctx._attribuer(e, code, "Identité", confiance="moyenne",
                       note="société très citée mais absente de l'en-tête : "
                            "autre partie probable, à vérifier")


def _contrat(ctx):
    definitions = _definitions(ctx)
    _roles_par_definitions(ctx, definitions)

    # Blocs « ENTRE ... ET ... » : ce qui reste sans definition prend le role
    # dont un mot figure dans le bloc (prestataire, client...), sinon PARTIE.
    replis = ctx.def_type.get("roles", [])
    for debut, fin, _ in _segments(ctx):
        tetes = [e for e in ctx.entites if e.type in TYPES_TETE and ctx.dans(e, debut, fin)]
        tetes.sort(key=lambda e: ctx.premiere_dans(e, debut, fin))
        tetes_libres = [t for t in tetes if not t.partie]
        if tetes_libres:
            bloc = ctx.T[debut:fin]
            for t in tetes_libres:
                code = None
                for r in replis:
                    m = _motif(r.get("mots", []))
                    if m and m.search(bloc) and r["code"] not in ctx.parties:
                        code = r["code"]
                        break
                generique = "PARTY" if modele_onglet.LANGUE == "en" else "PARTIE"
                ctx.pluriels.setdefault(generique, True)
                ctx.nouvelle_partie(code or generique, t)
                if not code:
                    t.note = (t.note + " ; rôle contractuel non trouvé : précisez la partie").strip(" ;")
                    t.confiance = "moyenne"
        _rattacher_dans_bloc(ctx, debut, fin, [t for t in tetes if t.partie])

    # Deux parties en tete, l'une est le Client (ou le Prestataire) et l'autre
    # n'a recu que le code generique PARTIE : elle est l'autre role, a verifier.
    if len(replis) >= 2:
        generiques = [c for c in ctx.parties if re.fullmatch(r"(?:PARTIE|PARTY)\d*", c)]
        roles_vus = set(ctx.parties) | {c for _, c, _ in definitions if c}
        nommees = [c for c in (replis[0]["code"], replis[1]["code"]) if c in roles_vus]
        if len(generiques) == 1 and len(nommees) == 1 and len(ctx.parties) <= 2:
            autre = replis[1]["code"] if nommees[0] == replis[0]["code"] else replis[0]["code"]
            tete = ctx.parties.pop(generiques[0])
            ctx.parties[autre] = tete
            ctx.compteurs[autre] += 1
            for e in ctx.entites:
                if e.partie == generiques[0]:
                    e.partie = autre
            tete.note = (tete.note.replace("rôle contractuel non trouvé : précisez la partie", "")
                         + f" ; rôle déduit : l'autre partie est {nommees[0]}, à vérifier").strip(" ;")

    codes = list({c for _, c, _ in definitions if c} | {r["code"] for r in replis})
    _signatures(ctx, codes)

    # « Le Prestataire désigne Madame X comme chef de projet » : une personne
    # citée juste après le nom d'une partie lui est rattachée, à vérifier.
    for e in ctx.entites:
        if e.partie or e.type not in ("personne_physique", "representant_legal"):
            continue
        avant = ctx.T[max(0, e.position - 70):e.position]
        for code in codes:
            if code in ctx.parties and re.search(
                    rf"(?<![A-Z])(?:LE|LA|L')\s*{re.escape(code.replace('_', ' '))}(?![A-Z])", avant):
                ctx.attribuer(e, code, "Personne liée", confiance="moyenne",
                              note=f"personne citée avec {code} : rôle à préciser")
                break


def _sens_courrier(ctx, sens):
    if sens in (SENS_EMIS, SENS_RECU):
        return sens
    avocats = [plier(a).strip() for a in (ctx.config.get("cabinet") or {}).get("avocats", []) if a.strip()]
    if not avocats:
        return ""
    for a in avocats:
        if re.search(rf"(?<![A-Z]){re.escape(a)}(?![A-Z])", ctx.T):
            return SENS_EMIS
    return SENS_RECU


def _est_du_cabinet(ctx, e):
    avocats = {plier(a).strip() for a in (ctx.config.get("cabinet") or {}).get("avocats", []) if a.strip()}
    return any(plier(v).strip() in avocats for v in e.variantes) or any(
        a and a in plier(e.canonique) for a in avocats)


_LIANT = re.compile(r"^[\s,:;\-]*(?:(?:DE|DU|DES|D')\s*)?(?:(?:LA|LE|LES|L')\s*)?"
                    r"(?:(?:SOCIETE|STE|ENTREPRISE|ASSOCIATION|COMMUNE)\s*)?$")


def _juste_apres(ctx, fin_marqueur, fenetre=70):
    """Tetes placees IMMEDIATEMENT apres un marqueur (« conseil de la societe X »,
    « ma cliente, Y ») : seuls des articles et « la societe » peuvent les separer.
    « Ma cliente constate que la societe Z » ne designe pas Z."""
    sortie = []
    for e in ctx.entites:
        if e.type not in TYPES_TETE or e.partie:
            continue
        a = ctx.premiere_dans(e, fin_marqueur, fin_marqueur + fenetre)
        if a is not None and _LIANT.match(ctx.T[fin_marqueur:a]):
            sortie.append((a, e))
    return [e for _, e in sorted(sortie, key=lambda x: x[0])]


def _code_commun(ctx, rang, defaut):
    """Code d'un role commun, par position dans le fichier de reglages
    (commissaire, expert, magistrat, greffier, tiers, salarie) : la version
    anglaise emploie d'autres codes (BAILIFF, EMPLOYEE...). Avant le 8 octobre
    2026, « COMMISSAIRE » et « SALARIE » etaient ecrits en dur."""
    codes = [r.get("code") for r in ctx.config.get("roles_communs", []) if r.get("code")]
    if defaut in codes:
        return defaut
    return codes[rang] if len(codes) > rang else defaut


def _codes_roles(ctx, defauts):
    """Codes des roles du type de document, dans l'ordre du fichier de
    reglages (la version anglaise emploie d'autres codes) ; a defaut, les
    codes historiques."""
    codes = [r.get("code") for r in ctx.def_type.get("roles", []) if r.get("code")]
    return tuple(codes[:len(defauts)]) if len(codes) >= len(defauts) else tuple(defauts)


def _mise_en_demeure(ctx, sens):
    nous, eux = _codes_roles(ctx, ("NOTRE_CLIENT", "PARTIE_ADVERSE"))
    represente = nous if sens == SENS_EMIS else eux
    destinataire = eux if sens == SENS_EMIS else nous

    # 1. Le client de l'avocat : « conseil de la societe X », « ma cliente, Y »
    motif = _motif(ctx.def_type.get("representation", []))
    if motif:
        for m in motif.finditer(ctx.T):
            tetes = _juste_apres(ctx, m.end())
            if tetes:
                t = tetes[0]
                code = represente if represente not in ctx.parties else ctx.nouvelle_partie(represente)
                if code == represente:
                    ctx.compteurs[represente] += 1
                ctx.parties[code] = t
                ctx._attribuer(t, code, "Identité")

    # 1 bis. « votre cliente, la societe X » : le client du destinataire
    motif = _motif(ctx.def_type.get("representation_destinataire", []))
    if motif:
        for m in motif.finditer(ctx.T):
            tetes = _juste_apres(ctx, m.end())
            if tetes and destinataire not in ctx.parties:
                ctx.parties[destinataire] = tetes[0]
                ctx.compteurs[destinataire] += 1
                ctx._attribuer(tetes[0], destinataire, "Identité")

    # 2. Les avocats. La liste du cabinet fait foi quand elle est remplie.
    #    A defaut : dans une lettre du cabinet, l'avocat qui signe est le notre ;
    #    dans une lettre recue, l'avocat nomme en tete (destinataire) est le notre.
    liste_cabinet = [a for a in (ctx.config.get("cabinet") or {}).get("avocats", []) if a.strip()]
    zone_signature = max(0, len(ctx.T) - 900)
    for e in ctx.entites:
        if e.type != "avocat" or e.partie:
            continue
        if _est_du_cabinet(ctx, e):
            proprietaire, conf = nous, None
        elif liste_cabinet:
            proprietaire, conf = eux, None
        else:
            # sans liste : l'avocat qui signe est l'auteur de la lettre
            signe = any(a >= zone_signature for a, _ in ctx.occurrences[id(e)])
            auteur = nous if sens == SENS_EMIS else eux
            autre = eux if sens == SENS_EMIS else nous
            proprietaire, conf = (auteur if signe else autre), "moyenne"
        ctx.attribuer(e, proprietaire, _role_attribut(e, ctx.config), confiance=conf,
                      note="" if conf is None else
                      "avocat attribué sans liste du cabinet (detection_config.json) : à vérifier")

    # 3. Le destinataire : la societe ou la personne de l'en-tete de la lettre
    fin = _fin_entete(ctx, 0) if ctx.def_type.get("fin_entete") else min(len(ctx.T), 1500)
    tetes = sorted((e for e in ctx.entites if e.type in TYPES_TETE and not e.partie
                    and ctx.dans(e, 0, fin)), key=lambda e: ctx.premiere_dans(e, 0, fin))
    if tetes and destinataire not in ctx.parties:
        t = tetes[0]
        ctx.parties[destinataire] = t
        ctx.compteurs[destinataire] += 1
        ctx._attribuer(t, destinataire, "Identité")
        _rattacher_dans_bloc(ctx, ctx.premiere_dans(t, 0, fin), fin, [t])
    if destinataire in ctx.parties:
        for e in tetes:
            if not e.partie and e.type != "personne_morale":
                ctx.attribuer(e, destinataire, "Personne liée", confiance="moyenne",
                              note="personne nommée avec le destinataire : rôle à préciser")


def _audit_rgpd(ctx):
    roles = {r["code"]: r for r in ctx.def_type.get("roles", [])}
    codes = [r["code"] for r in ctx.def_type.get("roles", [])]
    organisme = codes[0] if codes else "ORGANISME_AUDITE"
    # Roles de personnes physiques (DPO, referent) et de structures, par position
    # dans la liste : organisme, sous-traitant, responsable, DPO, referent, outil.
    structures = set(codes[0:3]) if len(codes) >= 3 else {"ORGANISME_AUDITE", "RESPONSABLE_TRAITEMENT", "SOUS_TRAITANT"}
    personnes = set(codes[3:5]) if len(codes) >= 5 else {"DPO", "REFERENT"}
    code_outil = codes[5] if len(codes) >= 6 else "OUTIL"

    # 1. Definitions : « ci-apres « la Societe » » designe en principe l'audite
    correspondance = {code_propre(x): organisme
                      for x in ctx.def_type.get("definitions_organisme", [])}
    _roles_par_definitions(ctx, _definitions(ctx), correspondance)

    # 2. Mots places pres d'un nom : « sous-traitant X », « DPO : Madame Y »
    for code, r in roles.items():
        motif = _motif(r.get("pres", []))
        if not motif:
            continue
        for m in motif.finditer(ctx.T):
            fenetre = (m.end(), m.end() + 90)
            for e in sorted(ctx.entites, key=lambda x: ctx.premiere_dans(x, *fenetre) or 10**9):
                if e.partie or e.type not in TYPES_TETE | {"representant_legal"} \
                        or not ctx.dans(e, *fenetre):
                    continue
                personne = e.type != "personne_morale"
                if code in personnes and not personne:
                    continue
                if code in structures and personne:
                    continue
                if not r.get("pluriel") and code in ctx.parties:
                    continue
                if r.get("pluriel"):
                    ctx.nouvelle_partie(code, e)
                else:
                    ctx.parties[code] = e
                    ctx.compteurs[code] += 1
                    ctx._attribuer(e, code, "Identité")
                break

    # 3. Repli : la societe la plus citee est l'organisme audite
    if organisme not in ctx.parties:
        libres = [e for e in ctx.entites if e.type == "personne_morale" and not e.partie]
        if libres:
            t = max(libres, key=lambda e: (e.occurrences, -e.position))
            ctx.parties[organisme] = t
            ctx.compteurs[organisme] += 1
            ctx._attribuer(t, organisme, "Identité", confiance="moyenne",
                           note="organisme audité déduit : c'est la société la plus citée")

    # 4. Outils et logiciels nommes
    for e in ctx.entites:
        if e.type == "outil" and not e.partie:
            ctx.nouvelle_partie(code_outil, e)

    # 5. Les autres personnes citees : des salaries de l'organisme, a verifier
    for e in ctx.entites:
        if e.type == "personne_physique" and not e.partie:
            ctx.nouvelle_partie(_code_commun(ctx, 5, "SALARIE"), e)
            e.confiance = "moyenne"
            e.note = (e.note + " ; personne citée dans l'audit : salarié présumé").strip(" ;")


MOTS_NON_OUTILS = {"RGPD", "CNIL", "LE", "LA", "LES", "UN", "UNE", "DES", "DU", "DE",
                   "CE", "CET", "CETTE", "SON", "SA", "SES", "LEUR", "NOTRE", "VOTRE",
                   "INTERNE", "EXTERNE", "DEDIE", "DEDIEE", "METIER", "MAISON"}


def detecter_outils(texte, config, entites_existantes=()):
    """Audit RGPD : logiciels et services nommes (« le logiciel de paie Silae »,
    « la solution CRM Salesforce »). Confiance moyenne : a relire."""
    from detection_core import EntiteDetectee, _contexte
    d = (config.get("types_documents") or {}).get("audit_rgpd", {})
    mots = d.get("mots_outils", [])
    if not mots:
        return []
    cles = "|".join(re.escape(m) for m in sorted(mots, key=len, reverse=True))
    motif = re.compile(
        rf"(?i:\b(?:{cles})s?\b)(?:\s+(?i:de|d')\s*[a-zà-ÿ]+)?\s*[«\"“]?\s*"
        rf"([A-Z][\w\-]*(?:\.[\w\-]+)*(?:[ \t]+[A-Z0-9][\w\-]*(?:\.[\w\-]+)*){{0,2}})")
    deja = {cle_variante(v) for e in entites_existantes for v in e.variantes}
    vus, sortie = set(), []
    for m in motif.finditer(texte):
        nom = m.group(1).strip(" .-")
        mots_nom = nom.split()
        cles_outils = {plier(x) for x in mots}
        while mots_nom and plier(mots_nom[0]) in cles_outils:
            mots_nom.pop(0)
        while mots_nom and plier(mots_nom[-1]) in MOTS_NON_OUTILS:
            mots_nom.pop()
        nom = " ".join(mots_nom)
        if len(nom) < 2 or plier(nom) in MOTS_NON_OUTILS or plier(nom) in {plier(x) for x in mots}:
            continue
        cle = cle_variante(nom)
        if cle in vus or cle in deja:
            continue
        vus.add(cle)
        e = EntiteDetectee(nom, "outil", m.start(1), _contexte(texte, m.start(1)),
                           confiance="moyenne")
        e.occurrences = len(re.findall(rf"(?<!\w){re.escape(nom)}(?!\w)", texte))
        e.note = "outil ou logiciel nommé : à confirmer"
        sortie.append(e)
    return sortie


# ==================================================================
# ROLES COMMUNS ET PASSE 2
# ==================================================================

def _roles_communs(ctx):
    communs = {r["code"]: r for r in ctx.config.get("roles_communs", [])}
    # Commissaires de justice et leur etude
    dernier_cj = None
    for e in sorted(ctx.entites, key=lambda x: x.position):
        if e.type != "commissaire_justice" or e.partie:
            continue
        if plier(e.canonique).split(" ")[0] in FORMES_STRUCTURE and dernier_cj:
            ctx.attribuer(e, dernier_cj, "Étude")
            continue
        dernier_cj = ctx.nouvelle_partie(_code_commun(ctx, 0, "COMMISSAIRE"), e)
    # Experts, magistrats, greffiers : reperes par le mot qui precede le nom
    for code, r in communs.items():
        motif = _motif(r.get("avant", []))
        if not motif:
            continue
        for e in ctx.entites:
            if e.partie or e.type != "personne_physique":
                continue
            a = e.position
            avant = ctx.T[max(0, a - 70):a]
            apres = ctx.T[a + len(e.canonique):a + len(e.canonique) + 30]
            if (motif.search(avant) and "SOCIETE" not in avant and " SON " not in avant) \
                    or re.match(r"\s*,?\s*(?:" + motif.pattern + ")", apres):
                ctx.nouvelle_partie(code, e)
    # References de procedure
    for e in ctx.entites:
        if e.type == "rg" and not e.partie:
            ctx.attribuer(e, COMMUN, "Référence RG")


# ==================================================================
# DIRIGEANTS NOMMES DANS LE CORPS DU TEXTE (15 octobre 2026)
# ==================================================================
# Une personne n'est rattachee comme dirigeant que sur une formule explicite
# (« X, president de la societe Y », « Y, representee par son gerant, X »),
# jamais par simple voisinage. L'entite etant unique pour tout le document
# (toutes ses ecritures sont regroupees), le rattachement vaut a chaque
# occurrence : X ne peut pas etre dirigeant de Y a un endroit et sans partie
# plus loin. Si le texte le presente comme dirigeant de deux societes
# differentes, rien n'est decide : la ligne reste a classer, avec les deux
# societes en remarque.

_FONCTION = (r"(?:PRESIDENTE?(?:[ -]+DIRECT(?:EUR|RICE)[ -]+GENERALE?|\s+DU\s+CONSEIL\s+D'ADMINISTRATION"
             r"|\s+DU\s+DIRECTOIRE)?|P\.?D\.?G\.?|DIRECT(?:EUR|RICE)\s+GENERALE?(?:\s+DELEGUEE?)?"
             r"|CO-?GERANTE?|GERANTE?|DIRIGEANTE?|REPRESENTANTE?\s+LEGALE?|FONDAT(?:EUR|RICE)"
             r"|DIRECTOR|CHAIRMAN|CHAIRWOMAN|MANAGER|LEGAL\s+REPRESENTATIVE|FOUNDER"
             r"|ASSOCIEE?\s+(?:UNIQUE|GERANTE?)|CEO|CHIEF\s+EXECUTIVE\s+OFFICER|MANAGING\s+DIRECTOR)")
_DIRIGEANT_APRES = re.compile(
    r"\s*,?\s*(?:\([^)\n]{0,60}\)\s*,?\s*)?"
    r"(?:(?:EN\s+SA\s+QUALITE\s+DE|ES\s+QUALITES?\s+DE|AGISSANT\s+EN\s+QUALITE\s+DE|SON|SA|"
    r"ANCIENNE?|ACTUELLE?|ALORS)\s+)*" + _FONCTION +
    r"\s+(?:DE\s+LA\s+SOCIETE|DE\s+LA|DE\s+L'|DU|DES|DE|D'|OF)\s*(?:SOCIETE\s+)?")
_DIRIGEANT_AVANT = re.compile(
    r"(?:REPRESENTEE?S?\s+PAR|PRISE?S?\s+EN\s+LA\s+PERSONNE\s+DE|AGISSANT\s+PAR)\s+"
    r"(?:(?:SON|SA|SES|L'UN\s+DE\s+SES|L'UNE\s+DE\s+SES)\s+)?(?:" + _FONCTION + r"S?\s*)?"
    r"(?:EN\s+EXERCICE\s*)?,?\s*(?:(?:MONSIEUR|MADAME|M\.|MME)\s*)?$")


def _societes_dirigees(ctx, e):
    """[(societe, extrait)] : societes dont e est dit dirigeant, occurrence par occurrence."""
    morales = [x for x in ctx.entites if x.type == "personne_morale" and x is not e]
    sortie = []
    for a, b in ctx.occurrences[id(e)]:
        m = _DIRIGEANT_APRES.match(ctx.T, b)
        if m and m.end() - b <= 140:
            fin = m.end()
            for x in morales:
                if any(c <= fin + 3 and d > fin for c, d in ctx.occurrences[id(x)]):
                    sortie.append((x, ctx.texte[a:min(len(ctx.texte), fin + 40)]))
                    break
            continue
        avant = ctx.T[max(0, a - 160):a]
        m = _DIRIGEANT_AVANT.search(avant)
        if not m:
            continue
        debut = max(0, a - 160) + m.start()
        meilleure = None
        for x in morales:
            for c, d in ctx.occurrences[id(x)]:
                if d <= debut and debut - d <= 300 and not re.search(r"\n[ \t]*\n", ctx.texte[d:debut]):
                    if meilleure is None or d > meilleure[1]:
                        meilleure = (x, d, c)
        if meilleure:
            x, d, c = meilleure
            sortie.append((x, ctx.texte[c:min(len(ctx.texte), a + len(e.canonique))]))
    return sortie


def _passe_dirigeants(ctx):
    for e in ctx.entites:
        if e.type not in ("personne_physique", "representant_legal"):
            continue
        liens = _societes_dirigees(ctx, e)
        if not liens:
            continue
        societes = []
        for x, _ in liens:
            if x not in societes:
                societes.append(x)
        extrait = " ".join(liens[0][1].split())[:110]
        if e.partie:
            autres = [x for x in societes if x.partie != e.partie]
            if autres:
                e.note = (e.note + " ; présenté aussi comme dirigeant de "
                          + ", ".join(x.canonique for x in autres) + " : à vérifier").strip(" ;")
            continue
        parties = []
        for x in societes:
            if x.partie and x.partie != COMMUN and x.partie not in parties:
                parties.append(x.partie)
        if len(societes) == 1 and len(parties) == 1:
            e.note = re.sub(r"\s*;?\s*dirigeant de [^;]*\(société sans partie\)[^;]*", "", e.note).strip(" ;")
            ctx.attribuer(e, parties[0], "Dirigeant",
                          note=f"dirigeant de {parties[0]} d'après le texte : « {extrait}… »")
        elif len(societes) > 1:
            e.note = (e.note + " ; présenté comme dirigeant de plusieurs sociétés ("
                      + ", ".join(x.canonique for x in societes) + ") : choisissez").strip(" ;")
        else:
            e.note = (e.note + f" ; dirigeant de {societes[0].canonique} d'après le texte "
                      "(société sans partie) : choisissez la partie de la société").strip(" ;")


# ==================================================================
# BALISES NEUTRES (15 octobre 2026)
# ==================================================================
# Anonymiser ne doit demander aucune connaissance du dossier. Tout terme que
# la detection ne sait pas rattacher avec certitude recoit une balise neutre
# selon sa nature, numerotee automatiquement et sans limite : une societe
# devient SOCIETE5, une personne PERSONNE_PHYSIQUE4, une adresse [ADRESSE3],
# un telephone [TELEPHONE2]... Rattacher a une partie reste possible ensuite
# (menu Partie du tableau, ou dans l'onglet), mais n'est jamais obligatoire.

def _codes_neutres(config):
    """(code societe, code personne physique) lus dans parties_menu « * »."""
    menu = modele_onglet.parties_menu(config)
    societe = next((r["code"] for r in menu if not r.get("personne_physique")), "SOCIETE")
    personne = next((r["code"] for r in menu if r.get("personne_physique")), "PERSONNE_PHYSIQUE")
    return societe, personne


NOTE_NEUTRE = "balise neutre (aucune partie certaine) : rattachez-la par le menu Partie si vous le souhaitez"


def _passe_neutre_societes(ctx):
    """Chaque societe restee sans partie devient une partie neutre (SOCIETE n),
    AVANT les rattachements : son dirigeant (« X, gerant de la societe Y »),
    son adresse ou son mail cites a cote d'elle la rejoignent."""
    societe, personne = _codes_neutres(ctx.config)
    ctx.pluriels[societe] = ctx.pluriels[personne] = True
    for e in sorted(ctx.entites, key=lambda x: x.position):
        if not e.partie and e.type == "personne_morale":
            ctx.nouvelle_partie(societe, e)
            e.note = (e.note + " ; " + NOTE_NEUTRE).strip(" ;")


def _passe_neutre(ctx):
    societe, personne = _codes_neutres(ctx.config)
    note = NOTE_NEUTRE
    # Les personnes et les attributs restes sans partie
    for e in sorted(ctx.entites, key=lambda x: x.position):
        if e.partie or e.type == "candidat":
            continue
        if e.type in ("personne_physique", "representant_legal"):
            ctx.nouvelle_partie(personne, e)
        else:
            role = _role_attribut(e, ctx.config)
            if not role or est_identite(role, ctx.config):
                continue
            ctx._attribuer(e, COMMUN, role)
        e.note = (e.note + " ; " + note).strip(" ;")


def _passe_proximite(ctx):
    """Un attribut isole (adresse, SIRET, courriel, dirigeant...) rejoint la
    partie citee juste avant lui, dans le meme paragraphe."""
    reperes = []
    for e in ctx.entites:
        if e.partie and (est_identite(e.role, ctx.config) or e.type in ("avocat", "commissaire_justice")):
            for a, b in ctx.occurrences[id(e)]:
                reperes.append((a, b, e))
    reperes.sort(key=lambda x: x[0])
    for e in sorted(ctx.entites, key=lambda x: x.position):
        # Un representant legal n'est jamais rattache par simple voisinage
        # (voir _passe_dirigeants) : une personne se classe sur formule explicite.
        if e.partie or e.type in TYPES_TETE | {"outil", "candidat", "expert", "representant_legal"} | TYPES_TITRES:
            continue
        role = _role_attribut(e, ctx.config)
        if not role:
            continue
        a = e.position
        avant = [r for r in reperes if r[1] <= a and a - r[1] <= FENETRE_PROXIMITE]
        avant = [r for r in avant if not re.search(r"\n[ \t]*\n", ctx.texte[r[1]:a])]
        if not avant:
            continue
        _, _, repere = avant[-1]
        if repere.type == "avocat":
            role = {"Adresse": "Adresse avocat", "E-mail": "E-mail avocat",
                    "Téléphone": "Téléphone avocat", "Site web": "Site web avocat"}.get(role, role)
        elif repere.type == "commissaire_justice" and role == "Adresse":
            role = "Adresse"
        a_repere = avant[-1][1]
        ecart = ctx.T[a_repere:a]
        explicite = len(ecart) <= 140 and re.search(
            r"SIEGE|DEMEURANT|DOMICILI|IMMATRICUL|\bSISE?\b|\bSIS\b|SIRET|SIREN|RCS", ecart)
        ctx.attribuer(e, repere.partie, role, confiance=None if explicite else "moyenne",
                      note="" if explicite else f"rattaché par proximité à {repere.partie}")


PROPRIETE = re.compile(r"TITULAIRE|PROPRIETAIRE|DEPOS|COMMERCIALIS|EXPLOIT|VEND|FABRIQU|DISTRIBU|"
                       r"CONCOIT|CONCU|PROPOS|OFFRE|EDITE|DEVELOPP|LANCE|SA MARQUE|SON BREVET|"
                       r"SES PRODUITS|SA GAMME|SES GAMMES|SON LOGICIEL|SON SYSTEME|SON PRODUIT")


def _passe_titres_et_produits(ctx):
    """Une marque, un brevet, un produit, un logiciel appartient a la partie
    qui en est dite titulaire ou qui le commercialise dans la meme phrase ;
    a defaut, il est COMMUN, avec sa balise : jamais « a classer » sans balise.
    Les experts nommes par leur fonction deviennent EXPERT1, EXPERT2..."""
    tetes = [e for e in ctx.entites if e.partie and e.partie != COMMUN
             and est_identite(e.role, ctx.config) and e.type in TYPES_TETE]
    for e in sorted(ctx.entites, key=lambda x: x.position):
        if e.partie:
            continue
        if e.type == "expert":
            ctx.nouvelle_partie("EXPERT", e)
            continue
        if e.type not in TYPES_TITRES:
            continue
        role = _role_attribut(e, ctx.config)
        if not role:
            continue
        proprietaire, meilleure = None, None
        for a, b in ctx.occurrences[id(e)]:
            debut = max(0, ctx.texte.rfind("\n", 0, a), ctx.texte.rfind(". ", 0, a) + 1)
            fin_p = ctx.texte.find("\n", b)
            fin_s = ctx.texte.find(". ", b)
            fin = min(len(ctx.texte), fin_p if fin_p != -1 else len(ctx.texte),
                      fin_s + 1 if fin_s != -1 else len(ctx.texte), b + 250)
            phrase = ctx.T[debut:fin]
            if not PROPRIETE.search(phrase):
                continue
            for t in tetes:
                for c, d in ctx.occurrences[id(t)]:
                    if debut <= c < fin:
                        dist = abs(c - a)
                        if meilleure is None or dist < meilleure:
                            proprietaire, meilleure = t, dist
            if proprietaire:
                break
        if proprietaire:
            ctx.attribuer(e, proprietaire.partie, role, confiance="moyenne",
                          note=f"rattaché à {proprietaire.partie} (titulaire / commercialise)")
        else:
            ctx.attribuer(e, COMMUN, role, confiance="moyenne")


# ==================================================================
# FILET : UNE BALISE NE REVELE JAMAIS LE NOM QU'ELLE MASQUE (19/10/2026)
# ==================================================================

def _sans_fuite(ctx):
    """Une partie dont le code contient un mot du nom d'un de ses membres
    (code AGICE pour « AGICE EDITIQUE », quel que soit le chemin qui y a
    mene) est renommee en partie neutre (SOCIETE n, PERSONNE_PHYSIQUE n)."""
    societe, personne = _codes_neutres(ctx.config)
    ctx.pluriels[societe] = ctx.pluriels[personne] = True
    for code in sorted({e.partie for e in ctx.entites if e.partie and e.partie != COMMUN}):
        membres = [e for e in ctx.entites if e.partie == code]
        variantes = [v for e in membres for v in e.variantes]
        fuite = mots_revelateurs(code, variantes)
        if not fuite:
            continue
        tete = next((e for e in membres if est_identite(e.role, ctx.config)), membres[0])
        base = personne if tete.type in ("personne_physique", "representant_legal") else societe
        nouveau = ctx.nouvelle_partie(base)
        ctx.parties[nouveau] = ctx.parties.pop(code, tete)
        for e in membres:
            e.partie = nouveau
        tete.note = (tete.note + f" ; le code « {code} » reprenait le nom ({', '.join(sorted(fuite))}) : "
                     f"remplacé par {nouveau}").strip(" ;")


# ==================================================================
# ALIGNEMENT SUR L'ONGLET DU DOSSIER
# ==================================================================

def _cle_souple(v):
    return modele_onglet.cle_souple(v)


def _aligner(ctx, existantes):
    """Un terme deja present dans l'onglet garde sa partie et sa balise ;
    une partie deja connue garde son code, et ses attributs la suivent.
    Une partie nouvelle prend un numero libre."""
    index, index_souple = {}, {}
    for x in existantes:
        for v in variantes_de(x["terme"]):
            if v.strip():
                index.setdefault(cle_variante(v), x)
                index_souple.setdefault(_cle_souple(v), x)
    codes_existants = {code_propre(x["partie"]) for x in existantes if x.get("partie")}

    renommage = {}
    for e in ctx.entites:
        trouve = next((index[cle_variante(v)] for v in e.variantes if cle_variante(v) in index), None)
        if trouve is None and e.type == "adresse":
            # « 3 rue Haute 59000 Lille » et « 3 rue Haute, 59000 Lille » : meme
            # adresse, la ponctuation en moins (15/10/2026)
            trouve = next((index_souple[_cle_souple(v)] for v in e.variantes
                           if _cle_souple(v) in index_souple), None)
        if trouve is None:
            continue
        e.existante = trouve
        connues = {cle_variante(v) for v in variantes_de(trouve["terme"])}
        e.variantes_nouvelles = [v for v in e.variantes if cle_variante(v) not in connues]
        if e.partie and est_identite(e.role, ctx.config) and trouve.get("partie"):
            renommage[e.partie] = code_propre(trouve["partie"])

    # Codes provisoires qui heurtent un code deja pris par une autre partie
    utilises = set(codes_existants) | set(renommage.values())
    for code in sorted({e.partie for e in ctx.entites if e.partie and e.partie != COMMUN}):
        if code in renommage:
            continue
        if code not in utilises:
            utilises.add(code)
            continue
        m = re.match(r"^(.*?)(\d*)$", code)
        base = m.group(1)
        pluriel = ctx.pluriels.get(base, bool(m.group(2)))
        # Numero suivant = plus grand numero deja pris + 1, comme le menu du
        # tableau : un trou (SOCIETE2 supprimee) n'est jamais rebouche, un
        # ancien document pourrait encore porter cette balise (15/10/2026).
        pris = [int(x.group(1) or 1) for u in utilises for x in [re.fullmatch(re.escape(base) + r"(\d*)", u)] if x]
        n = (max(pris) + 1) if pris else (1 if pluriel else 2)
        while (f"{base}{n}" if (pluriel or n > 1) else base) in utilises:
            n += 1
        nouveau = f"{base}{n}" if (pluriel or n > 1) else base
        renommage[code] = nouveau
        utilises.add(nouveau)

    for e in ctx.entites:
        if getattr(e, "existante", None):
            e.partie = e.existante.get("partie") or COMMUN
            e.role = e.existante.get("role") or e.role
            e.balise = e.existante["balise"]
        elif e.partie in renommage:
            e.partie = renommage[e.partie]


# ==================================================================
# POINT D'ENTREE
# ==================================================================

def attribuer_roles(texte, entites, config, type_force=None, sens=None,
                    existantes=None, demander_sens=None):
    """Rattache chaque entite a une partie et propose sa balise.

    type_force    : code d'un type de config, ou None / « automatique » ;
    sens          : « emis », « recu » ou None (mise en demeure) ;
    existantes    : entrees de l'onglet par partie vise (alignement) ;
    demander_sens : fonction appelee sans argument quand le sens d'un courrier
                    ne peut pas etre deduit ; doit renvoyer « emis » ou « recu ».

    Retourne un dictionnaire d'informations pour le classeur de relecture.
    """
    for e in entites:
        e.partie, e.role, e.balise = "", "", ""
        e.existante, e.variantes_nouvelles = None, []

    type_auto, scores = detecter_type(texte, config)
    type_doc = type_force if type_force and type_force != AUTOMATIQUE else type_auto
    ctx = _Contexte(texte, entites, config, type_doc)
    infos = {"type": type_doc, "type_auto": type_auto, "type_force": bool(
        type_force and type_force != AUTOMATIQUE), "scores": scores, "sens": "",
        "avertissements": []}

    if type_doc == "contentieux":
        _par_blocs(ctx)
        _parties_probables(ctx)
    elif type_doc == "contrat":
        _contrat(ctx)
    elif type_doc == "mise_en_demeure":
        s = _sens_courrier(ctx, sens)
        if not s and demander_sens:
            s = demander_sens()
        if s not in (SENS_EMIS, SENS_RECU):
            infos["avertissements"].append(
                "Sens du courrier inconnu : les parties n'ont pas été attribuées.")
        else:
            infos["sens"] = s
            _mise_en_demeure(ctx, s)
    elif type_doc == "audit_rgpd":
        _audit_rgpd(ctx)
    else:
        infos["avertissements"].append(
            "Type de document non reconnu : choisissez-le dans la fenêtre pour que "
            "les parties soient nommées.")

    _roles_communs(ctx)
    _passe_neutre_societes(ctx)
    _passe_dirigeants(ctx)
    _passe_proximite(ctx)
    _passe_titres_et_produits(ctx)
    _passe_neutre(ctx)

    for e in entites:
        if not e.partie:
            e.note = (e.note + " ; partie à choisir").strip(" ;")

    _sans_fuite(ctx)
    _aligner(ctx, existantes or [])

    # Balises proposees, comme l'injection les calculera
    prises = {x["balise"] for x in (existantes or []) if x.get("terme")}
    for e in sorted(entites, key=lambda x: x.position):
        if e.balise or not e.partie or not e.role:
            continue
        try:
            b, deja = calculer_balise(e.partie, e.role, prises, config)
        except ErreurModele:
            continue
        e.balise = b
        if deja:
            e.note = (e.note + f" ; même partie que {b} : sera ajouté comme variante").strip(" ;")
        prises.add(b)
    return infos


def proposer_a_retenir(e):
    """OUI pour toute donnee rangee (partie et role connus, balise neutre
    comprise) ; NON pour les candidats et pour un terme deja dans l'onglet
    sans ecriture nouvelle."""
    # 15 octobre 2026 : mieux vaut masquer un mot de trop que laisser passer
    # un nom. Restent a NON, meme pour une variante d'un terme deja dans
    # l'onglet : les « candidats » (mots repetes non reconnus), une ecriture
    # partagee entre deux termes (societe et marque du meme nom : une seule
    # balise possible, la decision vous revient) et tout terme lu par OCR
    # (orthographe possiblement fausse : il ne correspondrait pas au texte).
    if e.type == "candidat" or e.ambigu or getattr(e, "ocr", False):
        return "NON"
    if getattr(e, "existante", None):
        return "OUI" if e.variantes_nouvelles else "NON"
    return "OUI" if (e.partie and e.role) else "NON"
