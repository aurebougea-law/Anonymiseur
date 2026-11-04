"""
modele_onglet.py
----------------
Le modele d'onglet « par partie » et tout ce qui l'ecrit.

POURQUOI CE MODELE
    L'ancien gabarit offrait des cases numerotees a remplir a la file
    ([PM_DEM_1], [ADR_DEM_2]...). Rien n'y reliait une adresse, un SIRET ou un
    avocat a la partie concernee, et un LLM lisant le document anonymise ne
    pouvait pas savoir que [ADR_DEF_2] etait le siege de [PM_DEF_1].

    Ici, la balise se construit a partir de la PARTIE et du ROLE :
        partie DEFENDEUR1, role Identite   ->  [DEFENDEUR1]
        partie DEFENDEUR1, role Adresse    ->  [DEFENDEUR1_ADRESSE]
        une seconde adresse                ->  [DEFENDEUR1_ADRESSE2]
        partie COMMUN, role Reference RG   ->  [RG1]
    La balise se lit seule : c'est la seule information dont dispose le LLM.

COLONNES (identiques en position a l'ancien gabarit, pour le moteur)
    A Terme original | B Balise | C Role | D Partie | E Dossier | F Actif ? | G Notes
    Le moteur d'anonymisation ne lit que A, B et F : il n'est pas modifie.

REGLES D'ECRITURE
    - une balise existante n'est JAMAIS modifiee (la restitution en depend) ;
    - un terme deja present n'est jamais ecrit deux fois ;
    - il n'y a plus de cases vides : les lignes sont creees a la demande,
      regroupees par partie ; aucun debordement possible.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import re
import datetime
import unicodedata

import openpyxl
from openpyxl.styles import Font, Alignment

# Langue d'ecriture des onglets ("fr" ou "en"). La LECTURE accepte toujours
# les deux : un classeur francais se lit dans la version anglaise et inversement.
LANGUE = "fr"
TEXTES = {
    "fr": {
        "en_tetes": ["Terme original", "Balise", "Rôle", "Partie", "Dossier", "Actif ?", "Notes"],
        "sous_titre_balise": "(cliquez sur la flèche pour changer de balise)",
        "invite_balise": ("Balise", "Choisissez dans la liste (flèche à droite de la case), ou remplissez "
                          "Rôle et Partie : la balise s'écrit toute seule."),
        "titre_1": "TABLEAU DE CORRESPONDANCES — CONFIDENTIEL",
        "titre_2": ("Ne jamais transmettre ce fichier avec un document anonymisé  •  "
                    "Modèle par partie : une balise = une partie + un rôle"),
        "communs": "ÉLÉMENTS COMMUNS", "oui": "OUI",
        "autres": "AUTRES LIGNES (conservées telles quelles)",
    },
    "en": {
        "en_tetes": ["Original term", "Tag", "Role", "Party", "Matter", "Active?", "Notes"],
        "sous_titre_balise": "(click the arrow to change the tag)",
        "invite_balise": ("Tag", "Pick from the list (arrow on the right of the cell), or fill Role and "
                          "Party: the tag fills itself in."),
        "titre_1": "CORRESPONDENCE TABLE — CONFIDENTIAL",
        "titre_2": ("Never send this file together with an anonymised document  •  "
                    "Per-party layout: one tag = one party + one role"),
        "communs": "SHARED ITEMS", "oui": "YES",
        "autres": "OTHER ROWS (kept as they are)",
    },
}
EN_TETES_BALISE = ("Balise", "Tag")          # colonne B, dans les deux langues
EN_TETES_PARTIE = ("Partie", "Party")        # colonne D
LARGEURS = [48, 42, 22, 20, 14, 9, 44]
COMMUN = "COMMUN"              # valeur interne, identique dans les deux langues
COMMUN_EN = "SHARED"           # ce que voit l'utilisateur de la version anglaise


def partie_affichee(p):
    """Code de partie tel qu'ecrit dans le tableur : SHARED au lieu de COMMUN
    en version anglaise. La lecture accepte les deux (code_propre)."""
    return COMMUN_EN if (p == COMMUN and LANGUE == "en") else p


def T(cle):
    return TEXTES.get(LANGUE, TEXTES["fr"])[cle]


# Reglages de detection connus du module (poses par le programme principal) :
# servent a calculer une balise manquante et a nourrir les menus deroulants.
CONFIG = None


def config_courante():
    return CONFIG or {"attributs": [{"libelle": "Identité", "suffixe": ""}]}


NOMS_FEUILLE_ROLES = ("Rôles", "Roles")  # ancienne feuille partagee (avant le 8 oct. 2026)
NB_LIGNES_VIDES = 40


# ------------------------------------------------------------------
# COULEURS PAR PARTIE (partagees avec le classeur de detection)
# ------------------------------------------------------------------
# (teinte soutenue pour le titre du groupe, teinte claire pour ses lignes).
# Attribuees dans l'ordre d'apparition des parties ; COMMUN en gris.
PALETTE = [
    ("1F3864", "DCE6F2"),   # bleu nuit
    ("8B2E3F", "F6E0E4"),   # bordeaux / rose
    ("2E6B3F", "DFF0E3"),   # vert
    ("9A6400", "FBEFD3"),   # ambre
    ("5B3F8A", "E8E0F3"),   # violet
    ("1E6F75", "DAEFF1"),   # bleu canard
    ("7A4A1E", "F1E4D6"),   # brun
    ("3F5F8A", "E3EAF4"),   # bleu acier
]
COULEUR_COMMUN = ("595959", "EDEDED")
COULEUR_A_CLASSER = ("B45309", "FDE8CC")
COULEUR_CANDIDATS = ("6B7280", "F3F4F6")


def couleurs_parties(parties):
    """{code de partie: (fonce, clair)} dans l'ordre fourni."""
    table = {}
    i = 0
    for p in parties:
        if p in table:
            continue
        if p == COMMUN:
            table[p] = COULEUR_COMMUN
        else:
            table[p] = PALETTE[i % len(PALETTE)]
            i += 1
    return table


def hauteur_ligne(textes_et_largeurs, base=15.0, maxi=120.0):
    """Hauteur (points) pour que des cellules a retour a la ligne restent lisibles.
    textes_et_largeurs : [(texte, largeur de colonne en caracteres)]."""
    lignes = 1
    for texte, largeur in textes_et_largeurs:
        t = str(texte or "")
        if not t:
            continue
        n = sum(max(1, -(-len(part) // max(8, int(largeur * 1.1)))) for part in t.split("\n"))
        lignes = max(lignes, n)
    return min(maxi, base * lignes + (2 if lignes > 1 else 0))


def decomposer_balise(balise, config):
    """[PARTIE_AVOCAT_ADRESSE2] -> ("PARTIE", "Adresse avocat") ;
    [PARTIE] -> ("PARTIE", "Identité") ; [RG1] -> (COMMUN, "Référence RG").
    Retourne (partie, role interne) ou (None, None) si la balise est vide."""
    b = (balise or "").strip()
    if not (b.startswith("[") and b.endswith("]")):
        return None, None
    coeur = b[1:-1].strip()
    if not coeur:
        return None, None
    sans_numero = re.sub(r"\d+$", "", coeur)
    identite = next((l for l, sfx in attributs(config) if sfx == ""), "Identité")
    # suffixes les plus longs d'abord : AVOCAT_ADRESSE avant ADRESSE
    for libelle, suffixe in sorted(attributs(config), key=lambda x: -len(x[1])):
        if not suffixe:
            continue
        if sans_numero.endswith("_" + suffixe) and len(sans_numero) > len(suffixe) + 1:
            return sans_numero[:-(len(suffixe) + 1)], libelle
        if sans_numero == suffixe:
            return COMMUN, libelle
    # Role invente par l'utilisateur ([BAILLEUR_GARANT]) : la partie est un code
    # connu (BAILLEUR, DEMANDEUR2, NOTRE_CLIENT...), le reste est le role.
    codes = set()
    for d in (config.get("types_documents") or {}).values():
        codes |= {r.get("code", "") for r in d.get("roles", [])}
    codes |= {r.get("code", "") for r in config.get("roles_communs", [])}
    codes |= {r["code"] for r in parties_menu(config)}
    codes.discard("")
    for code in sorted(codes, key=len, reverse=True):
        m = re.match(rf"^({re.escape(code)}\d*)_(.+)$", sans_numero)
        if m:
            return m.group(1), m.group(2).replace("_", " ").capitalize()
    return coeur, identite


# Titres francais des anciens onglets (convertir_onglets.py).
TITRE_1 = TEXTES["fr"]["titre_1"]
TITRE_2 = TEXTES["fr"]["titre_2"]

CARACTERES_INTERDITS_ONGLET = set('[]:*?/\\')


class ErreurModele(RuntimeError):
    """Refus explique, destine a etre affiche tel quel."""


# ==================================================================
# OUTILS
# ==================================================================

def plier(texte):
    """Majuscules sans accents, LONGUEUR CONSERVEE caractere par caractere.

    Garder la longueur permet de chercher dans la version pliee et de relire
    les positions dans le texte d'origine.
    """
    sortie = []
    for c in str(texte):
        if c in "’‘`´":
            sortie.append("'")
            continue
        base = unicodedata.normalize("NFD", c)[0]
        u = base.upper()
        sortie.append(u if len(u) == 1 else c)
    return "".join(sortie)


def code_propre(texte):
    """« Sous-traitant 2 » -> SOUS_TRAITANT2. Compatible avec le motif des balises."""
    t = plier(texte or "").strip()
    t = re.sub(r"[^A-Z0-9]+", "_", t)
    t = re.sub(r"_+(\d+)$", r"\1", t)          # « PARTIE 1 » -> PARTIE1
    t = t.strip("_")
    return COMMUN if t == COMMUN_EN else t     # SHARED (version anglaise) = COMMUN


def _texte(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def variantes_de(terme):
    """Meme regle que le moteur : le « / » ne separe que s'il touche un espace
    (« 24/01234 » reste un seul terme)."""
    return [v.strip() for v in re.split(r"\s+/\s*|\s*/\s+", str(terme or "")) if v.strip()]


def cle_variante(v):
    return re.sub(r"\s+", " ", plier(v)).strip()


# Mots d'un terme qui ne revelent rien : formes sociales, civilites, mots
# generiques. Tout AUTRE mot du terme retrouve dans la balise est une fuite.
MOTS_SANS_RISQUE = {
    "SOCIETE", "SA", "SAS", "SASU", "SARL", "EURL", "SCI", "SNC", "SCP", "SELARL", "SELAS",
    "AARPI", "GIE", "LTD", "LLC", "INC", "PLC", "LLP", "GMBH", "CORP", "LIMITED", "CO",
    "MONSIEUR", "MADAME", "MAITRE", "MME", "MLLE", "MR", "MRS", "MS", "DR", "ME",
    "LE", "LA", "LES", "DE", "DU", "DES", "ET", "THE", "OF", "AND", "GROUPE", "GROUP",
    "COMPANY", "CABINET", "ETUDE", "AVOCAT", "AVOCATS", "PERSONNE", "PHYSIQUE", "MORALE",
    "CLIENT", "CLIENTE", "PRESTATAIRE", "FOURNISSEUR", "EDITEUR", "VENDEUR", "ACHETEUR",
    "COMMUN", "SHARED", "RUE", "AVENUE", "BOULEVARD", "CHEMIN", "PLACE", "BIS", "TER",
}


def mots_revelateurs(code, variantes):
    """Mots de la balise (ou du code de partie) qui figurent tels quels dans
    le terme : une balise [AGICE_EDITIQUE] pour « AGICE EDITIQUE » revele le
    nom qu'elle devait masquer. Retourne ces mots (vide si tout va bien).
    Les numeros et les mots sans risque (SOCIETE, SAS, MONSIEUR...) ne
    comptent pas (19 octobre 2026)."""
    mots_code = {m for m in re.split(r"[^A-Z]+", plier(str(code or ""))) if len(m) >= 3}
    mots_code -= MOTS_SANS_RISQUE
    if not mots_code:
        return set()
    mots_terme = set()
    for v in variantes or ():
        mots_terme |= {m for m in re.split(r"[^A-Z0-9]+", plier(str(v))) if len(m) >= 3}
    return mots_code & mots_terme


def cle_souple(v):
    """Cle de comparaison sans ponctuation (15 octobre 2026)."""
    return re.sub(r"\s+", " ", re.sub(r"[^\w]+", " ", plier(v))).strip()


# ==================================================================
# ATTRIBUTS (lus dans detection_config.json)
# ==================================================================

def attributs(config):
    """[(libelle interne, suffixe)]. Le libelle interne est celui que le code
    de detection emploie (en francais) ; « affichage » est ce que voit
    l'utilisateur dans le tableur (identique en francais, traduit en anglais)."""
    liste = config.get("attributs") or [{"libelle": "Identité", "suffixe": ""}]
    return [(a["libelle"], code_propre(a.get("suffixe", ""))) for a in liste]


def _attributs_complets(config):
    liste = config.get("attributs") or [{"libelle": "Identité", "suffixe": ""}]
    return [(a["libelle"], code_propre(a.get("suffixe", "")), a.get("affichage") or a["libelle"])
            for a in liste]


def libelles_attributs(config):
    """Libelles proposes a l'utilisateur (affichage)."""
    return [aff for _, _, aff in _attributs_complets(config)]


def _trouver_attribut(role, config):
    cle = plier(role or "").strip()
    for libelle, suffixe, affichage in _attributs_complets(config):
        if cle in (plier(libelle).strip(), plier(affichage).strip()):
            return libelle, suffixe, affichage
    return None


def affichage_role(role, config):
    """Libelle tel qu'il s'ecrit dans le tableur (traduit si besoin)."""
    trouve = _trouver_attribut(role, config)
    return trouve[2] if trouve else (role or "")


def suffixe_de(role, config):
    """Libelle de role (interne ou affiche) -> suffixe de balise. Un libelle
    inconnu devient son propre suffixe : l'utilisateur peut inventer un role
    sans toucher au code."""
    trouve = _trouver_attribut(role, config)
    return trouve[1] if trouve else code_propre(role)


def rang_role(role, config):
    cle = plier(role or "").strip()
    for i, (libelle, _, affichage) in enumerate(_attributs_complets(config)):
        if cle in (plier(libelle).strip(), plier(affichage).strip()):
            return i
    return 999


def est_identite(role, config):
    return bool(role) and suffixe_de(role, config) == ""


def calculer_balise(partie, role, prises, config):
    """Balise libre pour (partie, role), compte tenu des balises deja prises.

    Retourne (balise, deja_prise). deja_prise n'est vrai que pour l'Identite
    d'une partie qui en a deja une : c'est alors la meme partie, et le terme
    sera ajoute comme variante de la ligne existante.
    """
    partie = code_propre(partie)
    suffixe = suffixe_de(role, config)
    if not partie:
        raise ErreurModele("partie non renseignée")
    if partie == COMMUN:
        if not suffixe:
            raise ErreurModele("un élément COMMUN doit avoir un rôle autre que « Identité » "
                               "(par exemple Référence RG ou Montant)")
        n = 1
        while f"[{suffixe}{n}]" in prises:
            n += 1
        return f"[{suffixe}{n}]", False
    if not suffixe:
        b = f"[{partie}]"
        return b, b in prises
    base = f"{partie}_{suffixe}"
    if f"[{base}]" not in prises:
        return f"[{base}]", False
    n = 2
    while f"[{base}{n}]" in prises:
        n += 1
    return f"[{base}{n}]", False


# ==================================================================
# LECTURE D'UN ONGLET
# ==================================================================

def _ligne_entetes(ws, max_lignes=12):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_lignes, values_only=True), 1):
        if row and len(row) > 1 and _texte(row[1]).startswith(EN_TETES_BALISE):   # texte riche : valeur = texte brut
            return i, [_texte(v) for v in row[:7]]
    return None, None


def modele_de(ws):
    """« par_partie », « ancien » ou « inconnu »."""
    i, entetes = _ligne_entetes(ws)
    if not i:
        return "inconnu"
    if len(entetes) > 3 and entetes[3] in EN_TETES_PARTIE:
        return "par_partie"
    return "ancien"


def modele_onglet(chemin_excel, nom_onglet):
    wb = openpyxl.load_workbook(chemin_excel, read_only=True)
    try:
        if nom_onglet not in wb.sheetnames:
            raise ErreurModele(f"Onglet « {nom_onglet} » introuvable.")
        return modele_de(wb[nom_onglet])
    finally:
        wb.close()


def _lire_ws(ws, config=None):
    """Retourne (titres, entrees, autres).

    titres  : lignes au-dessus des en-tetes, reecrites telles quelles ;
    entrees : une par ligne portant une balise ;
    autres  : lignes remplies sans balise qui ne sont pas un titre de groupe
              (notes libres de l'utilisateur) : conservees en fin d'onglet.
    """
    i_ent, _ = _ligne_entetes(ws)
    titres, entrees, autres = [], [], []
    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
        neuf = [_texte(v) for v in (list(row or []) + [None] * 10)[:10]]
        origine, anciennes, ecrite = neuf[7], neuf[8], neuf[9]
        valeurs = ["" if v.startswith("=") else v for v in neuf[:7]]   # formules : vides
        if i_ent and i < i_ent:
            titres.append(valeurs)
            continue
        if i == i_ent:
            continue
        terme, balise, role, partie, dossier, actif, notes = valeurs
        if partie.strip().upper() == COMMUN_EN:
            partie = COMMUN                              # SHARED (version anglaise)
        if terme and not balise and role and partie:
            # ligne tapee a la main, balise non encore calculee (formule sans
            # valeur enregistree) : le programme la calcule a sa place
            prises = {e["balise"] for e in entrees}
            try:
                balise, _ = calculer_balise(partie, role, prises, config or config_courante())
            except ErreurModele:
                balise = ""
        if balise.startswith("[") and balise.endswith("]"):
            entrees.append({"ligne": i, "terme": terme, "balise": balise, "role": role,
                            "partie": partie, "dossier": dossier, "actif": actif,
                            "notes": notes, "origine": origine,
                            "anciennes": [b for b in re.split(r"[;\s]+", anciennes) if b.startswith("[")]})
            # Balise changee a la main dans la colonne B : l'ancienne reste reconnue
            if ecrite.startswith("[") and ecrite != balise and ecrite not in entrees[-1]["anciennes"]:
                entrees[-1]["anciennes"].append(ecrite)
        elif not balise and partie:
            continue                         # titre de groupe, regenere
        elif any(valeurs):
            autres.append(valeurs)
    appliquer_rattachements(entrees, config or config_courante())
    return titres, entrees, autres


# ------------------------------------------------------------------
# RATTACHEMENT APRES COUP (15 octobre 2026)
# ------------------------------------------------------------------
# Chaque ligne ecrite par le programme garde, en colonne H masquee, la partie
# et le role qu'elle avait a l'ecriture (« DEF1|Adresse »). Si l'utilisateur
# change ensuite la Partie ou le Role dans l'onglet (rattacher [ADRESSE3] a
# DEF1, par exemple), la ligne prend sa nouvelle balise ([DEF1_ADRESSE]) et
# l'ancienne est conservee en colonne I : la restitution, meme sans journal,
# reconnait les deux. Un document deja anonymise reste donc restituable.

def etat_ligne(partie, role):
    """« DEF1|Adresse » : partie et role tels qu'ecrits dans l'onglet."""
    return f"{code_propre(partie or COMMUN)}|{(role or '').strip()}"


def _etat_compare(texte):
    partie, _, role = (texte or "").partition("|")
    return code_propre(partie or COMMUN), plier(role).strip()


def appliquer_rattachements(entrees, config):
    """Lignes dont la partie ou le role a change depuis leur ecriture : nouvelle
    balise, l'ancienne rejoint la liste « anciennes ». Modifie sur place."""
    prises = {e["balise"] for e in entrees}
    for e in entrees:
        origine = e.get("origine") or ""
        if "|" not in origine or not e.get("terme"):
            continue
        p_orig, _, r_orig = origine.partition("|")
        # Seule une valeur ecrite par le programme compte : code de partie
        # propre, et balise de la ligne conforme a cet etat d'origine. Une
        # note tapee en colonne H (« voir | pièce 3 ») est ignoree.
        if p_orig != code_propre(p_orig) or not r_orig.strip():
            continue
        p_bal, r_bal = decomposer_balise(e.get("balise"), config)
        if not p_bal or code_propre(p_bal) != code_propre(p_orig) \
                or plier(suffixe_de(r_bal, config)) != plier(suffixe_de(r_orig, config)):
            continue
        actuel = etat_ligne(e.get("partie"), e.get("role"))
        if _etat_compare(origine) == _etat_compare(actuel):
            continue
        if not e.get("role") or not e.get("partie"):
            continue
        try:
            nouvelle, _ = calculer_balise(e["partie"], e["role"], prises, config)
        except ErreurModele:
            continue
        if nouvelle != e["balise"]:
            if e["balise"] not in e.setdefault("anciennes", []):
                e["anciennes"].append(e["balise"])
            e["balise"] = nouvelle
            prises.add(nouvelle)
        e["origine"] = actuel


def lire_onglet(chemin_excel, nom_onglet):
    """Entrees d'un onglet par partie (lecture seule)."""
    wb = openpyxl.load_workbook(chemin_excel, read_only=True, data_only=True)
    try:
        if nom_onglet not in wb.sheetnames:
            raise ErreurModele(f"Onglet « {nom_onglet} » introuvable.")
        ws = wb[nom_onglet]
        modele = modele_de(ws)
        _, entrees, _ = _lire_ws(ws)
        return modele, entrees
    finally:
        wb.close()


# ==================================================================
# ECRITURE
# ==================================================================

def _numero_balise(balise):
    m = re.search(r"(\d+)\]$", balise)
    return int(m.group(1)) if m else 0


def _ordonner(entrees, config):
    """Groupes dans l'ordre de premiere apparition, COMMUN en dernier ;
    dans un groupe : Identite, puis les roles dans l'ordre de la configuration."""
    ordre_parties = []
    for e in entrees:
        p = e["partie"] or COMMUN
        if p not in ordre_parties:
            ordre_parties.append(p)
    if COMMUN in ordre_parties:
        ordre_parties.remove(COMMUN)
        ordre_parties.append(COMMUN)
    groupes = []
    for p in ordre_parties:
        membres = [e for e in entrees if (e["partie"] or COMMUN) == p]
        membres.sort(key=lambda e: (rang_role(e["role"], config), e["role"],
                                    _numero_balise(e["balise"]), e["balise"]))
        groupes.append((p, membres))
    return groupes


def _nom_feuille_roles(nom_onglet):
    """« <onglet>_Rôles », tronque a 31 caracteres (limite Excel) en gardant
    le suffixe, qui fait reconnaitre la feuille comme technique."""
    sfx = "_Roles" if LANGUE == "en" else "_Rôles"
    return nom_onglet[:31 - len(sfx)] + sfx


def _feuille_roles(wb, config, parties, parties_onglet=(), nom_onglet=""):
    """Feuille masquee PROPRE A L'ONGLET : A role affiche | B suffixe |
    D parties proposees | F balises possibles. Nourrit les menus deroulants
    et la formule de balise. Recreee a chaque ecriture de l'onglet, donc
    toujours alignee sur ses reglages.

    Avant le 8 octobre 2026, une seule feuille « Rôles » servait a tout le
    classeur : chaque injection la regenerait avec les listes du dernier
    onglet ecrit, et les menus des autres onglets (contentieux, contrat...)
    pointaient sur de mauvaises listes. L'ancienne feuille partagee n'est pas
    supprimee : les onglets non reecrits depuis s'en servent encore."""
    nom = _nom_feuille_roles(nom_onglet)
    if nom in wb.sheetnames:
        del wb[nom]
    ws = wb.create_sheet(nom)
    ws["A1"], ws["B1"], ws["D1"] = ("Role", "Suffix", "Parties") if LANGUE == "en" else ("Rôle", "Suffixe", "Parties")
    roles = _attributs_complets(config)
    for i, (_, sfx, aff) in enumerate(roles, 2):
        ws.cell(row=i, column=1, value=aff)
        ws.cell(row=i, column=2, value=sfx)
    for i, p in enumerate(parties, 2):
        ws.cell(row=i, column=4, value=partie_affichee(p))
    balises = balises_possibles(config, _parties_balises(config, parties_onglet))
    ws["F1"] = "Tags" if LANGUE == "en" else "Balises"
    for i, b in enumerate(balises, 2):
        ws.cell(row=i, column=6, value=b)
    ws.sheet_state = "hidden"
    return nom, len(roles), len(parties), len(balises)


# Roles qui ont un sens pour un acteur commun (commissaire, expert, salarie...)
ROLES_ACTEUR_COMMUN = {"Identité", "Adresse", "E-mail", "Téléphone", "Étude"}
# Roles qui n'ont de sens que pour une personne (jamais pour COMMUN)
ROLES_PERSONNE = {"Identité", "Adresse", "SIRET", "Dirigeant", "Adresse dirigeant", "Signataire",
                  "Personne liée", "Salarié", "Enseigne", "E-mail", "Téléphone", "Site web",
                  "Avocat", "Cabinet d'avocat", "Adresse avocat", "E-mail avocat",
                  "Téléphone avocat", "Site web avocat", "Toque avocat", "Avocat postulant",
                  "Adresse avocat postulant", "Étude"}


def est_acteur_commun(partie, config):
    base = re.sub(r"\d+$", "", partie or "")
    return any(r["code"] == base for r in config.get("roles_communs", []))


# Roles sans objet pour une personne physique (menu allege)
ROLES_PERSONNE_MORALE_SEULEMENT = {"SIRET", "Dirigeant", "Adresse dirigeant", "Enseigne", "Site web"}


def parties_menu(config, type_doc=None):
    """Parties proposees dans les menus sans etre produites par la detection
    (reglage « parties_menu » : SOCIETE...). type_doc=None :
    toutes ; sinon celles du type demande et les transversales (« * »).
    Retourne la liste des reglages (dictionnaires), sans doublon de code."""
    menu = config.get("parties_menu") or {}
    sortie, vus = [], set()
    for cle, liste in menu.items():
        if type_doc is not None and cle not in ("*", type_doc):
            continue
        for r in liste or []:
            code = code_propre(r.get("code", ""))
            if code and code not in vus:
                vus.add(code)
                sortie.append(dict(r, code=code))
    return sortie


def codes_parties_menu(config, type_doc=None):
    """Codes numerotes des parties du menu : SOCIETE1 ... SOCIETE6."""
    sortie = []
    for r in parties_menu(config, type_doc):
        n = int(r.get("numeros") or 1)
        sortie.extend(f"{r['code']}{i}" for i in range(1, max(1, n) + 1))
    return sortie


# Roles qui decrivent le dossier plutot qu'une partie : proposes a COMMUN et
# aux parties detectees, pas aux parties du menu (liste deja longue).
ROLES_DOSSIER = {"Référence RG", "Montant", "Pièce", "Donnée", "Empreinte fichier", "URL dépôt",
                 "Clé de licence", "Version logiciel", "Dépôt INPI", "Dépôt européen"}


def est_partie_menu(partie, config):
    base = re.sub(r"\d+$", "", code_propre(partie or ""))
    return any(r["code"] == base for r in parties_menu(config))


def est_personne_physique(partie, config):
    """Vrai pour une partie du menu declaree « personne_physique » (menu de
    roles allege : pas de SIRET, dirigeant, enseigne, site web)."""
    base = re.sub(r"\d+$", "", code_propre(partie or ""))
    return any(r["code"] == base and r.get("personne_physique") for r in parties_menu(config))


def balises_possibles(config, parties, deja=()):
    """Balises proposees dans un menu, groupees par partie : pour une partie,
    l'identite puis ses roles de personne et ses titres ; pour un acteur
    commun (commissaire, expert...), l'identite et quelques roles ; pour
    COMMUN, les objets du dossier (marque, brevet, logiciel, RG...)."""
    sortie = []
    attributs = _attributs_complets(config)              # (libelle, suffixe, affichage)
    for p in parties:
        if p == COMMUN:
            for lib, sfx, _ in attributs:
                if sfx and lib not in ROLES_PERSONNE and f"[{sfx}1]" not in sortie:
                    sortie.append(f"[{sfx}1]")
            continue
        if f"[{p}]" not in sortie:
            sortie.append(f"[{p}]")
        commun = est_acteur_commun(p, config)
        du_menu = est_partie_menu(p, config)
        physique = est_personne_physique(p, config)
        for lib, sfx, _ in attributs:
            if not sfx:
                continue
            if commun and lib not in ROLES_ACTEUR_COMMUN:
                continue
            if du_menu and (lib in ROLES_DOSSIER or lib == "Étude"):
                continue
            if physique and lib in ROLES_PERSONNE_MORALE_SEULEMENT:
                continue
            if lib == "Étude" and not p.startswith(("COMMISSAIRE", "BAILIFF", "HUISSIER")):
                continue
            if f"[{p}_{sfx}]" not in sortie:
                sortie.append(f"[{p}_{sfx}]")
    for b in deja:
        if b and b not in sortie:
            sortie.append(b)
    return sortie


def _parties_balises(config, parties):
    """Parties du menu des balises (liste courte) : celles deja presentes dans
    l'onglet et le numero suivant de chacune ; a defaut, les parties du
    premier type de document ; les acteurs communs (n° 1) ; COMMUN. Les
    autres parties restent accessibles par Partie + Role."""
    codes_types = [r["code"] for d in (config.get("types_documents") or {}).values()
                   for r in d.get("roles", [])]
    communs = [r["code"] for r in config.get("roles_communs", [])]
    presentes = [p for p in parties if p != COMMUN and re.sub(r"\d+$", "", p) in codes_types]
    sortie = []
    for p in presentes:
        if p not in sortie:
            sortie.append(p)
    for p in presentes:
        m = re.fullmatch(r"(.*?)(\d*)", p)
        suivant = f"{m.group(1)}{int(m.group(2) or 1) + 1}"
        if suivant not in sortie:
            sortie.append(suivant)
    if not sortie:
        premier = next(iter((config.get("types_documents") or {}).values()), {})
        for r in premier.get("roles", []):
            sortie.append(f"{r['code']}1" if r.get("pluriel") else r["code"])
    for c in communs:
        if f"{c}1" not in sortie:
            sortie.append(f"{c}1")
    for c in codes_parties_menu(config):
        if c not in sortie:
            sortie.append(c)
    sortie.append(COMMUN)
    return sortie


def _parties_proposees(config, entrees):
    """Parties pour le menu deroulant : celles de l'onglet, puis les parties
    types de chaque type de document (numerotees), les acteurs communs, COMMUN."""
    sortie = []
    for e in entrees:
        p = e.get("partie") or COMMUN
        if p not in sortie:
            sortie.append(p)
    for d in (config.get("types_documents") or {}).values():
        for r in d.get("roles", []):
            code = r.get("code", "")
            for n in ((1, 2) if r.get("pluriel") else ("", 2)):
                c = f"{code}{n}"
                if c and c not in sortie:
                    sortie.append(c)
    for r in config.get("roles_communs", []):
        c = f"{r.get('code', '')}1"
        if c not in sortie:
            sortie.append(c)
    for c in codes_parties_menu(config):
        if c not in sortie:
            sortie.append(c)
    if COMMUN in sortie:
        sortie.remove(COMMUN)
    sortie.append(COMMUN)
    return sortie


def _formule_balise(r, premiere, feuille_roles):
    """Balise calculee par Excel pour une ligne saisie a la main :
    partie + suffixe du role (feuille Roles, sinon le role en majuscules) +
    numero si la meme partie a deja ce role au-dessus. COMMUN : suffixe + rang."""
    f = f"'{feuille_roles}'!$A:$B"
    # &"" : Excel renvoie 0 pour une cellule de suffixe vide (role Identite),
    # ce qui donnerait une balise terminee par « _0 » (8 octobre 2026).
    sfx = f'IFERROR(VLOOKUP(C{r},{f},2,FALSE)&"",UPPER(SUBSTITUTE(C{r}," ","_")))'
    deja = f"COUNTIFS(D${premiere}:D{r - 1},D{r},C${premiere}:C{r - 1},C{r})"
    return (f'=IF(OR(A{r}="",C{r}="",D{r}=""),"",'
            f'IF(D{r}="{partie_affichee(COMMUN)}","["&{sfx}&({deja}+1)&"]",'
            f'IF({sfx}="","["&D{r}&"]",'
            f'"["&D{r}&"_"&{sfx}&IF({deja}=0,"",{deja}+1)&"]")))')


def en_tete_balise(titre, sous_titre):
    """En-tete de la colonne Balise : « Balise » en gras, puis l'indication
    en corps 9 (texte riche). Les lecteurs voient la chaine complete, qui
    commence par « Balise » / « Tag »."""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    return CellRichText(TextBlock(InlineFont(b=True, color="FFFFFF"), titre + "  "),
                        TextBlock(InlineFont(sz=9, i=True, color="FFFFFF"), sous_titre))


def _ecrire_ws(ws, titres, entrees, autres, config):
    from openpyxl.styles import PatternFill, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import FormulaRule
    wb = ws.parent
    ws.delete_rows(1, ws.max_row + 1)
    ws.data_validations.dataValidation = []
    ws.conditional_formatting = type(ws.conditional_formatting)()
    ws.sheet_view.showGridLines = False
    blanc = "FFFFFF"
    fin = Side(style="thin", color="D0D5DB")
    fonce_titre = PALETTE[0][0]

    r = 1
    if not titres:
        titres = [[T("titre_1")] + [""] * 6, [T("titre_2")] + [""] * 6]
    for k, valeurs in enumerate(titres):
        for c, v in enumerate(valeurs, 1):
            if v:
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(bold=True, size=12, color=fonce_titre) if k == 0 else Font(italic=True, size=9, color="555555")
        r += 1
    epais = Side(style="medium", color=fonce_titre)
    sep_blanc = Side(style="thin", color=blanc)
    for c, (titre, largeur) in enumerate(zip(T("en_tetes"), LARGEURS), 1):
        cell = ws.cell(row=r, column=c, value=en_tete_balise(titre, T("sous_titre_balise")) if c == 2 else titre)
        cell.font = Font(bold=True, color=blanc)
        cell.fill = PatternFill("solid", fgColor=fonce_titre)
        cell.border = Border(left=sep_blanc, right=sep_blanc)   # separation fine entre colonnes
        cell.alignment = Alignment(vertical="center")
        ws.column_dimensions[cell.column_letter].width = largeur
    ws.row_dimensions[r].height = 22
    ws.freeze_panes = ws.cell(row=r + 1, column=1)
    r += 1
    premiere = r

    groupes = _ordonner(entrees, config)
    couleurs = couleurs_parties([p for p, _ in groupes])
    for partie, membres in groupes:
        sombre, pale = couleurs[partie]
        libelle = T("communs") if partie == COMMUN else partie
        cell = ws.cell(row=r, column=1, value=libelle)
        cell.font = Font(bold=True, color=blanc)
        ws.cell(row=r, column=4, value=partie_affichee(partie)).font = Font(bold=True, color=blanc)
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=sombre)
        ws.row_dimensions[r].height = 20
        r += 1
        for n, e in enumerate(membres):
            valeurs = [e["terme"], e["balise"], affichage_role(e["role"], config),
                       partie_affichee(e["partie"] or COMMUN),
                       e.get("dossier", ""), e.get("actif", "") or T("oui"), e.get("notes", "")]
            for c, v in enumerate(valeurs, 1):
                cell = ws.cell(row=r, column=c, value=v if v != "" else None)
                cell.alignment = Alignment(wrap_text=(c in (1, 2, 7)), vertical="top")
                cell.fill = PatternFill("solid", fgColor=pale if n % 2 == 0 else blanc)
                cell.border = Border(bottom=fin)
                if c == 2:
                    # Une vraie case de saisie : cadre epais, couleur de la partie
                    cell.font = Font(bold=True, color=sombre)
                    cell.border = Border(left=epais, right=epais, bottom=fin)
            # Colonnes masquees : etat a l'ecriture et anciennes balises
            ws.cell(row=r, column=8, value=f"{valeurs[3]}|{valeurs[2]}")   # tel qu'ecrit en D et C
            if e.get("anciennes"):
                ws.cell(row=r, column=9, value="; ".join(e["anciennes"]))
            ws.cell(row=r, column=10, value=e["balise"])
            ws.row_dimensions[r].height = hauteur_ligne(
                [(valeurs[0], LARGEURS[0]), (valeurs[1], LARGEURS[1]), (valeurs[6], LARGEURS[6])])
            r += 1

    if autres:
        r += 1
        ws.cell(row=r, column=1, value=T("autres")).font = Font(bold=True)
        r += 1
        for valeurs in autres:
            for c, v in enumerate(valeurs, 1):
                if v:
                    ws.cell(row=r, column=c, value=v)
            r += 1

    ws.column_dimensions["H"].hidden = True
    ws.column_dimensions["I"].hidden = True
    ws.column_dimensions["J"].hidden = True
    if r > premiere:
        # Partie ou role change a la main : la balise affichee sera mise a jour
        # au prochain usage (ambre) ; l'ancienne reste reconnue a la restitution.
        ws.conditional_formatting.add(
            f"B{premiere}:B{r - 1}",
            FormulaRule(formula=[f'AND($H{premiere}<>"",$D{premiere}&"|"&$C{premiere}<>$H{premiere})'],
                        fill=PatternFill("solid", fgColor="FFD8A8"), stopIfTrue=True))

    # Lignes vides pretes a l'emploi : terme + role + partie -> balise calculee
    parties = _parties_proposees(config, entrees)
    nom_roles, nb_roles, nb_parties, nb_balises = _feuille_roles(
        wb, config, parties, [e.get("partie") or COMMUN for e in entrees], ws.title)
    debut_vides = r
    for k in range(NB_LIGNES_VIDES):
        ws.cell(row=r, column=2, value=_formule_balise(r, premiere, nom_roles))
        ws.cell(row=r, column=2).font = Font(bold=True, color="333333")
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFF8E1")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="hair", color="E3E6EA"))
        ws.cell(row=r, column=2).border = Border(left=epais, right=epais, bottom=Side(style="hair", color="E3E6EA"))
        r += 1
    derniere = r - 1

    # Menu deroulant de la colonne Balise (toutes les balises possibles du
    # dossier) : choisir dans la liste remplace la formule, c'est voulu.
    dv_balise = DataValidation(type="list", formula1=f"'{nom_roles}'!$F$2:$F${nb_balises + 1}", allow_blank=True)
    dv_balise.showErrorMessage = False
    dv_balise.showInputMessage = True
    dv_balise.promptTitle, dv_balise.prompt = T("invite_balise")
    ws.add_data_validation(dv_balise)
    dv_balise.add(f"B{debut_vides}:B{derniere}")      # lignes deja ecrites : pas de menu (15/10/2026)

    # Menus deroulants : role (C), partie (D), actif (F)
    dv_role = DataValidation(type="list", formula1=f"'{nom_roles}'!$A$2:$A${nb_roles + 1}", allow_blank=True)
    dv_role.showErrorMessage = False
    dv_partie = DataValidation(type="list", formula1=f"'{nom_roles}'!$D$2:$D${nb_parties + 1}", allow_blank=True)
    dv_partie.showErrorMessage = False
    dv_actif = DataValidation(type="list", formula1=f'"{T("oui")},{"NO" if LANGUE == "en" else "NON"}"', allow_blank=True)
    for dv, col in ((dv_role, "C"), (dv_partie, "D"), (dv_actif, "F")):
        ws.add_data_validation(dv)
        dv.add(f"{col}{premiere}:{col}{derniere}")

    # Couleur automatique d'une ligne tapee a la main, selon sa partie
    for partie, (sombre, pale) in couleurs.items():
        ws.conditional_formatting.add(
            f"A{debut_vides}:G{derniere}",
            FormulaRule(formula=[f'$D{debut_vides}="{partie_affichee(partie)}"'], fill=PatternFill("solid", fgColor=pale)))


def _enregistrer(wb, chemin):
    try:
        wb.save(chemin)
    except PermissionError:
        raise ErreurModele(
            "Le classeur de correspondances est ouvert (dans Excel ?).\n\n"
            "Fermez-le, puis recommencez : rien n'a été modifié.")


def nom_onglet_valide(nom):
    nom = (nom or "").strip()
    if not nom:
        return "Le nom de l'onglet est vide."
    if len(nom) > 31:
        return "Excel limite le nom d'un onglet à 31 caractères."
    if any(c in CARACTERES_INTERDITS_ONGLET for c in nom):
        return "Le nom d'un onglet ne peut contenir aucun de ces signes :  [ ] : * ? / \\"
    return ""


def creer_onglet(chemin_excel, nom_onglet, config):
    """Cree un onglet vide au modele par partie. Refuse d'ecraser un onglet."""
    probleme = nom_onglet_valide(nom_onglet)
    if probleme:
        raise ErreurModele(probleme)
    nom_onglet = nom_onglet.strip()
    wb = openpyxl.load_workbook(chemin_excel)
    try:
        if nom_onglet in wb.sheetnames:
            raise ErreurModele(f"Un onglet « {nom_onglet} » existe déjà dans ce classeur.")
        ws = wb.create_sheet(nom_onglet)
        _ecrire_ws(ws, [], [], [], config)
        _enregistrer(wb, chemin_excel)
    finally:
        wb.close()
    return nom_onglet


def reecrire_onglet(chemin_excel, nom_onglet, entrees, config, titres=None):
    """Remplace tout le contenu d'un onglet par les entrees fournies
    (utilise par la conversion). Les autres onglets ne sont pas touches."""
    wb = openpyxl.load_workbook(chemin_excel)
    try:
        if nom_onglet not in wb.sheetnames:
            raise ErreurModele(f"Onglet « {nom_onglet} » introuvable.")
        ws = wb[nom_onglet]
        if titres is None:
            titres, _, _ = _lire_ws(ws)
        _ecrire_ws(ws, titres, entrees, [], config)
        _enregistrer(wb, chemin_excel)
    finally:
        wb.close()


# ==================================================================
# INJECTION
# ==================================================================

def injecter(chemin_excel, nom_onglet, retenues, config, mention=None):
    """Ecrit les termes retenus dans un onglet par partie.

    retenues : [{"terme", "partie", "role"}], dans l'ordre du classeur relu.
    La balise est TOUJOURS recalculee ici, d'apres la partie et le role :
    une correction faite a la main dans le classeur est donc respectee.

    Retourne (ecrits, refuses) :
        ecrits  : [(terme, balise, "nouvelle ligne" | "variantes ajoutées")]
        refuses : [(terme, raison)]
    """
    mention = mention or (f"added by detection on {datetime.date.today():%d/%m/%Y}" if LANGUE == "en"
                          else f"ajouté par la détection le {datetime.date.today():%d/%m/%Y}")
    wb = openpyxl.load_workbook(chemin_excel)
    try:
        if nom_onglet not in wb.sheetnames:
            raise ErreurModele(f"Onglet « {nom_onglet} » introuvable.")
        ws = wb[nom_onglet]
        if modele_de(ws) != "par_partie":
            raise ErreurModele(
                f"L'onglet « {nom_onglet} » suit l'ANCIEN modèle (cases numérotées).\n\n"
                "L'injection ne se fait que dans un onglet par partie. Deux solutions :\n"
                "  - créer un nouvel onglet (bouton « Nouvel onglet... ») ;\n"
                "  - convertir cet onglet (Convertir anciens onglets.bat).")
        titres, entrees, autres = _lire_ws(ws, config)

        index, index_souple = {}, {}
        for e in entrees:
            for v in variantes_de(e["terme"]):
                index.setdefault(cle_variante(v), e)
                index_souple.setdefault(cle_souple(v), e)
        prises = {e["balise"] for e in entrees if e["terme"]}
        ecrits, refuses = [], []

        def ajouter_variantes(cible, variantes):
            connues = {cle_variante(v) for v in variantes_de(cible["terme"])}
            nouvelles = [v for v in variantes if cle_variante(v) not in connues]
            if nouvelles:
                cible["terme"] = " / ".join(variantes_de(cible["terme"]) + nouvelles)
                for v in nouvelles:
                    index[cle_variante(v)] = cible
            return nouvelles

        # Mots de chaque partie (lignes de l'onglet + lignes a ecrire) : une
        # balise ne doit jamais contenir un mot du nom qu'elle masque.
        mots_par_partie = {}
        for e in entrees:
            if e["terme"]:
                mots_par_partie.setdefault(code_propre(e["partie"] or COMMUN), set()).update(variantes_de(e["terme"]))
        for r in retenues:
            mots_par_partie.setdefault(code_propre(r.get("partie", "")), set()).update(variantes_de(r["terme"]))

        for r in retenues:
            terme = r["terme"]
            variantes = variantes_de(terme)
            if not variantes:
                continue
            partie = code_propre(r.get("partie", ""))
            role = (r.get("role") or "").strip()
            fuite = mots_revelateurs(partie, mots_par_partie.get(partie, ()))
            if fuite:
                refuses.append((terme, "la partie « " + partie + " » reprend le nom qu'elle doit masquer ("
                                       + ", ".join(sorted(fuite)) + ") : choisissez une autre partie "
                                       "(SOCIETE, PRESTATAIRE, CLIENT…)"))
                continue

            deja = next((index[cle_variante(v)] for v in variantes
                         if cle_variante(v) in index), None)
            if deja is None:
                # Meme terme a la ponctuation pres (« 3 rue Haute 59000 Lille » et
                # « 3 rue Haute, 59000 Lille ») : une variante, pas une ligne de plus
                deja = next((index_souple[cle_souple(v)] for v in variantes
                             if cle_souple(v) in index_souple), None)
            if deja is not None:
                meme_place = (not partie or code_propre(deja["partie"] or COMMUN) == (partie or COMMUN)) \
                    and (not role or plier(deja["role"]) == plier(role))
                if not meme_place:
                    refuses.append((terme, f"déjà présent dans l'onglet sous {deja['balise']} "
                                           f"({deja['partie']} / {deja['role']}) : "
                                           "une balise existante n'est jamais déplacée"))
                    continue
                nouvelles = ajouter_variantes(deja, variantes)
                if nouvelles:
                    ecrits.append((" / ".join(nouvelles), deja["balise"], "variantes ajoutées"))
                else:
                    refuses.append((terme, f"déjà présent dans l'onglet ({deja['balise']})"))
                continue

            if not partie or partie in ("A_CLASSER", "ACLASSER"):
                refuses.append((terme, "partie à choisir (colonne Partie du classeur)"))
                continue
            if not role:
                refuses.append((terme, "rôle à choisir (colonne Rôle du classeur)"))
                continue
            # Ligne du modele complet (balise prete, terme vide) : on la remplit
            vide = next((e for e in entrees if not e["terme"]
                         and code_propre(e["partie"] or COMMUN) == partie
                         and plier(e["role"]) == plier(role)), None)
            if vide is not None:
                vide["terme"] = " / ".join(variantes)
                prises.add(vide["balise"])
                vide["notes"] = (vide["notes"] + " ; " + mention).strip(" ;") if vide.get("notes") else mention
                for v in variantes:
                    index[cle_variante(v)] = vide
                ecrits.append((vide["terme"], vide["balise"], "nouvelle ligne"))
                continue
            try:
                balise, deja_prise = calculer_balise(partie, role, prises, config)
            except ErreurModele as e:
                refuses.append((terme, str(e)))
                continue
            if deja_prise:
                cible = next(e for e in entrees if e["balise"] == balise)
                nouvelles = ajouter_variantes(cible, variantes)
                if nouvelles:
                    ecrits.append((" / ".join(nouvelles), balise, "variantes ajoutées"))
                else:
                    refuses.append((terme, f"déjà présent dans l'onglet ({balise})"))
                continue

            e = {"terme": " / ".join(variantes), "balise": balise, "role": role,
                 "partie": partie, "dossier": "", "actif": "OUI", "notes": mention}
            entrees.append(e)
            prises.add(balise)
            for v in variantes:
                index[cle_variante(v)] = e
            ecrits.append((e["terme"], balise, "nouvelle ligne"))

        if ecrits:
            _ecrire_ws(ws, titres, entrees, autres, config)
            _enregistrer(wb, chemin_excel)
        return ecrits, refuses
    finally:
        wb.close()
