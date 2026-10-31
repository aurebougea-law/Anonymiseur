"""
detection_classeur.py
---------------------
Classeur de relecture (detection ou conversion), sa relecture, et l'injection
des termes retenus dans un onglet par partie.

Aucun code d'interface ici : ce module est importable et testable sans ecran.

Le classeur propose, pour chaque terme, une PARTIE (DEFENDEUR1, PRESTATAIRE,
NOTRE_CLIENT...) et un ROLE (Identite, Adresse, Avocat...). Les deux se
corrigent dans des listes deroulantes. La balise affichee n'est qu'un apercu :
l'injection la recalcule d'apres la partie et le role relus, de sorte qu'une
correction faite a la main est toujours respectee.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import os
import re

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

import modele_onglet as modele
import textes
import detection_roles as roles

LANGUE = "fr"            # langue d'ecriture du classeur ; la lecture accepte les deux
TEXTES = {
    "fr": {
        "en_tetes": ["Terme et variantes", "Balise", "Garder ?",
                     "Partie (qui ?)", "Rôle (quoi ?)", "Nature", "Page", "Contexte"],
        "a_choisir": "▼ choisir Partie et Rôle", "sous_titre_balise": "(s'écrit toute seule)",
        "invite_titre": "Partie", "invite": "Qui ? Les parties déjà connues sont affichées avec leur nom ; « nouvelle partie » crée la suivante.",
        "titre": "PRÉ-DÉTECTION : à relire, puis Étape 2 · Injecter",
        "consigne": ("CHEMIN RAPIDE : rien à faire. Tout est déjà rangé ou reçoit une balise neutre "
                     "([ADRESSE3], [SOCIETE5]). Enregistrez, fermez, puis Étape 3 · Injecter.\n"
                     "CHEMIN PRÉCIS (si vous voulez dire qui est qui) : sur une ligne, choisissez la "
                     "PARTIE (qui ?) puis le RÔLE (quoi ?). La balise s'écrit toute seule.\n"
                     "Même numéro = même partie. « nouvelle partie » donne le numéro suivant tout seul. "
                     "Balise en ROUGE = deux identités différentes sous la même balise : choisissez une "
                     "nouvelle partie pour l'une des deux."),
        "nouvelle": "(nouvelle partie : numéro automatique)", "commun_menu": "(aucune partie : RG, montant…)",
        "deja_onglet": "déjà dans l'onglet", "propose": "proposé par la détection",
        "qui_statut": "Statut",
        "oui_non": "OUI,NON", "oui": "OUI", "non": "NON", "a_classer": "À CLASSER",
        "feuille": "Détection", "listes": ("Balises", "Garder", "Parties", "Rôles", "Suffixes"),
        "qui": "Qui est qui : chaque partie et ce qui lui est rattaché",
        "qui_colonnes": ["Partie", "Rôle", "Balise", "Terme", "Statut"],
        "deja": "déjà dans l'onglet ({balise})", "variantes": " ; variantes nouvelles : ",
        "non_reconnu": "non reconnu", "choisi": "choisi dans la fenêtre",
        "auto": "reconnu automatiquement", "type": "Type de document : {t} ({m})",
        "emis": "    |    Courrier : ÉMIS par le cabinet", "recu": "    |    Courrier : REÇU",
        "document": "Document : {d}    |    Destination : {t} / {o}",
        "ancien": ("ATTENTION : l'onglet visé suit l'ancien modèle ; l'injection y sera refusée "
                   "(créez un nouvel onglet)."),
        "attention": "ATTENTION : ",
        "groupe_a_classer": "À CLASSER : le programme n'a pas su rattacher ces termes, choisissez la Partie et le Rôle",
        "groupe_commun": "ÉLÉMENTS COMMUNS (n'appartiennent à aucune partie)",
        "groupe_candidats": "CANDIDATS : mots répétés non reconnus (produit ? personne ? lieu ?) : à vous de dire",
        "groupe_partie": "{p}",
    },
    "en": {
        "en_tetes": ["Term and variants", "Tag", "Keep?",
                     "Party (who?)", "Role (what?)", "Detected as", "Page", "Context"],
        "a_choisir": "▼ pick Party and Role", "sous_titre_balise": "(fills itself in)",
        "invite_titre": "Party", "invite": "Who? Known parties are shown with their name; « new party » creates the next one.",
        "titre": "PRE-DETECTION: review, then Step 2 · Inject",
        "consigne": ("QUICK PATH: nothing to do. Everything is already sorted or gets a neutral tag "
                     "([ADDRESS3], [COMPANY5]). Save, close, then Step 3 · Inject.\n"
                     "PRECISE PATH (if you want to say who is who): on a line, pick the PARTY (who?) "
                     "then the ROLE (what?). The tag fills itself in.\n"
                     "Same number = same party. « new party » gives the next number by itself. Tag in "
                     "RED = two different identities under the same tag: pick a new party for one of them."),
        "nouvelle": "(new party: automatic number)", "commun_menu": "(no party: case number, amount…)",
        "deja_onglet": "already in the sheet", "propose": "proposed by detection",
        "qui_statut": "Status",
        "oui_non": "YES,NO", "oui": "YES", "non": "NO", "a_classer": "TO SORT",
        "feuille": "Detection", "listes": ("Tags", "Keep", "Parties", "Roles", "Suffixes"),
        "qui": "Who is who: each party and what belongs to it",
        "qui_colonnes": ["Party", "Role", "Tag", "Term", "Status"],
        "deja": "already in the sheet ({balise})", "variantes": "; new variants: ",
        "non_reconnu": "not recognised", "choisi": "chosen in the window",
        "auto": "recognised automatically", "type": "Document type: {t} ({m})",
        "emis": "    |    Letter: SENT by the firm", "recu": "    |    Letter: RECEIVED",
        "document": "Document: {d}    |    Destination: {t} / {o}",
        "ancien": ("WARNING: the target sheet uses the old layout; injection will be refused "
                   "(create a new matter sheet)."),
        "attention": "WARNING: ",
        "groupe_a_classer": "TO SORT: the program could not assign these terms; pick the Party and the Role",
        "groupe_commun": "SHARED ITEMS (belong to no party)",
        "groupe_candidats": "CANDIDATES: repeated words not recognised (product? person? place?): your call",
        "groupe_partie": "{p}",
    },
}


def T(cle):
    return TEXTES.get(LANGUE, TEXTES["fr"])[cle]


LARGEURS = [46, 34, 10, 36, 22, 18, 7, 90]
COL_TERME, COL_BALISE, COL_RETENIR, COL_PARTIE, COL_ROLE, COL_NATURE, COL_PAGE = 1, 2, 3, 4, 5, 6, 7
NB_COLONNES = len(LARGEURS)
LIGNE_ENTETES = 7
VALEURS_A_CLASSER = ("À CLASSER", "A CLASSER", "TO SORT")
VALEURS_OUI = ("OUI", "O", "X", "1", "VRAI", "TRUE", "YES", "Y")
FEUILLES_DETECTION = ("Détection", "Detection")
ENTETES_BALISE = ("Balise", "Tag")


# ==================================================================
# MENU DES PARTIES (refonte du 15 octobre 2026)
# ==================================================================
# Avant : un menu de toutes les balises possibles (plus de 1 400 lignes) et,
# en concurrence, deux menus Partie / Role. Desormais un seul chemin : la
# PARTIE (liste courte, chaque partie connue affichee avec son nom, puis
# « nouvelle partie » pour la suivante) et le ROLE ; la balise s'ecrit seule.
# Le libelle du menu commence toujours par le code de la partie suivi de deux
# espaces : la formule et la relecture n'en gardent que le premier mot.

ROLES_ACTEUR_COMMUN = modele.ROLES_ACTEUR_COMMUN
ROLES_PERSONNE = modele.ROLES_PERSONNE
_est_acteur_commun = modele.est_acteur_commun
balises_possibles = modele.balises_possibles


def code_de_libelle(valeur):
    """« DEF1  =  la société X » -> DEF1 ; « SHARED  (...) » -> COMMUN."""
    t = modele._texte(valeur)
    if not t:
        return ""
    return modele.code_propre(t.split()[0])


def _terme_court(terme, n=32):
    """Nom affiche dans le menu : la plus longue ecriture, sans « la société »."""
    variantes = modele.variantes_de(terme) or [""]
    nues = [x for x in variantes
            if not modele.plier(x).startswith(("LA SOCIETE ", "SOCIETE ", "THE COMPANY "))]
    v = (nues or variantes)[0]
    return v if len(v) <= n else v[:n - 1].rstrip() + "…"


def _bases_parties(config, type_doc):
    """[(code de base, numerote ?)] : roles du type (tous les types s'il est
    inconnu), acteurs communs, parties du menu."""
    types = config.get("types_documents") or {}
    if type_doc in types:
        roles = types[type_doc].get("roles", [])
    else:
        roles = [r for d in types.values() for r in d.get("roles", [])]
    sortie, vus = [], set()
    for r in roles + config.get("roles_communs", []):
        c = modele.code_propre(r.get("code", ""))
        if c and c not in vus:
            vus.add(c)
            sortie.append((c, bool(r.get("pluriel"))))
    for r in modele.parties_menu(config, type_doc if type_doc in types else None):
        if r["code"] not in vus:
            vus.add(r["code"])
            sortie.append((r["code"], True))
    return sortie


def _suivant(base, numerote, codes_pris):
    """Prochain code libre pour une base : DEF3 si DEF1 et DEF2 sont pris ;
    PRESTATAIRE, puis PRESTATAIRE2, pour un role non numerote."""
    numeros = []
    for c in codes_pris:
        m = re.fullmatch(re.escape(base) + r"(\d*)", c)
        if m:
            numeros.append(int(m.group(1) or 1))
    if not numeros:
        return f"{base}1" if numerote else base
    return f"{base}{max(numeros) + 1}"


def menu_parties(config, type_doc, connues):
    """connues : [(code, terme d'identite ou "")] dans l'ordre (onglet puis
    detection). Retourne (parties connues [(code, libelle)],
    nouvelles parties [(base, libelle)] dont COMMUN en dernier).

    Une « nouvelle partie » n'a pas de numero dans le menu (« SOCIETE+ ») :
    le tableau donne le suivant tout seul, a chaque ligne ou on la choisit,
    sans limite (15 octobre 2026)."""
    sortie, vus = [], set()
    for code, terme in connues:
        code = modele.code_propre(code)
        if not code or code == modele.COMMUN or code in vus:
            continue
        vus.add(code)
        sortie.append((code, f"{code}  =  {_terme_court(terme)}" if terme else code))
    nouvelles = [(base, f"{base}+  {T('nouvelle')}") for base, _ in _bases_parties(config, type_doc)]
    commun = modele.partie_affichee(modele.COMMUN)
    nouvelles.append((modele.COMMUN, f"{commun}  {T('commun_menu')}"))
    return sortie, nouvelles


def parties_proposees(config, type_doc, existantes=(), detectees=()):
    """Codes de parties (conversion des anciens onglets)."""
    connues = [(c, "") for c in list(detectees) + list(existantes) if c not in VALEURS_A_CLASSER]
    return [c for c, _ in menu_parties(config, type_doc, connues)[0]]


def _numero(code):
    m = re.fullmatch(r"(.*?)(\d*)", code or "")
    return m.group(1), int(m.group(2) or 0)


def _cle_et_rang(balise, config):
    """[ADRESSE3] -> ("ADRESSE", 3) ; [DEF1_ADRESSE2] -> ("DEF1_ADRESSE", 2) ;
    [DEF1_ADRESSE] -> ("DEF1_ADRESSE", 1) ; [DEF1] -> ("DEF1", 1)."""
    partie, role = modele.decomposer_balise(balise, config)
    if not partie:
        return None, 0
    sfx = modele.suffixe_de(role, config)
    coeur = (balise or "").strip()[1:-1]
    if partie == modele.COMMUN:
        cle = sfx
    else:
        cle = f"{partie}_{sfx}" if sfx else partie
    reste = coeur[len(cle):] if coeur.startswith(cle) else ""
    return cle, int(reste) if reste.isdigit() else 1


# Colonnes de calcul masquees de la feuille Detection (loin des colonnes
# visibles et des colonnes supplementaires de la conversion).
C_CODE, C_BASE, C_EFF, C_SFX, C_CLE, C_RANG, C_MENU, C_MENU_N = range(27, 35)
NB_EMPLACEMENTS = 300            # parties creees dans le tableau, au plus


def _l(c):
    return openpyxl.utils.get_column_letter(c)


def _listes(wb, connues, nouvelles, config, prises_identite, max_bases, rangs, r0, rn):
    """Feuille masquee « Listes » :
       A  menu Partie : parties connues, puis celles creees dans le tableau
          (formules), puis les « nouvelle partie » et COMMUN ;
       B  oui/non ; D/E roles et suffixes ; F identites deja prises ;
       H  lignes « nouvelle partie » ; J/K plus grand numero connu par base ;
       M/N rang deja utilise par cle de balise (onglet) ; P1/P2 compteurs."""
    ws = wb.create_sheet("Listes")
    for col, t in enumerate(T("listes"), 1):
        ws.cell(row=1, column=col, value=t)
    for i, v in enumerate(T("oui_non").split(","), 2):
        ws.cell(row=i, column=2, value=v)
    roles_ = modele._attributs_complets(config)
    for i, (_, sfx, aff) in enumerate(roles_, 2):
        ws.cell(row=i, column=4, value=aff)
        ws.cell(row=i, column=5, value=sfx)
    ws.cell(row=1, column=6, value="Prises")
    for i, b in enumerate(sorted(prises_identite), 2):
        ws.cell(row=i, column=6, value=b)
    ws.cell(row=1, column=8, value="Nouvelles")
    for i, (_, libelle) in enumerate(nouvelles, 2):
        ws.cell(row=i, column=8, value=libelle)
    ws.cell(row=1, column=10, value="Base")
    ws.cell(row=1, column=11, value="Max")
    for i, (base, n) in enumerate(sorted(max_bases.items()), 2):
        ws.cell(row=i, column=10, value=base)
        ws.cell(row=i, column=11, value=n)
    ws.cell(row=1, column=13, value="Cle")
    ws.cell(row=1, column=14, value="Rang")
    for i, (cle, n) in enumerate(sorted(rangs.items()), 2):
        ws.cell(row=i, column=13, value=cle)
        ws.cell(row=i, column=14, value=n)

    feuille = f"'{T('feuille')}'"
    menu_n = f"{feuille}!${_l(C_MENU_N)}${r0}:${_l(C_MENU_N)}${rn}"
    menu_v = f"{feuille}!${_l(C_MENU)}${r0}:${_l(C_MENU)}${rn}"
    ws["P1"] = f'=COUNTIF({menu_v},"?*")'                  # parties creees dans le tableau
    k = len(connues)
    ws["P2"] = f"={k}+P1+{len(nouvelles)}"                 # longueur du menu
    for i, (_, libelle) in enumerate(connues, 2):
        ws.cell(row=i, column=1, value=libelle)
    for j in range(1, NB_EMPLACEMENTS + len(nouvelles) + 1):
        ws.cell(row=k + 1 + j, column=1, value=(
            f'=IF({j}<=$P$1,INDEX({menu_v},MATCH({j},{menu_n},0)),'
            f'IFERROR(INDEX($H$2:$H${len(nouvelles) + 1},{j}-$P$1),""))'))
    ws.sheet_state = "hidden"
    return ("OFFSET(Listes!$A$2,0,0,Listes!$P$2,1)", f"Listes!$D$2:$D${len(roles_) + 1}",
            f"Listes!$F$2:$F${max(2, len(prises_identite) + 1)}")


def _formules_ligne(ws, r, r0, existante, oui):
    """Colonnes de calcul d'une ligne et formule de la colonne Balise.

    Code choisi = premier mot du libelle. « SOCIETE+ » : nouvelle partie, dont
    le numero = plus grand numero connu + rang parmi les « SOCIETE+ » du
    tableau. Rang d'un attribut = deja pris dans l'onglet + lignes gardees
    au-dessus : la balise affichee est celle que l'injection ecrira
    ([DEF1_AVOCAT2], [ADRESSE3])."""
    commun = modele.partie_affichee(modele.COMMUN)
    D, E, C, A = f"D{r}", f"E{r}", f"C{r}", f"A{r}"
    code, base, eff, sfx, cle, rang, menu, menu_n = (f"{_l(c)}{r}" for c in range(C_CODE, C_MENU_N + 1))
    col = {c: _l(c) for c in range(C_CODE, C_MENU_N + 1)}
    ws[code] = f'=IF({D}="","",TRIM(LEFT({D},FIND(" ",{D}&" ")-1)))'
    ws[base] = f'=IF(RIGHT({code},1)="+",LEFT({code},LEN({code})-1),"")'
    ws[eff] = (f'=IF({base}="",{code},{base}&(IFERROR(VLOOKUP({base},Listes!$J:$K,2,FALSE),0)'
               f'+COUNTIF({col[C_BASE]}${r0}:{col[C_BASE]}{r},{base})))')
    ws[sfx] = f'=IF({E}="","",IFERROR(VLOOKUP({E},Listes!$D:$E,2,FALSE)&"",UPPER(SUBSTITUTE({E}," ","_"))))'
    f_cle = (f'IF(OR({eff}="",{E}=""),"",IF({eff}="{commun}",{sfx},'
             f'IF({sfx}="",{eff},{eff}&"_"&{sfx})))')
    if existante:
        # Terme deja dans l'onglet : tant que Partie et Role n'ont pas bouge, il
        # ne cree pas de balise (variantes ajoutees a sa ligne) et ne compte
        # donc pas dans le rang des lignes suivantes.
        lib_e, role_e = (str(x or "").replace('"', '""') for x in existante[:2])
        f_cle = f'IF(AND({D}="{lib_e}",{E}="{role_e}"),"",{f_cle})'
    ws[cle] = "=" + f_cle
    ws[rang] = (f'=IF({cle}="","",IFERROR(VLOOKUP({cle},Listes!$M:$N,2,FALSE),0)'
                f'+COUNTIFS({col[C_CLE]}${r0}:{col[C_CLE]}{r},{cle},C${r0}:C{r},"{oui}")'
                f'+IF(UPPER({C})="{oui}",0,1))')
    premier = f'LEFT({A},FIND(" / ",{A}&" / ")-1)'
    ws[menu] = (f'=IF({base}="","",{eff}&"  =  "&IF({sfx}="",LEFT({premier},32),'
                f'"("&{E}&") "&LEFT({premier},26)))')
    ws[menu_n] = f'=IF({menu}="","",COUNTIF({col[C_MENU]}${r0}:{col[C_MENU]}{r},"?*"))'

    a_choisir = '"' + T("a_choisir").replace('"', '""') + '"'
    calcul = (f'IF({eff}="{commun}",IF({sfx}="",{a_choisir},"["&{sfx}&{rang}&"]"),'
              f'IF({sfx}="","["&{eff}&"]","["&{cle}&IF({rang}<=1,"",{rang})&"]"))')
    if existante:
        libelle, role, b = (str(x or "").replace('"', '""') for x in existante)
        calcul = f'IF(AND({D}="{libelle}",{E}="{role}"),"{b}",{calcul})'
    return f'=IF(OR({D}="",{E}=""),{a_choisir},{calcul})'


def _validation(ws, formule, colonne, derniere, stricte=False, invite=None):
    dv = DataValidation(type="list", formula1=formule, allow_blank=True)
    dv.showErrorMessage = stricte       # saisie libre toleree (partie tapee a la main)
    if invite:
        dv.showInputMessage = True
        dv.promptTitle, dv.prompt = invite
    ws.add_data_validation(dv)
    lettre = openpyxl.utils.get_column_letter(colonne)
    dv.add(f"{lettre}{LIGNE_ENTETES + 1}:{lettre}{max(derniere, LIGNE_ENTETES + 300)}")


# ==================================================================
# PRODUCTION DU CLASSEUR
# ==================================================================

def _groupes(lignes):
    """[(cle, titre, couleurs, [lignes])] : A CLASSER, parties, COMMUN, candidats."""
    a_classer, candidats, communs, par_partie, ordre = [], [], [], {}, []
    for l in lignes:
        if l.get("nature_code") == "candidat":
            candidats.append(l)
        elif not l.get("partie"):
            a_classer.append(l)
        elif l["partie"] == modele.COMMUN:
            communs.append(l)
        else:
            if l["partie"] not in par_partie:
                par_partie[l["partie"]] = []
                ordre.append(l["partie"])
            par_partie[l["partie"]].append(l)
    couleurs = modele.couleurs_parties(ordre + [modele.COMMUN])
    groupes = []
    if a_classer:
        groupes.append(("A_CLASSER", T("groupe_a_classer"), modele.COULEUR_A_CLASSER, a_classer))
    for p in ordre:
        groupes.append((p, T("groupe_partie").format(p=p), couleurs[p], par_partie[p]))
    if communs:
        groupes.append((modele.COMMUN, T("groupe_commun"), couleurs[modele.COMMUN], communs))
    if candidats:
        groupes.append(("CANDIDATS", T("groupe_candidats"), modele.COULEUR_CANDIDATS, candidats))
    return groupes


def _identites(lignes, config):
    """[(code, terme)] des parties ayant une ligne Identite, dans l'ordre."""
    sortie = []
    for l in lignes:
        p = modele.code_propre(l.get("partie") or "")
        if p and p != modele.COMMUN and modele.est_identite(l.get("role"), config):
            sortie.append((p, l.get("terme", "")))
    for l in lignes:
        p = modele.code_propre(l.get("partie") or "")
        if p and p != modele.COMMUN and all(p != c for c, _ in sortie):
            sortie.append((p, ""))
    return sortie


def ecrire_classeur(chemin_sortie, lignes, config, entete, parties_liste=(), colonnes_sup=(),
                    titre=None, balises_existantes=(), parties_balises=None, existantes=(),
                    type_doc=None):
    """lignes : dictionnaires (terme, nature, nature_code, partie, role, balise,
    occurrences, confiance, retenir, contexte, remarque, existante ?).
    existantes : entrees de l'onglet vise (partie, role, balise, terme)."""
    from openpyxl.styles import PatternFill, Border, Side
    from openpyxl.formatting.rule import FormulaRule
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = T("feuille")
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"             # impression : tout en largeur
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    fonce, clair = modele.PALETTE[0]
    blanc = "FFFFFF"
    fin = Side(style="thin", color="D0D5DB")
    epais = Side(style="medium", color=fonce)

    # Menu des parties : celles de l'onglet, celles de la detection, puis les
    # « nouvelles parties » et COMMUN.
    connues = _identites(list(existantes), config) + _identites(lignes, config)
    for c in parties_liste or ():
        if c and c not in VALEURS_A_CLASSER and all(modele.code_propre(c) != x for x, _ in connues):
            connues.append((c, ""))
    connues_menu, nouvelles = menu_parties(config, type_doc, connues)
    libelle_de = {c: l for c, l in connues_menu}
    libelle_de[modele.COMMUN] = nouvelles[-1][1]
    # Plus grand numero connu par base (DEF -> 2, SOCIETE -> 5) et rang deja
    # utilise par cle de balise dans l'onglet (ADRESSE -> 3, DEF1_AVOCAT -> 2).
    max_bases = {}
    for c in [c for c, _ in connues_menu] + [modele.code_propre(x.get("partie") or "") for x in existantes]:
        b, n = _numero(c)
        if b and n:
            max_bases[b] = max(max_bases.get(b, 0), n)
    rangs = {}
    for x in existantes:
        if x.get("terme"):
            cle_b, n = _cle_et_rang(x.get("balise"), config)
            if cle_b:
                rangs[cle_b] = max(rangs.get(cle_b, 0), n)
    prises_identite = {x["balise"] for x in existantes
                       if x.get("terme") and modele.est_identite(x.get("role"), config)}

    ws["A1"] = titre or T("titre")
    ws["A1"].font = Font(bold=True, size=13, color=fonce)
    for i, texte in enumerate(entete[:3], 2):
        ws.cell(row=i, column=1, value=texte).font = Font(size=9, color="555555")
    if len(entete) > 2 and entete[2]:
        ws.cell(row=4, column=1).font = Font(size=10, bold=True, color="B00020")
    c = ws.cell(row=5, column=1, value=T("consigne"))
    c.font = Font(size=9, italic=True, color="7A4A1E")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=NB_COLONNES)
    ws.row_dimensions[5].height = 62
    for i in range(2, 5):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=NB_COLONNES)

    en_tetes = T("en_tetes") + [c for c, _ in colonnes_sup]
    largeurs = LARGEURS + [l for _, l in colonnes_sup]
    for i, (t, larg) in enumerate(zip(en_tetes, largeurs), start=1):
        cell = ws.cell(row=LIGNE_ENTETES, column=i,
                       value=modele.en_tete_balise(t, T("sous_titre_balise")) if i == COL_BALISE else t)
        cell.font = Font(bold=True, color=blanc)
        cell.fill = PatternFill("solid", fgColor=fonce)
        cell.border = Border(left=Side(style="thin", color=blanc), right=Side(style="thin", color=blanc))
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = larg
    ws.row_dimensions[LIGNE_ENTETES].height = 30
    ws.freeze_panes = ws.cell(row=LIGNE_ENTETES + 1, column=3)

    r0 = LIGNE_ENTETES + 1
    r = r0
    lignes_controlees = []
    for cle, titre_groupe, (sombre, pale), membres in _groupes(lignes):
        if cle in libelle_de:
            titre_groupe = libelle_de[cle]          # « DEF1  =  VEGA INDUSTRIES SARL »
        cell = ws.cell(row=r, column=1, value=titre_groupe)
        cell.font = Font(bold=True, color=blanc)
        for col in range(1, len(en_tetes) + 1):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=sombre)
        ws.row_dimensions[r].height = 20
        r += 1
        for n, l in enumerate(membres):
            retenir = l.get("retenir", "NON")
            if LANGUE == "en":
                retenir = "YES" if retenir == "OUI" else "NO"
            code = modele.code_propre(l.get("partie") or "")
            partie = libelle_de.get(code, modele.partie_affichee(code)) if code else ""
            role_aff = modele.affichage_role(l.get("role", ""), config) if l.get("role") else ""
            contexte = l.get("contexte", "")
            if l.get("remarque"):
                contexte = "⚠ " + l["remarque"] + "   |   " + contexte
            valeurs = [l.get("terme", ""), l.get("balise", ""), retenir, partie, role_aff,
                       l.get("nature", ""), l.get("page", ""), contexte]
            valeurs += [l.get(k, "") for k, _ in colonnes_sup]
            for i, v in enumerate(valeurs, start=1):
                if i == COL_BALISE:
                    v = _formules_ligne(ws, r, r0, (partie, role_aff, l.get("balise", ""))
                                        if l.get("existante") else None, T("oui"))
                cell = ws.cell(row=r, column=i, value=v if v != "" else None)
                cell.alignment = Alignment(vertical="top", wrap_text=(i in (COL_TERME, COL_BALISE, COL_PARTIE, 8)),
                                           horizontal="center" if i == COL_PAGE else None)
                cell.fill = PatternFill("solid", fgColor=pale if n % 2 == 0 else blanc)
                cell.border = Border(bottom=fin)
                if i == COL_BALISE:
                    # Apercu calcule : texte en gras, sans cadre de saisie
                    cell.font = Font(bold=True, color=sombre if l.get("balise") else "9A6400")
                    if not l.get("balise"):
                        cell.fill = PatternFill("solid", fgColor="FFF3C4")
                elif i in (COL_PARTIE, COL_ROLE):
                    # Les deux vraies cases de saisie : cadre epais
                    cell.font = Font(color="333333")
                    cell.border = Border(left=epais, right=epais, top=fin, bottom=fin)
                    cell.fill = PatternFill("solid", fgColor="FFF8E1" if cle == "A_CLASSER" or not partie else (pale if n % 2 == 0 else blanc))
            if not l.get("existante"):
                lignes_controlees.append(r)
            ws.row_dimensions[r].height = modele.hauteur_ligne(
                [(valeurs[0], LARGEURS[0]), (valeurs[7], LARGEURS[7]), (partie, LARGEURS[COL_PARTIE - 1])])
            r += 1

    rn = max(r - 1, LIGNE_ENTETES + 300)
    for c in range(C_CODE, C_MENU_N + 1):
        ws.column_dimensions[_l(c)].hidden = True
    f_parties, f_roles, f_prises = _listes(wb, connues_menu, nouvelles, config, prises_identite,
                                           max_bases, rangs, r0, rn)
    _validation(ws, f_parties, COL_PARTIE, r, invite=(T("invite_titre"), T("invite")))
    _validation(ws, '"' + T("oui_non") + '"', COL_RETENIR, r, stricte=True)
    _validation(ws, f_roles, COL_ROLE, r)

    # Balise en rouge : identite deja prise (onglet, ou plus haut dans le tableau)
    derniere = max(r, LIGNE_ENTETES + 300)
    plage_b = f"$B${LIGNE_ENTETES + 1}:$B${derniere}"
    plage_c = f"$C${LIGNE_ENTETES + 1}:$C${derniere}"
    for rr in lignes_controlees:
        ws.conditional_formatting.add(f"B{rr}", FormulaRule(
            formula=[f'AND(UPPER($C{rr})="{T("oui")}",${_l(C_SFX)}{rr}="",LEFT($B{rr},1)="[",'
                     f'OR(COUNTIF({f_prises},$B{rr})>0,COUNTIFS({plage_b},$B{rr},{plage_c},"{T("oui")}")>1))'],
            font=Font(bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor="C62828"),
            stopIfTrue=True))
        # Sur une ligne « a choisir », la balise choisie passe de l'ambre au fonce
        ws.conditional_formatting.add(f"B{rr}", FormulaRule(
            formula=[f'LEFT(B{rr},1)="["'], font=Font(bold=True, color=fonce)))

    # Vue d'ensemble : qui est qui (l'onglet d'abord, puis les propositions)
    ws2 = wb.create_sheet("Parties", 1)
    ws2["A1"] = T("qui")
    ws2["A1"].font = Font(bold=True, size=12)
    for i, t in enumerate(T("qui_colonnes"), 1):
        ws2.cell(row=3, column=i, value=t).font = Font(bold=True)
    for col, larg in zip("ABCDE", (16, 22, 32, 60, 24)):
        ws2.column_dimensions[col].width = larg
    rr = 4
    gris = PatternFill("solid", fgColor="EDEDED")
    for x in existantes:
        if not x.get("terme"):
            continue
        for col, v in enumerate((modele.partie_affichee(x.get("partie") or modele.COMMUN),
                                 modele.affichage_role(x.get("role", ""), config),
                                 x.get("balise"), x.get("terme"), T("deja_onglet")), 1):
            ws2.cell(row=rr, column=col, value=v).fill = gris
        rr += 1
    for cle, titre_groupe, (sombre, pale), membres in _groupes(lignes):
        if cle in ("A_CLASSER", "CANDIDATS"):
            continue
        for l in membres:
            if l.get("existante"):
                continue
            for col, v in enumerate((modele.partie_affichee(l.get("partie")),
                                     modele.affichage_role(l.get("role", ""), config),
                                     l.get("balise"), l.get("terme"), T("propose")), 1):
                ws2.cell(row=rr, column=col, value=v).fill = PatternFill("solid", fgColor=pale)
            rr += 1

    wb.save(chemin_sortie)


def lignes_depuis_entites(entites, config):
    lisibles = config.get("libelles_lisibles", {})
    lignes = []
    for e in entites:
        remarque = (e.note or "").replace("partie à choisir", "").strip(" ;")
        remarque = textes.traduire(remarque, LANGUE)
        existante = getattr(e, "existante", None)
        if existante:
            if e.variantes_nouvelles:
                remarque = (T("deja").format(balise=existante['balise']) + T("variantes")
                            + " / ".join(e.variantes_nouvelles) + " ; " + remarque).strip(" ;")
            else:
                remarque = (T("deja").format(balise=existante['balise'])
                            + (" ; " + remarque if remarque else ""))
        lignes.append({
            "terme": e.terme_pour_tableur(),
            "nature": lisibles.get(e.type, e.type),
            "nature_code": e.type,
            "partie": e.partie,
            "role": e.role,
            "balise": e.balise,
            "occurrences": e.occurrences,
            "confiance": e.confiance,
            "page": getattr(e, "page", "") or "",
            "retenir": roles.proposer_a_retenir(e),
            "contexte": e.contexte,
            "remarque": remarque,
            "existante": bool(existante),
        })
    return lignes


def ecrire_classeur_detection(chemin_sortie, entites, infos, config, document, table,
                              onglet, existantes=()):
    types = config.get("types_documents") or {}
    type_doc = infos.get("type") or ""
    libelle_type = types.get(type_doc, {}).get("libelle", T("non_reconnu"))
    mode = T("choisi") if infos.get("type_force") else T("auto")
    ligne_type = T("type").format(t=libelle_type, m=mode)
    if infos.get("sens"):
        ligne_type += T("emis") if infos["sens"] == roles.SENS_EMIS else T("recu")
    entete = [T("document").format(d=os.path.basename(document), t=table, o=onglet), ligne_type]
    avert = []
    if infos.get("modele_onglet") == "ancien":
        avert.append(T("ancien"))
    for a in infos.get("avertissements", []):
        avert.append(T("attention") + textes.traduire(a, LANGUE))
    entete.append("   ".join(avert))
    lignes = lignes_depuis_entites(entites, config)
    ecrire_classeur(chemin_sortie, lignes, config, entete,
                    balises_existantes=[x.get("balise", "") for x in existantes],
                    existantes=existantes, type_doc=type_doc)


# ==================================================================
# RELECTURE ET INJECTION
# ==================================================================

def _ligne_entetes(ws):
    for i in range(1, 15):
        valeurs = [str(c.value or "").strip() for c in ws[i]]
        if len(valeurs) > 1 and any(b.lower() in valeurs[1].lower() for b in ENTETES_BALISE):
            return i, valeurs
    return None, None


def etat_classeur(wb):
    """Compteur des « nouvelles parties » d'un classeur relu : plus grand
    numero connu par base (feuille Listes, J:K) et numeros deja donnes."""
    etat = {"max": {}, "nouvelles": {}}
    if "Listes" in wb.sheetnames:
        for b, n in wb["Listes"].iter_rows(min_row=2, min_col=10, max_col=11, values_only=True):
            if b and isinstance(n, (int, float)):
                etat["max"][str(b)] = int(n)
    return etat


def _partie_de_ligne(valeur, etat=None):
    """Code de partie d'une ligne relue. « SOCIETE+ » (nouvelle partie) :
    numero suivant, comme le calcule la formule du tableau."""
    t = modele._texte(valeur)
    premier = t.split()[0] if t else ""
    if premier.endswith("+") and etat is not None:
        base = modele.code_propre(premier[:-1])
        etat["nouvelles"][base] = etat["nouvelles"].get(base, 0) + 1
        return f"{base}{etat['max'].get(base, 0) + etat['nouvelles'][base]}"
    return code_de_libelle(valeur)


def balise_de_ligne(row, config=None, etat=None):
    """Balise d'une ligne relue : la colonne Balise (valeur calculee par Excel
    ou choisie dans la liste) ; sinon, si le tableur n'a pas calcule la
    formule, PARTIE + ROLE des colonnes de choix."""
    balise = str(row[COL_BALISE - 1] or "").strip()
    if balise.startswith("=") or not balise.startswith("["):
        balise = ""                                        # formule non calculee, ou « à choisir »
    partie = _partie_de_ligne(row[COL_PARTIE - 1], etat if not balise else None) \
        if len(row) >= COL_PARTIE else ""
    role = modele._texte(row[COL_ROLE - 1]) if len(row) >= COL_ROLE else ""
    if not balise and partie and role:
        try:
            balise, _ = modele.calculer_balise(partie, role, set(), config or modele.config_courante())
        except Exception:
            balise = ""
    return balise


def lire_classeur(chemin, colonnes_sup=(), config=None):
    """Lignes marquees OUI avec une balise : [{"terme", "partie", "role", "balise", + colonnes_sup}].
    La partie et le role se deduisent de la balise choisie."""
    wb = openpyxl.load_workbook(chemin, data_only=True)
    try:
        nom_feuille = next((f for f in FEUILLES_DETECTION if f in wb.sheetnames), None)
        if not nom_feuille:
            raise RuntimeError("Ce classeur n'est pas un classeur de détection "
                               "(onglet « Détection » absent). / Not a detection workbook.")
        ws = wb[nom_feuille]
        i_ent, entetes = _ligne_entetes(ws)
        if not i_ent:
            raise RuntimeError(
                "Ce classeur de détection a été produit par une ancienne version de l'outil.\n\n"
                "Relancez la détection sur le document.")
        if "Partie" in entetes or "Party" in entetes:
            raise RuntimeError(
                "Ce classeur de détection a été produit par l'ancienne version de l'outil "
                "(colonnes Partie et Rôle).\n\nRelancez la détection sur le document : le nouveau "
                "classeur propose la balise à choisir dans une liste.")
        idx = {t: i for i, t in enumerate(entetes)}
        etat = etat_classeur(wb)
        retenues = []
        for row in ws.iter_rows(min_row=i_ent + 1, values_only=True):
            if not row or not row[COL_TERME - 1]:
                continue
            if len(row) < COL_NATURE or not row[COL_NATURE - 1]:
                continue                                   # titre de groupe
            if str(row[COL_RETENIR - 1] or "").strip().upper() not in VALEURS_OUI:
                # une « nouvelle partie » non gardee prend tout de meme son numero
                _partie_de_ligne(row[COL_PARTIE - 1], etat)
                continue
            balise = balise_de_ligne(row, config, etat=etat)
            if not balise:
                continue
            ligne = {"terme": modele._texte(row[COL_TERME - 1]), "balise": balise,
                     "partie": "", "role": ""}
            for cle, _ in colonnes_sup:
                i = idx.get(cle)
                ligne[cle] = modele._texte(row[i]) if i is not None and i < len(row) else ""
            retenues.append(ligne)
        return retenues
    finally:
        wb.close()


def lire_classeur_detection(chemin, config=None):
    retenues = lire_classeur(chemin, config=config)
    if config is not None:
        for l in retenues:
            partie, role = modele.decomposer_balise(l.get("balise"), config)
            l["partie"], l["role"] = partie or "", role or ""
    return retenues


def injecter(chemin_excel, nom_onglet, retenues, config):
    """Voir modele_onglet.injecter. Une ligne relue par lire_classeur porte sa
    balise : la partie et le role en sont deduits ici si besoin."""
    for l in retenues:
        if l.get("balise") and not (l.get("partie") and l.get("role")):
            partie, role = modele.decomposer_balise(l["balise"], config)
            l["partie"], l["role"] = partie or "", role or ""
    return modele.injecter(chemin_excel, nom_onglet, retenues, config)
