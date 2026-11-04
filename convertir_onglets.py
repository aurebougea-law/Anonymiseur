"""
convertir_onglets.py
--------------------
Conversion d'un onglet de l'ANCIEN modele (cases numerotees [PM_DEF_1],
[ADR_DEM_2]...) vers le modele PAR PARTIE ([DEFENDEUR1], [DEMANDEUR1_ADRESSE]...).

DEUX TEMPS, comme la detection
    1. PREPARER  : lit l'onglet et produit un classeur _Conversion_<onglet>.xlsx
                   ou chaque ancienne ligne recoit une partie et un role
                   proposes. Vous relisez, corrigez, enregistrez.
    2. APPLIQUER : reecrit l'onglet au nouveau modele d'apres le classeur relu.

CE QUI PROTEGE LES DOCUMENTS DEJA ANONYMISES
    Avant toute reecriture, l'onglet est copie tel quel dans un classeur
    d'ARCHIVE place a cote de la table :
        <table>_ARCHIVE_anciennes_balises.xlsx
    Cette archive est ajoutee automatiquement a config.json : elle apparait
    dans la liste « Table » de l'outil. Pour restituer SANS journal un document
    anonymise avec les anciennes balises, choisissez cette table.
    La restitution AVEC journal n'est pas concernee : le journal suffit.

Lancement : double-clic sur « Convertir anciens onglets.bat »

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import os
import re
import sys
import json
import shutil
import pathlib

import openpyxl

SCRIPT_DIR = pathlib.Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

import modele_onglet as modele
import detection_classeur as classeur

SUFFIXE_ARCHIVE = "_ARCHIVE_anciennes_balises"
COLONNES_SUP = [("Ancienne balise", 24), ("Actif", 8), ("Notes", 34), ("Dossier", 14)]

MOTIF_CAMP = re.compile(r"^\[(?P<cat>[A-Z_]+?)_(?P<camp>DEM|DEF)(?:_(?P<n>\d+|PLAID|CONST))?\]$")
MOTIF_NEUTRE = re.compile(r"^\[(?P<cat>[A-Z_]+?)_?(?P<n>\d+)\]$")

# Ancienne categorie « de camp » -> role dans le nouveau modele
ROLES_CAMP = {
    "ADR": "Adresse", "SIRET": "SIRET", "RL": "Dirigeant", "MAIL": "E-mail",
    "SIT": "Site web", "TEL": "Téléphone", "LOG": "Logiciel", "MRQ": "Marque",
    "BRV": "Brevet", "DM": "Dessin et modèle", "PRD": "Produit", "ENS": "Enseigne",
}
AVOCATS = {("AV", "PLAID"): "Avocat", ("AV", "CONST"): "Avocat postulant",
           ("ADR_AV", "PLAID"): "Adresse avocat", ("ADR_AV", "CONST"): "Adresse avocat postulant"}

# Ancienne categorie neutre -> (role d'acteur numerote | COMMUN, role)
NEUTRES = {
    "CJ": ("COMMISSAIRE", "Identité"), "ADR_CJ": ("COMMISSAIRE", "Adresse"),
    "EXPERT": ("EXPERT", "Identité"), "MAGISTRAT": ("MAGISTRAT", "Identité"),
    "GREFFIER": ("GREFFIER", "Identité"),
    "TI": ("TIERS", "Identité"), "CLIENT_TIERS": ("TIERS", "Identité"),
    "SAL": ("SALARIE", "Identité"),
    "RG": (modele.COMMUN, "Référence RG"), "DEPOT_INPI": (modele.COMMUN, "Dépôt INPI"),
    "DEPOT_EUR": (modele.COMMUN, "Dépôt européen"),
    "VERSION_LOG": (modele.COMMUN, "Version logiciel"),
    "HASH_PIECE": (modele.COMMUN, "Empreinte fichier"),
    "URL_DEPOT": (modele.COMMUN, "URL dépôt"), "CLE_LICENCE": (modele.COMMUN, "Clé de licence"),
    "MONTANT": (modele.COMMUN, "Montant"), "PIECE": (modele.COMMUN, "Pièce"),
    "ADR_SUPP": (modele.COMMUN, "Adresse"), "DEPOT": (modele.COMMUN, "Dépôt INPI"),
    "MRQ": (modele.COMMUN, "Marque"), "BRV": (modele.COMMUN, "Brevet"),
    "DM": (modele.COMMUN, "Dessin et modèle"), "LOG": (modele.COMMUN, "Logiciel"),
    "PRD": (modele.COMMUN, "Produit"),
}


class ErreurConversion(RuntimeError):
    pass


# ==================================================================
# LECTURE DE L'ANCIEN ONGLET
# ==================================================================

def onglets_ancien_modele(chemin_table, config_outil=None):
    import anonymisation_core as coeur
    sortie = []
    wb = openpyxl.load_workbook(chemin_table, read_only=True)
    try:
        masques = set((config_outil or {}).get("onglets_masques", ())) | set(coeur.ONGLETS_MASQUES)
        suffixes = tuple((config_outil or {}).get("suffixes_onglets_masques", ())) \
            + coeur.SUFFIXES_ONGLETS_MASQUES
        for ws in wb.worksheets:
            if ws.title in masques or ws.title.endswith(suffixes):
                continue
            if modele.modele_de(ws) == "ancien":
                sortie.append(ws.title)
    finally:
        wb.close()
    return sortie


def lire_ancien_onglet(chemin_table, onglet):
    """Lignes REMPLIES de l'ancien onglet (les cases vides ne sont pas reprises)."""
    wb = openpyxl.load_workbook(chemin_table, read_only=True, data_only=True)
    try:
        if onglet not in wb.sheetnames:
            raise ErreurConversion(f"Onglet « {onglet} » introuvable.")
        ws = wb[onglet]
        if modele.modele_de(ws) != "ancien":
            raise ErreurConversion(f"L'onglet « {onglet} » n'est pas à l'ancien modèle "
                                   "(déjà converti ?).")
        lignes = []
        for row in ws.iter_rows(min_row=1, values_only=True):
            v = [modele._texte(x) for x in (list(row or []) + [None] * 7)[:7]]
            terme, balise = v[0], v[1]
            if terme and balise.startswith("[") and balise.endswith("]"):
                lignes.append({"terme": terme, "balise": balise, "categorie": v[2],
                               "camp": v[3], "dossier": v[4], "actif": v[5], "notes": v[6]})
        return lignes
    finally:
        wb.close()


# ==================================================================
# PROPOSITION DE CORRESPONDANCE
# ==================================================================

def _pluriel(code, config):
    for d in (config.get("types_documents") or {}).values():
        for r in d.get("roles", []):
            if r["code"] == code:
                return bool(r.get("pluriel"))
    for r in config.get("roles_communs", []):
        if r["code"] == code:
            return bool(r.get("pluriel"))
    return False


def _code(role, rang, config):
    role = modele.code_propre(role)
    if _pluriel(role, config):
        return f"{role}{rang}"
    return role if rang == 1 else f"{role}{rang}"


def proposer(lignes, role_dem, role_def, config):
    """Ancienne ligne -> {partie, role, confiance, remarque}. Mecanique et
    prudente : tout rattachement incertain est marque « à vérifier »."""
    roles_camp = {"DEM": role_dem, "DEF": role_def}
    analyse = []
    for l in lignes:
        m = MOTIF_CAMP.match(l["balise"])
        if m:
            analyse.append((l, "camp", m.group("cat"), m.group("camp"), m.group("n") or "1"))
            continue
        m = MOTIF_NEUTRE.match(l["balise"])
        if m:
            analyse.append((l, "neutre", m.group("cat"), None, m.group("n")))
            continue
        analyse.append((l, "inconnu", None, None, None))

    # Les parties de chaque camp : les societes d'abord, a defaut les personnes
    parties_camp = {}
    for camp in ("DEM", "DEF"):
        pm = sorted((a for a in analyse if a[1] == "camp" and a[3] == camp and a[2] == "PM"),
                    key=lambda a: int(a[4]) if a[4].isdigit() else 0)
        pp = sorted((a for a in analyse if a[1] == "camp" and a[3] == camp and a[2] == "PP"),
                    key=lambda a: int(a[4]) if a[4].isdigit() else 0)
        tetes = pm or pp
        parties_camp[camp] = [(id(a[0]), _code(roles_camp[camp], i, config))
                              for i, a in enumerate(tetes, 1)]
        if not tetes and any(a[1] == "camp" and a[3] == camp for a in analyse):
            parties_camp[camp] = [(None, _code(roles_camp[camp], 1, config))]

    sortie = []
    for l, genre, cat, camp, n in analyse:
        p = {"partie": "", "role": "", "confiance": "haute", "remarque": ""}
        if genre == "camp":
            parties = parties_camp[camp]
            codes = [c for _, c in parties]
            propre = next((c for i, c in parties if i == id(l)), None)
            plusieurs = len(parties) > 1
            if propre:
                p["partie"], p["role"] = propre, "Identité"
            elif cat in ("PM", "PP"):
                p["partie"], p["role"] = codes[0], "Personne liée"
                p["confiance"] = "à vérifier"
                p["remarque"] = ("personne citée avec la société de ce camp : dirigeant, "
                                 "signataire ou partie à part entière ? Corrigez Partie et Rôle")
            elif (cat, n) in AVOCATS or cat in ("AV", "ADR_AV"):
                p["partie"] = codes[0]
                p["role"] = AVOCATS.get((cat, n), "Avocat" if cat == "AV" else "Adresse avocat")
                if plusieurs:
                    p["confiance"] = "à vérifier"
                    p["remarque"] = "plusieurs parties dans ce camp : de qui est-ce l'avocat ?"
            elif cat in ROLES_CAMP:
                rang = int(n) if n.isdigit() else 1
                p["partie"] = codes[rang - 1] if plusieurs and rang <= len(codes) else codes[0]
                p["role"] = ROLES_CAMP[cat]
                if plusieurs:
                    p["confiance"] = "à vérifier"
                    p["remarque"] = "plusieurs parties dans ce camp : vérifiez le rattachement"
            else:
                p["partie"] = codes[0]
                p["confiance"] = "à vérifier"
                p["remarque"] = f"catégorie « {cat} » inconnue : choisissez le rôle"
        elif genre == "neutre" and cat in NEUTRES:
            acteur, role = NEUTRES[cat]
            p["partie"] = acteur if acteur == modele.COMMUN else f"{acteur}{n}"
            p["role"] = role
        else:
            p["confiance"] = "à vérifier"
            p["remarque"] = "balise non reconnue : choisissez la partie et le rôle"
        sortie.append((l, p))
    return sortie


# ==================================================================
# PREPARER
# ==================================================================

def _nom_fichier(texte):
    return re.sub(r'[\\/:*?"<>|]+', "_", texte).strip() or "onglet"


def preparer(chemin_table, onglet, role_dem, role_def, config, dossier_sortie=None):
    lignes = lire_ancien_onglet(chemin_table, onglet)
    if not lignes:
        raise ErreurConversion(f"L'onglet « {onglet} » ne contient aucun terme : rien à convertir.\n\n"
                               "Supprimez-le ou gardez-le tel quel ; pour ce dossier, créez un "
                               "nouvel onglet depuis la fenêtre de détection.")
    propositions = proposer(lignes, role_dem, role_def, config)
    prises, sortie = set(), []
    for l, p in propositions:
        balise = ""
        if p["partie"] and p["role"]:
            try:
                balise, deja = modele.calculer_balise(p["partie"], p["role"], prises, config)
                prises.add(balise)
            except modele.ErreurModele:
                balise = ""
        sortie.append({
            "terme": l["terme"], "nature": l["categorie"] or "", "partie": p["partie"],
            "role": p["role"], "balise": balise, "occurrences": "",
            "confiance": p["confiance"], "retenir": "OUI", "contexte": "",
            "remarque": p["remarque"], "Ancienne balise": l["balise"],
            "Actif": l["actif"], "Notes": l["notes"], "Dossier": l["dossier"],
        })
    dossier_sortie = dossier_sortie or os.path.join(os.path.dirname(os.path.abspath(chemin_table)),
                                                    "_Conversions")
    os.makedirs(dossier_sortie, exist_ok=True)
    chemin = os.path.join(dossier_sortie, f"_Conversion_{_nom_fichier(onglet)}.xlsx")
    entete = [f"Table : {os.path.basename(chemin_table)}    |    Onglet : {onglet}",
              f"Camp Demande -> {role_dem}    |    Camp Défense -> {role_def}",
              "Toutes les lignes remplies de l'ancien onglet sont reprises. Une ligne mise à NON "
              "sera retirée de l'onglet (elle restera dans l'archive).",
              "Ce classeur contient les vrais termes : supprimez-le après la conversion."]
    parties = classeur.parties_proposees(config, "", [], [s["partie"] for s in sortie if s["partie"]])
    classeur.ecrire_classeur(chemin, sortie, config, entete, parties, COLONNES_SUP,
                             titre="CONVERSION — proposition à relire avant application")
    return chemin, sortie


# ==================================================================
# ARCHIVE
# ==================================================================

def chemin_archive(chemin_table):
    base, ext = os.path.splitext(chemin_table)
    return base + SUFFIXE_ARCHIVE + ext


def archiver_onglet(chemin_table, onglet):
    """Copie l'onglet, dans son etat actuel, dans le classeur d'archive.
    Premiere fois : copie integrale de la table. Ensuite : l'onglet seul est
    remplace dans l'archive, les autres onglets archives restent intacts."""
    archive = chemin_archive(chemin_table)
    if not os.path.exists(archive):
        shutil.copy2(chemin_table, archive)
        return archive
    src = openpyxl.load_workbook(chemin_table, data_only=False)
    dst = openpyxl.load_workbook(archive)
    try:
        ws_src = src[onglet]
        if onglet in dst.sheetnames:
            del dst[onglet]
        ws = dst.create_sheet(onglet)
        for row in ws_src.iter_rows():
            for c in row:
                if c.value is not None:
                    ws.cell(row=c.row, column=c.column, value=c.value)
        for lettre, dim in ws_src.column_dimensions.items():
            ws.column_dimensions[lettre].width = dim.width
        try:
            dst.save(archive)
        except PermissionError:
            raise ErreurConversion("Le classeur d'archive est ouvert dans Excel : fermez-le puis "
                                   "recommencez. Rien n'a été modifié.")
    finally:
        src.close()
        dst.close()
    return archive


def declarer_archive(chemin_config, nom_table, archive):
    """Ajoute l'archive a config.json si elle n'y est pas deja."""
    with open(chemin_config, encoding="utf-8") as f:
        brut = json.load(f)
    tables = brut.setdefault("tables", [])
    cible = os.path.normcase(os.path.abspath(archive))
    for t in tables:
        fichier = t.get("fichier", "") if isinstance(t, dict) else str(t)
        if os.path.normcase(os.path.abspath(fichier)) == cible:
            return False
    tables.append({"nom": f"{nom_table} - ARCHIVE anciennes balises", "fichier": archive})
    with open(chemin_config, "w", encoding="utf-8") as f:
        json.dump(brut, f, ensure_ascii=False, indent=2)
    return True


# ==================================================================
# APPLIQUER
# ==================================================================

def appliquer(chemin_table, onglet, chemin_classeur, config, chemin_config=None, nom_table=""):
    """Reecrit l'onglet au modele par partie. Tout ou rien : a la moindre
    ligne incomplete, rien n'est ecrit."""
    lignes_onglet = lire_ancien_onglet(chemin_table, onglet)
    toutes = _toutes_lignes(chemin_classeur, config)
    attendu = sorted((l["balise"], l["terme"]) for l in lignes_onglet)
    vu = sorted((t["Ancienne balise"], t["terme"]) for t in toutes)
    if attendu != vu:
        raise ErreurConversion(
            "L'onglet a changé depuis la préparation du classeur de conversion (ou le classeur "
            "ne correspond pas à cet onglet).\n\nRelancez « Préparer » : rien n'a été modifié.")

    retenues = [t for t in toutes if t["retenir"]]
    problemes = [f"  - {t['terme'][:40]} ({t['Ancienne balise']}) : "
                 + ("partie manquante" if not t["partie"] else "rôle manquant")
                 for t in retenues if not t["partie"] or not t["role"]]
    if problemes:
        raise ErreurConversion("Lignes incomplètes dans le classeur (partie ou rôle) :\n\n"
                               + "\n".join(problemes[:15])
                               + "\n\nComplétez-les ou mettez-les à NON. Rien n'a été modifié.")

    prises, entrees = set(), []
    for t in retenues:
        try:
            balise, deja = modele.calculer_balise(t["partie"], t["role"], prises, config)
        except modele.ErreurModele as e:
            raise ErreurConversion(f"{t['terme'][:40]} : {e}. Rien n'a été modifié.")
        if deja:
            raise ErreurConversion(
                f"Deux lignes ont la même partie avec le rôle « Identité » ({balise}) :\n"
                f"  {t['terme'][:50]}\n\nDonnez-leur deux parties différentes (ex. DEMANDEUR1 et "
                "DEMANDEUR2), ou un autre rôle. Rien n'a été modifié.")
        prises.add(balise)
        notes = (t["Notes"] + " ; " if t["Notes"] else "") + f"ancienne balise {t['Ancienne balise']}"
        entrees.append({"terme": t["terme"], "balise": balise, "role": t["role"],
                        "partie": modele.code_propre(t["partie"]), "dossier": t["Dossier"],
                        "actif": t["Actif"] or "OUI", "notes": notes})

    archive = archiver_onglet(chemin_table, onglet)
    declaree = declarer_archive(chemin_config, nom_table, archive) if chemin_config else False
    titres = [[modele.TITRE_1] + [""] * 6, [modele.TITRE_2] + [""] * 6]
    modele.reecrire_onglet(chemin_table, onglet, entrees, config, titres=titres)
    retirees = len(toutes) - len(retenues)
    return {"archive": archive, "archive_declaree": declaree, "ecrites": len(entrees),
            "retirees": retirees, "correspondances": [(t["Ancienne balise"], e["balise"])
                                                      for t, e in zip(retenues, entrees)]}


def _toutes_lignes(chemin_classeur, config):
    wb = openpyxl.load_workbook(chemin_classeur, data_only=True)
    try:
        if "Détection" not in wb.sheetnames:
            raise ErreurConversion("Ce fichier n'est pas un classeur de conversion.")
        ws = wb["Détection"]
        entetes = [str(c.value or "").strip() for c in ws[classeur.LIGNE_ENTETES]]
        if "Ancienne balise" not in entetes:
            raise ErreurConversion("Ce fichier n'est pas un classeur de CONVERSION "
                                   "(colonne « Ancienne balise » absente).")
        idx = {t: i for i, t in enumerate(entetes)}
        etat = classeur.etat_classeur(wb)
        sortie = []
        for row in ws.iter_rows(min_row=classeur.LIGNE_ENTETES + 1, values_only=True):
            if not row or not row[classeur.COL_TERME - 1]:
                continue
            if len(row) < classeur.COL_NATURE or not row[classeur.COL_NATURE - 1]:
                continue                                   # titre de groupe
            balise = classeur.balise_de_ligne(row, config, etat=etat)
            partie, role = modele.decomposer_balise(balise, config)
            ligne = {"terme": modele._texte(row[classeur.COL_TERME - 1]),
                     "partie": partie or "",
                     "role": role or "",
                     "retenir": modele._texte(row[classeur.COL_RETENIR - 1]).upper()
                                in classeur.VALEURS_OUI}
            for cle, _ in COLONNES_SUP:
                ligne[cle] = modele._texte(row[idx[cle]]) if idx.get(cle) is not None else ""
            sortie.append(ligne)
        return sortie
    finally:
        wb.close()


# ==================================================================
# INTERFACE
# ==================================================================

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    import anonymisation_core as coeur
    import detection_core as det

    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    BG, BLEU, VERT, ROUGE = "#F5F5F5", "#1A5276", "#1E8449", "#C0392B"
    # Installation neuve : config.json et detection_config.json sont dans le
    # sous-dossier Données (8 octobre 2026) ; sinon a cote du programme.
    dossier_donnees = SCRIPT_DIR
    if not (SCRIPT_DIR / "config.json").exists():
        for d in ("Données", "Data"):
            if (SCRIPT_DIR / d / "config.json").exists():
                dossier_donnees = SCRIPT_DIR / d
                break
    chemin_config = dossier_donnees / "config.json"
    try:
        config_outil = coeur.charger_config(chemin_config)
    except Exception as e:
        tk.Tk().withdraw()
        messagebox.showerror("Configuration", str(e))
        return
    config_det = det.charger_config_detection(dossier_donnees / "detection_config.json")
    tables, _ = coeur.verifier_tables(config_outil)
    tables = [t for t in tables if SUFFIXE_ARCHIVE not in t["fichier"]]
    types = config_det.get("types_documents") or {}

    app = tk.Tk()
    app.title("Conversion des anciens onglets")
    app.configure(bg=BG)
    v_table, v_onglet, v_type = tk.StringVar(), tk.StringVar(), tk.StringVar()
    v_dem, v_def, v_classeur = tk.StringVar(), tk.StringVar(), tk.StringVar()
    PAD = 18

    def table_courante():
        return next((t for t in tables if t["nom"] == v_table.get()), None)

    def maj_onglets(*_):
        t = table_courante()
        if not t:
            return
        try:
            noms = onglets_ancien_modele(t["fichier"], config_outil)
        except Exception as e:
            messagebox.showerror("Lecture impossible", str(e))
            return
        c_onglet.config(values=noms)
        v_onglet.set(noms[0] if noms else "")
        statut.config(text="" if noms else "Aucun onglet à l'ancien modèle dans cette table.",
                      fg="#555")

    def maj_roles(*_):
        code = next((c for c, d in types.items() if d.get("libelle") == v_type.get()), "")
        codes = [r["code"] for r in types.get(code, {}).get("roles", [])]
        tous = codes + [r["code"] for r in config_det.get("roles_communs", [])]
        c_dem.config(values=tous)
        c_def.config(values=tous)
        if len(codes) >= 2:
            v_dem.set(codes[0])
            v_def.set(codes[1])

    def ligne(texte):
        f = tk.Frame(app, bg=BG)
        f.pack(fill="x", padx=PAD, pady=(8, 0))
        tk.Label(f, text=texte, font=("Segoe UI", 10), bg=BG, width=24, anchor="w").pack(side="left")
        return f

    tk.Label(app, text="Conversion des anciens onglets", font=("Segoe UI", 13, "bold"),
             bg=BG, fg=BLEU).pack(pady=(PAD, 2))
    tk.Label(app, text="Une archive de l'onglet est faite avant toute modification.",
             font=("Segoe UI", 8), bg=BG, fg="#777").pack()
    ttk.Separator(app, orient="horizontal").pack(fill="x", padx=PAD, pady=10)

    c_table = ttk.Combobox(ligne("Table :"), textvariable=v_table,
                           values=[t["nom"] for t in tables], state="readonly", width=34)
    c_table.pack(side="left")
    c_table.bind("<<ComboboxSelected>>", maj_onglets)
    c_onglet = ttk.Combobox(ligne("Onglet à convertir :"), textvariable=v_onglet,
                            values=[], state="readonly", width=34)
    c_onglet.pack(side="left")
    c_type = ttk.Combobox(ligne("Type de dossier :"), textvariable=v_type,
                          values=[d.get("libelle", c) for c, d in types.items()],
                          state="readonly", width=44)
    c_type.pack(side="left")
    c_type.bind("<<ComboboxSelected>>", maj_roles)
    c_dem = ttk.Combobox(ligne("Le camp Demande devient :"), textvariable=v_dem, width=26)
    c_dem.pack(side="left")
    c_def = ttk.Combobox(ligne("Le camp Défense devient :"), textvariable=v_def, width=26)
    c_def.pack(side="left")

    ttk.Separator(app, orient="horizontal").pack(fill="x", padx=PAD, pady=14)

    def preparer_clic():
        t = table_courante()
        if not t or not v_onglet.get() or not v_dem.get() or not v_def.get():
            messagebox.showwarning("Incomplet", "Choisissez la table, l'onglet, le type de "
                                   "dossier et ce que deviennent les deux camps.")
            return
        try:
            chemin, lignes = preparer(t["fichier"], v_onglet.get(), v_dem.get(), v_def.get(),
                                      config_det)
        except Exception as e:
            messagebox.showerror("Préparation impossible", str(e))
            return
        v_classeur.set(chemin)
        a_verifier = sum(1 for l in lignes if l["confiance"] != "haute")
        statut.config(text=f"Classeur prêt : {os.path.basename(chemin)}", fg=VERT)
        messagebox.showinfo("Classeur de conversion prêt",
                            f"{len(lignes)} ligne(s) reprise(s), dont {a_verifier} à vérifier.\n\n"
                            f"Classeur :\n{chemin}\n\nOuvrez-le, corrigez la Partie et le Rôle des "
                            "lignes « à vérifier », enregistrez, fermez-le, puis cliquez sur "
                            "« 2. Appliquer ».")

    def choisir_classeur():
        c = filedialog.askopenfilename(title="Classeur de conversion relu",
                                       filetypes=[("Classeurs Excel", "*.xlsx")])
        if c:
            v_classeur.set(c)

    def appliquer_clic():
        t = table_courante()
        if not t or not v_onglet.get() or not v_classeur.get():
            messagebox.showwarning("Incomplet", "Choisissez la table, l'onglet et le classeur relu.")
            return
        if not messagebox.askyesno(
                "Confirmer la conversion",
                f"L'onglet « {v_onglet.get()} » va être réécrit au modèle par partie.\n\n"
                "Son état actuel sera d'abord copié dans l'archive :\n"
                f"{os.path.basename(chemin_archive(t['fichier']))}\n\nContinuer ?"):
            return
        try:
            res = appliquer(t["fichier"], v_onglet.get(), v_classeur.get(), config_det,
                            str(chemin_config), t["nom"])
        except Exception as e:
            statut.config(text="Conversion non faite — voir le message", fg=ROUGE)
            messagebox.showerror("Conversion non faite", str(e))
            return
        exemples = "\n".join(f"    {a}  ->  {n}" for a, n in res["correspondances"][:12])
        msg = (f"Onglet converti : {res['ecrites']} ligne(s).\n\n{exemples}\n\n"
               f"Archive : {os.path.basename(res['archive'])}\n")
        if res["archive_declaree"]:
            msg += ("L'archive a été ajoutée à la liste des tables. Pour restituer SANS journal "
                    "un ancien document de ce dossier, choisissez-la comme Table.\n")
        if res["retirees"]:
            msg += f"\n{res['retirees']} ligne(s) mise(s) à NON ont été retirées (elles sont dans l'archive).\n"
        msg += "\nPensez à supprimer le classeur de conversion : il contient les vrais termes."
        statut.config(text=f"OK — onglet « {v_onglet.get()} » converti", fg=VERT)
        messagebox.showinfo("Conversion terminée", msg)
        maj_onglets()

    tk.Button(app, text="1.  Préparer la conversion (classeur à relire)", command=preparer_clic,
              font=("Segoe UI", 11, "bold"), bg=BLEU, fg="white", relief="flat", pady=8,
              cursor="hand2").pack(fill="x", padx=PAD)
    f = ligne("Classeur relu :")
    tk.Entry(f, textvariable=v_classeur, font=("Segoe UI", 9), state="readonly",
             relief="solid").pack(side="left", fill="x", expand=True, padx=(0, 6))
    tk.Button(f, text="Parcourir...", command=choisir_classeur, font=("Segoe UI", 9),
              bg="#E8E8E8", relief="solid", bd=1, cursor="hand2").pack(side="left")
    tk.Button(app, text="2.  Appliquer la conversion", command=appliquer_clic,
              font=("Segoe UI", 11, "bold"), bg=VERT, fg="white", relief="flat", pady=8,
              cursor="hand2").pack(fill="x", padx=PAD, pady=(8, 0))
    statut = tk.Label(app, text="", font=("Segoe UI", 9), bg=BG, fg="#555", wraplength=620)
    statut.pack(pady=(12, 8))

    if tables:
        v_table.set(tables[0]["nom"])
        maj_onglets()
    if types:
        v_type.set(next(iter(types.values())).get("libelle", ""))
        maj_roles()
    app.update_idletasks()
    w, h = 680, 520
    app.geometry(f"{w}x{h}+{(app.winfo_screenwidth()-w)//2}+{(app.winfo_screenheight()-h)//2}")
    app.mainloop()


if __name__ == "__main__":
    main()
