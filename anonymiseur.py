"""
anonymiseur.py
--------------
Programme unique : preparer la table (pre-detection), anonymiser, controler,
restituer. Une seule fenetre, un seul executable.

Toute la logique metier est dans les modules du meme dossier :
    anonymisation_core.py   anonymiser / restituer / controler
    detection_core.py       lire le document, reconnaitre les termes (OCR compris)
    detection_roles.py      rattacher chaque terme a une partie et un role
    detection_classeur.py   classeur de detection a relire, injection
    modele_onglet.py        le modele d'onglet « par partie »
    classeur_modele.py      le classeur remis a l'utilisateur, le document d'exemple

Ce fichier ne fait que l'affichage, les confirmations et le premier demarrage
(creation de la table, des reglages et du dossier Journaux a cote du programme).

Le document ne quitte jamais le poste : aucune IA, aucun appel reseau.

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

import os
import sys
import json
import pathlib
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog

VERSION = "2026.10.1"

# Mise a l'echelle Windows (125 %, 150 %) : evite une fenetre qui deborde.
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

# ------------------------------------------------------------------
# OU SONT LES CHOSES
# ------------------------------------------------------------------
# DOSSIER_APP : a cote de l'executable (ou du script). C'est la que vivent
#               la table, les reglages, les journaux : ce que l'utilisateur voit.
# RESSOURCES  : les fichiers livres avec le programme (reglages par defaut,
#               icone). Dans l'executable, PyInstaller les depose dans un
#               dossier temporaire ; en script, c'est le dossier du script.

if getattr(sys, "frozen", False):
    DOSSIER_APP = pathlib.Path(sys.executable).resolve().parent
    RESSOURCES = pathlib.Path(getattr(sys, "_MEIPASS", DOSSIER_APP))
else:
    DOSSIER_APP = pathlib.Path(__file__).resolve().parent
    RESSOURCES = DOSSIER_APP
sys.path.insert(0, str(RESSOURCES))
if not getattr(sys, "frozen", False):
    # En version .exe, tous les modules sont embarques : le dossier de l'exe
    # n'est PAS ajoute au chemin d'import, pour qu'un fichier .py depose a
    # cote de l'exe ne puisse jamais se substituer a un module du programme.
    sys.path.insert(0, str(DOSSIER_APP))

# Langue : fichier langue.txt livre avec l'executable (« en » pour la version
# anglaise), sinon francais.
import textes
for _candidat in (DOSSIER_APP / "langue.txt", RESSOURCES / "langue.txt"):
    try:
        textes.LANGUE = _candidat.read_text(encoding="utf-8").strip().lower()[:2] or "fr"
        break
    except OSError:
        pass
T = textes.T



def _texte_alterees(alterees):
    """Liste des balises alterees par l'IA, laissees en place a la restitution."""
    if not alterees:
        return ""
    lignes = [f"    {t['forme']}  ->  {t['balise']}  " + T("fois").format(n=t['occurrences'])
              for t in alterees[:10]]
    if len(alterees) > 10:
        lignes.append(T("et_autres").format(n=len(alterees) - 10))
    return T("alterees_restitution").format(n=len(alterees), d="\n".join(lignes))

class _MessagesTraduits:
    """messagebox dont le texte passe par textes.traduire() : en version
    anglaise, les messages des moteurs (ecrits en francais) sont traduits."""
    def __getattr__(self, nom):
        fonction = getattr(_messagebox_tk, nom)
        if not callable(fonction):
            return fonction

        def appel(*args, **kw):
            args = [textes.traduire(a) if isinstance(a, str) else a for a in args]
            if "message" in kw:
                kw["message"] = textes.traduire(kw["message"])
            return fonction(*args, **kw)
        return appel


_messagebox_tk = messagebox
messagebox = _MessagesTraduits()
NOM_PROGRAMME = T("nom")

# Ce que le programme cree au premier lancement (reglages, table, journaux,
# exemple) est range dans un sous-dossier « Données » (« Data » en anglais),
# pour que le dossier du programme reste lisible. Une installation existante,
# dont le config.json est deja a cote du programme, garde son organisation :
# rien n'est jamais deplace.
def _dossier_donnees():
    if (DOSSIER_APP / "config.json").exists():
        return DOSSIER_APP
    return DOSSIER_APP / T("dossier_donnees")


DOSSIER_DONNEES = _dossier_donnees()
CHEMIN_CONFIG = DOSSIER_DONNEES / "config.json"
CHEMIN_CONFIG_DETECTION = DOSSIER_DONNEES / "detection_config.json"
CHEMIN_TABLE_DEFAUT = DOSSIER_DONNEES / T("table_defaut")
DOSSIER_JOURNAUX = DOSSIER_DONNEES / T("dossier_journaux")
DOSSIER_EXEMPLES = DOSSIER_DONNEES / T("dossier_exemples")
CHEMIN_MODE_EMPLOI = DOSSIER_APP / T("mode_emploi_pdf")
DOSSIER_TESSERACT = DOSSIER_APP / "Tesseract"

try:
    import anonymisation_core as coeur
    import detection_core as det
    import detection_classeur as classeur
    import detection_roles as roles
    import modele_onglet as modele
    import classeur_modele
    for _module in (modele, classeur, classeur_modele, coeur):
        _module.LANGUE = textes.LANGUE       # textes des classeurs et des comptes rendus
except Exception as e:                       # pragma: no cover
    tk.Tk().withdraw()
    messagebox.showerror(T("demarrage_impossible"),
                         T("demarrage_detail").format(d=DOSSIER_APP, e=e))
    sys.exit(1)


# ------------------------------------------------------------------
# PREMIER DEMARRAGE
# ------------------------------------------------------------------

def _ecrire_json(chemin, contenu):
    with open(chemin, "w", encoding="utf-8") as f:
        json.dump(contenu, f, ensure_ascii=False, indent=2)


def preparer_dossier():
    """Cree ce qui manque a cote du programme. Ne touche jamais a ce qui existe.
    Retourne la liste de ce qui a ete cree (pour le message de bienvenue)."""
    crees = []
    DOSSIER_DONNEES.mkdir(parents=True, exist_ok=True)

    if not CHEMIN_CONFIG_DETECTION.exists():
        source = RESSOURCES / "detection_config.json"
        if source.exists() and source != CHEMIN_CONFIG_DETECTION:
            CHEMIN_CONFIG_DETECTION.write_bytes(source.read_bytes())
            crees.append(T("cree_config_det"))

    config_det = det.charger_config_detection(CHEMIN_CONFIG_DETECTION)

    if not CHEMIN_CONFIG.exists():
        if not CHEMIN_TABLE_DEFAUT.exists():
            classeur_modele.creer_classeur_modele(str(CHEMIN_TABLE_DEFAUT), config_det)
            crees.append(T("cree_table"))
        _ecrire_json(CHEMIN_CONFIG, {
            "_lisez_moi": T("config_lisez_moi"),
            "tables": [{"nom": T("nom_table_defaut"), "fichier": T("table_defaut")}],
            "dossier_journaux": T("dossier_journaux"),
            "masquer_proprietes": True,
        })
        crees.append(T("cree_config"))

    if not DOSSIER_JOURNAUX.exists():
        DOSSIER_JOURNAUX.mkdir(parents=True, exist_ok=True)
        crees.append(T("cree_journaux"))

    exemple = DOSSIER_EXEMPLES / T("doc_exemple")
    if not DOSSIER_EXEMPLES.exists():
        DOSSIER_EXEMPLES.mkdir(parents=True, exist_ok=True)
        try:
            classeur_modele.creer_document_exemple(str(exemple))
            crees.append(T("cree_exemple"))
        except Exception:
            pass

    # Tesseract embarque a cote du programme : prioritaire sur tout le reste.
    exe = "tesseract.exe" if sys.platform == "win32" else "tesseract"
    det.EMPLACEMENTS_SUPPLEMENTAIRES[:] = [str(DOSSIER_TESSERACT / exe)]
    det.TESSDATA_SUPPLEMENTAIRES[:] = [str(DOSSIER_TESSERACT / "tessdata")]

    if DOSSIER_DONNEES != DOSSIER_APP:
        crees = [f"{DOSSIER_DONNEES.name}\\{c}" for c in crees]
    return crees


def ouvert_dans_excel(chemin):
    """Vrai si Excel a ce classeur ouvert : il pose alors un fichier de verrou
    « ~$<nom> » a cote. Sans verrou, on ne peut rien affirmer (False)."""
    try:
        dossier, nom = os.path.split(os.path.abspath(str(chemin)))
        return bool(nom) and os.path.exists(os.path.join(dossier, "~$" + nom))
    except OSError:
        return False


def ouvrir(chemin):
    """Ouvre un fichier ou un dossier avec le programme associe (Excel, PDF...)."""
    chemin = str(chemin)
    try:
        if sys.platform == "win32":
            os.startfile(chemin)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", chemin])
        else:
            subprocess.Popen(["xdg-open", chemin])
    except Exception as e:
        messagebox.showerror(T("ouverture_impossible"), f"{chemin}\n\n{e}")


# ------------------------------------------------------------------
# APPARENCE (legal design : une teinte, des gris, une couleur par action)
# ------------------------------------------------------------------

FOND = "#F3F5F7"
CARTE = "#FFFFFF"
TEXTE = "#1F2933"
TEXTE_DOUX = "#5F6B76"
BORD = "#D9DFE5"
BLEU = "#1B3A5C"      # identite du programme, anonymiser
BLEU_PALE = "#E8EEF4"
VERT = "#2E7D4F"      # restituer (voie sure)
AMBRE = "#B8741A"     # restituer sans journal (voie degradee)
ROUGE = "#B23B3B"     # controler
TEAL = "#2A6F77"      # preparer la table

POLICE = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"


def police(taille=10, gras=False, italique=False):
    styles = []
    if gras:
        styles.append("bold")
    if italique:
        styles.append("italic")
    return (POLICE, taille, " ".join(styles)) if styles else (POLICE, taille)


class Carte(tk.Frame):
    """Un bloc blanc a bord fin, avec un titre optionnel."""

    def __init__(self, parent, titre=None, **kw):
        super().__init__(parent, bg=CARTE, highlightbackground=BORD,
                         highlightthickness=1, bd=0, **kw)
        if titre:
            tk.Label(self, text=titre, font=police(10, gras=True), bg=CARTE,
                     fg=BLEU, anchor="w").pack(fill="x", padx=16, pady=(12, 4))


def bouton_action(parent, texte, couleur, commande, taille=11):
    return tk.Button(parent, text=texte, command=commande, font=police(taille, gras=True),
                     bg=couleur, fg="white", activebackground=couleur,
                     activeforeground="white", relief="flat", bd=0, pady=10,
                     cursor="hand2", disabledforeground="#DDDDDD")


LARGEUR_BOUTON = 20          # en caracteres : tous les boutons secondaires s'alignent


def bouton_discret(parent, texte, commande, largeur=LARGEUR_BOUTON):
    return tk.Button(parent, text=texte, command=commande, font=police(9),
                     bg="#F7F9FA", fg=TEXTE, activebackground=BLEU_PALE, relief="solid",
                     bd=1, highlightthickness=0, padx=6, pady=4, cursor="hand2",
                     # Largeur commune pour l'alignement, elargie si le libelle
                     # est plus long : il n'est jamais coupe (9 octobre 2026).
                     width=max(largeur, len(texte) + 2))


def paragraphe(parent, texte, couleur=TEXTE_DOUX, taille=9, largeur=700, **kw):
    return tk.Label(parent, text=texte, font=police(taille), bg=parent["bg"],
                    fg=couleur, wraplength=largeur, justify="left", anchor="w", **kw)


# ------------------------------------------------------------------
# FENETRE DE CHOIX D'UN JOURNAL
# ------------------------------------------------------------------

class FenetreChoixJournal(tk.Toplevel):
    def __init__(self, parent, candidats):
        super().__init__(parent)
        self.title(T("choisir_journal"))
        self.configure(bg=FOND)
        self.resultat = None
        self.transient(parent)
        self.grab_set()

        tk.Label(self, text=T("plusieurs_journaux"),
                 font=police(12, gras=True), bg=FOND, fg=BLEU).pack(pady=(14, 6), padx=16)
        tk.Label(self, text=T("selectionnez_journal"),
                 font=police(9), bg=FOND, fg=TEXTE_DOUX).pack(padx=16)

        self.libelles = []
        for chemin, meta in candidats:
            meta = meta or {}
            table = meta.get("table", T("table_inconnue"))
            onglet = meta.get("onglet", T("onglet_inconnu"))
            date = meta.get("horodatage", "")[:16].replace("T", " ")
            self.libelles.append(f"{table} / {onglet}   ({date})" if meta
                                 else os.path.basename(chemin) + T("format_ancien"))
        self.candidats = candidats
        self.choix = tk.StringVar(value=self.libelles[0])
        ttk.Combobox(self, textvariable=self.choix, values=self.libelles,
                     state="readonly", width=58).pack(padx=16, pady=10)

        boutons = tk.Frame(self, bg=FOND)
        boutons.pack(pady=(2, 12))
        bouton_discret(boutons, T("annuler"), self.destroy).pack(side="left", padx=6)
        bouton_action(boutons, T("utiliser_journal"), VERT, self._ok, 10).pack(side="left", padx=6)

        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 60}+{parent.winfo_rooty() + 80}")
        self.wait_window()

    def _ok(self):
        self.resultat = self.candidats[self.libelles.index(self.choix.get())]
        self.destroy()


# ------------------------------------------------------------------
# FENETRE PRINCIPALE
# ------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self, config, tables, problemes, crees):
        super().__init__()
        self.config_outil = config
        self.tables = tables
        self._taches_de_fond = 0           # taches de fond en cours
        self._demandes = []                # boites de dialogue demandees par les taches
        self._verrou_demandes = threading.Lock()
        self.config_det = det.charger_config_detection(CHEMIN_CONFIG_DETECTION)
        modele.CONFIG = self.config_det          # balises calculees, menus des onglets

        self.title(f"{NOM_PROGRAMME}")
        self.configure(bg=FOND)
        self.minsize(760, 640)
        self._icone()

        self.chemin_doc = tk.StringVar()
        self.chemin_detection = tk.StringVar()
        self.nom_table = tk.StringVar()
        self.nom_onglet = tk.StringVar()
        types = self.config_det.get("types_documents") or {}
        self.types = [(T("automatique"), roles.AUTOMATIQUE)] + [
            (d.get("libelle", code), code) for code, d in types.items()]
        self.type_doc = tk.StringVar(value=self.types[0][0])
        self.sens_choix = [(T("automatique"), None), (T("emis"), roles.SENS_EMIS),
                           (T("recu"), roles.SENS_RECU)]
        self.sens = tk.StringVar(value=self.sens_choix[0][0])
        self.boutons_actions = []
        self.toiles = {}            # zone defilante de chaque onglet -> sa barre

        self._style()
        self._construire()
        self.bind_all("<MouseWheel>", self._molette)
        # Molette au-dessus d'un menu deroulant : la page defile, le choix
        # ne change pas par accident.
        self.bind_class("TCombobox", "<MouseWheel>", self._molette)
        self._centrer(860, 760)

        if self.tables:
            self.nom_table.set(self.tables[0]["nom"])
            self._maj_onglets()

        if crees:
            self.after(300, lambda: messagebox.showinfo(
                T("bienvenue_titre"),
                T("bienvenue").format(liste="\n".join("   •  " + c for c in crees))))
        if problemes:
            self.after(400, lambda: messagebox.showwarning(
                T("tables_introuvables"),
                T("tables_introuvables_detail").format(p="\n\n".join(problemes))))

    # ---------- construction ----------

    def _icone(self):
        try:
            if sys.platform == "win32" and (RESSOURCES / "icone.ico").exists():
                self.iconbitmap(str(RESSOURCES / "icone.ico"))
            elif (RESSOURCES / "icone.png").exists():
                self._icone_png = tk.PhotoImage(file=str(RESSOURCES / "icone.png"))
                self.iconphoto(True, self._icone_png)
        except Exception:
            pass

    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TNotebook", background=FOND, borderwidth=0, tabmargins=(0, 0, 0, 0))
        st.configure("TNotebook.Tab", font=police(10, gras=True), padding=(16, 8),
                     background="#E3E8EE", foreground=TEXTE_DOUX, borderwidth=0)
        st.map("TNotebook.Tab", background=[("selected", CARTE)],
               foreground=[("selected", BLEU)], expand=[("selected", (0, 0, 0, 0))])
        st.configure("TCombobox", font=police(10), padding=4, fieldbackground=CARTE,
                     background=CARTE, arrowsize=16)
        st.map("TCombobox", fieldbackground=[("readonly", CARTE)],
               selectbackground=[("readonly", CARTE)], selectforeground=[("readonly", TEXTE)])
        st.configure("TSeparator", background=BORD)

    def _centrer(self, w, h):
        self.update_idletasks()
        w = min(w, self.winfo_screenwidth() - 40)
        h = min(h, self.winfo_screenheight() - 80)
        x = (self.winfo_screenwidth() - w) // 2
        y = max(10, (self.winfo_screenheight() - h) // 2 - 20)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _construire(self):
        # Bandeau
        bandeau = tk.Frame(self, bg=BLEU)
        bandeau.pack(fill="x")
        ligne1 = tk.Frame(bandeau, bg=BLEU)
        ligne1.pack(fill="x", padx=22, pady=(14, 0))
        tk.Label(ligne1, text=NOM_PROGRAMME, font=police(18, gras=True), bg=BLEU,
                 fg="white").pack(side="left")
        tk.Label(ligne1, text=T("sous_titre"), font=police(10), bg=BLEU,
                 fg="#C9D6E3").pack(side="left", padx=(12, 0), pady=(6, 0))
        tk.Label(ligne1, text=T("version").format(v=VERSION), font=police(8), bg=BLEU,
                 fg="#8FA6BC").pack(side="right", pady=(6, 0))
        tk.Label(bandeau, text=T("bandeau"),
                 font=police(9, italique=True), bg=BLEU, fg="#C9D6E3",
                 anchor="w").pack(fill="x", padx=22, pady=(0, 12))

        # Pied : statut (place avant le corps pour rester visible quoi qu'il arrive)
        pied = tk.Frame(self, bg=FOND)
        pied.pack(side="bottom", fill="x", padx=18, pady=(8, 10))
        self.statut = tk.Label(pied, text=T("pret"), font=police(9), bg=FOND, fg=TEXTE_DOUX,
                               wraplength=640, justify="left", anchor="w")
        self.statut.pack(side="left", fill="x", expand=True)
        tk.Label(pied, text=T("developpe").format(a=coeur.AUTEUR), font=police(8, italique=True),
                 bg=FOND, fg="#9AA5B1").pack(side="right")

        corps = tk.Frame(self, bg=FOND)
        corps.pack(fill="both", expand=True, padx=18, pady=(14, 0))

        self._carte_contexte(corps)

        self.onglets = ttk.Notebook(corps)
        self.onglets.pack(fill="both", expand=True, pady=(12, 0))
        self._onglet_preparer()
        self._onglet_anonymiser()
        self._onglet_controler()
        self._onglet_restituer()
        self._onglet_aide()
        self.onglets.select(1)       # l'usage courant commence par Anonymiser


    def _carte_contexte(self, parent):
        carte = Carte(parent, T("contexte_titre"))
        carte.pack(fill="x")
        grille = tk.Frame(carte, bg=CARTE)
        grille.pack(fill="x", padx=16, pady=(0, 12))
        grille.columnconfigure(1, weight=1)

        def etiquette(ligne, texte):
            tk.Label(grille, text=texte, font=police(10), bg=CARTE, fg=TEXTE,
                     anchor="w", width=16).grid(row=ligne, column=0, sticky="w", pady=5)

        def separation(ligne):
            ttk.Separator(grille, orient="horizontal").grid(
                row=ligne, column=0, columnspan=4, sticky="ew", pady=3)

        etiquette(0, T("document"))
        tk.Entry(grille, textvariable=self.chemin_doc, font=police(9), state="readonly",
                 relief="flat", readonlybackground="#F7F9FA", fg=TEXTE,
                 highlightthickness=1, highlightbackground=BORD).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(0, 8), ipady=5)
        bouton_discret(grille, T("choisir"), self._choisir_doc).grid(row=0, column=3, sticky="ew")

        separation(1)

        etiquette(2, T("table"))
        self.combo_table = ttk.Combobox(grille, textvariable=self.nom_table,
                                        values=[t["nom"] for t in self.tables],
                                        state="readonly")
        self.combo_table.grid(row=2, column=1, sticky="ew", padx=(0, 8))
        self.combo_table.bind("<<ComboboxSelected>>", lambda e: self._maj_onglets())
        bouton_discret(grille, T("ouvrir_excel"), self._ouvrir_table).grid(
            row=2, column=2, sticky="ew", padx=(0, 8))
        bouton_discret(grille, T("ajouter_table"), self._ajouter_table).grid(
            row=2, column=3, sticky="ew")

        separation(3)

        etiquette(4, T("dossier_onglet"))
        self.combo_onglet = ttk.Combobox(grille, textvariable=self.nom_onglet, values=[],
                                         state="readonly")
        self.combo_onglet.grid(row=4, column=1, sticky="ew", padx=(0, 8))
        self.combo_onglet.bind("<Button-1>",
                               lambda e: self._maj_onglets(garder_choix=True, silencieux=True))
        bouton_discret(grille, T("nouveau_dossier"), self._nouvel_onglet).grid(
            row=4, column=2, sticky="ew", padx=(0, 8))
        bouton_discret(grille, T("actualiser"), self._actualiser_tout).grid(
            row=4, column=3, sticky="ew")

    def _page(self, titre, couleur, intro):
        page = tk.Frame(self.onglets, bg=CARTE)
        self.onglets.add(page, text=titre)
        tk.Frame(page, bg=couleur, height=4).pack(fill="x")

        # Zone defilante : si la fenetre est trop petite (ecran reduit, mise a
        # l'echelle Windows), le contenu defile au lieu de se chevaucher.
        toile = tk.Canvas(page, bg=CARTE, highlightthickness=0, bd=0)
        barre = ttk.Scrollbar(page, orient="vertical", command=toile.yview)
        toile.configure(yscrollcommand=barre.set)
        toile.pack(side="left", fill="both", expand=True)
        contenu = tk.Frame(toile, bg=CARTE)
        fenetre = toile.create_window((0, 0), window=contenu, anchor="nw")

        def ajuster(event=None):
            toile.itemconfigure(fenetre, width=toile.winfo_width())
            toile.configure(scrollregion=toile.bbox("all"))
            if contenu.winfo_reqheight() > toile.winfo_height():
                barre.pack(side="right", fill="y")
            else:
                barre.pack_forget()
                toile.yview_moveto(0)
        toile.bind("<Configure>", ajuster)
        contenu.bind("<Configure>", ajuster)

        # La molette est geree une fois pour toute la fenetre (_molette) : un
        # <Leave> etait envoye des que la souris passait sur un bouton ou un
        # texte de la page, et la molette decrochait.
        self.toiles[str(toile)] = barre

        interieur = tk.Frame(contenu, bg=CARTE)
        interieur.pack(fill="both", expand=True, padx=22, pady=(14, 14))
        paragraphe(interieur, intro, TEXTE, 10, 760).pack(fill="x", pady=(0, 10))
        return interieur

    def _molette(self, event):
        """Fait defiler l'onglet sous la souris, quel que soit l'element survole."""
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
        except (KeyError, tk.TclError):
            widget = None
        # Une zone de texte avec sa propre barre (onglet Aide) defile seule.
        if widget is not None and widget.winfo_class() in ("Text", "Listbox"):
            return None
        while widget is not None:
            barre = self.toiles.get(str(widget))
            if barre is not None:
                if barre.winfo_ismapped():
                    widget.yview_scroll(-1 if event.delta > 0 else 1, "units")
                break
            widget = widget.master
        return "break"

    def _onglet_preparer(self):
        c = self._page(T("onglet_preparer"), TEAL, T("preparer_intro"))
        grille = tk.Frame(c, bg=CARTE)
        grille.pack(fill="x", pady=(0, 8))
        tk.Label(grille, text=T("type_document"), font=police(10), bg=CARTE,
                 fg=TEXTE).grid(row=0, column=0, sticky="w", pady=3)
        ttk.Combobox(grille, textvariable=self.type_doc, values=[l for l, _ in self.types],
                     state="readonly", width=40).grid(row=0, column=1, sticky="w", padx=(12, 0))
        tk.Label(grille, text=T("sens_courrier"), font=police(10), bg=CARTE,
                 fg=TEXTE).grid(row=1, column=0, sticky="w", pady=3)
        ttk.Combobox(grille, textvariable=self.sens, values=[l for l, _ in self.sens_choix],
                     state="readonly", width=22).grid(row=1, column=1, sticky="w", padx=(12, 0))

        b = bouton_action(c, T("etape1"), TEAL, self._detecter)
        b.pack(fill="x", pady=(6, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("etape1_note")).pack(fill="x")

        ligne = tk.Frame(c, bg=CARTE)
        ligne.pack(fill="x", pady=(14, 0))
        tk.Label(ligne, text=T("classeur_relu"), font=police(10), bg=CARTE, fg=TEXTE).pack(side="left")
        tk.Entry(ligne, textvariable=self.chemin_detection, font=police(9), state="readonly",
                 relief="flat", readonlybackground="#F7F9FA", highlightthickness=1,
                 highlightbackground=BORD).pack(side="left", fill="x", expand=True,
                                                padx=(12, 8), ipady=5)
        bouton_discret(ligne, T("choisir"), self._choisir_detection, 12).pack(side="left")

        b = bouton_action(c, T("etape2"), AMBRE, self._relire)
        b.pack(fill="x", pady=(8, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("etape2_note")).pack(fill="x")

        b = bouton_action(c, T("etape3"), VERT, self._injecter)
        b.pack(fill="x", pady=(14, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("etape3_note")).pack(fill="x")

    def _onglet_anonymiser(self):
        c = self._page(T("onglet_anonymiser"), BLEU, T("anonymiser_intro"))
        b = bouton_action(c, T("anonymiser_bouton"), BLEU, lambda: self._lancer("anonymiser"), 11)
        b.pack(fill="x", pady=(10, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("anonymiser_note")).pack(fill="x")
        self._encadre(c, T("avant_anonymiser"), T("avant_anonymiser_texte"))

    def _onglet_controler(self):
        c = self._page(T("onglet_controler"), ROUGE, T("controler_intro"))
        b = bouton_action(c, T("controler_bouton"), ROUGE, lambda: self._lancer("controle"), 11)
        b.pack(fill="x", pady=(10, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("controler_note")).pack(fill="x")
        self._encadre(c, T("bonne_question"), T("bonne_question_texte"))

    def _onglet_restituer(self):
        c = self._page(T("onglet_restituer"), VERT, T("restituer_intro"))
        b = bouton_action(c, T("restituer_journal"), VERT, lambda: self._lancer("journal"), 11)
        b.pack(fill="x", pady=(10, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("restituer_journal_note")).pack(fill="x")

        b = bouton_action(c, T("restituer_sans"), AMBRE, lambda: self._lancer("sans_journal"), 11)
        b.pack(fill="x", pady=(16, 2))
        self.boutons_actions.append(b)
        paragraphe(c, T("restituer_sans_note")).pack(fill="x")
        paragraphe(c, T("restituer_cree")).pack(fill="x", pady=(8, 0))

    def _onglet_aide(self):
        c = self._page(T("onglet_aide"), TEXTE_DOUX, T("aide_intro"))
        ligne = tk.Frame(c, bg=CARTE)
        ligne.pack(fill="x", pady=(0, 8))
        bouton_discret(ligne, T("ouvrir_mode_emploi"),
                       lambda: ouvrir(CHEMIN_MODE_EMPLOI) if CHEMIN_MODE_EMPLOI.exists()
                       else messagebox.showinfo(T("mode_emploi"),
                                                T("mode_emploi_absent"))).pack(side="left")
        bouton_discret(ligne, T("ouvrir_dossier_programme"),
                       lambda: ouvrir(DOSSIER_DONNEES)).pack(side="left", padx=(8, 0))
        bouton_discret(ligne, T("ouvrir_journaux"),
                       lambda: ouvrir(self.config_outil["dossier_journaux"])).pack(side="left", padx=(8, 0))

        cadre = tk.Frame(c, bg=CARTE, highlightbackground=BORD, highlightthickness=1)
        cadre.pack(fill="both", expand=True)
        texte = tk.Text(cadre, font=police(9), bg="#FBFCFD", fg=TEXTE, relief="flat",
                        wrap="word", padx=14, pady=10, spacing1=2, spacing3=4)
        barre = ttk.Scrollbar(cadre, command=texte.yview)
        texte.configure(yscrollcommand=barre.set)
        barre.pack(side="right", fill="y")
        texte.pack(fill="both", expand=True)
        texte.tag_configure("titre", font=police(10, gras=True), foreground=BLEU, spacing1=10)
        for genre, contenu in T("aide"):
            texte.insert("end", contenu + "\n", genre)
        texte.configure(state="disabled")

    def _encadre(self, parent, titre, corps):
        cadre = tk.Frame(parent, bg=BLEU_PALE)
        cadre.pack(fill="x", pady=(16, 0))
        tk.Label(cadre, text=titre, font=police(9, gras=True), bg=BLEU_PALE, fg=BLEU,
                 anchor="w").pack(fill="x", padx=14, pady=(10, 2))
        paragraphe(cadre, corps, TEXTE, 9, 740).pack(fill="x", padx=14, pady=(0, 12))

    # ---------- contexte ----------

    def _table_courante(self):
        return next((t for t in self.tables if t["nom"] == self.nom_table.get()), None)

    def _maj_onglets(self, garder_choix=False, silencieux=False):
        table = self._table_courante()
        if not table:
            return
        ancien = self.nom_onglet.get()
        try:
            onglets = coeur.lister_onglets(table["fichier"], self.config_outil)
        except Exception as e:
            if not silencieux:
                messagebox.showerror(T("table_illisible"), str(e))
            return
        self.combo_onglet.config(values=onglets)
        if garder_choix and ancien in onglets:
            self.nom_onglet.set(ancien)
        else:
            self.nom_onglet.set(onglets[0] if onglets else "")

    def _actualiser_tout(self):
        ancienne_table, ancien_onglet = self.nom_table.get(), self.nom_onglet.get()
        try:
            config = coeur.charger_config(CHEMIN_CONFIG)
            tables_ok, problemes = coeur.verifier_tables(config)
        except coeur.ErreurAnonymisation as e:
            messagebox.showerror(T("configuration"), str(e))
            return
        if not tables_ok:
            messagebox.showerror(T("aucune_table"), "\n\n".join(problemes))
            return
        self.config_outil, self.tables = config, tables_ok
        noms = [t["nom"] for t in self.tables]
        self.combo_table.config(values=noms)
        self.nom_table.set(ancienne_table if ancienne_table in noms else noms[0])
        self._maj_onglets(garder_choix=True)
        if ancien_onglet in self.combo_onglet["values"]:
            self.nom_onglet.set(ancien_onglet)
        self._statut(T("actualise").format(t=len(noms), d=len(self.combo_onglet['values']),
                                           n=self.nom_table.get()), VERT)
        if problemes:
            messagebox.showwarning(T("tables_introuvables"), "\n\n".join(problemes))

    def _ouvrir_table(self):
        table = self._table_courante()
        if table:
            ouvrir(table["fichier"])
            self._statut(T("excel_conseil"), TEXTE_DOUX)

    def _ajouter_table(self):
        chemin = filedialog.askopenfilename(
            title=T("choisir_classeur"),
            filetypes=[(T("classeurs_excel"), "*.xlsx")])
        if not chemin:
            return
        nom = simpledialog.askstring(
            T("nom_table"), T("nom_table_question"), parent=self,
            initialvalue=os.path.splitext(os.path.basename(chemin))[0])
        if not nom:
            return
        try:
            with open(CHEMIN_CONFIG, encoding="utf-8") as f:
                brut = json.load(f)
            fichier = chemin
            try:
                # chemin relatif au dossier de config.json, comme le lit charger_config
                fichier = str(pathlib.Path(chemin).relative_to(CHEMIN_CONFIG.parent))
            except ValueError:
                pass
            brut.setdefault("tables", []).append({"nom": nom, "fichier": fichier})
            _ecrire_json(CHEMIN_CONFIG, brut)
        except Exception as e:
            messagebox.showerror(T("ajout_impossible"), str(e))
            return
        self._actualiser_tout()
        self.nom_table.set(nom)
        self._maj_onglets()

    def _nouvel_onglet(self):
        table = self._table_courante()
        if not table:
            return
        nom = simpledialog.askstring(
            T("nouveau_dossier_titre"),
            T("nouveau_dossier_question").format(t=table['nom']), parent=self)
        if not nom:
            return
        try:
            nom = modele.creer_onglet(table["fichier"], nom, self.config_det)
        except Exception as e:
            messagebox.showerror(T("creation_impossible"), str(e))
            return
        self._maj_onglets()
        self.nom_onglet.set(nom)
        self._statut(T("dossier_cree").format(n=nom), VERT)

    def _choisir_doc(self):
        chemin = filedialog.askopenfilename(
            title=T("choisir_document"),
            initialdir=str(DOSSIER_EXEMPLES) if not self.chemin_doc.get() and DOSSIER_EXEMPLES.exists()
            else None,
            filetypes=[(T("docs_word_pdf"), "*.docx *.pdf *.txt *.md"), ("Word", "*.docx"),
                       (T("pdf_predetection"), "*.pdf"), (T("texte_txt_md"), "*.txt *.md")])
        if chemin:
            self.chemin_doc.set(chemin)
            self._statut(T("document_statut").format(n=os.path.basename(chemin)), TEXTE_DOUX)

    def _choisir_detection(self):
        chemin = filedialog.askopenfilename(
            title=T("choisir_detection"),
            filetypes=[(T("classeurs_excel"), "*.xlsx")])
        if chemin:
            self.chemin_detection.set(chemin)

    # ---------- helpers ----------

    def _statut(self, texte, couleur=TEXTE_DOUX):
        # Appelable depuis une tache de fond : la fenetre n'est touchee que
        # par le fil principal.
        if threading.current_thread() is not threading.main_thread():
            with self._verrou_demandes:
                self._demandes.append({"fonction": lambda: self._statut(texte, couleur),
                                       "fait": threading.Event()})
            return
        self.statut.config(text=texte, fg=couleur)
        self.update_idletasks()

    def _boutons(self, etat):
        for b in self.boutons_actions:
            b.config(state=etat)

    def _type_choisi(self):
        return next(c for l, c in self.types if l == self.type_doc.get())

    def _sens_choisi(self):
        return next(c for l, c in self.sens_choix if l == self.sens.get())

    def _exige_docx(self):
        if not self.chemin_doc.get():
            messagebox.showwarning(T("document_manquant"), T("choisir_dabord"))
            return False
        if not self.chemin_doc.get().lower().endswith((".docx", ".txt", ".md")):
            messagebox.showwarning(T("word_attendu"), T("word_attendu_detail"))
            return False
        return True

    def _exige_onglet(self):
        if not self.nom_onglet.get():
            messagebox.showwarning(T("dossier_manquant"), T("choisir_table_dossier"))
            return False
        return True

    # ---------- lancement ----------

    def _lancer(self, mode):
        if not self._exige_docx():
            return
        if mode in ("anonymiser", "sans_journal", "controle") and not self._exige_onglet():
            return
        if mode == "sans_journal":
            self._restituer_sans_journal()
        elif mode == "journal":
            self._restituer_avec_journal()
        elif mode == "controle":
            self._controler()
        else:
            self._verifier_onglet_avant_anonymisation()

    # ---------- taches de fond ----------
    #
    # Les traitements longs (gros document, OCR) tournent dans un fil separe :
    # la fenetre reste vivante au lieu d'afficher « Ne repond pas ». Seul le
    # fil principal touche a la fenetre : il surveille la tache par after()
    # et execute lui-meme la suite (resultat ou erreur) ainsi que les boites
    # de dialogue que la tache demande par _depuis_la_tache() (8 octobre 2026).

    def _en_arriere_plan(self, travail, puis, sur_erreur=None):
        """Lance travail() dans un fil de fond, boutons desactives ; a la fin,
        appelle puis(resultat) ou sur_erreur(exception) dans le fil principal,
        puis reactive les boutons."""
        self._boutons("disabled")
        self._taches_de_fond = getattr(self, "_taches_de_fond", 0) + 1
        boite = {}

        def courir():
            try:
                boite["resultat"] = travail()
            except Exception as e:
                boite["erreur"] = e

        fil = threading.Thread(target=courir, daemon=True)
        fil.start()

        def surveiller():
            self._executer_demandes()
            if fil.is_alive():
                self.after(100, surveiller)
                return
            try:
                if "erreur" in boite:
                    (sur_erreur or self._echec)(boite["erreur"])
                else:
                    puis(boite.get("resultat"))
            finally:
                # Si la suite a elle-meme lance une autre tache (verification
                # du dossier puis anonymisation), les boutons restent inactifs.
                self._taches_de_fond -= 1
                if self._taches_de_fond == 0:
                    self._boutons("normal")

        self.after(100, surveiller)

    def _depuis_la_tache(self, fonction):
        """Depuis une tache de fond : fait executer fonction() (boite de
        dialogue) par le fil principal et en attend la reponse."""
        if threading.current_thread() is threading.main_thread():
            return fonction()
        demande = {"fonction": fonction, "fait": threading.Event()}
        with self._verrou_demandes:
            self._demandes.append(demande)
        demande["fait"].wait()
        if "erreur" in demande:
            raise demande["erreur"]
        return demande.get("resultat")

    def _executer_demandes(self):
        """Fil principal : execute les demandes deposees par les taches."""
        with self._verrou_demandes:
            demandes, self._demandes = self._demandes, []
        for d in demandes:
            try:
                d["resultat"] = d["fonction"]()
            except Exception as e:
                d["erreur"] = e
            finally:
                d["fait"].set()

    def _verifier_onglet_avant_anonymisation(self):
        """Le dossier choisi correspond-il au document ? (tache de fond, puis
        anonymisation si la reponse est oui)"""
        chemin, nom_table, onglet = self.chemin_doc.get(), self.nom_table.get(), self.nom_onglet.get()
        self._statut(T("verification_dossier"))

        def travail():
            try:
                return coeur.analyser_pertinence(chemin, self.tables, self.config_outil)
            except Exception:
                return None

        def puis(classement):
            self._statut("")
            if self._onglet_convient(classement, nom_table, onglet):
                self._anonymiser()

        self._en_arriere_plan(travail, puis)

    def _onglet_convient(self, classement, nom_table, onglet):
        if classement is None:
            return True
        choisi = next((r for r in classement if r["table"] == nom_table
                       and r["onglet"] == onglet), None)
        if not choisi:
            return True
        meilleur = classement[0] if classement else None
        if choisi["trouvees"] == 0:
            sugg = ""
            if meilleur and meilleur["trouvees"] > 0:
                sugg = T("suggestion").format(t=meilleur['table'], o=meilleur['onglet'],
                                              n=meilleur['trouvees'])
            messagebox.showerror(
                T("dossier_sans_rapport"),
                T("dossier_sans_rapport_detail").format(o=onglet, s=sugg))
            self._statut(T("interrompue_sans_rapport"), ROUGE)
            return False
        if meilleur and meilleur["trouvees"] > choisi["trouvees"]:
            return messagebox.askyesno(
                T("dossier_errone"),
                T("dossier_errone_detail").format(
                    t1=choisi['table'], o1=choisi['onglet'], n1=choisi['trouvees'],
                    t2=meilleur['table'], o2=meilleur['onglet'], n2=meilleur['trouvees']),
                icon="warning", default="no")
        return True

    # ---------- actions ----------

    def _anonymiser(self):
        table = self._table_courante()
        chemin, onglet = self.chemin_doc.get(), self.nom_onglet.get()
        self._statut(T("anonymisation_en_cours"))

        def travail():
            entrees = coeur.charger_entrees(table["fichier"], onglet)
            return coeur.anonymiser_document(chemin, entrees, self.config_outil,
                                             table["nom"], onglet)

        def puis(res):
            invisibles = res.get("total_zones_invisibles", 0)
            detail = T("invisibles").format(n=invisibles) if invisibles else ""
            annexes = []
            if res.get("total_proprietes"):
                annexes.append(T("proprietes").format(n=res['total_proprietes']))
            if res.get("total_liens"):
                annexes.append(T("liens").format(n=res['total_liens']))
            if annexes:
                detail += T("egalement").format(a=", ".join(annexes))
            msg = T("anonymise_resultat").format(n=res['total'], d=detail,
                                                 s=res['chemin_sortie'], j=res['chemin_journal'])
            for a in res.get("avertissements", [])[:8]:
                msg += T("attention").format(a=a)
            self._succes(res["total"], os.path.basename(res["chemin_sortie"]), msg)

        self._en_arriere_plan(travail, puis)

    def _restituer_avec_journal(self):
        try:
            candidats = coeur.journaux_candidats(self.config_outil, self.chemin_doc.get())
            if not candidats:
                messagebox.showerror(
                    T("journal_introuvable"),
                    T("journal_introuvable_detail").format(d=self.config_outil['dossier_journaux']))
                return
            if len(candidats) == 1:
                chemin_j = candidats[0][0]
            else:
                f = FenetreChoixJournal(self, candidats)
                if not f.resultat:
                    self._statut(T("restitution_annulee"))
                    return
                chemin_j = f.resultat[0]
        except Exception as e:
            self._echec(e)
            return
        chemin = self.chemin_doc.get()
        self._statut(T("verification_journal"))

        def travail():
            try:
                return coeur.desanonymiser_avec_journal(chemin, chemin_j)
            except coeur.ErreurConcordance as ec:
                if ec.taux < 0.5:
                    raise
                taux, details = ec.taux, ec.details
                accepte = self._depuis_la_tache(lambda: messagebox.askyesno(
                    T("concordance_partielle"),
                    T("concordance_detail").format(p=taux * 100, d=details),
                    icon="warning", default="no"))
                if not accepte:
                    return None
                return coeur.desanonymiser_avec_journal(chemin, chemin_j, tolerer_partiel=True)

        def puis(res):
            if res is None:
                self._statut(T("restitution_annulee"))
                return
            meta = res.get("meta") or {}
            origine = (T("table_origine").format(t=meta.get('table'), o=meta.get('onglet'))
                       if meta.get("onglet") else "")
            self._succes(res["total"], os.path.basename(res["chemin_sortie"]),
                         T("restitue_resultat").format(n=res['total'], s=res['chemin_sortie'],
                                                       o=origine)
                         + _texte_alterees(res.get("balises_alterees")))

        self._en_arriere_plan(travail, puis)

    def _restituer_sans_journal(self):
        table = self._table_courante()
        chemin, onglet = self.chemin_doc.get(), self.nom_onglet.get()
        self._statut(T("restitution_en_cours"))

        def travail():
            entrees = coeur.charger_entrees(table["fichier"], onglet, anciennes=True)
            try:
                return coeur.desanonymiser_sans_journal(chemin, entrees, confirme=True,
                                                        config=self.config_outil)
            except coeur.ErreurAnonymisation as e:
                # Quel onglet couvrirait le document ? Calcule ici, dans la
                # tache, pour ne pas figer la fenetre a l'affichage de l'erreur.
                e.complement = self._ou_trouver_les_balises(chemin)
                raise

        def puis(res):
            self._succes(res["total"], os.path.basename(res["chemin_sortie"]),
                         T("restitue_sans_resultat").format(n=res['total'], s=res['chemin_sortie'],
                                                            t=table['nom'], o=onglet)
                         + _texte_alterees(res.get("balises_alterees")))

        def sur_erreur(e):
            self._echec(e, complement=getattr(e, "complement", ""))

        self._en_arriere_plan(travail, puis, sur_erreur)

    def _ou_trouver_les_balises(self, chemin):
        try:
            prep = coeur.preparer_restitution_sans_journal(chemin, self.tables, self.config_outil)
        except Exception:
            return ""
        if prep["verdict"] == "aucune_balise":
            return T("aucune_balise")
        couvrants = [f"{a['table']} / {a['onglet']}" for a in prep["candidats"]]
        if not couvrants:
            return T("aucun_dossier_couvre")
        if len(couvrants) == 1:
            return T("dossier_couvre").format(d=couvrants[0])
        return T("dossiers_couvrent").format(d="\n    ".join(couvrants[:6]))

    def _controler(self):
        table = self._table_courante()
        chemin, onglet = self.chemin_doc.get(), self.nom_onglet.get()
        self._statut(T("lecture_parties"))

        def travail():
            entrees = coeur.charger_entrees(table["fichier"], onglet)
            res = coeur.controler(chemin, entrees, self.config_outil)
            rapport = coeur.ecrire_rapport(chemin, res, table["nom"], onglet)
            return res, rapport

        def puis(resultat):
            res, rapport = resultat
            etats = {coeur.ETAT_ANONYMISE: T("etat_anonymise"), coeur.ETAT_RESTITUE: T("etat_restitue"),
                     coeur.ETAT_INDETERMINE: T("etat_indetermine")}
            entete = T("etat_document").format(e=etats.get(res['etat']), m=res['motif_etat'])
            detail = []
            if res["termes_en_clair"]:
                total = sum(t["occurrences"] for t in res["termes_en_clair"])
                detail.append(T("termes_en_clair").format(n=len(res['termes_en_clair']), t=total))
                zone = None
                for t in res["termes_en_clair"][:8]:
                    if t["zone"] != zone:
                        zone = t["zone"]
                        detail.append(f"    {zone}")
                    detail.append(f"        {t['balise']} : {t['terme']}  "
                                  + T("fois").format(n=t['occurrences']))
                if len(res["termes_en_clair"]) > 8:
                    detail.append(T("et_autres").format(n=len(res['termes_en_clair']) - 8))
            for cle, cle_texte in (("hors_table_alertes", "hors_table_alerte"),
                                   ("hors_table_a_verifier", "hors_table_verifier")):
                liste = res.get(cle) or []
                if liste:
                    detail.append(T(cle_texte).format(n=len(liste)))
                    for t in liste[:8]:
                        detail.append(f"    {t['type']} : {t['terme']}  "
                                      + T("fois").format(n=t['occurrences']))
                    if len(liste) > 8:
                        detail.append(T("et_autres").format(n=len(liste) - 8))
            if res.get("balises_revelatrices"):
                detail.append(T("revelatrices_controle").format(n=len(res["balises_revelatrices"])))
                for t in res["balises_revelatrices"][:8]:
                    detail.append(f"    {t['balise']}  ({', '.join(t['mots'])})")
            if res.get("balises_alterees"):
                detail.append(T("alterees_controle").format(n=len(res["balises_alterees"])))
                for t in res["balises_alterees"][:8]:
                    detail.append(f"    {t['forme']}  ->  {t['balise']}  "
                                  + T("fois").format(n=t['occurrences']))
                if len(res["balises_alterees"]) > 8:
                    detail.append(T("et_autres").format(n=len(res['balises_alterees']) - 8))
            if res["balises_restantes"]:
                total = sum(t["occurrences"] for t in res["balises_restantes"])
                suffixe = T("normal_anonymise") if res["etat"] == coeur.ETAT_ANONYMISE else ""
                detail.append(T("balises_presentes").format(n=len(res['balises_restantes']),
                                                            t=total, s=suffixe))
                for t in res["balises_restantes"][:8]:
                    detail.append(f"    {t['balise']} : {t['zone']} "
                                  + T("fois").format(n=t['occurrences']))
                if len(res["balises_restantes"]) > 8:
                    detail.append(T("et_autres").format(n=len(res['balises_restantes']) - 8))
            corps = (entete + res["conseil"] + "\n\n" + ("\n".join(detail) + "\n\n" if detail else "")
                     + T("rapport").format(r=rapport))
            if res["alerte"]:
                self._statut(T("alerte").format(t=res['titre']), ROUGE)
                messagebox.showwarning(res["titre"], corps)
            else:
                self._statut(T("zones_examinees").format(t=res['titre'], n=len(res['zones'])), VERT)
                messagebox.showinfo(res["titre"], corps)

        self._en_arriere_plan(travail, puis)

    def _detecter(self):
        if not self.chemin_doc.get():
            messagebox.showwarning(T("document_manquant"), T("choisir_document_word_pdf"))
            return
        if not self._exige_onglet():
            return
        table = self._table_courante()
        chemin, onglet = self.chemin_doc.get(), self.nom_onglet.get()
        type_doc, sens = self._type_choisi(), self._sens_choisi()
        # Onglet a l'ancien modele : l'injection y serait refusee. On le dit
        # AVANT de detecter (15/10/2026). Les onglets crees par le programme
        # ne sont jamais concernes : rien ne change pour les autres utilisateurs.
        try:
            ancien = modele.modele_onglet(table["fichier"], onglet) == "ancien"
        except Exception:
            ancien = False
        if ancien:
            outil_conv = DOSSIER_APP / "Convertir anciens onglets.bat"
            msg = T("ancien_bloque").format(o=onglet)
            if outil_conv.is_file():
                if messagebox.askyesno(T("ancien_bloque_titre"), msg + T("ancien_convertir")):
                    ouvrir(outil_conv)
            else:
                messagebox.showwarning(T("ancien_bloque_titre"), msg)
            return
        self._statut(T("analyse"))

        def progression(page, total):
            self._statut(T("ocr_page").format(p=page, t=total), AMBRE)

        def demander_sens():
            return self._depuis_la_tache(self._demander_sens)

        def travail():
            entites, infos, _, source = det.analyser_document(
                chemin, table["fichier"], onglet, str(CHEMIN_CONFIG_DETECTION),
                progression=progression, type_doc=type_doc, sens=sens,
                demander_sens=demander_sens)
            if not entites:
                return None
            base = os.path.splitext(os.path.basename(chemin))[0]
            # Classeur EN CLAIR (vrais noms) : range dans le sous-dossier _Prive.
            sortie = os.path.join(coeur.dossier_prive(chemin), f"_Detection_{base}.xlsx")
            _, existantes = modele.lire_onglet(table["fichier"], onglet)
            classeur.ecrire_classeur_detection(sortie, entites, infos, self.config_det,
                                               chemin, table["nom"], onglet, existantes)
            return entites, infos, source, sortie

        def puis(resultat):
            if resultat is None:
                messagebox.showinfo(T("aucun_terme"), T("aucun_terme_detail"))
                self._statut("")
                return
            entites, infos, source, sortie = resultat
            self.chemin_detection.set(sortie)
            retenus = sum(1 for e in entites if roles.proposer_a_retenir(e) == "OUI")
            a_classer = sum(1 for e in entites if not e.partie)
            types = self.config_det.get("types_documents") or {}
            libelle_type = types.get(infos.get("type"), {}).get("libelle", T("non_reconnu"))
            if source == "ocr":
                msg = T("ocr_avertissement")
            elif source == "mixte":
                msg = T("ocr_avertissement_mixte").format(
                    p=det.plages_pages(infos.get("pages_ocr")))
            else:
                msg = ""
            msg += T("detection_resume").format(t=libelle_type, n=len(entites), r=retenus,
                                                v=len(entites) - retenus, c=a_classer)
            for a in infos.get("avertissements", []):
                msg += T("attention").format(a=a)
            if infos.get("modele_onglet") == "ancien":
                msg += T("ancien_modele")
            msg += T("classeur_a_relire").format(s=sortie)
            self._statut(T("detection_ok").format(n=len(entites)), VERT)
            if messagebox.askyesno(T("detection_terminee"), msg + T("ouvrir_classeur_question")):
                ouvrir(sortie)

        def sur_erreur(e):
            if isinstance(e, det.OcrIndisponible):
                self._statut(T("ocr_indisponible"), ROUGE)
                messagebox.showerror(T("ocr_impossible"), str(e))
            else:
                self._statut(T("detection_interrompue"), ROUGE)
                messagebox.showerror(T("detection_interrompue_titre"), str(e))

        self._en_arriere_plan(travail, puis, sur_erreur)

    def _demander_sens(self):
        oui = messagebox.askyesno(T("sens_question_titre"), T("sens_question"))
        return roles.SENS_EMIS if oui else roles.SENS_RECU

    def _relire(self):
        """Etape 2 : ouvre le classeur de detection dans Excel pour relecture."""
        if not self.chemin_detection.get():
            messagebox.showwarning(T("classeur_manquant"), T("classeur_manquant_detail"))
            return
        ouvrir(self.chemin_detection.get())
        self._statut(T("relecture_en_cours"), AMBRE)

    def _injecter(self):
        if not self.chemin_detection.get():
            messagebox.showwarning(T("classeur_manquant"), T("classeur_manquant_detail"))
            return
        if not self._exige_onglet():
            return
        table = self._table_courante()
        # Excel encore ouvert ? Le classeur de detection non enregistre serait lu
        # dans son ancien etat ; la table, elle, ne peut pas etre ecrite.
        if ouvert_dans_excel(self.chemin_detection.get()):
            messagebox.showwarning(T("excel_ouvert_titre"),
                                   T("classeur_ouvert_excel").format(s=self.chemin_detection.get()))
            return
        if ouvert_dans_excel(table["fichier"]):
            messagebox.showwarning(T("excel_ouvert_titre"),
                                   T("table_ouverte_excel").format(s=table["fichier"]))
            return
        try:
            retenues = classeur.lire_classeur_detection(self.chemin_detection.get(), self.config_det)
        except Exception as e:
            messagebox.showerror(T("classeur_illisible"), str(e))
            return
        if not retenues:
            messagebox.showinfo(T("rien_a_injecter"), T("rien_a_injecter_detail"))
            return
        if not messagebox.askyesno(
                T("confirmer"),
                T("confirmer_injection").format(n=len(retenues), t=table['nom'],
                                                o=self.nom_onglet.get())):
            return
        try:
            ecrits, refuses = classeur.injecter(table["fichier"], self.nom_onglet.get(),
                                                retenues, self.config_det)
        except Exception as e:
            self._statut(T("injection_interrompue"), ROUGE)
            messagebox.showerror(T("injection_interrompue_titre"), str(e))
            return
        detail = "\n".join(f"    {b}  <-  {t[:44]}" + (T("variantes") if q != "nouvelle ligne" else "")
                           for t, b, q in ecrits[:15])
        if len(ecrits) > 15:
            detail += "\n" + T("et_autres").format(n=len(ecrits) - 15)
        msg = T("ecrits").format(n=len(ecrits), d=detail)
        if refuses:
            msg += T("non_places").format(d="\n".join(f"    {t[:36]} : {r}" for t, r in refuses[:10]))
        msg += T("dossier_pret")
        self._statut(T("injection_ok").format(e=len(ecrits), r=len(refuses)),
                     VERT if not refuses else AMBRE)
        messagebox.showinfo(T("injection_terminee"), msg)

    # ---------- retours ----------

    def _succes(self, total, nom_sortie, msg):
        self._statut(T("succes").format(n=total, s=nom_sortie), VERT)
        messagebox.showinfo(T("termine"), msg)

    def _echec(self, e, complement=""):
        self._statut(T("interrompue"), ROUGE)
        messagebox.showerror(T("interrompue_titre"), str(e) + (complement or ""))


# ------------------------------------------------------------------
# TEXTE DE L'AIDE
# ------------------------------------------------------------------



# ------------------------------------------------------------------

def verifier_installation(sortie):
    """Mode  Anonymiseur.exe --verifier <fichier.txt>  : rejoue la recette du
    paquet dans un dossier temporaire, ecrit le compte rendu, et dit le resultat."""
    import test_paquet
    code = test_paquet.main(sortie)
    try:
        racine = tk.Tk()
        racine.withdraw()
        if code == 0:
            messagebox.showinfo(T("verif_ok"), T("verif_ok_detail").format(s=sortie))
        else:
            messagebox.showerror(T("verif_ko"), T("verif_ko_detail").format(s=sortie))
    except Exception:
        pass
    sys.exit(code)


def main():
    if len(sys.argv) >= 2 and sys.argv[1] == "--verifier":
        verifier_installation(sys.argv[2] if len(sys.argv) > 2
                              else str(DOSSIER_APP / "_verification.txt"))
        return
    try:
        crees = preparer_dossier()
    except Exception as e:
        tk.Tk().withdraw()
        messagebox.showerror(T("premier_demarrage_impossible"),
                             T("premier_demarrage_detail").format(d=DOSSIER_APP, e=e))
        sys.exit(1)
    try:
        config = coeur.charger_config(CHEMIN_CONFIG)
    except coeur.ErreurAnonymisation as e:
        tk.Tk().withdraw()
        messagebox.showerror(T("configuration"), str(e))
        sys.exit(1)
    tables_ok, problemes = coeur.verifier_tables(config)
    if not tables_ok:
        tk.Tk().withdraw()
        messagebox.showerror(T("aucune_table"),
                             T("aucune_table_detail").format(p="\n\n".join(problemes)))
        sys.exit(1)
    App(config, tables_ok, problemes, crees).mainloop()


if __name__ == "__main__":
    main()
