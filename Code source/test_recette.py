"""Licence : GNU Affero General Public License v3 (LICENSE_AGPL-3.0.txt).

test_recette.py
---------------
Recette automatique de l'outil d'anonymisation.

Ne touche JAMAIS aux dossiers reels : le script fabrique lui-meme, dans un
sous-dossier temporaire, des tables de correspondances et des documents
entierement fictifs, puis verifie le comportement attendu.

Lancement : double-clic sur Recette.bat
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

import anonymisation_core as coeur


OK, KO = [], []


def verifier(libelle, condition, detail=""):
    if condition:
        OK.append(libelle)
        print(f"  [OK ] {libelle}")
    else:
        KO.append((libelle, detail))
        print(f"  [KO ] {libelle}" + (f"\n         {detail}" if detail else ""))


def attend_refus(libelle, fonction, extrait_attendu=None):
    """Verifie qu'une operation interdite leve bien une erreur metier."""
    try:
        fonction()
    except coeur.ErreurAnonymisation as e:
        if extrait_attendu and extrait_attendu.lower() not in str(e).lower():
            verifier(libelle, False, f"refus obtenu mais message inattendu : {e}")
        else:
            verifier(libelle, True)
        return
    except Exception as e:
        verifier(libelle, False, f"mauvais type d'erreur : {type(e).__name__} : {e}")
        return
    verifier(libelle, False, "AUCUN refus : l'operation a ete autorisee")


# ==================================================================
# JEU D'ESSAI FICTIF
# ==================================================================

EN_TETES = ["Terme original", "Balise", "Catégorie", "Camp", "Dossier", "Actif ?", "Notes"]

# Noms volontairement imaginaires.
DOSSIER_1 = [
    ("Aldebaran Systemes / Aldebaran", "[SOC_DEM_1]", "OUI"),
    ("Pierre Vaillantin", "[PP_DEM_1]", "OUI"),
    ("Quentin Morsange", "[PP_DEF_1]", "OUI"),
    ("81234567800019", "[SIRET_1]", "OUI"),
    ("Table Ignoree SARL", "[SOC_DEF_9]", "NON"),
]
DOSSIER_2 = [
    ("Vega Industries", "[SOC_DEM_1]", "OUI"),
    ("Berthe Nolarge", "[PP_DEM_1]", "OUI"),
    ("Come Vanterre", "[PP_DEF_1]", "OUI"),
    ("49876543200025", "[SIRET_1]", "OUI"),
]
DOSSIER_3 = [
    ("Sirius Conseil", "[SOC_DEM_1]", "OUI"),
    ("Louis Perthuis", "[PP_DEM_1]", "OUI"),
]


def ecrire_table(chemin, onglets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Mode d'emploi")
    ws["A1"] = "Onglet technique : doit etre masque par l'outil"
    for nom, lignes in onglets.items():
        ws = wb.create_sheet(nom)
        ws["A1"] = "TABLEAU DE CORRESPONDANCES — FICTIF"
        ws["A2"] = "Jeu d'essai, aucune donnee reelle"
        for i, valeur in enumerate(EN_TETES, start=1):
            ws.cell(row=3, column=i, value=valeur)
        r = 4
        for terme, balise, actif in lignes:
            ws.cell(row=r, column=1, value=terme)
            ws.cell(row=r, column=2, value=balise)
            ws.cell(row=r, column=6, value=actif)
            r += 1
    ws = wb.create_sheet("Dossiers_DataSrc")
    ws["A1"] = "Colonne1"
    wb.save(chemin)


def ecrire_document(chemin, lignes, avec_tableau=True, avec_entete=True):
    doc = Document()
    for ligne in lignes:
        doc.add_paragraph(ligne)
    if avec_tableau:
        t = doc.add_table(rows=2, cols=2)
        t.cell(0, 0).text = "Demandeur"
        t.cell(0, 1).text = "Aldebaran Systemes"
        t.cell(1, 0).text = "SIRET"
        t.cell(1, 1).text = "81234567800019"
    if avec_entete:
        doc.sections[0].header.paragraphs[0].text = "Dossier Pierre Vaillantin contre Quentin Morsange"
    doc.save(chemin)


def texte_complet(chemin):
    return coeur.texte_du_document(chemin)



# ------------------------------------------------------------------
# Un .docx fabrique a la main, pour loger un terme dans chacune des zones
# que python-docx n'expose pas. C'est la seule facon de verifier que le
# controle avant envoi les atteint vraiment.
# ------------------------------------------------------------------

def docx_piege(chemin, nom_corps, nom_note, nom_zone, nom_commentaire,
               nom_proprietes, url_client):
    import zipfile

    def para(texte):
        return f"<w:p><w:r><w:t>{texte}</w:t></w:r></w:p>"

    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {para(f"Le present acte concerne {nom_corps}.")}
    <w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml">
      <v:textbox><w:txbxContent>{para(f"Encadre : {nom_zone}")}</w:txbxContent></v:textbox>
    </v:shape></w:pict></w:r></w:p>
    <w:p><w:hyperlink r:id="rId9"><w:r><w:t>notre site</w:t></w:r></w:hyperlink></w:p>
  </w:body>
</w:document>"""

    footnotes = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="2">{para(f"Voir le contrat conclu avec {nom_note}.")}</w:footnote>
</w:footnotes>"""

    comments = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1">{para(f"A verifier avec {nom_commentaire}.")}</w:comment>
</w:comments>"""

    core = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>{nom_proprietes}</dc:creator>
  <cp:lastModifiedBy>{nom_proprietes}</cp:lastModifiedBy>
</cp:coreProperties>"""

    rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="{url_client}" TargetMode="External"/>
</Relationships>"""

    with zipfile.ZipFile(chemin, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", document)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/comments.xml", comments)
        z.writestr("docProps/core.xml", core)
        z.writestr("word/_rels/document.xml.rels", rels)



# ------------------------------------------------------------------
# Un VRAI .docx (valide, ouvrable par python-docx) portant un terme dans
# chacune des zones que python-docx n'expose pas. Les parties ajoutees sont
# declarees dans les content types et dans les relations, sans quoi elles
# seraient perdues a la premiere sauvegarde.
# ------------------------------------------------------------------

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def docx_zones_invisibles(chemin, texte_corps, texte_note, texte_note_fin,
                          texte_commentaire, texte_zone):
    import zipfile
    base = str(chemin) + ".base"
    d = Document()
    d.add_paragraph(texte_corps)
    d.save(base)

    def para(t):
        return f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>"

    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    parties = {
        "word/footnotes.xml":
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes {ns}><w:footnote w:id="2">{para(texte_note)}</w:footnote></w:footnotes>',
        "word/endnotes.xml":
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:endnotes {ns}><w:endnote w:id="2">{para(texte_note_fin)}</w:endnote></w:endnotes>',
        "word/comments.xml":
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:comments {ns}><w:comment w:id="1">{para(texte_commentaire)}</w:comment></w:comments>',
    }
    types = {"word/footnotes.xml": ("footnotes", "footnotes"),
             "word/endnotes.xml": ("endnotes", "endnotes"),
             "word/comments.xml": ("comments", "comments")}

    zone_texte = (
        '<w:p><w:r><w:pict>'
        '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="zt1" type="#_x0000_t202"'
        ' style="width:200pt;height:50pt">'
        f'<v:textbox><w:txbxContent>{para(texte_zone)}</w:txbxContent></v:textbox>'
        '</v:shape></w:pict></w:r></w:p>')

    with zipfile.ZipFile(base) as z1, \
         zipfile.ZipFile(chemin, "w", zipfile.ZIP_DEFLATED) as z2:
        for it in z1.infolist():
            data = z1.read(it.filename)
            if it.filename == "[Content_Types].xml":
                ajouts = "".join(
                    f'<Override PartName="/{nom}" ContentType="{CT}.{t}+xml"/>'
                    for nom, (t, _) in types.items())
                data = data.replace(b"</Types>", ajouts.encode() + b"</Types>")
            elif it.filename == "word/_rels/document.xml.rels":
                ajouts = "".join(
                    f'<Relationship Id="rId{900+i}" Type="{REL}/{r}"'
                    f' Target="{os.path.basename(nom)}"/>'
                    for i, (nom, (_, r)) in enumerate(types.items()))
                data = data.replace(b"</Relationships>",
                                    ajouts.encode() + b"</Relationships>")
            elif it.filename == "word/document.xml":
                data = data.replace(b"</w:body>", zone_texte.encode() + b"</w:body>")
            z2.writestr(it, data)
        for nom, contenu in parties.items():
            z2.writestr(nom, contenu)

    os.remove(base)


def textes_zones_invisibles(chemin):
    """Texte de chaque zone invisible, pour comparer avant et apres."""
    import zipfile
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    resultat = {}
    with zipfile.ZipFile(chemin) as z:
        for nom in z.namelist():
            if nom in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
                racine = etree.fromstring(z.read(nom))
                resultat[nom] = "".join(t.text or "" for t in racine.iter(f"{W}t"))
            elif nom == "word/document.xml":
                racine = etree.fromstring(z.read(nom))
                zones = []
                for boite in racine.iter(f"{W}txbxContent"):
                    zones.append("".join(t.text or "" for t in boite.iter(f"{W}t")))
                if zones:
                    resultat["zone de texte"] = " ".join(zones)
    return resultat



def ecrire_document_mise_en_forme(chemin, paragraphes):
    """paragraphes : liste de listes de (texte, gras, italique, souligne)."""
    doc = Document()
    for fragments in paragraphes:
        para = doc.add_paragraph()
        for texte, gras, italique, souligne in fragments:
            r = para.add_run(texte)
            r.bold, r.italic, r.underline = gras, italique, souligne
    doc.save(chemin)


def lire_fragments(chemin, index_para=0):
    """Fragments non vides d'un paragraphe, avec leur mise en forme."""
    para = Document(chemin).paragraphs[index_para]
    return [(r.text, r.bold is True, r.italic is True, r.underline is True)
            for r in para.runs if r.text]


def mise_en_forme_de(fragments, fragment_texte):
    for texte, gras, italique, souligne in fragments:
        if texte == fragment_texte:
            return (gras, italique, souligne)
    return None



def docx_avec_proprietes_et_lien(chemin, texte, auteur, societe, titre, url):
    """Un .docx valide portant un auteur, une societe, un titre et un lien
    hypertexte externe : les trois vecteurs invisibles."""
    import zipfile
    base = str(chemin) + ".base"
    d = Document()
    d.add_paragraph(texte)
    d.core_properties.author = auteur
    d.core_properties.last_modified_by = auteur
    d.core_properties.title = titre
    d.save(base)

    APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           f'<Company>{societe}</Company></Properties>')

    with zipfile.ZipFile(base) as z1, \
         zipfile.ZipFile(chemin, "w", zipfile.ZIP_DEFLATED) as z2:
        vu_app = False
        for it in z1.infolist():
            data = z1.read(it.filename)
            if it.filename == "word/_rels/document.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    f'<Relationship Id="rIdLien" Type="{REL}/hyperlink"'
                    f' Target="{url}" TargetMode="External"/></Relationships>'.encode())
            elif it.filename == "docProps/app.xml":
                data = APP.encode(); vu_app = True
            elif it.filename == "[Content_Types].xml" and b"app.xml" not in data:
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/docProps/app.xml" ContentType="application/'
                    b'vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')
            z2.writestr(it, data)
        if not vu_app:
            z2.writestr("docProps/app.xml", APP)
    os.remove(base)


def lire_annexes(chemin):
    """Auteur, societe, titre et cibles de liens d'un .docx."""
    import zipfile
    from lxml import etree
    res = {"auteur": None, "societe": None, "titre": None, "liens": []}
    with zipfile.ZipFile(chemin) as z:
        noms = z.namelist()
        if "docProps/core.xml" in noms:
            racine = etree.fromstring(z.read("docProps/core.xml"))
            for el in racine.iter():
                if el.tag.endswith("}creator"):
                    res["auteur"] = el.text
                elif el.tag.endswith("}title"):
                    res["titre"] = el.text
        if "docProps/app.xml" in noms:
            racine = etree.fromstring(z.read("docProps/app.xml"))
            for el in racine.iter():
                if el.tag.endswith("}Company"):
                    res["societe"] = el.text
        for n in noms:
            if n.endswith(".rels"):
                racine = etree.fromstring(z.read(n))
                for el in racine.iter():
                    if el.get("TargetMode") == "External":
                        res["liens"].append(el.get("Target"))
    return res


# ==================================================================
# RECETTE
# ==================================================================

# ==================================================================
# RECETTE DES CORRECTIONS DE L'AUDIT D'OCTOBRE 2026
# ==================================================================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _docx_corps_xml(chemin, corps_xml, entete_tableau=None, propriete_perso=None):
    """Un vrai .docx dont le corps est ecrit a la main (XML), pour loger des
    termes la ou python-docx ne regarde pas : liens, suivi des modifications,
    controle de contenu, tableau imbrique, appel de note dans le meme run."""
    import zipfile
    base = str(chemin) + ".base"
    d = Document()
    d.add_paragraph("PLACEHOLDER")
    if entete_tableau:
        t = d.sections[0].header.add_table(rows=1, cols=1, width=d.sections[0].page_width)
        t.cell(0, 0).text = entete_tableau
    d.save(base)
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(chemin, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                texte = data.decode("utf-8")
                debut = texte.index("<w:body>") + len("<w:body>")
                fin = texte.index("<w:sectPr")
                texte = texte[:debut] + corps_xml + texte[fin:]
                data = texte.encode("utf-8")
            if propriete_perso and item.filename == "[Content_Types].xml":
                data = data.decode("utf-8").replace(
                    "</Types>",
                    '<Override PartName="/docProps/custom.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>').encode("utf-8")
            if propriete_perso and item.filename == "_rels/.rels":
                data = data.decode("utf-8").replace(
                    "</Relationships>",
                    '<Relationship Id="rIdPerso" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/custom-properties" '
                    'Target="docProps/custom.xml"/></Relationships>').encode("utf-8")
            zout.writestr(item, data)
        if propriete_perso:
            zout.writestr("docProps/custom.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client">'
                f'<vt:lpwstr>{propriete_perso}</vt:lpwstr></property></Properties>'))
    os.remove(base)


def _lire_xml(chemin, partie="word/document.xml"):
    import zipfile
    with zipfile.ZipFile(chemin) as z:
        return z.read(partie).decode("utf-8") if partie in z.namelist() else ""


def recette_audit_2026_10(racine):
    dossier = racine / "audit_2026_10"
    dossier.mkdir()
    table = dossier / "Table_audit.xlsx"
    lignes = [
        ("Aldebaran", "[SOC_DEM_1]", "OUI"),
        ("Pierre Vaillantin", "[PP_DEM_1]", "OUI"),
        ("SA", "[FORME_1]", "OUI"),
        ("59000 Lille", "[ADRESSE_1]", "OUI"),
        ("L'Orbe Conseil", "[SOC_DEF_1]", "OUI"),
        ("Quentin Morsange", "[PP_DEF_1]", "OUI"),
    ]
    ecrire_table(table, {"Audit": lignes, "Audit bis": lignes})
    config_path = dossier / "config.json"
    config_path.write_text(json.dumps({
        "tables": [{"nom": "Table audit", "fichier": str(table)}],
        "dossier_journaux": "Journaux"}, ensure_ascii=False), encoding="utf-8")
    config = coeur.charger_config(config_path)
    entrees = coeur.charger_entrees(table, "Audit")

    print("\n12. Audit 10/2026 - point 1 : frontières de mot, variantes courtes")
    doc1 = dossier / "Mots voisins.docx"
    ecrire_document(doc1, ["Le style Aldebaranesque n'est pas Aldebaran.",
                           "La SAS reste une SAS ; la SA disparaît. SAINT reste."],
                    avec_tableau=False, avec_entete=False)
    r1 = coeur.anonymiser_document(doc1, entrees, config, "Table audit", "Audit")
    t1 = texte_complet(r1["chemin_sortie"])
    verifier("mot voisin non touché (« Aldebaranesque » intact)",
             "Aldebaranesque" in t1 and "]esque" not in t1, t1)
    verifier("le mot isolé est bien remplacé", "n'est pas [SOC_DEM_1]." in t1, t1)
    verifier("variante courte « SA » : « SAS » et « SAINT » intacts",
             t1.count("SAS") == 2 and "SAINT" in t1 and "la [FORME_1] disparaît" in t1, t1)
    verifier("variante courte signalée en avertissement",
             any("« SA »" in a for a in r1.get("avertissements", [])), str(r1.get("avertissements")))
    ctl1 = coeur.controler(r1["chemin_sortie"], entrees, config)
    verifier("le contrôle applique la même règle (aucun faux positif sur le mot voisin)",
             not ctl1["termes_en_clair"], str(ctl1["termes_en_clair"]))

    print("\n13. Audit 10/2026 - point 5 : espaces insécables et apostrophes")
    doc2 = dossier / "Typographie.docx"
    ecrire_document(doc2, ["Siège : 59000\u00a0Lille, chez L\u2019Orbe Conseil.",
                           "Adresse : 59000   Lille."], avec_tableau=False, avec_entete=False)
    r2 = coeur.anonymiser_document(doc2, entrees, config, "Table audit", "Audit")
    t2 = texte_complet(r2["chemin_sortie"])
    verifier("« 59000[insécable]Lille » retrouvé avec la saisie « 59000 Lille »",
             "Siège : [ADRESSE_1]," in t2 and "Adresse : [ADRESSE_1]." in t2, t2)
    verifier("apostrophe typographique retrouvée avec la saisie droite",
             "chez [SOC_DEF_1]." in t2, t2)
    ctl2 = coeur.controler(r2["chemin_sortie"], entrees, config)
    verifier("contrôle : plus aucun terme en clair", not ctl2["termes_en_clair"],
             str(ctl2["termes_en_clair"]))
    rr2 = coeur.desanonymiser_avec_journal(r2["chemin_sortie"], r2["chemin_journal"])
    verifier("restitution exacte de la typographie d'origine (insécable, apostrophe)",
             texte_complet(rr2["chemin_sortie"]) == texte_complet(doc2),
             texte_complet(rr2["chemin_sortie"]))

    print("\n14. Audit 10/2026 - point 3 : balises altérées par l'IA")
    reponse = dossier / "Reponse IA.docx"
    ecrire_document(reponse, [
        "Synthèse : [PP_DEM_1] reproche à [soc_dem_1] un retard.",
        "[SOC DEM 1] conteste ; SOC_DEM_1 invoque la force majeure ; [PP_DEM_1 ] insiste.",
        "Une personne physique et la société restent des mots ordinaires."],
        avec_tableau=False, avec_entete=False)
    rr3 = coeur.desanonymiser_sans_journal(reponse, entrees, confirme=True, config=config)
    t3 = texte_complet(rr3["chemin_sortie"])
    formes = {a["forme"] for a in rr3["balises_alterees"]}
    verifier("la balise exacte est restituée", t3.startswith("Synthèse : Pierre Vaillantin"), t3)
    verifier("les 4 formes altérées sont listées",
             formes == {"[soc_dem_1]", "[SOC DEM 1]", "SOC_DEM_1", "[PP_DEM_1 ]"}, str(formes))
    verifier("aucune forme altérée restituée en silence",
             "[soc_dem_1]" in t3 and "SOC_DEM_1 invoque" in t3, t3)
    verifier("les mots ordinaires ne sont pas pris pour des balises",
             not any("personne" in f.lower() or "société" in f.lower() for f in formes), str(formes))
    ctl3 = coeur.controler(rr3["chemin_sortie"], entrees, config)
    verifier("contrôle d'un _RESTITUE avec balises altérées : alerte",
             ctl3["alerte"] and len(ctl3["balises_alterees"]) == 4,
             f"{ctl3['titre']} / {ctl3['balises_alterees']}")
    rapport = coeur.ecrire_rapport(rr3["chemin_sortie"], ctl3, "Table audit", "Audit")
    verifier("le rapport liste les balises altérées",
             "BALISES ALTÉRÉES" in open(rapport, encoding="utf-8").read())

    print("\n15. Audit 10/2026 - point 7 : fichiers en clair rangés dans _Privé")
    verifier("le rapport de contrôle est rangé dans le sous-dossier _Privé",
             os.path.basename(os.path.dirname(rapport)) == coeur.DOSSIER_PRIVE[coeur.LANGUE]
             and os.path.dirname(os.path.dirname(rapport)) == str(dossier), rapport)

    print("\n16. Audit 10/2026 - point 6 : liens, suivi des modifications, contrôles de contenu")
    w = f'xmlns:w="{W_NS}"'
    corps = (
        '<w:p><w:hyperlink w:anchor="ancre"><w:r><w:t>Site de Aldebaran</w:t></w:r></w:hyperlink></w:p>'
        '<w:p><w:ins w:id="1" w:author="Relecteur" w:date="2026-10-07T10:00:00Z"><w:r>'
        '<w:t>ajout pour Quentin Morsange</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="Pierre Vaillantin" w:date="2026-10-07T10:00:00Z"><w:r>'
        '<w:delText>retrait de Pierre Vaillantin</w:delText></w:r></w:del></w:p>'
        '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Champ : Aldebaran</w:t></w:r></w:p>'
        '</w:sdtContent></w:sdt>'
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cellule</w:t>'
        '</w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r>'
        '<w:t>Imbriqué : Pierre Vaillantin</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>'
        '<w:p><w:r><w:t xml:space="preserve">Voir Aldebaran </w:t><w:footnoteReference w:id="5"/>'
        '</w:r></w:p>')
    doc4 = dossier / "Zones cachees.docx"
    _docx_corps_xml(doc4, corps, entete_tableau="En-tête : Quentin Morsange",
                    propriete_perso="Dossier Aldebaran")
    r4 = coeur.anonymiser_document(doc4, entrees, config, "Table audit", "Audit")
    xml4 = _lire_xml(r4["chemin_sortie"])
    entete4 = "".join(_lire_xml(r4["chemin_sortie"], p) for p in
                      ("word/header1.xml", "word/header2.xml", "word/header3.xml"))
    perso4 = _lire_xml(r4["chemin_sortie"], "docProps/custom.xml")
    for libelle, absent in (("texte d'un hyperlien", "Site de Aldebaran"),
                            ("insertion suivie (w:ins)", "ajout pour Quentin"),
                            ("suppression suivie (w:delText)", "retrait de Pierre"),
                            ("contrôle de contenu (w:sdt)", "Champ : Aldebaran"),
                            ("tableau imbriqué", "Imbriqué : Pierre"),
                            ("texte partageant un appel de note", "Voir Aldebaran")):
        verifier(f"anonymisé : {libelle}", absent not in xml4)
    verifier("anonymisé : tableau d'en-tête", "Quentin Morsange" not in entete4 and entete4, entete4[:200])
    verifier("anonymisé : propriété personnalisée (custom.xml)",
             "Aldebaran" not in perso4 and "[SOC_DEM_1]" in perso4, perso4)
    verifier("l'appel de note partageant le run d'un terme est conservé",
             "footnoteReference" in xml4)
    ctl4 = coeur.controler(r4["chemin_sortie"], entrees, config)
    zones4 = {t["zone"] for t in ctl4["termes_en_clair"]}
    # Depuis le 7/10/2026, l'auteur d'une revision est masque a l'anonymisation
    verifier("l'auteur d'une révision (w:author) est masqué, le contrôle n'a plus rien à y voir",
             "Pierre Vaillantin" not in " ".join(v for vs in _auteurs_xml(r4["chemin_sortie"]).values()
                                                 for v in vs)
             and not any(z.startswith(("AUTEURS", "AUTHORS")) for z in zones4), str(zones4))
    rr4 = coeur.desanonymiser_avec_journal(r4["chemin_sortie"], r4["chemin_journal"])
    xml_r = _lire_xml(rr4["chemin_sortie"])
    verifier("restitution avec journal : liens, révisions, sdt, imbriqué, note",
             all(x in xml_r for x in ("Site de Aldebaran", "ajout pour Quentin Morsange",
                                      "retrait de Pierre Vaillantin", "Champ : Aldebaran",
                                      "Imbriqué : Pierre Vaillantin", "Voir Aldebaran ",
                                      "footnoteReference")), xml_r[:400])
    verifier("restitution : propriété personnalisée remise",
             "Dossier Aldebaran" in _lire_xml(rr4["chemin_sortie"], "docProps/custom.xml"))
    rr4b = coeur.desanonymiser_sans_journal(r4["chemin_sortie"], entrees, confirme=True, config=config)
    xml_rb = _lire_xml(rr4b["chemin_sortie"])
    verifier("restitution sans journal : révisions et zones nouvelles restituées",
             "retrait de Pierre Vaillantin" in xml_rb and "Imbriqué : Pierre Vaillantin" in xml_rb
             and "[" not in texte_complet(rr4b["chemin_sortie"]), xml_rb[:300])

    print("\n17. Audit 10/2026 - point 8 : journal homonyme, cache, collision de variante")
    for sous in ("A", "B"):
        (dossier / sous).mkdir()
        ecrire_document(dossier / sous / "Conclusions.docx", ["Aldebaran contre Quentin Morsange."],
                        avec_tableau=False, avec_entete=False)
    ra = coeur.anonymiser_document(dossier / "A" / "Conclusions.docx", entrees, config,
                                   "Table audit", "Audit")
    rb = coeur.anonymiser_document(dossier / "B" / "Conclusions.docx", entrees, config,
                                   "Table audit", "Audit")
    verifier("journal homonyme : le précédent est conservé, et signalé",
             len(coeur.journaux_candidats(config, dossier / "A" / "Conclusions_ANONYMISE.docx")) == 2
             and any("conservé" in a or "kept" in a for a in rb["avertissements"]),
             str(rb["avertissements"]))

    avant = coeur.balises_declarees(config)
    wb = openpyxl.load_workbook(table)
    wb["Audit"].append(["Nouveau Terme", "[NOUVELLE_1]", None, None, None, "OUI"])
    wb.save(table)
    st = os.stat(table)
    os.utime(table, (st.st_atime, st.st_mtime + 5))
    verifier("balise ajoutée dans Excel en cours de session prise en compte (cache)",
             "[NOUVELLE_1]" not in avant and "[NOUVELLE_1]" in coeur.balises_declarees(config))

    collision = [("[SOC_DEM_1]", ["Aldebaran"]), ("[SOC_DEF_2]", ["ALDEBARAN"])]
    verifier("collision de variante entre deux balises signalée",
             any("ALDEBARAN" in a or "Aldebaran" in a for a in coeur.avertissements_entrees(collision)),
             str(coeur.avertissements_entrees(collision)))



def recette_accents_2026_10(racine):
    """Accents et ligatures (7 octobre 2026) : la saisie de la table et le
    document peuvent differer par les accents, dans un sens comme dans l'autre."""
    dossier = racine / "accents_2026_10"
    dossier.mkdir()
    table = dossier / "Table_accents.xlsx"
    lignes = [
        ("Société Dupré", "[SOC_DEM_1]", "OUI"),
        ("Helene Lefevre", "[PP_DEM_1]", "OUI"),
        ("Lœuvre", "[PP_DEF_1]", "OUI"),
        ("Aldebaran", "[SOC_DEF_1]", "OUI"),
    ]
    ecrire_table(table, {"Accents": lignes})
    config_path = dossier / "config.json"
    config_path.write_text(json.dumps({
        "tables": [{"nom": "Table accents", "fichier": str(table)}],
        "dossier_journaux": "Journaux"}, ensure_ascii=False), encoding="utf-8")
    config = coeur.charger_config(config_path)
    entrees = coeur.charger_entrees(table, "Accents")

    print("\n18. Accents et ligatures (7/10/2026)")
    doc = dossier / "Accents.docx"
    ecrire_document(doc, [
        "ENTRE LA SOCIETE DUPRE, demanderesse,",
        "et Mme Hélène LEFÈVRE, puis Hélène Lefèvre ; la Société Dupré.",
        "M. LOEUVRE répond ; Monsieur Lœuvre confirme.",
        "Le style Aldébaranesque reste intact, Aldébaran non.",
    ], avec_tableau=False, avec_entete=False)
    r = coeur.anonymiser_document(doc, entrees, config, "Table accents", "Accents")
    t = texte_complet(r["chemin_sortie"])
    verifier("table accentuée : « SOCIETE DUPRE » en capitales sans accents est remplacé",
             "ENTRE LA [SOC_DEM_1]," in t and "la [SOC_DEM_1]." in t, t)
    verifier("table sans accents : « Hélène LEFÈVRE » et « Hélène Lefèvre » sont remplacés",
             "Mme [PP_DEM_1], puis [PP_DEM_1] ;" in t, t)
    verifier("ligature : « Lœuvre » trouve « LOEUVRE » et inversement",
             "M. [PP_DEF_1] répond ; Monsieur [PP_DEF_1] confirme." in t, t)
    verifier("frontière de mot conservée avec les accents (« Aldébaranesque » intact)",
             "Aldébaranesque" in t and "[SOC_DEF_1] non." in t, t)
    ctl = coeur.controler(r["chemin_sortie"], entrees, config)
    verifier("contrôle : plus aucun terme en clair", not ctl["termes_en_clair"],
             str(ctl["termes_en_clair"]))
    rr = coeur.desanonymiser_avec_journal(r["chemin_sortie"], r["chemin_journal"])
    verifier("restitution avec journal : la forme exacte revient (capitales sans accents comprises)",
             texte_complet(rr["chemin_sortie"]) == texte_complet(doc),
             texte_complet(rr["chemin_sortie"]))

    oubli = dossier / "Oubli.docx"
    ecrire_document(oubli, ["Synthèse : la SOCIETE DUPRE est restée en clair."],
                    avec_tableau=False, avec_entete=False)
    ctl2 = coeur.controler(oubli, entrees, config)
    verifier("contrôle : un oubli écrit sans accents est signalé",
             bool(ctl2["termes_en_clair"]), str(ctl2["termes_en_clair"]))

    nom_fichier = dossier / "Conclusions Societe Dupre_ANONYMISE.docx"
    shutil.copy(r["chemin_sortie"], nom_fichier)
    verifier("nom de fichier sans accents signalé (« Conclusions Societe Dupre »)",
             bool(coeur.controler_chemin(nom_fichier, entrees)),
             str(coeur.controler_chemin(nom_fichier, entrees)))

    for nom in ("Conclusions_Societe_Dupre.docx", "Pieces-Société-Dupré.docx",
                "Note.Societe.Dupre.docx"):
        f = dossier / nom
        shutil.copy(r["chemin_sortie"], f)
        verifier(f"nom de fichier avec séparateurs signalé (« {nom} »)",
                 bool(coeur.controler_chemin(f, entrees)), nom)
    neutre = dossier / "Conclusions_Societes_Dupreville.docx"
    shutil.copy(r["chemin_sortie"], neutre)
    verifier("pas de faux positif sur un nom voisin (« Societes_Dupreville »)",
             not coeur.controler_chemin(neutre, entrees),
             str(coeur.controler_chemin(neutre, entrees)))

    doublon = [("[PP_DEM_1]", ["Lefèvre"]), ("[PP_DEM_2]", ["Lefevre"])]
    verifier("deux noms ne différant que par un accent : avertissement, sans blocage",
             any("Lefevre" in a and "[PP_DEM_1]" in a for a in coeur.avertissements_entrees(doublon)),
             str(coeur.avertissements_entrees(doublon)))


def recette_texte_2026_10(racine):
    """Fichiers .txt et .md (7 octobre 2026) : anonymisation, controle,
    restitution avec et sans journal, encodage et fins de ligne conserves."""
    dossier = racine / "texte_2026_10"
    dossier.mkdir()
    table = dossier / "Table_texte.xlsx"
    lignes = [
        ("Société Dupré", "[SOC_DEM_1]", "OUI"),
        ("Hélène Lefèvre", "[PP_DEM_1]", "OUI"),
        ("Lefèvre", "[PP_DEM_1]", "OUI"),
        ("Aldebaran", "[SOC_DEF_1]", "OUI"),
    ]
    ecrire_table(table, {"Texte": lignes, "Autre": [("Zorglub", "[PP_DEF_9]", "OUI")]})
    config_path = dossier / "config.json"
    config_path.write_text(json.dumps({
        "tables": [{"nom": "Table texte", "fichier": str(table)}],
        "dossier_journaux": "Journaux"}, ensure_ascii=False), encoding="utf-8")
    config = coeur.charger_config(config_path)
    entrees = coeur.charger_entrees(table, "Texte")

    print("\n19. Fichiers texte .txt et .md (7/10/2026)")
    contenu = ("# Note pour la SOCIETE DUPRE\r\n\r\n"
               "Mme Hélène Lefèvre reproche à Aldebaran un retard ; Mme LEFEVRE insiste.\r\n"
               "- point **Aldebaran** à vérifier\r\n")
    for nom, encodage in (("Note.md", "utf-8"), ("Note ancienne.txt", "cp1252")):
        src = dossier / nom
        src.write_bytes(contenu.encode(encodage))
        r = coeur.anonymiser_document(src, entrees, config, "Table texte", "Texte")
        sortie = pathlib.Path(r["chemin_sortie"])
        brut = sortie.read_bytes()
        t = brut.decode(encodage)
        verifier(f"{nom} : fichier anonymisé dans le même format",
                 sortie.suffix == src.suffix and "_ANONYMISE" in sortie.name, sortie.name)
        verifier(f"{nom} : termes remplacés, accents et capitales compris",
                 "la [SOC_DEM_1]" in t and "Mme [PP_DEM_1] reproche à [SOC_DEF_1]" in t
                 and "Mme [PP_DEM_1] insiste" in t and "**[SOC_DEF_1]**" in t, t)
        verifier(f"{nom} : encodage et fins de ligne conservés",
                 brut.count(b"\r\n") == contenu.count("\r\n") and "é".encode(encodage) in brut,
                 repr(brut[:80]))
        ctl = coeur.controler(sortie, entrees, config)
        verifier(f"{nom} : contrôle sans terme en clair, document reconnu anonymisé",
                 not ctl["termes_en_clair"] and ctl["etat"] == coeur.ETAT_ANONYMISE,
                 str(ctl["termes_en_clair"]))
        attend_refus(f"{nom} : pas de seconde anonymisation",
                     lambda: coeur.anonymiser_document(sortie, entrees, config,
                                                       "Table texte", "Texte"))
        rr = coeur.desanonymiser_avec_journal(sortie, r["chemin_journal"])
        verifier(f"{nom} : restitution avec journal identique à l'original, octet pour octet",
                 pathlib.Path(rr["chemin_sortie"]).read_bytes() == src.read_bytes(),
                 repr(pathlib.Path(rr["chemin_sortie"]).read_bytes()[:120]))

    # Reponse d'une IA collee dans un .txt : restitution SANS journal.
    reponse = dossier / "Reponse IA.txt"
    reponse.write_text("Synthèse : [PP_DEM_1] reproche à [SOC_DEF_1] un retard ; "
                       "[SOC DEM 1] conteste.\n", encoding="utf-8")
    rs = coeur.desanonymiser_sans_journal(reponse, entrees, confirme=True, config=config)
    t = pathlib.Path(rs["chemin_sortie"]).read_text(encoding="utf-8")
    verifier("réponse IA en .txt : restitution sans journal",
             "Hélène Lefèvre reproche à Aldebaran" in t, t)
    verifier("réponse IA en .txt : balise altérée laissée en place et listée",
             "[SOC DEM 1]" in t and any(a["forme"] == "[SOC DEM 1]"
                                         for a in rs["balises_alterees"]),
             str(rs["balises_alterees"]))
    autre = coeur.charger_entrees(table, "Autre")
    attend_refus("réponse IA en .txt : une table qui ne couvre pas les balises bloque",
                 lambda: coeur.desanonymiser_sans_journal(reponse, autre, confirme=True,
                                                          config=config), "ne couvre pas")

    oubli = dossier / "Oubli_ANONYMISE.txt"
    oubli.write_text("[PP_DEM_1] et la Societe Dupre.\n", encoding="utf-8")
    verifier("contrôle d'un .txt : un oubli est signalé",
             bool(coeur.controler(oubli, entrees, config)["termes_en_clair"]))
    import detection_core as det
    texte_det, source_det = det.extraire_texte_ou_ocr(str(dossier / "Note.md"))
    verifier("pré-détection : un .md est lu comme du texte",
             source_det == "texte" and "Aldebaran" in texte_det, source_det)


def recette_hors_table_2026_10(racine):
    """Controle des donnees hors table (7 octobre 2026) : le detecteur relit le
    document anonymise. ALERTE pour IBAN, NIR, courriel, telephone ; simple
    signalement pour les noms ; rien quand le document est propre."""
    dossier = racine / "hors_table_2026_10"
    dossier.mkdir()
    table = dossier / "Table_hors.xlsx"
    lignes = [
        ("Société Dupré", "[SOC_DEM_1]", "OUI"),
        ("Hélène Lefèvre", "[PP_DEM_1]", "OUI"),
        ("Lefèvre", "[PP_DEM_1]", "OUI"),
    ]
    ecrire_table(table, {"Hors": lignes})
    config_path = dossier / "config.json"
    config_path.write_text(json.dumps({
        "tables": [{"nom": "Table hors", "fichier": str(table)}],
        "dossier_journaux": "Journaux"}, ensure_ascii=False), encoding="utf-8")
    config = coeur.charger_config(config_path)
    entrees = coeur.charger_entrees(table, "Hors")

    print("\n20. Contrôle des données hors table (7/10/2026)")
    propre = dossier / "Propre.docx"
    ecrire_document(propre, [
        "Monsieur le Président, Mesdames et Messieurs les Conseillers,",
        "La Société Dupré, représentée par Madame Hélène Lefèvre, demeurant à Lille CEDEX.",
        "Vu l'article 1103 du Code civil et CJUE, 3 octobre 2019, C-18/18, RG n° 24/01234.",
        "La somme de 15 000,00 € TTC, facture n° 2024-0156 du 03.02.2024.",
        "Le logiciel Amazon WEB SERVICES, version 2.3.1. Madame la Présidente a statué.",
    ], avec_tableau=False, avec_entete=False)
    r = coeur.anonymiser_document(propre, entrees, config, "Table hors", "Hors")
    ctl = coeur.controler(r["chemin_sortie"], entrees, config)
    verifier("document propre : ni alerte, ni nom à vérifier",
             not ctl["alerte"] and not ctl["hors_table_alertes"]
             and not ctl["hors_table_a_verifier"],
             str(ctl["hors_table_alertes"] + ctl["hors_table_a_verifier"]))

    oublis = dossier / "Oublis.docx"
    ecrire_document(oublis, [
        "La Société Dupré, représentée par Madame Hélène Lefèvre.",
        "Virement sur le compte FR76 3000 6000 0112 3456 7890 189.",
        "Numéro de sécurité sociale : 2 55 08 14 168 025 38.",
        "Contact : accueil@cabinet-tiers.fr ou 06 12 34 56 78.",
        "Témoin : Monsieur Gaspard Fontanel. Copie à Jean-Pierre MARTIN.",
    ], avec_tableau=False, avec_entete=False)
    r = coeur.anonymiser_document(oublis, entrees, config, "Table hors", "Hors")
    ctl = coeur.controler(r["chemin_sortie"], entrees, config)
    types = {t["type"] for t in ctl["hors_table_alertes"]}
    verifier("IBAN, NIR, courriel et téléphone hors table : ALERTE",
             ctl["alerte"] and len(ctl["hors_table_alertes"]) == 4
             and not ctl["termes_en_clair"], str(types))
    noms = {t["terme"] for t in ctl["hors_table_a_verifier"]}
    verifier("civilité + nom et Prénom NOM : « à vérifier »",
             any("Gaspard Fontanel" in n for n in noms) and "Jean-Pierre MARTIN" in noms,
             str(noms))
    verifier("un nom remplacé par sa balise n'est pas signalé",
             not any("Lefèvre" in n or "[" in n for n in noms), str(noms))
    rapport = pathlib.Path(coeur.ecrire_rapport(r["chemin_sortie"], ctl, "Table hors", "Hors"))
    contenu = rapport.read_text(encoding="utf-8")
    verifier("le rapport liste les deux niveaux",
             "ALERTE, DONNÉES HORS TABLE" in contenu and "À VÉRIFIER (non bloquant)" in contenu)

    auteurs = dossier / "Auteurs_ANONYMISE.docx"
    _docx_corps_xml(auteurs, '<w:p><w:ins w:id="1" w:author="Gaspard Fontanel" '
                    'w:date="2026-10-07T10:00:00Z"><w:r><w:t>La [SOC_DEM_1] '
                    'conteste.</w:t></w:r></w:ins></w:p>')
    ctl = coeur.controler(auteurs, entrees, config)
    verifier("auteurs des révisions : non signalés (le programme ne peut pas les masquer)",
             ctl["etat"] == coeur.ETAT_ANONYMISE and not ctl["hors_table_alertes"] and not ctl["hors_table_a_verifier"],
             str(ctl["hors_table_a_verifier"]))

    seulement_noms = dossier / "Noms.docx"
    ecrire_document(seulement_noms, ["La Société Dupré a écrit à Monsieur Gaspard Fontanel."],
                    avec_tableau=False, avec_entete=False)
    r = coeur.anonymiser_document(seulement_noms, entrees, config, "Table hors", "Hors")
    ctl = coeur.controler(r["chemin_sortie"], entrees, config)
    verifier("des noms seuls à vérifier ne déclenchent pas d'alerte",
             not ctl["alerte"] and ctl["hors_table_a_verifier"])

    texte = dossier / "Note.txt"
    texte.write_text("La Société Dupré, IBAN FR76 3000 6000 0112 3456 7890 189.\n",
                     encoding="utf-8")
    r = coeur.anonymiser_document(texte, entrees, config, "Table hors", "Hors")
    ctl = coeur.controler(r["chemin_sortie"], entrees, config)
    verifier("fichier .txt : l'IBAN hors table est aussi signalé",
             ctl["alerte"] and len(ctl["hors_table_alertes"]) == 1,
             str(ctl["hors_table_alertes"]))

    restitue = coeur.desanonymiser_avec_journal(r["chemin_sortie"], r["chemin_journal"])
    ctl = coeur.controler(restitue["chemin_sortie"], entrees, config)
    verifier("document restitué : pas de contrôle hors table (les données y sont normales)",
             not ctl["hors_table_alertes"] and not ctl["hors_table_a_verifier"])


def _auteurs_xml(chemin):
    """Valeurs des attributs author / initials / userId, partie par partie."""
    import zipfile
    from lxml import etree
    vus = {}
    with zipfile.ZipFile(chemin) as z:
        for nom in z.namelist():
            if nom.startswith("word/") and nom.endswith(".xml"):
                racine = etree.fromstring(z.read(nom))
                for el in racine.iter():
                    if not isinstance(el.tag, str):
                        continue
                    for k, v in el.attrib.items():
                        if k.rsplit("}", 1)[-1] in ("author", "initials", "userId"):
                            vus.setdefault(nom, []).append(v)
    return vus


def recette_auteurs_2026_10(racine):
    """Auteurs des commentaires et des modifications suivies (7 octobre 2026) :
    masques a l'anonymisation, remis a la restitution avec journal."""
    import zipfile
    dossier = racine / "auteurs_2026_10"
    dossier.mkdir()
    table = dossier / "Table_auteurs.xlsx"
    ecrire_table(table, {"Auteurs": [("Aldebaran", "[SOC_DEM_1]", "OUI")]})

    def config_avec(masquer):
        chemin = dossier / f"config_{masquer}.json"
        chemin.write_text(json.dumps({
            "tables": [{"nom": "Table auteurs", "fichier": str(table)}],
            "dossier_journaux": "Journaux", "masquer_proprietes": masquer},
            ensure_ascii=False), encoding="utf-8")
        return coeur.charger_config(chemin)

    entrees = coeur.charger_entrees(table, "Auteurs")

    def fabriquer(chemin):
        base = dossier / "base.docx"
        docx_zones_invisibles(base, "Aldebaran conteste.", "Note neutre.", "Fin neutre.",
                              "Commentaire sur Aldebaran.", "Encadre neutre.")
        revision = ('<w:p><w:ins w:id="7" w:author="Gaspard Fontanel" w:date="2026-10-07T10:00:00Z">'
                    '<w:r><w:t>Ajout de Aldebaran.</w:t></w:r></w:ins>'
                    '<w:del w:id="8" w:author="Lucie Ormesson" w:date="2026-10-07T11:00:00Z">'
                    '<w:r><w:delText>Texte retire.</w:delText></w:r></w:del></w:p>')
        with zipfile.ZipFile(base) as z1, zipfile.ZipFile(chemin, "w", zipfile.ZIP_DEFLATED) as z2:
            for it in z1.infolist():
                data = z1.read(it.filename)
                if it.filename == "word/comments.xml":
                    data = data.replace(b'<w:comment w:id="1">',
                                        '<w:comment w:id="1" w:author="Lucie Ormesson" '
                                        'w:initials="LO" w:date="2026-10-07T09:00:00Z">'.encode())
                elif it.filename == "word/document.xml":
                    data = data.replace(b"</w:body>", revision.encode() + b"</w:body>")
                z2.writestr(it, data)
        os.remove(base)

    print("\n21. Auteurs des commentaires et des modifications suivies (7/10/2026)")
    doc = dossier / "Revisions.docx"
    fabriquer(doc)
    avant = _auteurs_xml(doc)
    r = coeur.anonymiser_document(doc, entrees, config_avec(True), "Table auteurs", "Auteurs")
    apres = _auteurs_xml(r["chemin_sortie"])
    tout = " ".join(v for vs in apres.values() for v in vs)
    verifier("anonymisation : plus aucun nom d'auteur ni initiales",
             "Fontanel" not in tout and "Ormesson" not in tout and "LO" not in tout.split(), tout)
    verifier("anonymisation : auteurs numérotés « Auteur 1 », « Auteur 2 », même numéro partout",
             apres["word/document.xml"] == ["Auteur 1", "Auteur 2"]
             and apres["word/comments.xml"] == ["Auteur 2", "A2"], str(apres))
    verifier("anonymisation : le texte du commentaire reste anonymisé",
             "[SOC_DEM_1]" in textes_zones_invisibles(r["chemin_sortie"]).get("word/comments.xml", ""))
    ctl = coeur.controler(r["chemin_sortie"], entrees, config_avec(True))
    verifier("contrôle : ni terme en clair ni donnée hors table",
             not ctl["alerte"] and not ctl["hors_table_a_verifier"],
             str(ctl["termes_en_clair"]) + str(ctl["hors_table_a_verifier"]))

    rr = coeur.desanonymiser_avec_journal(r["chemin_sortie"], r["chemin_journal"])
    verifier("restitution avec journal : chaque auteur retrouve son nom et ses initiales",
             _auteurs_xml(rr["chemin_sortie"]) == avant,
             f"{_auteurs_xml(rr['chemin_sortie'])} / {avant}")
    verifier("restitution avec journal : le texte revient aussi",
             "Aldebaran" in textes_zones_invisibles(rr["chemin_sortie"]).get("word/comments.xml", ""))

    doc2 = dossier / "Revisions sans masquage.docx"
    fabriquer(doc2)
    r2 = coeur.anonymiser_document(doc2, entrees, config_avec(False), "Table auteurs", "Auteurs")
    verifier("masquer_proprietes à false : les auteurs sont laissés tels quels",
             _auteurs_xml(r2["chemin_sortie"]) == _auteurs_xml(doc2))
    table_nom = dossier / "Table_nom.xlsx"
    ecrire_table(table_nom, {"Auteurs": [("Aldebaran", "[SOC_DEM_1]", "OUI"),
                                         ("Lucie Ormesson", "[PP_DEM_1]", "OUI")]})
    ctl2 = coeur.controler(r2["chemin_sortie"], coeur.charger_entrees(table_nom, "Auteurs"),
                           config_avec(False))
    verifier("sans masquage, le contrôle voit encore un nom de la table dans les auteurs",
             any(t["zone"].startswith(("AUTEURS", "AUTHORS")) for t in ctl2["termes_en_clair"]),
             str([t["zone"] for t in ctl2["termes_en_clair"]]))



def recette_revue_2026_10_08(racine):
    """Revue du 8 octobre 2026 : sauts de ligne et tabulations dans un
    paragraphe, titre recopie dans app.xml, rapports de controle distincts,
    sortie jamais confondue avec l'entree."""
    import zipfile
    dossier = racine / "revue_2026_10_08"
    dossier.mkdir()
    table = dossier / "Table_revue.xlsx"
    ecrire_table(table, {"Revue": [("Société Dupré", "[SOC_1]", "OUI"), ("Dupré", "[SOC_1]", "OUI"),
                                   ("Marie Lefebvre", "[PP_1]", "OUI")]})
    chemin_config = dossier / "config.json"
    chemin_config.write_text(json.dumps({
        "tables": [{"nom": "Table revue", "fichier": str(table)}],
        "dossier_journaux": "Journaux"}, ensure_ascii=False), encoding="utf-8")
    config = coeur.charger_config(chemin_config)
    entrees = coeur.charger_entrees(table, "Revue")

    print("\n22. Sauts de ligne (Maj+Entrée) et tabulations dans un paragraphe (8/10/2026)")
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Société Dupré").add_break()
    p.add_run("12 rue des Lilas").add_break()
    p.add_run("59000 Lille")
    p2 = doc.add_paragraph()
    p2.add_run("Société")
    p2.add_run().add_tab()
    p2.add_run("Dupré et Marie Lefebvre")
    doc.core_properties.title = "Conclusions Dupré"
    src = dossier / "Bloc adresse.docx"
    doc.save(src)
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>Microsoft Office Word</Application><TitlesOfParts><vt:vector size="1" '
           'baseType="lpstr"><vt:lpstr>Conclusions Dupré</vt:lpstr></vt:vector></TitlesOfParts></Properties>')
    coeur._reecrire_archive(str(src), {"docProps/app.xml": app.encode("utf-8")})
    original = texte_complet(src)
    verifier("le texte lu garde le saut de ligne et la tabulation",
             "Dupré\n12 rue" in original and "Société\tDupré" in original, repr(original))
    r = coeur.anonymiser_document(src, entrees, config, "Table revue", "Revue")
    t = texte_complet(r["chemin_sortie"])
    verifier("le nom en fin de ligne (avant Maj+Entrée) est remplacé",
             "[SOC_1]\n12 rue des Lilas\n59000 Lille" in t, repr(t))
    verifier("le nom séparé par une tabulation est remplacé",
             "[SOC_1] et [PP_1]" in t and "Dupré" not in t, repr(t))
    ctl = coeur.controler(r["chemin_sortie"], entrees, config)
    verifier("le contrôle ne trouve rien en clair (app.xml TitlesOfParts compris)",
             not ctl["termes_en_clair"], str(ctl["termes_en_clair"]))
    with zipfile.ZipFile(r["chemin_sortie"]) as zf:
        app_ano = zf.read("docProps/app.xml").decode("utf-8")
    verifier("le titre recopié dans app.xml (TitlesOfParts) est anonymisé",
             "Dupré" not in app_ano and "[SOC_1]" in app_ano, app_ano)
    rr = coeur.desanonymiser_avec_journal(r["chemin_sortie"], r["chemin_journal"])
    verifier("la restitution redonne exactement le texte d'origine",
             texte_complet(rr["chemin_sortie"]) == original, repr(texte_complet(rr["chemin_sortie"])))
    with zipfile.ZipFile(rr["chemin_sortie"]) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
        app_res = zf.read("docProps/app.xml").decode("utf-8")
    verifier("la tabulation restituée redevient un vrai w:tab (pas un caractère dans le texte)",
             "<w:tab/>" in xml and "\t" not in xml)
    verifier("le titre d'app.xml est restitué", "<vt:lpstr>Conclusions Dupré</vt:lpstr>" in app_res, app_res)

    print("\n23. Rapports de contrôle et sorties (8/10/2026)")
    coeur.ecrire_rapport(r["chemin_sortie"], ctl, "Table revue", "Revue")
    coeur.ecrire_rapport(rr["chemin_sortie"], coeur.controler(rr["chemin_sortie"], entrees, config),
                         "Table revue", "Revue")
    prive = pathlib.Path(coeur.dossier_prive(str(src)))
    rapports = sorted(f.name for f in prive.glob("_Contr*"))
    verifier("le rapport du restitué n'écrase pas celui de l'anonymisé",
             len(rapports) == 2, str(rapports))
    piege = dossier / "Piege_RESTITUE.docx"
    shutil.copy(r["chemin_sortie"], piege)
    try:
        coeur.desanonymiser_sans_journal(piege, entrees, confirme=True, config=config)
        verifier("restituer un fichier nommé _RESTITUE est refusé (il s'écraserait lui-même)", False)
    except coeur.ErreurAnonymisation as e:
        verifier("restituer un fichier nommé _RESTITUE est refusé (il s'écraserait lui-même)",
                 "meme nom" in str(e) or "same name" in str(e), str(e))
    verifier("le fichier soumis est intact", piege.is_file() and "[SOC_1]" in texte_complet(piege))


def recette_lot3_2026_10_08(racine):
    """Revue du 8 octobre 2026, lot 3 : couverture sans journal mesuree sur
    tout le document, suffixes en fin de jeton, comptage des occurrences
    insensible a la casse et aux accents, balise portee par deux termes,
    echec de la recherche hors table signale, journal lu une seule fois."""
    dossier = racine / "lot3_2026_10_08"
    dossier.mkdir()
    table = dossier / "Table_lot3.xlsx"
    ecrire_table(table, {
        "Dossier A": [("Aldebaran Systemes / ALDEBARAN SYSTEMES / Aldébaran Systèmes", "[SOC_DEM_1]", "OUI"),
                      ("Pierre Vaillantin", "[PP_DEM_1]", "OUI"),
                      ("Quentin Morsange", "[PP_DEF_1]", "OUI")],
        "Dossier B": [("Vega Industries", "[SOC_DEM_1]", "OUI"),
                      ("Berthe Nolarge", "[PP_DEM_1]", "OUI")],
        "Doublon": [("Alpha SAS", "[SOC_1]", "OUI"), ("Beta SARL", "[SOC_1]", "OUI"),
                    ("Gamma", "[SOC_2]", "OUI")],
    })
    chemin_config = dossier / "config.json"
    chemin_config.write_text(json.dumps({
        "tables": [{"nom": "Table lot 3", "fichier": str(table)}],
        "dossier_journaux": "Journaux"}, ensure_ascii=False), encoding="utf-8")
    config = coeur.charger_config(chemin_config)
    entrees_a = coeur.charger_entrees(table, "Dossier A")

    print("\n24. Couverture sans journal : tout le document, marqueurs écartés (8/10/2026)")
    # Le corps ne porte que des balises du dossier A ; la note de bas de page
    # porte une balise que seul le dossier A couvre aussi, et un marqueur.
    piege = dossier / "Piege_ANONYMISE.docx"
    docx_zones_invisibles(piege, "[SOC_DEM_1] et [PP_DEM_1]", "[PP_DEF_1] le [DATE]",
                          "[PP_DEM_1]", "[SOC_DEM_1]", "[PP_DEM_1]")
    prep = coeur.preparer_restitution_sans_journal(piege, config["tables"], config)
    verifier("les balises des notes et zones de texte comptent dans la couverture",
             "[PP_DEF_1]" in prep["balises"], str(prep["balises"]))
    verifier("un marqueur de rédaction ([DATE]) ne compte pas comme balise à couvrir",
             "[DATE]" not in prep["balises"], str(prep["balises"]))
    verifier("seul l'onglet qui couvre aussi la note est candidat",
             [a["onglet"] for a in prep["candidats"]] == ["Dossier A"],
             str([(a["onglet"], sorted(a["manquantes"])) for a in prep["candidats"] + prep["autres"]]))
    verifier("le verdict est « candidat unique »", prep["verdict"] == "candidat_unique", prep["verdict"])
    r = coeur.desanonymiser_sans_journal(piege, entrees_a, confirme=True, config=config)
    verifier("et la restitution avec cet onglet est acceptée", os.path.isfile(r["chemin_sortie"]))

    print("\n25. Le suffixe _ANONYMISE n'est reconnu qu'en fin de jeton (8/10/2026)")
    verifier("« Lettre_anonymisee.docx » est un nom ordinaire",
             coeur._base_propre("Lettre_anonymisee.docx") == "Lettre_anonymisee"
             and not coeur.MOTIF_ANONYMISE.search("Lettre_anonymisee.docx"))
    verifier("« Note_restituee » aussi", not coeur.MOTIF_RESTITUE.search("Note_restituee.docx"))
    verifier("« Lettre_ANONYMISE.docx » et « Lettre_ANONYMISE (1).docx » sont bien reconnus",
             coeur._base_propre("Lettre_ANONYMISE.docx") == "Lettre"
             and coeur.MOTIF_ANONYMISE.search("Lettre_ANONYMISE (1).docx") is not None)
    verifier("« Lettre_ANONYMISED_RESTORED.docx » (version anglaise) aussi",
             coeur._base_propre("Lettre_ANONYMISED_RESTORED.docx") == "Lettre")
    ordinaire = dossier / "Lettre_anonymisee.docx"
    ecrire_document(ordinaire, ["Courrier de Pierre Vaillantin à Aldebaran Systemes."],
                    avec_tableau=False, avec_entete=False)
    try:
        ra = coeur.anonymiser_document(ordinaire, entrees_a, config, "Table lot 3", "Dossier A")
        verifier("un fichier nommé « _anonymisee » peut être anonymisé (pas pris pour déjà fait)",
                 os.path.isfile(ra["chemin_sortie"]), ra["chemin_sortie"])
    except coeur.ErreurAnonymisation as e:
        verifier("un fichier nommé « _anonymisee » peut être anonymisé (pas pris pour déjà fait)",
                 False, str(e))

    print("\n26. Occurrences comptées une fois par terme, quelles que soient casse et accents (8/10/2026)")
    src = dossier / "Casse.docx"
    ecrire_document(src, ["Aldebaran Systemes, puis ALDEBARAN SYSTEMES, puis Aldébaran Systèmes."],
                    avec_tableau=False, avec_entete=False)
    trouvailles, _ = coeur.controler_document(src, entrees_a)
    corps = [t for t in trouvailles if t["balise"] == "[SOC_DEM_1]"]
    verifier("trois variantes qui ne diffèrent que par casse et accents : une seule trouvaille",
             len(corps) == 1, str([(t["terme"], t["occurrences"]) for t in corps]))
    verifier("qui compte les trois occurrences",
             corps and corps[0]["occurrences"] == 3, str([(t["terme"], t["occurrences"]) for t in corps]))

    print("\n27. Une même balise portée par deux termes : avertissement non bloquant (8/10/2026)")
    entrees_d = coeur.charger_entrees(table, "Doublon")
    avert = coeur.avertissements_entrees(entrees_d)
    verifier("l'avertissement nomme la balise et les deux termes",
             any("[SOC_1]" in a and "Alpha SAS" in a and "Beta SARL" in a for a in avert), str(avert))
    verifier("une balise à un seul terme ne déclenche rien",
             not any("[SOC_2]" in a for a in avert), str(avert))
    doublon = dossier / "Doublon.docx"
    ecrire_document(doublon, ["Alpha SAS contre Beta SARL et Gamma."], avec_tableau=False, avec_entete=False)
    rd = coeur.anonymiser_document(doublon, entrees_d, config, "Table lot 3", "Doublon")
    verifier("l'anonymisation passe quand même (non bloquant) et remonte l'avertissement",
             os.path.isfile(rd["chemin_sortie"]) and any("[SOC_1]" in a for a in rd["avertissements"]),
             str(rd["avertissements"]))

    print("\n28. Recherche hors table en échec : dit dans le compte rendu (8/10/2026)")
    sauve = coeur.donnees_hors_table

    def en_panne(*a, **k):
        raise RuntimeError("panne simulée")
    coeur.donnees_hors_table = en_panne
    try:
        ctl = coeur.controler(ra["chemin_sortie"], entrees_a, config)
    finally:
        coeur.donnees_hors_table = sauve
    verifier("l'échec est consigné", "panne simulée" in ctl.get("hors_table_erreur", ""), str(ctl.get("hors_table_erreur")))
    verifier("le conseil ne prétend plus « ni IBAN, ni courriel » et renvoie à une relecture",
             "panne simulée" in ctl["conseil"] and "ni IBAN" not in ctl["conseil"], ctl["conseil"])
    verifier("le verdict reste non bloquant (pas d'alerte sur une table propre)", not ctl["alerte"])
    ctl_ok = coeur.controler(ra["chemin_sortie"], entrees_a, config)
    verifier("sans panne, aucune erreur consignée", ctl_ok.get("hors_table_erreur") == "")

    print("\n29. Journal lu une seule fois à la restitution (8/10/2026)")
    compteur = {"n": 0}
    charger = coeur._charger_journal

    def compter(chemin):
        compteur["n"] += 1
        return charger(chemin)
    coeur._charger_journal = compter
    try:
        rj = coeur.desanonymiser_avec_journal(ra["chemin_sortie"], ra["chemin_journal"])
    finally:
        coeur._charger_journal = charger
    verifier("une seule lecture du fichier journal", compteur["n"] == 1, str(compteur["n"]))
    verifier("et le texte revient à l'identique",
             texte_complet(rj["chemin_sortie"]) == texte_complet(ordinaire))
    meta, rempl = coeur.lire_journal(ra["chemin_journal"])
    verifier("lire_journal accepte toujours un chemin (compatibilité)",
             meta.get("onglet") == "Dossier A" and len(rempl) > 0)
    verifier("l'en-tête du journal ne porte plus d'empreinte de table (jamais relue)",
             "empreinte_table" not in meta, str(sorted(meta)))


def main():
    racine = pathlib.Path(tempfile.mkdtemp(prefix="recette_anonymisation_"))
    print(f"\nDossier de recette (fictif, supprime a la fin) :\n  {racine}\n")

    try:
        # --- tables fictives ---
        table_a = racine / "Table_A.xlsx"
        table_b = racine / "Table_B.xlsx"
        ecrire_table(table_a, {"Dossier fictif 1": DOSSIER_1, "Dossier fictif 3": DOSSIER_3})
        ecrire_table(table_b, {"Dossier fictif 2": DOSSIER_2})

        config_path = racine / "config.json"
        config_path.write_text(json.dumps({
            "tables": [{"nom": "Table A", "fichier": str(table_a)},
                       {"nom": "Table B", "fichier": str(table_b)}],
            "dossier_journaux": "Journaux",
        }, ensure_ascii=False, indent=2), encoding="utf-8")

        config = coeur.charger_config(config_path)

        print("1. Configuration et lecture des tables")
        verifier("2 tables chargees", len(config["tables"]) == 2)
        verifier("dossier des journaux resolu depuis le dossier du programme",
                 config["dossier_journaux"] == str(racine / "Journaux"),
                 config["dossier_journaux"])
        onglets_a = coeur.lister_onglets(table_a, config)
        verifier("onglets techniques masques (Mode d'emploi, _DataSrc)",
                 onglets_a == ["Dossier fictif 1", "Dossier fictif 3"], str(onglets_a))

        entrees_1 = coeur.charger_entrees(table_a, "Dossier fictif 1")
        verifier("ligne inactive (colonne Actif = NON) ecartee",
                 "[SOC_DEF_9]" not in coeur.balises_de(entrees_1))
        verifier("SIRET stocke en nombre relu comme texte",
                 any("81234567800019" in v for _, vs in entrees_1 for v in vs))
        verifier("variantes separees par / prises en compte",
                 any(len(vs) == 2 for _, vs in entrees_1))

        print("\n2. Aller-retour anonymisation puis restitution avec journal")
        source = racine / "Piece fictive.docx"
        lignes = [
            "La societe Aldebaran Systemes, immatriculee sous le SIRET 81234567800019,",
            "represente par Pierre Vaillantin, assigne Quentin Morsange.",
            "Aldebaran a produit ses pieces. ALDEBARAN SYSTEMES conteste.",
            "Paragraphe sans aucun terme sensible.",
        ]
        ecrire_document(source, lignes)
        texte_origine = texte_complet(source)

        res = coeur.anonymiser_document(source, entrees_1, config, "Table A", "Dossier fictif 1")
        anonymise = pathlib.Path(res["chemin_sortie"])
        verifier("document anonymise cree", anonymise.is_file())
        verifier("journal cree", os.path.isfile(res["chemin_journal"]))
        verifier("nom du journal porte l'empreinte de l'onglet (jamais son nom)",
                 coeur._empreinte_onglet("Dossier_fictif_1") in os.path.basename(res["chemin_journal"])
                 and "Dossier_fictif_1" not in os.path.basename(res["chemin_journal"]),
                 os.path.basename(res["chemin_journal"]))

        texte_ano = texte_complet(anonymise)
        for terme in ("Aldebaran", "Pierre Vaillantin", "Quentin Morsange", "81234567800019"):
            verifier(f"terme « {terme} » absent du document anonymise",
                     terme.lower() not in texte_ano.lower())
        verifier("balises presentes dans le corps, le tableau et l'en-tete",
                 texte_ano.count("[SOC_DEM_1]") >= 3 and "[SIRET_1]" in texte_ano,
                 texte_ano)

        res_r = coeur.desanonymiser_avec_journal(anonymise, res["chemin_journal"])
        restitue = pathlib.Path(res_r["chemin_sortie"])
        verifier("restitution avec journal identique a l'original",
                 texte_complet(restitue) == texte_origine,
                 "differences detectees entre l'original et le document restitue")
        verifier("en-tete du journal renseigne (table et onglet)",
                 res_r["meta"].get("table") == "Table A"
                 and res_r["meta"].get("onglet") == "Dossier fictif 1")

        print("\n3. Garde-fou anti double anonymisation")
        attend_refus("refus d'anonymiser un document deja anonymise",
                     lambda: coeur.anonymiser_document(anonymise, entrees_1, config,
                                                       "Table A", "Dossier fictif 1"),
                     "deja anonymise")

        print("\n4. Restitution sans journal : verrous")
        entrees_3 = coeur.charger_entrees(table_a, "Dossier fictif 3")
        attend_refus("refus sans confirmation explicite",
                     lambda: coeur.desanonymiser_sans_journal(anonymise, entrees_1, confirme=False),
                     "non confirmee")
        attend_refus("refus si la table ne couvre pas toutes les balises",
                     lambda: coeur.desanonymiser_sans_journal(anonymise, entrees_3, confirme=True),
                     "ne couvre pas")

        prep = coeur.preparer_restitution_sans_journal(anonymise, config["tables"], config)
        verifier("ambiguite detectee entre deux tables concurrentes",
                 prep["verdict"] == "ambigu", f"verdict = {prep['verdict']}")
        noms_candidats = sorted(f"{a['table']}/{a['onglet']}" for a in prep["candidats"])
        verifier("les deux onglets couvrants sont candidats",
                 noms_candidats == ["Table A/Dossier fictif 1", "Table B/Dossier fictif 2"],
                 str(noms_candidats))
        verifier("l'onglet non couvrant est ecarte des candidats",
                 all(a["onglet"] != "Dossier fictif 3" for a in prep["candidats"]))
        verifier("divergences de termes exposees pour arbitrage",
                 len(prep["divergences"]) >= 3, f"{len(prep['divergences'])} divergence(s)")

        copie = racine / "Piece fictive_ANONYMISE_copie.docx"
        shutil.copy(anonymise, copie)
        res_s = coeur.desanonymiser_sans_journal(copie, entrees_1, confirme=True)
        verifier("restitution sans journal autorisee apres confirmation",
                 os.path.isfile(res_s["chemin_sortie"]) and res_s["total"] > 0)

        print("\n5. Confusion de dossiers : le scenario a empecher")
        doc_dossier_2 = racine / "Piece dossier 2.docx"
        ecrire_document(doc_dossier_2,
                        ["Vega Industries, SIRET 49876543200025, representee par Berthe Nolarge,",
                         "assigne Come Vanterre."],
                        avec_tableau=False, avec_entete=False)
        entrees_2 = coeur.charger_entrees(table_b, "Dossier fictif 2")
        res2 = coeur.anonymiser_document(doc_dossier_2, entrees_2, config,
                                         "Table B", "Dossier fictif 2")
        ano2 = pathlib.Path(res2["chemin_sortie"])

        # Restituer le document du dossier 2 avec la table du dossier 1 :
        # techniquement possible (memes balises), juridiquement une fuite.
        # L'outil doit l'autoriser seulement apres confirmation, et l'ecran de
        # confirmation doit signaler l'ambiguite et les divergences.
        prep2 = coeur.preparer_restitution_sans_journal(ano2, config["tables"], config)
        verifier("document d'un autre dossier : ambiguite signalee, pas de choix automatique",
                 prep2["verdict"] == "ambigu" and len(prep2["divergences"]) >= 3,
                 f"verdict = {prep2['verdict']}, divergences = {len(prep2['divergences'])}")

        # Avec journal, en revanche, aucune confusion possible : le journal fait foi.
        res2_r = coeur.desanonymiser_avec_journal(ano2, res2["chemin_journal"])
        verifier("avec journal, restitution du bon dossier",
                 "Vega Industries" in texte_complet(res2_r["chemin_sortie"])
                 and "Aldebaran" not in texte_complet(res2_r["chemin_sortie"]))

        print("\n6. Journaux : homonymes et mauvais journal")
        sous = racine / "autre_dossier"
        sous.mkdir()
        homonyme = sous / "Piece fictive.docx"
        ecrire_document(homonyme, ["Vega Industries et Berthe Nolarge."],
                        avec_tableau=False, avec_entete=False)
        res_h = coeur.anonymiser_document(homonyme, entrees_2, config, "Table B", "Dossier fictif 2")
        verifier("deux documents homonymes n'ecrasent pas leur journal",
                 os.path.isfile(res["chemin_journal"])
                 and os.path.isfile(res_h["chemin_journal"])
                 and res["chemin_journal"] != res_h["chemin_journal"])
        candidats_j = coeur.journaux_candidats(config, homonyme)
        verifier("les journaux homonymes sont tous proposes au choix",
                 len(candidats_j) >= 2, f"{len(candidats_j)} journal/journaux trouve(s)")

        attend_refus("refus d'un journal etranger, malgre des balises identiques",
                     lambda: coeur.desanonymiser_avec_journal(ano2, res["chemin_journal"]),
                     "ne decrit pas ce document")

        taux, attendues, satisfaites, _ = coeur.mesurer_concordance(
            ano2, coeur.lire_journal(res["chemin_journal"])[1])
        verifier("concordance mesuree et faible sur un journal etranger",
                 taux < 0.5, f"taux = {taux:.0%} ({satisfaites}/{attendues})")

        taux_bon, _, _, _ = coeur.mesurer_concordance(
            anonymise, coeur.lire_journal(res["chemin_journal"])[1])
        verifier("concordance totale sur le bon journal", taux_bon == 1.0, f"taux = {taux_bon:.0%}")

        # Document retouche apres anonymisation : refus par defaut, mais
        # restitution possible si l'utilisateur l'accepte en connaissance de cause.
        retouche = racine / "Piece fictive_ANONYMISE_retouche.docx"
        shutil.copy(anonymise, retouche)
        d = Document(retouche)
        d.paragraphs[1].runs[0].text = "Paragraphe reecrit par l'utilisateur."
        d.save(retouche)
        attend_refus("document retouche depuis l'anonymisation : refus par defaut",
                     lambda: coeur.desanonymiser_avec_journal(retouche, res["chemin_journal"]),
                     "ne decrit pas ce document")
        res_tol = coeur.desanonymiser_avec_journal(retouche, res["chemin_journal"],
                                                  tolerer_partiel=True)
        verifier("document retouche : restitution possible apres acceptation",
                 res_tol["total"] > 0 and res_tol["concordance"] < 1.0,
                 f"total = {res_tol['total']}, concordance = {res_tol['concordance']:.0%}")

        print("\n6 bis. Ordre des variantes d'une meme balise")
        multi = racine / "Variantes multiples.docx"
        ecrire_document(multi,
                        ["Aldebaran Systemes puis Aldebaran puis ALDEBARAN SYSTEMES puis aldebaran."],
                        avec_tableau=False, avec_entete=False)
        origine_multi = texte_complet(multi)
        res_m = coeur.anonymiser_document(multi, entrees_1, config, "Table A", "Dossier fictif 1")
        res_m_r = coeur.desanonymiser_avec_journal(res_m["chemin_sortie"], res_m["chemin_journal"])
        verifier("quatre variantes d'une meme balise restituees chacune a sa place",
                 texte_complet(res_m_r["chemin_sortie"]) == origine_multi,
                 f"attendu : {origine_multi!r}\n         obtenu : "
                 f"{texte_complet(res_m_r['chemin_sortie'])!r}")

        print("\n7. Compatibilite avec l'existant")
        ancien = racine / "Journaux" / "Ancien format_JOURNAL.json"
        _, remplacements = coeur.lire_journal(res["chemin_journal"])
        ancien.write_text(json.dumps(remplacements, ensure_ascii=False), encoding="utf-8")
        meta_a, rempl_a = coeur.lire_journal(ancien)
        verifier("journal au format historique (liste simple) toujours lisible",
                 meta_a == {} and len(rempl_a) == len(remplacements))

        copie2 = racine / "Piece fictive_ANONYMISE_copie2.docx"
        shutil.copy(anonymise, copie2)
        res_anc = coeur.desanonymiser_avec_journal(copie2, ancien)
        verifier("restitution possible depuis un journal au format historique",
                 res_anc["total"] > 0)

        config_ancienne = racine / "config_ancienne.json"
        config_ancienne.write_text(json.dumps(
            {"fichier_excel": "Table_A.xlsx", "dossier_json": "Journaux"},
            ensure_ascii=False, indent=2), encoding="utf-8")
        cfg_anc = coeur.charger_config(config_ancienne)
        verifier("config.json au format historique toujours acceptee",
                 len(cfg_anc["tables"]) == 1
                 and cfg_anc["tables"][0]["fichier"] == str(racine / "Table_A.xlsx"),
                 str(cfg_anc["tables"]))

        cfg_manque = racine / "config_manquante.json"
        cfg_manque.write_text(json.dumps(
            {"tables": [{"nom": "Absente", "fichier": str(racine / "nexistepas.xlsx")}]},
            ensure_ascii=False), encoding="utf-8")
        ok_t, pbs = coeur.verifier_tables(coeur.charger_config(cfg_manque))
        verifier("table declaree mais absente : signalee, pas de plantage",
                 ok_t == [] and len(pbs) == 1)

        print("\n8. Controle de l'onglet AVANT anonymisation")
        classement = coeur.analyser_pertinence(source, config["tables"], config)
        tete = classement[0]
        verifier("l'onglet du bon dossier arrive en tete",
                 tete["table"] == "Table A" and tete["onglet"] == "Dossier fictif 1",
                 f"tete = {tete['table']} / {tete['onglet']} ({tete['trouvees']} terme(s))")
        etranger = next(r for r in classement if r["onglet"] == "Dossier fictif 2")
        verifier("un onglet d'un autre dossier ne trouve aucun terme",
                 etranger["trouvees"] == 0, f"{etranger['trouvees']} terme(s) trouve(s)")

        neuf = racine / "Autre piece dossier 1.docx"
        ecrire_document(neuf, ["Aldebaran Systemes et Pierre Vaillantin, encore."],
                        avec_tableau=False, avec_entete=False)
        attend_refus("refus d'anonymiser avec un onglet sans aucun terme present",
                     lambda: coeur.anonymiser_document(neuf, entrees_2, config,
                                                       "Table B", "Dossier fictif 2"),
                     "aucun terme")
        verifier("aucun fichier _ANONYMISE trompeur cree lors de ce refus",
                 not (racine / "Autre piece dossier 1_ANONYMISE.docx").is_file())

        print("\n14. Contrôle avant envoi : les zones que l'anonymisation ne voit pas")
        piege = racine / "Piece piegee.docx"
        docx_piege(piege,
                   nom_corps="Aldebaran Systemes",
                   nom_note="Pierre Vaillantin",
                   nom_zone="Quentin Morsange",
                   nom_commentaire="Aldebaran",
                   nom_proprietes="Pierre Vaillantin",
                   url_client="https://www.aldebaran-systemes-exemple.fr/contrat")

        trouvailles, zones = coeur.controler_document(piege, entrees_1)
        zones_touchees = {t["zone"] for t in trouvailles}

        for libelle, attendu in [
            ("note de bas de page", "NOTES DE BAS DE PAGE"),
            ("zone de texte ou forme", "ZONES DE TEXTE ET FORMES"),
            ("commentaire", "COMMENTAIRES"),
            ("propriétés du fichier", "PROPRIÉTÉS DU DOCUMENT (auteur, titre)"),
            ("corps du document", "corps du document"),
            ("cible d'un lien hypertexte", "LIENS HYPERTEXTE"),
        ]:
            verifier(f"terme en clair détecté dans : {libelle}",
                     attendu in zones_touchees,
                     f"zones où quelque chose a été trouvé : {sorted(zones_touchees)}")

        verifier("le contrôle nomme le terme et la balise concernés",
                 all(t["terme"] and t["balise"] for t in trouvailles))
        verifier("le contrôle fournit un extrait pour situer chaque trouvaille",
                 all(len(t["extrait"]) > 5 for t in trouvailles))
        verifier("les zones invisibles sont classées avant le corps du document",
                 trouvailles[0]["zone"] != "corps du document",
                 f"première trouvaille : {trouvailles[0]['zone']}")

        print("\n15. Contrôle sur un document réellement anonymisé")
        trouvailles_ok, zones_ok = coeur.controler_document(anonymise, entrees_1)
        verifier("aucune alerte sur un document correctement anonymisé",
                 trouvailles_ok == [],
                 str([(t["zone"], t["terme"]) for t in trouvailles_ok]))
        trouvailles_src, _ = coeur.controler_document(source, entrees_1)
        verifier("le document d'origine, lui, déclenche bien l'alerte",
                 len(trouvailles_src) > 0,
                 "un document non anonymisé doit être signalé")

        verifier("un mot ordinaire contenant un terme n'est pas signalé",
                 all(t["terme"].lower() != "aldebaranesque" for t in trouvailles_src))

        print("\n16. Contrôle dans les deux sens")

        # --- Etat du document ---
        etat_a, motif_a = coeur.etat_document(anonymise, config)
        verifier("un document _ANONYMISE est reconnu comme tel",
                 etat_a == coeur.ETAT_ANONYMISE, f"{etat_a} ({motif_a})")
        restitue_ok = pathlib.Path(res_r["chemin_sortie"])
        etat_r, motif_r = coeur.etat_document(restitue_ok, config)
        verifier("un document _RESTITUE est reconnu comme tel",
                 etat_r == coeur.ETAT_RESTITUE, f"{etat_r} ({motif_r})")

        neutre = racine / "Note interne sans lien.docx"
        ecrire_document(neutre, ["Note de travail sans aucun nom ni balise."],
                        avec_tableau=False, avec_entete=False)
        etat_n, motif_n = coeur.etat_document(neutre, config)
        verifier("un document sans balise ni journal reste indéterminé",
                 etat_n == coeur.ETAT_INDETERMINE, f"{etat_n} ({motif_n})")

        # --- Sens 1 : document anonymise ---
        r_anon = coeur.controler(anonymise, entrees_1, config)
        verifier("anonymisé propre : aucune alerte",
                 not r_anon["alerte"],
                 str([(t["zone"], t["terme"]) for t in r_anon["termes_en_clair"]]))
        verifier("anonymisé : les balises présentes sont relevées sans alarmer",
                 len(r_anon["balises_restantes"]) > 0 and not r_anon["alerte"])
        verifier("anonymisé propre : le conseil autorise le dépôt dans une IA",
                 "peut être déposé" in r_anon["conseil"], r_anon["conseil"][:60])

        r_source = coeur.controler(source, entrees_1, config)
        verifier("document jamais traité : alerte sur les termes en clair",
                 r_source["alerte"] and len(r_source["termes_en_clair"]) > 0,
                 f"état={r_source['etat']}, alerte={r_source['alerte']}")
        verifier("un original n'est JAMAIS pris pour un restitué",
                 r_source["etat"] != coeur.ETAT_RESTITUE,
                 "le classer restitué le déclarerait bon pour l'IA")
        verifier("en cas de doute, le conseil interdit le dépôt dans une IA",
                 "aucun outil d'ia" in r_source["conseil"].lower(),
                 r_source["conseil"][:90])

        # --- Sens 2 : document restitue ---
        r_rest = coeur.controler(restitue_ok, entrees_1, config)
        verifier("restitué complet : aucune alerte malgré les vrais noms",
                 not r_rest["alerte"],
                 f"balises restantes : {[t['balise'] for t in r_rest['balises_restantes']]}")
        verifier("restitué : les vrais noms sont attendus, pas signalés comme faute",
                 len(r_rest["termes_en_clair"]) > 0 and not r_rest["alerte"])
        verifier("restitué : le conseil interdit explicitement le dépôt dans une IA",
                 "AUCUN outil d'IA" in r_rest["conseil"], r_rest["conseil"][:80])

        # Restitution incomplete : une balise a survecu
        rest_incomplet = racine / "Piece bancale_RESTITUE.docx"
        shutil.copy(restitue_ok, rest_incomplet)
        d = Document(rest_incomplet)
        d.add_paragraph("Reste une balise oubliée : [SOC_DEM_1] doit disparaître.")
        d.save(rest_incomplet)

        r_bancal = coeur.controler(rest_incomplet, entrees_1, config)
        verifier("restitué incomplet : ALERTE sur la balise oubliée",
                 r_bancal["alerte"] and len(r_bancal["balises_restantes"]) > 0,
                 f"alerte={r_bancal['alerte']}, "
                 f"balises={[t['balise'] for t in r_bancal['balises_restantes']]}")
        verifier("restitué incomplet : le titre nomme la restitution, pas l'IA",
                 "RESTITUTION INCOMPLÈTE" in r_bancal["titre"], r_bancal["titre"])
        verifier("restitué incomplet : le conseil dit de ne pas envoyer",
                 "Ne l'envoyez pas" in r_bancal["conseil"], r_bancal["conseil"][:70])

        # --- Balises cherchees partout, pas seulement dans le corps ---
        piege_balise = racine / "Restitution piegee_RESTITUE.docx"
        docx_piege(piege_balise,
                   nom_corps="texte neutre",
                   nom_note="[SOC_DEM_1]",
                   nom_zone="texte neutre",
                   nom_commentaire="texte neutre",
                   nom_proprietes="texte neutre",
                   url_client="https://exemple.invalid/")
        r_piege = coeur.controler(piege_balise, entrees_1, config)
        zones_balises = {t["zone"] for t in r_piege["balises_restantes"]}
        verifier("une balise oubliée en note de bas de page est trouvée",
                 "NOTES DE BAS DE PAGE" in zones_balises, str(sorted(zones_balises)))

        # --- Rapport ---
        rapport2 = coeur.ecrire_rapport(rest_incomplet, r_bancal,
                                        "Table A", "Dossier fictif 1")
        texte_rapport = open(rapport2, encoding="utf-8").read()
        verifier("le rapport indique l'état du document",
                 "RESTITUÉ" in texte_rapport)
        verifier("le rapport liste les balises restantes",
                 "BALISES PRÉSENTES" in texte_rapport)
        verifier("le titre du rapport parle de transmission ET de dépôt dans une IA",
                 "DÉPÔT DANS UN OUTIL D'IA" in texte_rapport)

        print("\n17. Le nom du fichier et du dossier")
        dossier_nomme = racine / "Aldebaran Systemes"
        dossier_nomme.mkdir(exist_ok=True)
        piece_nommee = dossier_nomme / "Conclusions_Aldebaran_v3.docx"
        ecrire_document(piece_nommee, ["Texte neutre."], avec_tableau=False, avec_entete=False)

        tr_chemin = coeur.controler_chemin(str(piece_nommee), entrees_1)
        zones_chemin = {t["zone"] for t in tr_chemin}
        verifier("un terme dans le nom du fichier est détecté",
                 "nom du fichier" in zones_chemin, str(sorted(zones_chemin)))
        verifier("un terme dans le nom d'un dossier est détecté",
                 "nom d'un dossier" in zones_chemin)
        verifier("le tiret bas est bien traité comme un séparateur",
                 any(t["terme"].lower() == "aldebaran" and t["zone"] == "nom du fichier"
                     for t in tr_chemin))

        neutre2 = racine / "Aldebaranesque et autres mots.docx"
        ecrire_document(neutre2, ["Texte neutre."], avec_tableau=False, avec_entete=False)
        verifier("un mot plus long contenant le terme n'est pas signalé",
                 not any(t["zone"] == "nom du fichier"
                         for t in coeur.controler_chemin(str(neutre2), entrees_1)))

        r_nom = coeur.controler(piece_nommee, entrees_1, config)
        verifier("le contrôle intègre le nom du fichier dans ses alertes",
                 r_nom["alerte"] and any(t["zone"].startswith("nom ")
                                         for t in r_nom["termes_en_clair"]))
        verifier("le nom du fichier est classé avant les zones internes",
                 r_nom["termes_en_clair"][0]["zone"].startswith("nom "),
                 r_nom["termes_en_clair"][0]["zone"])

        print("\n18. Anonymisation des zones invisibles")
        piege2 = racine / "Piece avec zones.docx"
        docx_zones_invisibles(
            piege2,
            texte_corps="Le present acte concerne Aldebaran Systemes.",
            texte_note="Note de bas de page citant Pierre Vaillantin.",
            texte_note_fin="Note de fin citant Quentin Morsange.",
            texte_commentaire="Commentaire citant Aldebaran.",
            texte_zone="Encadre citant Pierre Vaillantin et Aldebaran Systemes.")
        avant = textes_zones_invisibles(piege2)
        verifier("le document de test porte bien des termes dans chaque zone",
                 len(avant) == 4 and all("Aldebaran" in v or "Vaillantin" in v
                                         or "Morsange" in v for v in avant.values()),
                 str(list(avant)))

        res_z = coeur.anonymiser_document(piege2, entrees_1, config,
                                          "Table A", "Dossier fictif 1")
        anon_z = pathlib.Path(res_z["chemin_sortie"])
        verifier("des remplacements ont eu lieu dans les zones invisibles",
                 res_z["total_zones_invisibles"] > 0,
                 f"{res_z['total_zones_invisibles']} remplacement(s)")

        apres = textes_zones_invisibles(anon_z)
        for cle, libelle in [("word/footnotes.xml", "note de bas de page"),
                             ("word/endnotes.xml", "note de fin"),
                             ("word/comments.xml", "commentaire"),
                             ("zone de texte", "zone de texte")]:
            contenu = apres.get(cle, "")
            verifier(f"plus aucun nom en clair dans : {libelle}",
                     contenu and not any(n in contenu for n in
                                         ("Aldebaran", "Vaillantin", "Morsange")),
                     f"{cle} = {contenu!r}")
            verifier(f"une balise a bien remplacé le nom dans : {libelle}",
                     "[" in contenu and "]" in contenu, f"{cle} = {contenu!r}")

        verifier("le document produit reste un .docx valide",
                 len(Document(anon_z).paragraphs) > 0)

        tr_z, _ = coeur.controler_document(anon_z, entrees_1)
        verifier("le contrôle ne trouve plus rien après anonymisation complète",
                 tr_z == [], str([(t["zone"], t["terme"]) for t in tr_z]))

        print("\n19. Restitution des zones invisibles")
        xml_journal = coeur.lire_journal_xml(res_z["chemin_journal"])
        verifier("le journal enregistre les remplacements des zones invisibles",
                 len(xml_journal) > 0 and all("partie" in e and "para_index" in e
                                              for e in xml_journal))
        res_zr = coeur.desanonymiser_avec_journal(anon_z, res_z["chemin_journal"])
        rendu = textes_zones_invisibles(res_zr["chemin_sortie"])
        for cle in avant:
            verifier(f"restitution exacte de : {cle}",
                     rendu.get(cle) == avant[cle],
                     f"attendu {avant[cle]!r}\n         obtenu {rendu.get(cle)!r}")

        print("\n20. Cas limites des zones invisibles")
        # Un document dont SEULES les notes portent des termes ne doit pas etre refuse.
        note_seule = racine / "Termes en note seulement.docx"
        docx_zones_invisibles(
            note_seule,
            texte_corps="Un texte parfaitement neutre, sans aucun nom.",
            texte_note="Mais la note cite Pierre Vaillantin.",
            texte_note_fin="Neutre.", texte_commentaire="Neutre.", texte_zone="Neutre.")
        res_ns = coeur.anonymiser_document(note_seule, entrees_1, config,
                                           "Table A", "Dossier fictif 1")
        verifier("un document dont seules les notes portent des termes est traité",
                 res_ns["total"] > 0 and res_ns["total_zones_invisibles"] > 0)
        verifier("le journal d'un tel document reste exploitable",
                 coeur.desanonymiser_avec_journal(
                     res_ns["chemin_sortie"], res_ns["chemin_journal"])["total"] > 0)

        # Restitution sans journal : les zones invisibles doivent l'etre aussi.
        copie_z = racine / "Piece avec zones_ANONYMISE_copie.docx"
        shutil.copy(anon_z, copie_z)
        res_zs = coeur.desanonymiser_sans_journal(copie_z, entrees_1, confirme=True)
        rendu_s = textes_zones_invisibles(res_zs["chemin_sortie"])
        verifier("la restitution sans journal traite aussi les zones invisibles",
                 res_zs.get("total_zones_invisibles", 0) > 0
                 and "Vaillantin" in rendu_s.get("word/footnotes.xml", ""),
                 str(rendu_s.get("word/footnotes.xml")))

        # Une balise cachee en note doit bloquer une restitution non couvrante.
        attend_refus("une balise en note de bas de page compte dans la couverture",
                     lambda: coeur.desanonymiser_sans_journal(anon_z, entrees_3,
                                                              confirme=True),
                     "ne couvre pas")

        print("\n21. Préservation de la mise en forme")
        mef = racine / "Piece mise en forme.docx"
        ecrire_document_mise_en_forme(mef, [[
            ("Le contrat conclu avec ", False, False, False),
            ("Aldebaran Systemes", True, False, False),
            (" a ete resilie, comme ", False, False, False),
            ("Pierre Vaillantin", False, True, False),
            (" l'a confirme.", False, False, False),
        ]])
        origine_mef = lire_fragments(mef)
        verifier("le document de test a bien cinq fragments distincts",
                 len(origine_mef) == 5, str([f[0] for f in origine_mef]))

        res_m = coeur.anonymiser_document(mef, entrees_1, config,
                                          "Table A", "Dossier fictif 1")
        apres_mef = lire_fragments(res_m["chemin_sortie"])

        verifier("la balise hérite du gras du terme remplacé",
                 mise_en_forme_de(apres_mef, "[SOC_DEM_1]") == (True, False, False),
                 str(apres_mef))
        verifier("la balise hérite de l'italique du terme remplacé",
                 mise_en_forme_de(apres_mef, "[PP_DEM_1]") == (False, True, False),
                 str(apres_mef))
        verifier("le texte alentour garde sa mise en forme d'origine",
                 mise_en_forme_de(apres_mef, "Le contrat conclu avec ") == (False, False, False)
                 and mise_en_forme_de(apres_mef, " l'a confirme.") == (False, False, False),
                 str(apres_mef))
        verifier("les fragments ne sont pas fusionnés en un seul",
                 len(apres_mef) == 5, f"{len(apres_mef)} fragment(s) : {[f[0] for f in apres_mef]}")

        res_mr = coeur.desanonymiser_avec_journal(res_m["chemin_sortie"],
                                                  res_m["chemin_journal"])
        rendu_mef = lire_fragments(res_mr["chemin_sortie"])
        verifier("après restitution, texte ET mise en forme sont identiques à l'origine",
                 rendu_mef == origine_mef,
                 f"attendu {origine_mef}\n         obtenu {rendu_mef}")

        print("\n22. Mise en forme, cas difficiles")
        # Terme coupe en deux fragments de mise en forme differente : Word produit
        # cela apres une correction orthographique. Le texte doit rester juste.
        coupe = racine / "Terme coupe.docx"
        ecrire_document_mise_en_forme(coupe, [[
            ("Voir ", False, False, False),
            ("Aldeba", False, False, False),
            ("ran Systemes", True, False, False),
            (" pour la suite.", False, False, False),
        ]])
        res_c = coeur.anonymiser_document(coupe, entrees_1, config,
                                          "Table A", "Dossier fictif 1")
        texte_c = texte_complet(res_c["chemin_sortie"])
        verifier("un terme à cheval sur deux fragments est quand même remplacé",
                 "[SOC_DEM_1]" in texte_c and "Aldeba" not in texte_c, texte_c)
        res_cr = coeur.desanonymiser_avec_journal(res_c["chemin_sortie"],
                                                  res_c["chemin_journal"])
        verifier("et il est restitué avec le bon texte",
                 "Aldebaran Systemes" in texte_complet(res_cr["chemin_sortie"]),
                 texte_complet(res_cr["chemin_sortie"]))

        # Plusieurs occurrences du meme terme, mises en forme differentes.
        multi_mef = racine / "Occurrences multiples.docx"
        ecrire_document_mise_en_forme(multi_mef, [[
            ("Aldebaran", True, False, False),
            (" puis ", False, False, False),
            ("Aldebaran", False, True, False),
            (" puis ", False, False, False),
            ("Aldebaran", False, False, True),
            (".", False, False, False),
        ]])
        origine_multi = lire_fragments(multi_mef)
        res_mm = coeur.anonymiser_document(multi_mef, entrees_1, config,
                                           "Table A", "Dossier fictif 1")
        apres_multi = lire_fragments(res_mm["chemin_sortie"])
        formes = [f[1:] for f in apres_multi if f[0] == "[SOC_DEM_1]"]
        verifier("trois occurrences du même terme gardent chacune leur mise en forme",
                 formes == [(True, False, False), (False, True, False), (False, False, True)],
                 str(apres_multi))
        res_mmr = coeur.desanonymiser_avec_journal(res_mm["chemin_sortie"],
                                                   res_mm["chemin_journal"])
        verifier("restitution identique sur les occurrences multiples",
                 lire_fragments(res_mmr["chemin_sortie"]) == origine_multi,
                 str(lire_fragments(res_mmr["chemin_sortie"])))

        # Restitution sans journal : la mise en forme doit aussi survivre.
        copie_mef = racine / "Piece mise en forme_ANONYMISE_copie.docx"
        shutil.copy(res_m["chemin_sortie"], copie_mef)
        res_ms = coeur.desanonymiser_sans_journal(copie_mef, entrees_1, confirme=True)
        rendu_sans = lire_fragments(res_ms["chemin_sortie"])
        verifier("la restitution sans journal préserve aussi la mise en forme",
                 mise_en_forme_de(rendu_sans, "Aldebaran Systemes") == (True, False, False),
                 str(rendu_sans))

        print("\n23. Mise en forme dans les zones invisibles")
        zone_mef = racine / "Zones et mise en forme.docx"
        docx_zones_invisibles(
            zone_mef,
            texte_corps="Corps neutre.",
            texte_note="La note cite Pierre Vaillantin ici.",
            texte_note_fin="Neutre.", texte_commentaire="Neutre.", texte_zone="Neutre.")
        res_zm = coeur.anonymiser_document(zone_mef, entrees_1, config,
                                           "Table A", "Dossier fictif 1")
        note = textes_zones_invisibles(res_zm["chemin_sortie"])["word/footnotes.xml"]
        verifier("dans une note, seul le terme est remplacé, le reste du texte demeure",
                 note.startswith("La note cite ") and note.endswith(" ici.")
                 and "[PP_DEM_1]" in note, repr(note))

        print("\n24. Propriétés du fichier et cibles des liens")
        annexe = racine / "Piece avec annexes.docx"
        docx_avec_proprietes_et_lien(
            annexe,
            texte="Le present acte concerne Aldebaran Systemes.",
            auteur="Pierre Vaillantin",
            societe="Aldebaran Systemes",
            titre="Conclusions pour Aldebaran Systemes",
            url="https://www.aldebaran-systemes-exemple.fr/dossier")
        avant_annexes = lire_annexes(annexe)
        verifier("le document de test porte bien auteur, société, titre et lien",
                 avant_annexes["auteur"] == "Pierre Vaillantin"
                 and avant_annexes["societe"] == "Aldebaran Systemes"
                 and any("aldebaran" in (l or "").lower()
                         for l in avant_annexes["liens"]),
                 str(avant_annexes))

        res_a = coeur.anonymiser_document(annexe, entrees_1, config,
                                          "Table A", "Dossier fictif 1")
        apres_annexes = lire_annexes(res_a["chemin_sortie"])

        verifier("l'auteur du fichier est masqué",
                 apres_annexes["auteur"] == "Auteur masqué", str(apres_annexes["auteur"]))
        verifier("la société est masquée",
                 apres_annexes["societe"] == "Société masquée", str(apres_annexes["societe"]))
        verifier("le titre du document est anonymisé terme par terme",
                 apres_annexes["titre"] and "[SOC_DEM_1]" in apres_annexes["titre"]
                 and "Aldebaran" not in apres_annexes["titre"],
                 str(apres_annexes["titre"]))
        verifier("la cible du lien ne porte plus le nom du client",
                 all("aldebaran" not in (l or "").lower() for l in apres_annexes["liens"]),
                 str(apres_annexes["liens"]))
        verifier("la cible neutralisée reste une adresse valide",
                 any((l or "").startswith("https://") for l in apres_annexes["liens"]),
                 str(apres_annexes["liens"]))
        verifier("le compte rendu chiffre les propriétés et les liens traités",
                 res_a["total_proprietes"] >= 3 and res_a["total_liens"] == 1,
                 f"{res_a['total_proprietes']} propriété(s), {res_a['total_liens']} lien(s)")

        tr_a = coeur.controler(res_a["chemin_sortie"], entrees_1, config)
        verifier("le contrôle ne signale plus rien sur ce document",
                 not tr_a["alerte"],
                 str([(t["zone"], t["terme"]) for t in tr_a["termes_en_clair"]]))

        print("\n25. Restitution des propriétés et des liens")
        res_ar = coeur.desanonymiser_avec_journal(res_a["chemin_sortie"],
                                                  res_a["chemin_journal"])
        rendu_annexes = lire_annexes(res_ar["chemin_sortie"])
        verifier("l'auteur d'origine est restitué",
                 rendu_annexes["auteur"] == avant_annexes["auteur"],
                 str(rendu_annexes["auteur"]))
        verifier("la société d'origine est restituée",
                 rendu_annexes["societe"] == avant_annexes["societe"])
        verifier("le titre d'origine est restitué",
                 rendu_annexes["titre"] == avant_annexes["titre"],
                 str(rendu_annexes["titre"]))
        verifier("l'adresse du lien d'origine est restituée",
                 sorted(rendu_annexes["liens"]) == sorted(avant_annexes["liens"]),
                 str(rendu_annexes["liens"]))

        print("\n26. L'option de masquage et la compatibilité")
        config_sans = dict(config)
        config_sans["masquer_proprietes"] = False
        annexe2 = racine / "Piece sans masquage.docx"
        docx_avec_proprietes_et_lien(
            annexe2, texte="Acte citant Aldebaran Systemes.",
            auteur="Pierre Vaillantin", societe="Aldebaran Systemes",
            titre="Dossier", url="https://exemple.invalid/neutre")
        res_sm = coeur.anonymiser_document(annexe2, entrees_1, config_sans,
                                           "Table A", "Dossier fictif 1")
        verifier("l'option masquer_proprietes=false laisse l'auteur intact",
                 lire_annexes(res_sm["chemin_sortie"])["auteur"] == "Pierre Vaillantin",
                 str(lire_annexes(res_sm["chemin_sortie"])["auteur"]))

        verifier("un journal sans propriétés ni liens reste exploitable",
                 coeur.lire_journal_annexes(ancien) == ([], []))
        copie3 = racine / "Piece fictive_ANONYMISE_copie3.docx"
        shutil.copy(anonymise, copie3)
        verifier("et la restitution depuis un tel journal fonctionne toujours",
                 coeur.desanonymiser_avec_journal(copie3, ancien)["total"] > 0)

        print("\n27. Le garde-fou de l'onglet résiste au masquage")
        annexe3 = racine / "Onglet errone.docx"
        docx_avec_proprietes_et_lien(
            annexe3, texte="Texte sans aucun terme de la table B.",
            auteur="Quelqu'un", societe="Une societe", titre="Titre",
            url="https://exemple.invalid/x")
        attend_refus("le masquage des propriétés ne fait pas passer un onglet erroné",
                     lambda: coeur.anonymiser_document(annexe3, entrees_2, config,
                                                       "Table B", "Dossier fictif 2"),
                     "aucun terme")
        verifier("et aucun fichier n'est laissé derrière",
                 not (racine / "Onglet errone_ANONYMISE.docx").is_file())

        print("\n28. Marqueurs de rédaction ([DATE], [MONTANT]...)")
        declarees = coeur.balises_declarees(config, forcer=True)
        verifier("les balises déclarées sont lues dans les onglets",
                 "[SOC_DEM_1]" in declarees and "[PP_DEM_1]" in declarees,
                 f"{len(declarees)} balise(s) déclarée(s)")
        verifier("un marqueur de rédaction n'est pas une balise déclarée",
                 "[DATE]" not in declarees)

        reelles, marqueurs = coeur.trier_balises(
            {"[SOC_DEM_1]", "[DATE]", "[MONTANT]"}, config)
        verifier("le tri sépare balises réelles et marqueurs",
                 reelles == {"[SOC_DEM_1]"} and marqueurs == {"[DATE]", "[MONTANT]"},
                 f"réelles={reelles}, marqueurs={marqueurs}")

        # Le cas rencontre en usage reel : un acte anonymise qui contient aussi
        # un marqueur de redaction laisse par le redacteur.
        avec_marqueur = racine / "Acte avec marqueur.docx"
        ecrire_document(avec_marqueur,
                        ["La societe Aldebaran Systemes, le [DATE], a ecrit a",
                         "Pierre Vaillantin pour un montant de [MONTANT] euros."],
                        avec_tableau=False, avec_entete=False)
        res_mq = coeur.anonymiser_document(avec_marqueur, entrees_1, config,
                                           "Table A", "Dossier fictif 1")
        texte_mq = texte_complet(res_mq["chemin_sortie"])
        verifier("les marqueurs de rédaction traversent l'anonymisation intacts",
                 "[DATE]" in texte_mq and "[MONTANT]" in texte_mq, texte_mq)

        res_mqs = coeur.desanonymiser_sans_journal(
            res_mq["chemin_sortie"], entrees_1, confirme=True, config=config)
        verifier("un marqueur de rédaction ne bloque plus la restitution sans journal",
                 res_mqs["total"] > 0,
                 "c'est le refus rencontré en usage réel sur [DATE]")
        verifier("les marqueurs laissés en place sont listés dans le compte rendu",
                 set(res_mqs["marqueurs_laisses"]) == {"[DATE]", "[MONTANT]"},
                 str(res_mqs["marqueurs_laisses"]))
        texte_rendu = texte_complet(res_mqs["chemin_sortie"])
        verifier("et ils sont toujours là après restitution",
                 "[DATE]" in texte_rendu and "Aldebaran" in texte_rendu, texte_rendu)

        # Une VRAIE balise manquante doit toujours bloquer : le document
        # principal porte [PP_DEF_1], que l'onglet 3 ne declare pas.
        attend_refus("une balise réelle absente de la table bloque toujours",
                     lambda: coeur.desanonymiser_sans_journal(
                         anonymise, entrees_3, confirme=True, config=config),
                     "ne couvre pas")

        # Et le controle ne doit pas crier a la restitution incomplete.
        rest_marqueur = racine / "Copie controlee_RESTITUE.docx"
        shutil.copy(res_mqs["chemin_sortie"], rest_marqueur)
        r_mq = coeur.controler(rest_marqueur, entrees_1, config)
        verifier("un restitué contenant un marqueur n'est pas jugé incomplet",
                 not r_mq["alerte"],
                 f"balises retenues : {[b['balise'] for b in r_mq['balises_restantes']]}")
        verifier("les marqueurs sont signalés à part, pour information",
                 {b["balise"] for b in r_mq["marqueurs_redaction"]} == {"[DATE]", "[MONTANT]"},
                 str([b["balise"] for b in r_mq["marqueurs_redaction"]]))

        recette_audit_2026_10(racine)
        recette_accents_2026_10(racine)
        recette_texte_2026_10(racine)
        recette_hors_table_2026_10(racine)
        recette_auteurs_2026_10(racine)
        recette_revue_2026_10_08(racine)
        recette_lot3_2026_10_08(racine)

    finally:
        print("\n" + "=" * 68)
        print(f"RESULTAT : {len(OK)} controle(s) reussi(s), {len(KO)} echec(s)")
        if KO:
            print("\nEchecs a corriger :")
            for libelle, detail in KO:
                print(f"  - {libelle}" + (f"\n      {detail}" if detail else ""))
        else:
            print("\nTous les controles sont passes.")
        print("=" * 68 + "\n")
        shutil.rmtree(racine, ignore_errors=True)

    return 1 if KO else 0


if __name__ == "__main__":
    sys.exit(main())
