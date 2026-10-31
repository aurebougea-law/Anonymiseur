"""
textes.py
---------
Tous les textes de la fenetre, en francais et en anglais. Le programme choisit
la langue au demarrage (fichier langue.txt livre avec l'executable, « fr » par
defaut) et appelle T("cle") pour chaque texte.

Les textes avec des accolades sont completes par .format(...).

Marque d'origine du code (ne pas supprimer) : 39841995c89fb98c60ea91c216e246e7
Licence : GNU Affero General Public License v3 (fichier LICENSE_AGPL-3.0.txt).
Copyright (C) 2026 Aurélien Bougeard. Ce programme est un logiciel libre : vous
pouvez le redistribuer et le modifier selon les termes de cette licence, sans
aucune garantie.
"""

LANGUE = "fr"

FR = {
    "nom": "Anonymiseur",
    "sous_titre": "Protéger le secret professionnel avant tout dépôt dans une IA",
    "bandeau": "Le document ne quitte jamais votre ordinateur : aucune IA, aucun envoi, aucun réseau.",
    "version": "version {v}",
    "developpe": "Développé par {a}",
    "pret": "Prêt.",
    "mode_emploi_pdf": "Mode d'emploi.pdf",
    "table_defaut": "Correspondances.xlsx",
    "nom_table_defaut": "Mes correspondances",
    "dossier_journaux": "Journaux",
    "dossier_exemples": "Exemples",
    "dossier_donnees": "Données",
    "doc_exemple": "Conclusions d'exemple.docx",

    # premier demarrage
    "cree_config_det": "detection_config.json (réglages de la pré-détection)",
    "cree_table": "Correspondances.xlsx (votre table, avec un dossier « Exemple »)",
    "cree_config": "config.json (liste de vos tables)",
    "cree_journaux": "Journaux\\ (les journaux de restitution)",
    "cree_exemple": "Exemples\\Conclusions d'exemple.docx (document fictif pour essayer)",
    "config_lisez_moi": ("Liste de vos tables de correspondances. Le programme ajoute une table "
                         "par le bouton « Ajouter une table… » ; vous n'avez normalement pas a "
                         "modifier ce fichier."),
    "bienvenue_titre": "Bienvenue",
    "bienvenue": ("Le programme a préparé son dossier de travail (sous-dossier Données, "
                  "à côté du programme). Il a créé :\n\n{liste}\n\n"
                  "Pour essayer : choisissez le document du dossier Exemples, le dossier "
                  "« Exemple », puis cliquez sur Anonymiser.\n\n"
                  "Le mode d'emploi complet est dans l'onglet Aide."),
    "demarrage_impossible": "Démarrage impossible",
    "demarrage_detail": ("Un fichier du programme manque ou un composant n'est pas installé.\n\n"
                         "Dossier : {d}\n\nDétail : {e}"),
    "premier_demarrage_impossible": "Premier démarrage impossible",
    "premier_demarrage_detail": (
        "Le programme n'a pas pu préparer son dossier.\n\nDossier : {d}\n\nDétail : {e}\n\n"
        "Vérifiez que le programme est dans un dossier où vous avez le droit d'écrire "
        "(par exemple Documents), pas dans un fichier zip ni dans Program Files."),
    "configuration": "Configuration",
    "aucune_table": "Aucune table disponible",
    "aucune_table_detail": ("Aucune table de correspondances n'a pu être ouverte.\n\n{p}\n\n"
                            "Supprimez config.json : le programme recréera une table au "
                            "prochain démarrage."),
    "tables_introuvables": "Table(s) introuvable(s)",
    "tables_introuvables_detail": "Les tables suivantes sont déclarées mais absentes du disque :\n\n{p}",
    "ouverture_impossible": "Ouverture impossible",

    # contexte
    "contexte_titre": "Le document et le dossier concernés",
    "document": "Document",
    "table": "Table",
    "dossier_onglet": "Dossier (onglet)",
    "choisir": "Choisir…",
    "ouvrir_excel": "Ouvrir dans Excel",
    "ajouter_table": "Ajouter une table…",
    "nouveau_dossier": "Nouveau dossier…",
    "actualiser": "Actualiser",

    # onglets
    "onglet_preparer": "1. Préparer la table",
    "preparer_intro": ("Le programme lit le document, reconnaît les noms, adresses, SIRET, avocats, "
                       "marques… et les range par partie dans un classeur à relire. Vous validez, "
                       "puis il remplit l'onglet du dossier. Rien n'est écrit sans votre relecture."),
    "type_document": "Type de document",
    "sens_courrier": "Sens du courrier",
    "automatique": "Automatique",
    "emis": "Émis par le cabinet",
    "recu": "Reçu",
    "etape1": "Étape 1  ·  Détecter et produire le classeur à relire",
    "etape1_note": ("Crée  _Detection_<document>.xlsx  dans le sous-dossier « _Privé » à côté du document (il contient les vrais noms : ne le joignez jamais). Word et PDF ; un PDF "
                    "scanné est lu par reconnaissance de caractères (plus lent, à vérifier ligne "
                    "par ligne)."),
    "classeur_relu": "Classeur de détection",
    "etape2": "Étape 2  ·  Relire et valider dans Excel",
    "etape2_note": ("Ouvre le classeur de détection dans Excel. Pour chaque ligne, vérifiez trois "
                    "choses : « Garder ? » (OUI = le terme sera anonymisé), la partie, et la balise. "
                    "Une balise à changer se choisit par la flèche du menu déroulant de la case ; "
                    "dans Microsoft 365, tapez ses premières lettres (sans accent : « defe », "
                    "« soc ») pour ne garder que les choix qui les contiennent, puis Entrée. "
                    "Enregistrez (Ctrl+S) et fermez le classeur : c'est la version enregistrée "
                    "qui sera injectée."),
    "relecture_en_cours": "Classeur ouvert dans Excel : relisez, enregistrez, fermez, puis Étape 3.",
    "excel_ouvert_titre": "Classeur encore ouvert dans Excel",
    "classeur_ouvert_excel": ("Le classeur de détection est encore ouvert dans Excel :\n{s}\n\n"
                              "Enregistrez-le (Ctrl+S) et fermez-le, puis cliquez de nouveau sur "
                              "Étape 3. Rien n'a été écrit."),
    "table_ouverte_excel": ("La table de correspondances est ouverte dans Excel :\n{s}\n\n"
                            "Fermez-la, puis cliquez de nouveau sur Étape 3 : le programme doit "
                            "y écrire. Rien n'a été modifié."),
    "etape3": "Étape 3  ·  Injecter les lignes retenues dans le dossier",
    "etape3_note": ("Seules les lignes marquées OUI dans la colonne « Garder ? » sont écrites, "
                    "groupées par partie. Une balise existante n'est jamais modifiée."),
    "onglet_anonymiser": "2. Anonymiser",
    "anonymiser_intro": ("Remplace chaque terme du dossier par sa balise, partout : corps, tableaux, "
                         "en-têtes, notes de bas de page, commentaires, zones de texte, propriétés "
                         "du fichier. La mise en forme est conservée."),
    "anonymiser_bouton": "Anonymiser le document",
    "anonymiser_note": ("Crée  <document>_ANONYMISE.docx  à côté de l'original, et un journal dans "
                        "le dossier Journaux : c'est lui qui permettra de restituer exactement les "
                        "vrais noms."),
    "avant_anonymiser": "Avant d'anonymiser",
    "avant_anonymiser_texte": (
        "Le programme vérifie que le dossier choisi correspond bien au document. Si aucun terme "
        "du dossier n'y figure, il s'arrête ; si un autre dossier correspond mieux, il vous le "
        "dit.\n\nEnsuite, passez par l'onglet « 3. Contrôler » avant tout dépôt dans une IA : "
        "c'est la vérification indépendante."),
    "onglet_controler": "3. Contrôler",
    "controler_intro": ("Relit TOUTES les parties du fichier, y compris celles qu'on ne voit pas à "
                        "l'écran, ainsi que le nom du fichier et de ses dossiers. Il constate, il ne "
                        "corrige rien : c'est un contrôle indépendant."),
    "controler_bouton": "Vérifier ce document avant tout envoi ou dépôt dans une IA",
    "controler_note": ("Écrit un rapport  _Controle_<document>.txt  dans le sous-dossier « _Privé » "
                       "à côté du document : la trace de la vérification (il cite les vrais noms)."),
    "bonne_question": "La bonne question selon l'état du document",
    "bonne_question_texte": (
        "Document anonymisé : reste-t-il un vrai nom ? Si oui, alerte.\n"
        "Document restitué : reste-t-il une balise ? Si oui, alerte. Si non, il peut partir, "
        "mais il contient les vrais noms : jamais dans une IA.\n"
        "Document indéterminé : par prudence, tout terme en clair est une alerte."),
    "onglet_restituer": "4. Restituer",
    "restituer_intro": ("Remet les vrais noms dans un document balisé : la réponse de l'IA "
                        "enregistrée en Word (elle a travaillé avec les balises), ou un document "
                        "relu par un confrère. L'original n'est jamais modifié : un nouveau fichier "
                        "est créé à côté."),
    "restituer_journal": "Restituer avec le journal   (votre document d'origine, exacte)",
    "restituer_journal_note": ("Pour retrouver votre propre document anonymisé sur ce poste. "
                               "Le journal créé lors de l'anonymisation fait foi : aucun risque de "
                               "mélanger deux dossiers. Le programme refuse d'écrire si le journal "
                               "ne correspond pas au document."),
    "restituer_sans": "Restituer sans journal   (via la table du dossier)",
    "restituer_sans_note": ("Pour la réponse de l'IA enregistrée en Word, ou un document dont vous n'avez pas le journal (autre ordinateur, "
                            "document d'un collaborateur, balises tapées à la main). La table et le "
                            "dossier choisis en haut font foi. Seule la première forme de chaque "
                            "balise est restituée : relisez le résultat."),
    "restituer_cree": "Crée  <document>_RESTITUE.docx  à côté du document.",
    "onglet_aide": "Aide",
    "aide_intro": ("L'essentiel pour bien utiliser le programme. Le mode d'emploi complet est dans "
                   "le fichier PDF livré avec lui."),
    "ouvrir_mode_emploi": "Ouvrir le mode d'emploi (PDF)",
    "ouvrir_dossier_programme": "Ouvrir le dossier Données",
    "ouvrir_journaux": "Ouvrir le dossier Journaux",
    "mode_emploi": "Mode d'emploi",
    "mode_emploi_absent": "Le fichier « Mode d'emploi.pdf » n'est pas à côté du programme.",

    # contexte : actions
    "table_illisible": "Table illisible",
    "actualise": "Actualisé : {t} table(s), {d} dossier(s) dans « {n} ».",
    "excel_conseil": ("Modifiez, puis enregistrez dans Excel : le programme relit la table à chaque "
                      "clic. Fermez Excel avant Nouveau dossier ou Injecter."),
    "choisir_classeur": "Choisir un classeur de correspondances",
    "classeurs_excel": "Classeurs Excel",
    "nom_table": "Nom de la table",
    "nom_table_question": "Nom qui s'affichera dans la liste :",
    "ajout_impossible": "Ajout impossible",
    "nouveau_dossier_titre": "Nouveau dossier",
    "nouveau_dossier_question": ("Nom du nouveau dossier (onglet) dans la table « {t} » :\n\n"
                                 "31 caractères au plus, sans  [ ] : * ? / \\"),
    "creation_impossible": "Création impossible",
    "dossier_cree": ("Dossier « {n} » créé. Remplissez-le par l'onglet « 1. Préparer la table » ou "
                     "directement dans Excel."),
    "choisir_document": "Choisir un document",
    "docs_word_pdf": "Documents Word et PDF",
    "pdf_predetection": "PDF (pré-détection seulement)",
    "texte_txt_md": "Texte (.txt, .md)",
    "document_statut": "Document : {n}",
    "choisir_detection": "Choisir le classeur de détection relu",

    # verifications
    "document_manquant": "Document manquant",
    "choisir_dabord": "Choisissez d'abord un document.",
    "word_attendu": "Document Word attendu",
    "word_attendu_detail": ("L'anonymisation, le contrôle et la restitution travaillent sur des "
                            "documents Word (.docx).\n\nUn PDF ne sert qu'à la pré-détection "
                            "(onglet « 1. Préparer la table »)."),
    "dossier_manquant": "Dossier manquant",
    "choisir_table_dossier": "Choisissez la table et le dossier concernés.",
    "anonymisation_en_cours": "Anonymisation en cours…",
    "verification_dossier": "Vérification du dossier…",
    "dossier_sans_rapport": "Dossier sans rapport avec ce document",
    "dossier_sans_rapport_detail": ("Aucun terme du dossier « {o} » n'apparaît dans ce document.{s}"
                                    "\n\nAucun fichier n'a été créé."),
    "suggestion": ("\n\nEn revanche, le dossier « {t} / {o} » trouve {n} terme(s) dans ce document."),
    "interrompue_sans_rapport": "Anonymisation interrompue : dossier sans rapport.",
    "dossier_errone": "Dossier probablement erroné",
    "dossier_errone_detail": ("Vous avez choisi « {t1} / {o1} », qui trouve {n1} terme(s) dans ce "
                              "document.\n\nLe dossier « {t2} / {o2} » en trouve {n2}.\n\n"
                              "Anonymiser quand même avec le dossier choisi ?"),

    # resultats
    "invisibles": "\n\nDont {n} dans les notes de bas de page, commentaires ou zones de texte.",
    "proprietes": "{n} propriété(s) du fichier",
    "liens": "{n} cible(s) de lien hypertexte",
    "egalement": "\n\nÉgalement traité : {a}.",
    "anonymise_resultat": ("{n} remplacement(s).{d}\n\nDocument anonymisé :\n{s}\n\nJournal :\n{j}"
                           "\n\nÉtape suivante conseillée : onglet « 3. Contrôler »."),
    "journal_introuvable": "Journal introuvable",
    "journal_introuvable_detail": ("Aucun journal ne correspond à ce document.\n\nLes journaux sont "
                                   "cherchés dans :\n{d}\n\nSi le document vient d'un autre "
                                   "ordinateur, utilisez « Restituer sans journal »."),
    "restitution_annulee": "Restitution annulée.",
    "verification_journal": "Vérification du journal…",
    "concordance_partielle": "Concordance partielle",
    "concordance_detail": ("Le journal ne retrouve que {p:.0f} % de ses remplacements dans ce "
                           "document.\n\nCela arrive si le document a été modifié depuis son "
                           "anonymisation (paragraphe ajouté, supprimé ou déplacé).\n\n{d}\n\n"
                           "Restituer quand même ? Les balises non couvertes resteront en place, "
                           "et le résultat devra être relu."),
    "table_origine": "\n\nTable d'origine : {t} / {o}",
    "restitue_resultat": ("{n} remplacement(s).\n\nDocument restitué :\n{s}{o}\n\n"
                          "Ce document contient les vrais noms : ne le déposez dans aucune IA."),
    "restitution_en_cours": "Restitution en cours…",
    "restitue_sans_resultat": ("{n} remplacement(s).\n\nDocument restitué :\n{s}\n\nTable utilisée : "
                               "{t} / {o}\n\nRestitution non exacte (première forme de chaque "
                               "balise) : relisez le document."),
    "aucune_balise": "\n\nAucune balise n'a été trouvée dans ce document.",
    "aucun_dossier_couvre": ("\n\nAucun dossier déclaré ne couvre ces balises : il manque "
                             "vraisemblablement une ou plusieurs lignes dans la table."),
    "dossier_couvre": "\n\nLe dossier qui couvre ce document est : {d}",
    "dossiers_couvrent": "\n\nDossiers qui couvriraient ce document :\n    {d}",
    "lecture_parties": "Lecture de toutes les parties du document…",
    "etat_anonymise": "anonymisé", "etat_restitue": "restitué", "etat_indetermine": "indéterminé",
    "etat_document": "Document {e} ({m}).\n\n",
    "termes_en_clair": "Termes en clair : {n} terme(s), {t} occurrence(s)",
    "fois": "({n} fois)",
    "et_autres": "    … et {n} autre(s).",
    "balises_presentes": "\nBalises présentes : {n} balise(s), {t} occurrence(s){s}",
    "normal_anonymise": " (normal dans un document anonymisé)",
    "rapport": "Rapport :\n{r}",
    "alerte": "ALERTE : {t}",
    "zones_examinees": "{t} : {n} zone(s) examinée(s)",

    # detection
    "choisir_document_word_pdf": "Choisissez un document (Word ou PDF).",
    "analyse": "Analyse en cours…",
    "ocr_page": "Lecture par reconnaissance de caractères : page {p} sur {t}…",
    "aucun_terme": "Aucun terme détecté",
    "aucun_terme_detail": "Aucun terme n'a été reconnu dans ce document.",
    "non_reconnu": "non reconnu",
    "ocr_avertissement": ("Ce PDF était scanné : il a été lu par reconnaissance de caractères. Le "
                          "texte obtenu est approché : toutes les lignes sont à vérifier, aucune "
                          "n'est proposée à OUI.\n\n"),
    "ocr_avertissement_mixte": ("Pages {p} de ce PDF lues par reconnaissance de caractères "
                                "(pages scannées) : le texte obtenu est approché, les termes "
                                "trouvés sur ces pages sont à vérifier.\n\n"),
    "detection_resume": ("Type de document : {t}\n\n{n} terme(s) détecté(s) : {r} proposé(s) à OUI, "
                         "{v} à vérifier (dont {c} sans partie : « À CLASSER »).\n"),
    "attention": "\nATTENTION : {a}\n",
    "alterees_restitution": ("\n\nBALISES ALTÉRÉES NON RESTITUÉES ({n}) : l'outil d'IA les a "
                             "recopiées de travers. Elles sont laissées telles quelles ; corrigez-les "
                             "à la main :\n{d}"),
    "alterees_controle": "\nBalises altérées (jamais restituées automatiquement) : {n} forme(s)",
    "hors_table_alerte": "\nALERTE, données hors table en clair : {n}",
    "revelatrices_controle": "\nALERTE, balises qui RÉVÈLENT le nom qu'elles masquent : {n}",
    "hors_table_verifier": "\nÀ vérifier (non bloquant), noms hors table : {n}",
    "ancien_modele": ("\nATTENTION : le dossier choisi suit un ancien modèle d'onglet. Créez un "
                      "nouveau dossier (bouton « Nouveau dossier… »).\n"),
    "classeur_a_relire": ("\nClasseur à relire :\n{s}\n\nÉtape 2 : ouvrez-le ; pour chaque ligne, "
                          "choisissez la Partie puis le Rôle et mettez « Garder ? » à OUI ou NON "
                          "(une balise en rouge = partie qui a déjà une identité) ; enregistrez, "
                          "fermez-le. Puis Étape 3 : Injecter."),
    "ancien_bloque_titre": "Onglet à l'ancien modèle",
    "ancien_bloque": ("L'onglet « {o} » suit l'ancien modèle (cases numérotées) : les termes "
                      "détectés ne pourraient pas y être injectés.\n\nDeux solutions :\n"
                      "  - convertir cet onglet (outil « Convertir anciens onglets ») ;\n"
                      "  - choisir un autre onglet, ou en créer un (« Nouveau dossier… »).\n\n"
                      "L'anonymisation et la restitution avec cet onglet restent possibles."),
    "ancien_convertir": "\n\nOuvrir l'outil de conversion maintenant ?",
    "detection_ok": "OK : {n} terme(s) détecté(s). Relisez le classeur.",
    "detection_terminee": "Détection terminée",
    "ouvrir_classeur_question": "\n\nOuvrir le classeur maintenant ?",
    "ocr_indisponible": "Reconnaissance de caractères indisponible.",
    "ocr_impossible": "Lecture du PDF scanné impossible",
    "detection_interrompue": "Détection interrompue.",
    "detection_interrompue_titre": "Détection interrompue",
    "sens_question_titre": "Sens du courrier",
    "sens_question": ("Ce document est une mise en demeure ou un courrier.\n\nA-t-il été ÉMIS par "
                      "votre cabinet ?\n\nOui : vous écrivez pour votre client.\nNon : le courrier a "
                      "été reçu (votre client en est le destinataire)."),
    "classeur_manquant": "Classeur manquant",
    "classeur_manquant_detail": "Choisissez le classeur de détection que vous avez relu.",
    "classeur_illisible": "Classeur illisible",
    "rien_a_injecter": "Rien à injecter",
    "rien_a_injecter_detail": "Aucune ligne n'est marquée OUI dans la colonne « Garder ? ».",
    "confirmer": "Confirmer",
    "confirmer_injection": ("{n} terme(s) vont être écrits dans :\n\n{t} / {o}\n\nLes lignes sont "
                            "ajoutées dans le groupe de leur partie ; aucune balise existante n'est "
                            "modifiée. Continuer ?"),
    "injection_interrompue": "Injection interrompue.",
    "injection_interrompue_titre": "Injection interrompue",
    "variantes": "   (variantes)",
    "ecrits": "{n} terme(s) écrit(s) :\n\n{d}\n",
    "non_places": "\nNon placés :\n{d}",
    "dossier_pret": "\n\nLe dossier est prêt : passez à l'onglet « 2. Anonymiser ».",
    "injection_ok": "OK : {e} écrit(s), {r} non placé(s).",
    "injection_terminee": "Injection terminée",

    # retours
    "succes": "OK : {n} remplacement(s) : {s}",
    "termine": "Terminé",
    "interrompue": "Opération interrompue : voir le message.",
    "interrompue_titre": "Opération interrompue",

    # journal
    "choisir_journal": "Choisir le journal",
    "plusieurs_journaux": "Plusieurs journaux correspondent à ce document",
    "selectionnez_journal": "Sélectionnez celui du dossier concerné.",
    "table_inconnue": "table inconnue", "onglet_inconnu": "onglet inconnu",
    "format_ancien": "   (format ancien)",
    "annuler": "Annuler",
    "utiliser_journal": "Utiliser ce journal",

    # verification
    "verif_ok": "Vérification réussie",
    "verif_ok_detail": "Tous les contrôles sont passés.\n\nCompte rendu :\n{s}",
    "verif_ko": "Vérification en échec",
    "verif_ko_detail": "Au moins un contrôle a échoué.\n\nCompte rendu :\n{s}",

    "aide": [
        ("titre", "Le principe en trois phrases"),
        ("", "Votre table de correspondances (Correspondances.xlsx) dit quel vrai nom devient "
             "quelle balise : « Société Horizon Numérique » devient [DEM1]. Le programme "
             "remplace, vous déposez le document anonymisé dans l'IA, et le programme remet "
             "les vrais noms dans la réponse. La table est votre clé : elle ne sort jamais."),
        ("titre", "Un onglet par dossier, pas par document"),
        ("", "L'anonymisation ne se prépare pas document par document, mais une fois pour tout le "
             "dossier. L'onglet est la liste des noms du dossier et de leurs balises : une fois "
             "rempli et bien complété, il sert à toutes les pièces, passées et à venir. Pour une "
             "nouvelle pièce : choisir le document, choisir l'onglet, cliquer ; inutile de refaire "
             "une détection. La détection n'est qu'une aide pour remplir l'onglet au départ, puis "
             "l'enrichir : un nom nouveau = une ligne de plus, et toutes les pièces en profitent. "
             "La même société reste [DEF1] d'un document à l'autre : l'IA suit le dossier "
             "sans jamais voir un nom."),
        ("titre", "Deux façons de préparer un dossier"),
        ("", "Le chemin rapide, pour donner vite des pièces à une IA et obtenir une synthèse : "
             "Détecter, puis Injecter, puis Anonymiser. Vous n'avez rien à savoir du dossier. Ce que le "
             "programme reconnaît avec certitude est rangé par partie ([DEF1], [DEF1_ADRESSE]). Le reste "
             "reçoit une balise neutre selon sa nature ([SOCIETE5], [ADRESSE3], [TELEPHONE2]). Tout est "
             "masqué.\n"
             "Le chemin précis, pour des écritures ou un dossier où il faut savoir qui est qui : avant "
             "d'injecter, ouvrez le tableau (Étape 2). Sur chaque ligne à préciser, choisissez la Partie, "
             "puis le Rôle. Les parties connues sont affichées avec leur nom. « nouvelle partie » crée la "
             "suivante, numérotée toute seule.\n"
             "Rattacher plus tard. Vous découvrez dans une autre pièce à qui appartient [ADRESSE3] ? Dans "
             "l'onglet du dossier, changez la Partie de cette ligne. La case Balise passe en orange : la "
             "ligne prendra [DEF1_ADRESSE] au prochain usage. Les documents déjà anonymisés avec "
             "[ADRESSE3] restent restituables, même sans journal.\n"
             "Quatre règles à retenir :\n"
             "  - Même numéro = même partie.\n"
             "  - « nouvelle partie » donne le numéro suivant tout seul, sans limite.\n"
             "  - Rouge = deux identités différentes sous la même balise.\n"
             "  - Une balise déjà utilisée n'est jamais perdue : la restitution la reconnaît toujours."),
        ("titre", "L'ordre habituel"),
        ("", "1. Préparer la table, en trois étapes : Détecter (le programme repère les termes "
             "d'un document et les range par partie dans un classeur de relecture) ; Relire et "
             "valider dans Excel (pour chaque ligne : la Partie, qui ?, puis le Rôle, quoi ?, dans "
             "deux menus courts, la balise s'écrivant toute seule ; OUI ou NON dans « Garder ? » ; "
             "une balise en rouge signale une partie qui a déjà une identité ; enregistrer, fermer) ; Injecter (les lignes retenues entrent dans "
             "l'onglet du dossier). Ou vous remplissez l'onglet vous-même dans Excel.\n"
             "2. Anonymiser : crée <document>_ANONYMISE.docx et son journal.\n"
             "3. Contrôler : vérification indépendante avant tout dépôt dans une IA.\n"
             "4. Restituer : remet les vrais noms dans un document anonymisé (avec le journal, "
             "ou sans journal via la table)."),
        ("titre", "La table : un fichier Excel ordinaire, que vous pouvez remplir à la main"),
        ("", "Un onglet par dossier. Créez-le depuis le programme (Nouveau dossier…) : il arrive "
             "au bon format (ou copiez l'onglet « Modèle complet », qui porte déjà toutes les "
             "balises, et remplissez la colonne Terme). Ouvrez-le avec « Ouvrir dans Excel » : quarante lignes sont prêtes "
             "sous le tableau. Tapez le vrai nom, puis choisissez le rôle et la partie dans les "
             "menus déroulants : la balise s'écrit toute seule, numéro compris "
             "([DEM1_LOGICIEL], puis [DEM1_LOGICIEL2]), et la ligne prend la couleur "
             "de sa partie. Un rôle propre à votre spécialité se tape librement : la balise se "
             "forme avec. Commissaire de justice, expert, salarié : partie COMMISSAIRE1, EXPERT1, "
             "SALARIE1 et rôle Identité. Les parties ont des codes courts numérotés : DEM1, DEM2 (demandeurs), "
             "DEF1 (défendeur), APP1, INTIME1, INTERV1, CLI (notre client), ADV (partie adverse), "
             "et pour les autres documents SOCIETE1, PERSONNE_PHYSIQUE1, PERSONNE_MORALE1 ; ce qui "
             "appartient à une partie s'y rattache : [DEF1_ADRESSE], [DEF1_DIRIGEANT]. La liste "
             "s'allonge dans detection_config.json (réglage parties_menu). Plusieurs façons "
             "d'écrire un même nom : une ligne par écriture avec la même balise, ou sur une "
             "ligne séparées par « / ». Ne modifiez jamais une balise déjà utilisée dans un "
             "document anonymisé.\n"
             "Enregistrez dans Excel : c'est tout. Le programme relit la table à chaque clic, "
             "sans qu'il soit nécessaire de le relancer. Si un nouvel onglet n'apparaît pas dans "
             "la liste Dossier, cliquez sur Actualiser. Fermez Excel seulement quand le programme "
             "doit écrire dans la table (Nouveau dossier, Injecter)."),
        ("titre", "Ce que l'anonymisation couvre"),
        ("", "Corps, tableaux, en-têtes et pieds de page, notes de bas de page et de fin, "
             "commentaires, zones de texte, propriétés du fichier (auteur, société), cibles des "
             "liens hypertexte. La mise en forme (gras, italique) est conservée. Le nom du "
             "fichier n'est pas modifié : le contrôle vous le signale s'il contient un nom."),
        ("titre", "Les marqueurs de rédaction"),
        ("", "[DATE], [MONTANT], [A COMPLETER]… ressemblent à des balises mais n'en sont pas : "
             "seule une balise déclarée dans la table en est une. Les autres traversent "
             "l'anonymisation et la restitution sans être touchés."),
        ("titre", "Trois règles d'or"),
        ("", "1. Contrôlez toujours avant de déposer dans une IA.\n"
             "2. Un document restitué contient les vrais noms : il ne va jamais dans une IA.\n"
             "3. Relisez le document produit. Le programme vérifie beaucoup, il ne lit pas à "
             "votre place."),
        ("titre", "Aucune intelligence artificielle dans le programme"),
        ("", "Le programme n'utilise aucune IA, ni pour détecter, ni pour anonymiser, ni pour "
             "contrôler ou restituer : uniquement des règles (civilités, formes de sociétés, "
             "adresses, numéros) et votre table, sur votre ordinateur, sans réseau. L'IA n'est "
             "que la destination du document anonymisé, et c'est vous qui l'y déposez."),
        ("titre", "Où sont les fichiers"),
        ("", "Votre travail est dans le sous-dossier Données\\, à côté du programme : "
             "Correspondances.xlsx (votre table), Journaux\\ (les journaux de restitution), "
             "config.json et detection_config.json (réglages), Exemples\\ (un document fictif). "
             "Pour sauvegarder votre travail, copiez le dossier Données\\ ; pour changer "
             "d'ordinateur, copiez-le au même endroit à côté du nouveau programme. "
             "(Une installation antérieure, dont les fichiers étaient directement à côté du "
             "programme, garde son organisation : rien n'est déplacé.)"),
        ("", "À côté du programme lui-même : Mode d'emploi.pdf, Tesseract\\ (lecture des PDF "
             "scannés), Code source\\ et LICENCE.txt (logiciel libre, licence AGPL v3). "
             "Les fichiers produits (document anonymisé, restitué) sont écrits à côté du "
             "document d'origine ; le classeur de détection et le rapport de contrôle, qui "
             "contiennent les vrais noms, dans son sous-dossier _Privé\\."),
    ],
}

EN = {
    "nom": "Anonymizator",
    "sous_titre": "Protect client confidentiality before anything goes into an AI",
    "bandeau": "Your document never leaves your computer: no AI, no upload, no network.",
    "version": "version {v}",
    "developpe": "Developed by {a}",
    "pret": "Ready.",
    "mode_emploi_pdf": "User guide.pdf",
    "table_defaut": "Correspondences.xlsx",
    "nom_table_defaut": "My correspondences",
    "dossier_journaux": "Logs",
    "dossier_exemples": "Examples",
    "dossier_donnees": "Data",
    "doc_exemple": "Sample pleadings.docx",

    "cree_config_det": "detection_config.json (pre-detection settings)",
    "cree_table": "Correspondences.xlsx (your table, with an \"Example\" matter)",
    "cree_config": "config.json (list of your tables)",
    "cree_journaux": "Logs\\ (the restoration logs)",
    "cree_exemple": "Examples\\Sample pleadings.docx (fictional document to try the program)",
    "config_lisez_moi": ("List of your correspondence tables. The program adds a table through the "
                         "\"Add a table...\" button; you normally never need to edit this file."),
    "bienvenue_titre": "Welcome",
    "bienvenue": ("The program has prepared its working folder (Data subfolder, next to "
                  "the program). It created:\n\n{liste}\n\n"
                  "To try it: pick the document in the Examples folder, the matter \"Example\", "
                  "then click Anonymise.\n\nThe full user guide is in the Help tab."),
    "demarrage_impossible": "Cannot start",
    "demarrage_detail": ("A program file is missing or a component is not installed.\n\n"
                         "Folder: {d}\n\nDetail: {e}"),
    "premier_demarrage_impossible": "First start failed",
    "premier_demarrage_detail": (
        "The program could not prepare its folder.\n\nFolder: {d}\n\nDetail: {e}\n\n"
        "Make sure the program sits in a folder where you are allowed to write "
        "(for instance Documents), not inside a zip file or in Program Files."),
    "configuration": "Configuration",
    "aucune_table": "No table available",
    "aucune_table_detail": ("No correspondence table could be opened.\n\n{p}\n\n"
                            "Delete config.json: the program will create a new table at the "
                            "next start."),
    "tables_introuvables": "Table(s) not found",
    "tables_introuvables_detail": "The following tables are declared but missing from disk:\n\n{p}",
    "ouverture_impossible": "Cannot open",

    "contexte_titre": "The document and the matter",
    "document": "Document",
    "table": "Table",
    "dossier_onglet": "Matter (sheet)",
    "choisir": "Choose…",
    "ouvrir_excel": "Open in Excel",
    "ajouter_table": "Add a table…",
    "nouveau_dossier": "New matter…",
    "actualiser": "Refresh",

    "onglet_preparer": "1. Prepare the table",
    "preparer_intro": ("The program reads the document, recognises names, addresses, company "
                       "numbers, lawyers, trademarks… and sorts them by party in a workbook for "
                       "you to review. You validate, then it fills the matter sheet. Nothing is "
                       "written without your review."),
    "type_document": "Document type",
    "sens_courrier": "Letter: sent or received",
    "automatique": "Automatic",
    "emis": "Sent by the firm",
    "recu": "Received",
    "etape1": "Step 1  ·  Detect and produce the workbook to review",
    "etape1_note": ("Creates  _Detection_<document>.xlsx  in the \"_Private\" subfolder next to the document (it holds the real names: never attach it). Word and PDF; a "
                    "scanned PDF is read by text recognition (OCR): slower, and every row must "
                    "be checked."),
    "classeur_relu": "Detection workbook",
    "etape2": "Step 2  ·  Review and validate in Excel",
    "etape2_note": ("Opens the detection workbook in Excel. For each row, check three things: "
                    "\"Keep?\" (YES = the term will be anonymised), the party, and the tag. To "
                    "change a tag, click the arrow of the cell's drop-down list; in Microsoft 365, "
                    "type its first letters (\"defe\", \"comp\") to keep only the entries containing "
                    "them, then Enter. Save (Ctrl+S) and close the workbook: the saved version is "
                    "what gets injected."),
    "relecture_en_cours": "Workbook open in Excel: review, save, close, then Step 3.",
    "excel_ouvert_titre": "Workbook still open in Excel",
    "classeur_ouvert_excel": ("The detection workbook is still open in Excel:\n{s}\n\n"
                              "Save it (Ctrl+S) and close it, then click Step 3 again. Nothing "
                              "was written."),
    "table_ouverte_excel": ("The correspondence table is open in Excel:\n{s}\n\n"
                            "Close it, then click Step 3 again: the program must write to it. "
                            "Nothing was changed."),
    "etape3": "Step 3  ·  Inject the kept rows into the matter sheet",
    "etape3_note": ("Only rows marked YES in the \"Keep?\" column are written, grouped by party. "
                    "An existing tag is never modified."),
    "onglet_anonymiser": "2. Anonymise",
    "anonymiser_intro": ("Replaces every term of the matter with its tag, everywhere: body, tables, "
                         "headers, footnotes, comments, text boxes, file properties. Formatting "
                         "is preserved."),
    "anonymiser_bouton": "Anonymise the document",
    "anonymiser_note": ("Creates  <document>_ANONYMISED.docx  next to the original, plus a log in "
                        "the Logs folder: the log is what restores the real names exactly."),
    "avant_anonymiser": "Before anonymising",
    "avant_anonymiser_texte": (
        "The program checks that the chosen matter matches the document. If no term of the "
        "matter appears in it, it stops; if another matter matches better, it tells you.\n\n"
        "Then run the \"3. Check\" tab before anything goes into an AI: it is the "
        "independent verification."),
    "onglet_controler": "3. Check",
    "controler_intro": ("Re-reads EVERY part of the file, including what is not visible on screen, "
                        "plus the file name and folder path. It reports and changes nothing: "
                        "an independent check."),
    "controler_bouton": "Check this document before sending it or putting it into an AI",
    "controler_note": ("Writes a report  _Check_<document>.txt  in the \"_Private\" subfolder next "
                       "to the document: the record of the check (it quotes the real names)."),
    "bonne_question": "The right question depends on the state of the document",
    "bonne_question_texte": (
        "Anonymised document: is any real name left? If yes, alert.\n"
        "Restored document: is any tag left? If yes, alert. If not, it can go out, but it "
        "contains the real names: never into an AI.\n"
        "Document of unclear status: as a precaution, any unmasked term is an alert."),
    "onglet_restituer": "4. Restore",
    "restituer_intro": ("Puts the real names back into a tagged document: the AI's answer saved as "
                        "a Word document (it worked with the tags), or a document reviewed by a colleague. The "
                        "original is never modified: a new file is created next to it."),
    "restituer_journal": "Restore with the log   (your original document, exact)",
    "restituer_journal_note": ("To get back your own document anonymised on this computer. "
                               "The log created at anonymisation is the reference: no risk of mixing up "
                               "two matters. The program refuses to write if the log does not "
                               "match the document."),
    "restituer_sans": "Restore without the log   (using the matter's table)",
    "restituer_sans_note": ("For the AI's answer saved as a Word document, or a document whose log you do not have (another computer, a "
                            "colleague's document, tags typed by hand). The table and matter "
                            "chosen above are used. Each tag is restored to its first listed "
                            "spelling: proofread the result."),
    "restituer_cree": "Creates  <document>_RESTORED.docx  next to the document.",
    "onglet_aide": "Help",
    "aide_intro": "The essentials. The full user guide is in the PDF delivered with the program.",
    "ouvrir_mode_emploi": "Open the user guide (PDF)",
    "ouvrir_dossier_programme": "Open the Data folder",
    "ouvrir_journaux": "Open the Logs folder",
    "mode_emploi": "User guide",
    "mode_emploi_absent": "The file \"User guide.pdf\" is not next to the program.",

    "table_illisible": "Unable to read the table",
    "actualise": "Refreshed: {t} table(s), {d} matter(s) in \"{n}\".",
    "excel_conseil": ("Edit, then save in Excel: the program re-reads the table at every click. "
                      "Close Excel before New matter or Inject."),
    "choisir_classeur": "Choose a correspondence workbook",
    "classeurs_excel": "Excel workbooks",
    "nom_table": "Table name",
    "nom_table_question": "Name shown in the list:",
    "ajout_impossible": "Cannot add",
    "nouveau_dossier_titre": "New matter",
    "nouveau_dossier_question": ("Name of the new matter (sheet) in table \"{t}\":\n\n"
                                 "31 characters at most, without  [ ] : * ? / \\"),
    "creation_impossible": "Cannot create",
    "dossier_cree": ("Matter \"{n}\" created. Fill it through the \"1. Prepare the table\" tab or "
                     "directly in Excel."),
    "choisir_document": "Choose a document",
    "docs_word_pdf": "Word and PDF documents",
    "pdf_predetection": "PDF (pre-detection only)",
    "texte_txt_md": "Text (.txt, .md)",
    "document_statut": "Document: {n}",
    "choisir_detection": "Choose the reviewed detection workbook",

    "document_manquant": "No document",
    "choisir_dabord": "Choose a document first.",
    "word_attendu": "Word document expected",
    "word_attendu_detail": ("Anonymisation, check and restoration work on Word documents (.docx)."
                            "\n\nA PDF is only used for pre-detection (tab \"1. Prepare the "
                            "table\")."),
    "dossier_manquant": "No matter",
    "choisir_table_dossier": "Choose the table and the matter.",
    "anonymisation_en_cours": "Anonymising…",
    "verification_dossier": "Checking the matter…",
    "dossier_sans_rapport": "Matter unrelated to this document",
    "dossier_sans_rapport_detail": ("No term of matter \"{o}\" appears in this document.{s}"
                                    "\n\nNo file was created."),
    "suggestion": "\n\nHowever, matter \"{t} / {o}\" finds {n} term(s) in this document.",
    "interrompue_sans_rapport": "Anonymisation stopped: unrelated matter.",
    "dossier_errone": "Probably the wrong matter",
    "dossier_errone_detail": ("You chose \"{t1} / {o1}\", which finds {n1} term(s) in this "
                              "document.\n\nMatter \"{t2} / {o2}\" finds {n2}.\n\n"
                              "Anonymise anyway with the chosen matter?"),

    "invisibles": "\n\nOf which {n} in footnotes, comments or text boxes.",
    "proprietes": "file properties ({n})",
    "liens": "hyperlink targets ({n})",
    "egalement": "\n\nAlso processed: {a}.",
    "anonymise_resultat": ("{n} replacement(s).{d}\n\nAnonymised document:\n{s}\n\nLog:\n{j}"
                           "\n\nRecommended next step: the \"3. Check\" tab."),
    "journal_introuvable": "Log not found",
    "journal_introuvable_detail": ("No log matches this document.\n\nLogs are looked for in:\n{d}"
                                   "\n\nIf the document comes from another computer, use "
                                   "\"Restore without the log\"."),
    "restitution_annulee": "Restoration cancelled.",
    "verification_journal": "Checking the log…",
    "concordance_partielle": "Partial match",
    "concordance_detail": ("The log only finds {p:.0f}% of its replacements in this document.\n\n"
                           "This happens when the document was edited after anonymisation "
                           "(paragraph added, removed or moved).\n\n{d}\n\nRestore anyway? "
                           "Tags the log does not cover will stay as they are; proofread the result."),
    "table_origine": "\n\nOriginal table: {t} / {o}",
    "restitue_resultat": ("{n} replacement(s).\n\nRestored document:\n{s}{o}\n\n"
                          "This document contains the real names: never put it into an AI."),
    "restitution_en_cours": "Restoring…",
    "restitue_sans_resultat": ("{n} replacement(s).\n\nRestored document:\n{s}\n\nTable used: "
                               "{t} / {o}\n\nNot an exact restoration (each tag gets its first listed "
                               "spelling): proofread the document."),
    "aucune_balise": "\n\nNo tag was found in this document.",
    "aucun_dossier_couvre": ("\n\nNo declared matter covers these tags: one or more rows are "
                             "probably missing from the table."),
    "dossier_couvre": "\n\nThe matter covering this document is: {d}",
    "dossiers_couvrent": "\n\nMatters that would cover this document:\n    {d}",
    "lecture_parties": "Reading every part of the document…",
    "etat_anonymise": "anonymised", "etat_restitue": "restored", "etat_indetermine": "undetermined",
    "etat_document": "Document {e} ({m}).\n\n",
    "termes_en_clair": "Unmasked terms: {n} term(s), {t} occurrence(s)",
    "fois": "(×{n})",
    "et_autres": "    … and {n} more.",
    "balises_presentes": "\nTags present: {n} tag(s), {t} occurrence(s){s}",
    "normal_anonymise": " (normal in an anonymised document)",
    "rapport": "Report:\n{r}",
    "alerte": "ALERT: {t}",
    "zones_examinees": "{t}: {n} area(s) examined",

    "choisir_document_word_pdf": "Choose a document (Word or PDF).",
    "analyse": "Analysing…",
    "ocr_page": "Text recognition (OCR): page {p} of {t}…",
    "aucun_terme": "No term detected",
    "aucun_terme_detail": "No term was recognised in this document.",
    "non_reconnu": "not recognised",
    "ocr_avertissement": ("This PDF was scanned: it was read by text recognition (OCR). The "
                          "text is approximate: every row must be checked, none is proposed as "
                          "YES.\n\n"),
    "ocr_avertissement_mixte": ("Pages {p} of this PDF were read by text recognition (OCR) "
                                "(scanned pages): the text is approximate, terms found on "
                                "these pages must be checked.\n\n"),
    "detection_resume": ("Document type: {t}\n\n{n} term(s) detected: {r} proposed as YES, "
                         "{v} to check (of which {c} without a party: \"TO SORT\").\n"),
    "attention": "\nWARNING: {a}\n",
    "alterees_restitution": ("\n\nALTERED TAGS NOT RESTORED ({n}): the AI tool miscopied them. "
                             "They are left as they are; fix them by hand:\n{d}"),
    "alterees_controle": "\nAltered tags (never restored automatically): {n} form(s)",
    "hors_table_alerte": "\nALERT, data not in the table, unmasked: {n}",
    "revelatrices_controle": "\nALERT, tags that REVEAL the name they mask: {n}",
    "hors_table_verifier": "\nTo check (not blocking), names not in the table: {n}",
    "ancien_modele": ("\nWARNING: the chosen matter follows an old sheet model. Create a new matter "
                      "(\"New matter…\" button).\n"),
    "classeur_a_relire": ("\nWorkbook to review:\n{s}\n\nStep 2: open it; for each row pick the "
                          "Party then the Role and set \"Keep?\" to YES or NO (a tag in red = a party "
                          "that already has an identity); save, close it. Then Step 3: Inject."),
    "ancien_bloque_titre": "Old-layout sheet",
    "ancien_bloque": ("The sheet \"{o}\" uses the old layout (numbered boxes): detected terms "
                      "could not be injected into it.\n\nTwo options:\n"
                      "  - convert this sheet (\"Convert old sheets\" tool);\n"
                      "  - pick another sheet, or create one (\"New matter…\").\n\n"
                      "Anonymising and restoring with this sheet still work."),
    "ancien_convertir": "\n\nOpen the conversion tool now?",
    "detection_ok": "OK: {n} term(s) detected. Review the workbook.",
    "detection_terminee": "Detection finished",
    "ouvrir_classeur_question": "\n\nOpen the workbook now?",
    "ocr_indisponible": "Text recognition (OCR) unavailable.",
    "ocr_impossible": "Cannot read the scanned PDF",
    "detection_interrompue": "Detection stopped.",
    "detection_interrompue_titre": "Detection stopped",
    "sens_question_titre": "Sent or received?",
    "sens_question": ("This document is a letter (demand letter, letter before action…).\n\nWas it SENT by your firm?"
                      "\n\nYes: you write for your client.\nNo: the letter was received (your "
                      "client is the addressee)."),
    "classeur_manquant": "No workbook",
    "classeur_manquant_detail": "Choose the detection workbook you reviewed.",
    "classeur_illisible": "Unable to read the workbook",
    "rien_a_injecter": "Nothing to inject",
    "rien_a_injecter_detail": "No row is marked YES in the \"Keep?\" column.",
    "confirmer": "Confirm",
    "confirmer_injection": ("{n} term(s) will be written into:\n\n{t} / {o}\n\nRows are added in "
                            "their party's group; no existing tag is modified. Continue?"),
    "injection_interrompue": "Injection stopped.",
    "injection_interrompue_titre": "Injection stopped",
    "variantes": "   (variants)",
    "ecrits": "{n} term(s) written:\n\n{d}\n",
    "non_places": "\nNot placed:\n{d}",
    "dossier_pret": "\n\nThe matter is ready: go to the \"2. Anonymise\" tab.",
    "injection_ok": "OK: {e} written, {r} not placed.",
    "injection_terminee": "Injection finished",

    "succes": "OK: {n} replacement(s): {s}",
    "termine": "Done",
    "interrompue": "Operation stopped: see the message.",
    "interrompue_titre": "Operation stopped",

    "choisir_journal": "Choose the log",
    "plusieurs_journaux": "Several logs match this document",
    "selectionnez_journal": "Select the one for the right matter.",
    "table_inconnue": "unknown table", "onglet_inconnu": "unknown sheet",
    "format_ancien": "   (old format)",
    "annuler": "Cancel",
    "utiliser_journal": "Use this log",

    "verif_ok": "Verification passed",
    "verif_ok_detail": "All checks passed.\n\nReport:\n{s}",
    "verif_ko": "Verification failed",
    "verif_ko_detail": "At least one check failed.\n\nReport:\n{s}",

    "aide": [
        ("titre", "The principle in three sentences"),
        ("", "Your correspondence table (Correspondences.xlsx) says which real name becomes "
             "which tag: \"Horizon Digital Ltd\" becomes [CLM1]. The program makes the "
             "replacements, you give the anonymised document to the AI, and the program puts the "
             "real names back into its answer. The table is your key: it never leaves your computer."),
        ("titre", "One sheet per matter, not per document"),
        ("", "Anonymisation is not prepared document by document, but once for the whole matter. "
             "The sheet is the list of the matter's names and their tags: once filled in and "
             "completed, it serves every document, past and future. For a new document: pick the "
             "document, pick the sheet, click; no need to run detection again. Detection is only "
             "an aid to fill the sheet at the start, then to enrich it: a new name = one more row, "
             "and every document benefits. The same company stays [DEF1] from one document "
             "to the next: the AI follows the matter without ever seeing a name."),
        ("titre", "Two ways to prepare a matter"),
        ("", "The quick path, to hand documents to an AI fast and get a summary: Detect, then Inject, "
             "then Anonymise. You need to know nothing about the matter. What the program recognises "
             "with certainty is sorted by party ([DEF1], [DEF1_ADDRESS]). The rest gets a neutral tag by "
             "its nature ([COMPANY5], [ADDRESS3], [PHONE2]). Everything is masked.\n"
             "The precise path, for pleadings or a matter where who is who matters: before injecting, "
             "open the workbook (Step 2). On each line to refine, pick the Party, then the Role. Known "
             "parties are shown with their name. « new party » creates the next one, numbered by itself.\n"
             "Attaching later. You find out in another document who [ADDRESS3] belongs to? In the matter "
             "sheet, change the Party of that row. The Tag cell turns orange: the row will take "
             "[DEF1_ADDRESS] at the next use. Documents already anonymised with [ADDRESS3] can still be "
             "restored, even without the log.\n"
             "Four rules to remember:\n"
             "  - Same number = same party.\n"
             "  - « new party » gives the next number by itself, without limit.\n"
             "  - Red = two different identities under the same tag.\n"
             "  - A tag already used is never lost: restoration always recognises it."),
        ("titre", "The usual order"),
        ("", "1. Prepare the table, in three steps: Detect (the program finds the terms of a "
             "document and sorts them by party in a review workbook); Review and validate in "
             "Excel (for each row: the Party, who?, then the Role, what?, from two short lists, the "
             "tag filling itself in; YES or NO in \"Keep?\"; a tag in red flags a party that "
             "already has an identity; save, close); Inject "
             "(the kept rows go into the matter sheet). Or you fill the sheet yourself in Excel.\n"
             "2. Anonymise: creates <document>_ANONYMISED.docx and its log.\n"
             "3. Check: independent verification before anything goes into an AI.\n"
             "4. Restore: puts the real names back into an anonymised document (with the log, "
             "or without it using the table)."),
        ("titre", "The table: an ordinary Excel file you can fill in by hand"),
        ("", "One sheet per matter. Create it from the program (New matter…): it comes in the "
             "right format (or copy the \"Full template\" sheet, which already carries every "
             "tag, and fill the Term column). Open it with \"Open in Excel\": forty rows are ready below the "
             "table. Type the real name, then pick the role and the party from the drop-down lists: "
             "the tag fills itself in, number included ([CLM1_SOFTWARE], then "
             "[CLM1_SOFTWARE2]), and the row takes its party's colour. A role of your own "
             "speciality can be typed freely: the tag is built from it. Bailiff, expert, "
             "employee: party BAILIFF1, EXPERT1, EMPLOYEE1 with role Identity. Parties have short numbered codes: CLM1, CLM2 (claimants), DEF1 "
             "(defendant), APP1, RESP1, INTERV1, CLI (our client), OPP (opposing party), and for "
             "other documents COMPANY1, INDIVIDUAL1, LEGAL_ENTITY1; what belongs to a party hangs "
             "off it: [DEF1_ADDRESS], [DEF1_DIRECTOR]. The list grows in detection_config.json "
             "(parties_menu setting). Several spellings of the "
             "same name: one row per spelling with the same tag, or on one row separated by "
             "\" / \". Never change a tag already used in an anonymised document.\n"
             "Save in Excel and you are done. The program re-reads the table at every click, no "
             "need to restart it. If a new sheet does not show in the Matter list, click "
             "Refresh. Close Excel only when the program must write into the table (New "
             "matter, Inject)."),
        ("titre", "What anonymisation covers"),
        ("", "Body, tables, headers and footers, footnotes and endnotes, comments, text boxes, "
             "file properties (author, company), hyperlink targets. Formatting (bold, italic) "
             "is preserved. The file name is not changed: the check warns you if it contains "
             "a name."),
        ("titre", "Drafting placeholders"),
        ("", "[DATE], [AMOUNT], [TO COMPLETE]… look like tags but are not: only a tag declared "
             "in the table is one. The others go through anonymisation and restoration "
             "untouched."),
        ("titre", "Three golden rules"),
        ("", "1. Always check before putting a document into an AI.\n"
             "2. A restored document contains the real names: it never goes into an AI.\n"
             "3. Proofread the produced document. The program checks a lot; it does not read "
             "for you."),
        ("titre", "No artificial intelligence inside the program"),
        ("", "The program uses no AI at all: not to detect, not to anonymise, not to check or "
             "restore. Only rules (titles, company forms, addresses, numbers) and your table, on "
             "your computer, with no network. The AI is only where the anonymised document goes, "
             "and you are the one who puts it there."),
        ("titre", "Where the files are"),
        ("", "Your work is in the Data\\ subfolder, next to the program: Correspondences.xlsx "
             "(your table), Logs\\ (the restoration logs), config.json and detection_config.json "
             "(settings), Examples\\ (a fictional document). To back up your work, copy the "
             "Data\\ folder; to move to another computer, copy it to the same place next to "
             "the new program. (An earlier installation, whose files sat directly next to the "
             "program, keeps its layout: nothing is moved.)"),
        ("", "Next to the program itself: User guide.pdf, Tesseract\\ (reading scanned PDFs), "
             "Source code\\ and LICENSE.txt (free software, AGPL v3). The files produced "
             "(anonymised, restored document) are written next to the original document; the "
             "detection workbook and the check report, which contain the real names, go into "
             "its _Private\\ subfolder."),
        ("titre", "A note on pre-detection"),
        ("", "Pre-detection was built for French legal documents (French titles such as Maître, company forms, "
             "French addresses and company numbers). On English documents it still finds "
             "e-mails, websites, numbers and most capitalised company names, but it will miss "
             "some people and roles: review its workbook carefully, or fill the table in by hand. "
             "Anonymisation, check and restoration work in any language."),
    ],
}


# ------------------------------------------------------------------
# MESSAGES DES MOTEURS (ecrits en francais) : traduction a l'affichage
# ------------------------------------------------------------------
# Les moteurs (anonymisation, detection, table) produisent leurs messages en
# francais. En version anglaise, tout message montre a l'utilisateur (fenetre,
# classeur de relecture) passe par traduire(). Un message non prevu reste en
# francais : c'est un oubli a ajouter ici, jamais une erreur de fonctionnement.
_TR = [
    # anonymisation_core
    (r"config\.json introuvable :", "config.json not found:"),
    (r"Aucune table de correspondances declaree dans config\.json\.\n\nAttendu : une liste \"tables\", "
     r"chaque entree ayant un \"nom\" et un \"fichier\"\.",
     "No correspondence table is declared in config.json.\n\nExpected: a \"tables\" list, each "
     "entry having a \"nom\" (name) and a \"fichier\" (file)."),
    (r"Table de correspondances introuvable :", "Correspondence table not found:"),
    (r"« (.+?) » : fichier introuvable", r'"\1": file not found'),
    (r"Onglet « (.+?) » introuvable\.\n\nOnglets du classeur : ", r'Sheet "\1" not found.\n\nSheets in the workbook: '),
    (r"Onglet « (.+?) » introuvable\.", r'Sheet "\1" not found.'),
    (r"Aucune correspondance active dans l'onglet « (.+?) »\.\n\nVerifiez la colonne A \(terme original\) "
     r"et la colonne F \(actif\)\.",
     r'Sheet "\1" has no active row.\n\nCheck column A (original term) and column F (active).'),
    (r"Journal introuvable :", "Log not found:"),
    (r"Journal illisible \(format inattendu\) :", "Log cannot be read (unexpected format):"),
    (r"Le journal ne contient aucun remplacement :", "The log contains no replacement:"),
    (r"Rien a restituer\.", "Nothing to restore."),
    (r"REFUS : la table selectionnee ne couvre pas tout le document\.",
     "REFUSED: the selected table does not cover the whole document."),
    (r"balise\(s\) du document sont absentes de cette table :", "tag(s) of the document are missing from this table:"),
    (r"Restituer malgre tout laisserait des balises non resolues, ou pire,\nmelangerait deux dossiers\. "
     r"Aucune ecriture effectuee\.",
     "Restoring anyway would leave unresolved tags or, worse,\nmix two matters. Nothing was written."),
    (r"Restitution sans journal non confirmee : aucune ecriture effectuee\.\n\nCette voie relit la table "
     r"de correspondances et non le journal :\nelle doit etre validee a l'ecran avant execution\.",
     "Restoration without the log was not confirmed: nothing was written.\n\nThis route reads the "
     "correspondence table, not the log:\nit must be confirmed on screen first."),
    (r"La table s[eé]lectionn[eé]e ne contient aucune correspondance active :\nil n'y a rien à vérifier\.",
     "The selected table has no active row:\nthere is nothing to check."),
    (r"La table s[eé]lectionn[eé]e ne contient aucune correspondance active\.",
     "The selected table has no active row."),
    (r"Document introuvable :", "Document not found:"),
    # detection_core
    (r"detection_config\.json introuvable :", "detection_config.json not found:"),
    (r"Ce PDF ne contient quasiment aucun texte : c'est vraisemblablement\nun document scanne\. "
     r"Il faut d'abord le passer par l'OCR\.",
     "This PDF contains almost no text: it is most likely\na scanned document. It must go through "
     "text recognition first."),
    (r"Ce PDF ne contient quasiment aucun texte : c'est un document scanne\.\nActivez l'OCR pour le lire\.",
     "This PDF contains almost no text: it is a scanned document.\nTurn on text recognition to read it."),
    (r"Aucune bibliotheque de lecture PDF n'a fonctionne\.", "No PDF reading library worked."),
    (r"Installez-en une :", "Install one:"),
    (r"Tesseract est introuvable sur ce poste\.\n\nSi vous l'avez installe ailleurs, indiquez son chemin "
     r"dans\ndetection_config\.json, rubrique « ocr », cle « tesseract »\.",
     "Tesseract cannot be found on this computer.\n\nIf it is installed elsewhere, enter its path "
     "in\ndetection_config.json, section \"ocr\", key \"tesseract\"."),
    (r"Le pack de langue « (.+?) » est introuvable\.", r'Language pack "\1" not found.'),
    (r"Tesseract n'installe pas le francais par defaut\. Deposez le fichier", "Tesseract does not install every language by default. Put the file"),
    (r"Tesseract a echoue sur la page (\d+)\.", r"Tesseract failed on page \1."),
    (r"Format non pris en charge : (\S+)\nAttendu : \.docx ou \.pdf", r"Unsupported format: \1\nExpected: .docx or .pdf"),
    # notes de la detection (classeur de relecture)
    (r"clé de contrôle incorrecte \(numéro erroné dans le document \?\)",
     "check digit is wrong (number mistyped in the document?)"),
    (r"mot répété, non reconnu : produit, personne, lieu \? à vous de dire",
     "repeated word, not recognised: product, person, place? Your call"),
    (r"forme « (.+?) » partagée avec « (.+?) »", r'spelling "\1" shared with "\2"'),
    (r"issu d'un OCR : vérifiez l'orthographe exacte avant injection",
     "read by text recognition: check the exact spelling before injecting"),
    (r"même partie que (\[[A-Z0-9_]+\]) : sera ajouté comme variante", r"same party as \1: will be added as a variant"),
    (r"société très citée mais absente de l'en-tête : autre partie probable, à vérifier",
     "company cited often but absent from the heading: probably another party, to be checked"),
    (r"rôle contractuel non trouvé : précisez la partie", "contractual role not found: specify the party"),
    (r"personne citée avec (\S+) : rôle à préciser", r"person named with \1: role to be specified"),
    (r"personne nommée avec le destinataire : rôle à préciser", "person named with the addressee: role to be specified"),
    (r"avocat attribué sans liste du cabinet \(detection_config\.json\) : à vérifier",
     "lawyer attributed without the firm's list (detection_config.json): to be checked"),
    (r"organisme audité déduit : c'est la société la plus citée", "audited entity inferred: the most cited company"),
    (r"personne citée dans l'audit : salarié présumé", "person named in the audit: presumably an employee"),
    (r"outil ou logiciel nommé : à confirmer", "tool or software named: to be confirmed"),
    (r"rattaché par proximité à (\S+)", r"attached to \1 by proximity"),
    # 15 octobre 2026 : balises neutres, dirigeants
    (r"balise neutre \(aucune partie certaine\) : rattachez-la par le menu Partie si vous le souhaitez",
     "neutral tag (no certain party): attach it through the Party menu if you wish"),
    (r"dirigeant de (\S+) d'après le texte : « (.*?)… »", r"director of \1 according to the text: « \2… »"),
    (r"présenté comme dirigeant de plusieurs sociétés \((.*?)\) : choisissez",
     r"presented as director of several companies (\1): your call"),
    (r"dirigeant de (.*?) d'après le texte \(société sans partie\) : choisissez la partie de la société",
     r"director of \1 according to the text (company without party): pick the company's party"),
    (r"présenté aussi comme dirigeant de (.*?) : à vérifier", r"also presented as director of \1: to check"),
    (r"rattaché à (\S+) \(titulaire / commercialise\)", r"attached to \1 (owner / seller)"),
    (r"déjà dans l'onglet", "already in the sheet"),
    (r"partie à choisir(?! \()", "party to be chosen"),
    (r"Sens du courrier inconnu : les parties n'ont pas été attribuées\.",
     "Letter direction unknown: the parties were not assigned."),
    (r"Type de document non reconnu : choisissez-le dans la fenêtre pour que les parties soient nommées\.",
     "Document type not recognised: choose it in the window so that the parties can be named."),
    # detection_classeur
    (r"Ce classeur n'est pas un classeur de détection \(onglet « Détection » absent\)\. / Not a detection workbook\.",
     "Not a detection workbook (no \"Detection\" sheet)."),
    (r"Ce classeur de détection a été produit par une ancienne version de l'outil\.\n\nRelancez la détection sur le document\.",
     "This detection workbook was produced by an older version of the program.\n\nRun the detection again on the document."),
    (r"Ce classeur de détection a été produit par l'ancienne version de l'outil \(colonnes Partie et Rôle\)\."
     r"\n\nRelancez la détection sur le document : le nouveau classeur propose la balise à choisir dans une liste\.",
     "This detection workbook was produced by an older version of the program.\n\nRun the detection "
     "again on the document: the new workbook offers the tag in a drop-down list."),
    # modele_onglet (table, injection)
    (r"Le classeur de correspondances est ouvert \(dans Excel \?\)\.\n\nFermez-le, puis recommencez : "
     r"rien n'a été modifié\.",
     "The correspondence workbook is open (in Excel?).\n\nClose it, then try again: nothing was changed."),
    (r"Le nom de l'onglet est vide\.", "The sheet name is empty."),
    (r"Excel limite le nom d'un onglet à 31 caractères\.", "Excel limits a sheet name to 31 characters."),
    (r"Le nom d'un onglet ne peut contenir aucun de ces signes :", "A sheet name cannot contain any of these characters:"),
    (r"Un onglet « (.+?) » existe déjà dans ce classeur\.", r'A sheet "\1" already exists in this workbook.'),
    (r"L'onglet « (.+?) » suit l'ANCIEN modèle \(cases numérotées\)\.[\s\S]*",
     r'Sheet "\1" follows the OLD layout (numbered boxes).\n\nInjection only works in a per-party '
     r'sheet: create a new matter ("New matter…" button).'),
    (r"déjà présent dans l'onglet sous (\[[A-Z0-9_]+\]) \((.*?) / (.*?)\) : une balise existante n'est jamais déplacée",
     r"already in the sheet as \1 (\2 / \3): an existing tag is never moved"),
    (r"déjà présent dans l'onglet \((\[[A-Z0-9_]+\])\)", r"already in the sheet (\1)"),
    (r"partie à choisir \(colonne Partie du classeur\)", "party to be chosen (Party column of the workbook)"),
    (r"rôle à choisir \(colonne Rôle du classeur\)", "role to be chosen (Role column of the workbook)"),
    (r"partie non renseignée", "party missing"),
    (r"un élément COMMUN doit avoir un rôle autre que « Identité » \(par exemple Référence RG ou Montant\)",
     "a SHARED item needs a role other than \"Identity\" (for instance Case number or Amount)"),
]
_TR_COMPILES = None


def traduire(message, langue=None):
    """Message d'un moteur, traduit en anglais si la version est anglaise."""
    global _TR_COMPILES
    if (langue or LANGUE) != "en" or not isinstance(message, str) or not message:
        return message
    if _TR_COMPILES is None:
        import re
        _TR_COMPILES = [(re.compile(a), b) for a, b in _TR]
    for motif, rempl in _TR_COMPILES:
        message = motif.sub(rempl, message)
    return message


def T(cle):
    table = EN if LANGUE == "en" else FR
    return table.get(cle, FR.get(cle, cle))
