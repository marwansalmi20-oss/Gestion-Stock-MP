import os
import re
import sys
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# Les données sont enregistrées à côté du programme (ou du fichier .py).
# Chaque PC aura donc son propre fichier Gestion_Stock_MP.xlsx.
if getattr(sys, "frozen", False):
    DOSSIER_APP = Path(sys.executable).resolve().parent
else:
    DOSSIER_APP = Path(__file__).resolve().parent

FICHIER_EXCEL = DOSSIER_APP / "Gestion_Stock_MP.xlsx"

ENTETES = {
    "Articles": ["Code MP", "Désignation", "Unité", "Stock initial", "Stock minimum"],
    "Mouvements": ["Date", "Type", "Code MP", "Désignation", "Quantité",
                   "Référence / Origine", "Observation"],
    "Stock": ["Code MP", "Désignation", "Unité", "Stock initial", "Total entrées",
              "Total sorties", "Stock actuel", "Stock minimum", "État"],
}


def nombre(valeur, defaut=0.0):
    if valeur is None or valeur == "":
        return defaut
    if isinstance(valeur, (int, float)):
        return float(valeur)
    try:
        return float(str(valeur).replace(",", ".").strip())
    except (ValueError, TypeError):
        return defaut


def vers_entier_si_possible(x):
    return int(x) if float(x).is_integer() else x


def est_en_alerte(stock, minimum):
    return minimum > 0 and stock <= minimum


# Import de l'ancien fichier de stock
PREFIXES_FEUILLES = [
    ("REACTIF", "REA"), ("CULTURE", "MC"), ("M,S", "MS"),
    ("PC", "PC"), ("PARA", "PA"), ("(PO)", "PO"),
]


def _norm(texte):
    if texte is None:
        return ""
    t = str(texte).lower()
    t = (t.replace("é", "e").replace("è", "e").replace("ê", "e")
           .replace("à", "a").replace("ù", "u").replace("ç", "c"))
    return re.sub(r"[^a-z0-9]", "", t)


def _date_texte(valeur):
    if valeur is None:
        return ""
    if hasattr(valeur, "strftime"):
        return valeur.strftime("%d/%m/%Y")
    return str(valeur).strip()


def _prefixe(nom_feuille):
    nom = nom_feuille.upper()
    for mot, prefixe in PREFIXES_FEUILLES:
        if mot in nom:
            return prefixe
    return "ART"


def _colonnes(ligne_entete):
    cols = {}
    for j, titre in enumerate(ligne_entete):
        n = _norm(titre)
        if not n:
            continue
        if n.startswith("article"):
            cols["article"] = j
        elif n.startswith("capacite") or n == "cond":
            cols["cap"] = j
        elif n.startswith("date"):
            cols["date"] = j
        elif n.startswith("factu"):
            cols["fact"] = j
        elif n.startswith("inv"):
            cols["inv"] = j
        elif n.startswith("entree"):
            cols["in"] = j
        elif n.startswith("sortie"):
            cols["out"] = j
        elif "final" in n:
            cols["final"] = j
        elif n.startswith("bl"):
            cols["bl"] = j
        elif n.startswith("fourn"):
            cols["four"] = j
    return cols


def lire_fichier_stock(chemin):
    classeur = load_workbook(chemin, data_only=True)
    articles, mouvements = [], []
    nb_sans_stock = 0

    for feuille in classeur.worksheets:
        lignes = list(feuille.iter_rows(
            min_row=1, max_row=min(feuille.max_row, 20000),
            max_col=16, values_only=True
        ))
        idx_entete = next(
            (i for i, l in enumerate(lignes[:15])
             if l and _norm(l[0]) == "article"), None
        )
        if idx_entete is None:
            continue

        cols = _colonnes(lignes[idx_entete])
        if "article" not in cols:
            continue
        prefixe = _prefixe(feuille.title)

        def val(ligne, cle):
            j = cols.get(cle)
            return ligne[j] if j is not None and j < len(ligne) else None

        liste = []
        courant = None
        for ligne in lignes[idx_entete + 1:]:
            if all(v is None for v in ligne):
                continue
            nom = val(ligne, "article")
            if nom is not None and str(nom).strip():
                courant = {
                    "nom": re.sub(r"\s+", " ", str(nom)).strip(),
                    "cap": str(val(ligne, "cap") or "").strip(),
                    "inv": 0.0,
                    "date": _date_texte(val(ligne, "date")),
                    "mvts": [],
                    "final": None,
                }
                liste.append(courant)
            if courant is None:
                continue

            inv = nombre(val(ligne, "inv"))
            if inv:
                courant["inv"] += inv
            entree = nombre(val(ligne, "in"))
            sortie = nombre(val(ligne, "out"))
            if entree or sortie:
                courant["mvts"].append({
                    "date": _date_texte(val(ligne, "date")),
                    "in": entree, "out": sortie,
                    "fact": str(val(ligne, "fact") or "").strip(),
                    "bl": str(val(ligne, "bl") or "").strip(),
                    "four": str(val(ligne, "four") or "").strip(),
                })
            stock_final = val(ligne, "final")
            if stock_final is not None and str(stock_final).strip() != "":
                courant["final"] = nombre(stock_final)

        compteur = 0
        for a in liste:
            net = sum(m["in"] - m["out"] for m in a["mvts"])
            # On conserve aussi les articles dont le stock final est exactement zéro.
            a_du_stock = bool(a["inv"] or a["mvts"] or a["final"] is not None)
            if not a_du_stock:
                nb_sans_stock += 1
                continue

            compteur += 1
            code = f"{prefixe}{compteur:04d}"
            designation = a["nom"]
            if a["cap"] and a["cap"].lower() not in designation.lower():
                designation += f" ({a['cap']})"

            initial = a["final"] - net if a["final"] is not None else a["inv"]
            articles.append([
                code, designation, "pièce",
                vers_entier_si_possible(initial), 0
            ])

            for m in a["mvts"]:
                reference = " / ".join(x for x in (
                    ("FAC " + m["fact"]) if m["fact"] else "",
                    ("BL " + m["bl"]) if m["bl"] else ""
                ) if x)
                if m["in"]:
                    mouvements.append([
                        m["date"] or a["date"], "Entrée", code, designation,
                        vers_entier_si_possible(m["in"]), reference, m["four"]
                    ])
                if m["out"]:
                    mouvements.append([
                        m["date"] or a["date"], "Sortie", code, designation,
                        vers_entier_si_possible(m["out"]), reference, m["four"]
                    ])

    return articles, mouvements, nb_sans_stock


def creer_fichier_excel_si_absent():
    if FICHIER_EXCEL.exists():
        classeur = load_workbook(FICHIER_EXCEL)
    else:
        classeur = Workbook()
        classeur.remove(classeur.active)

    for nom, colonnes in ENTETES.items():
        feuille = classeur[nom] if nom in classeur.sheetnames else classeur.create_sheet(nom)
        if feuille.max_row == 1 and all(
            feuille.cell(1, c).value is None for c in range(1, len(colonnes) + 1)
        ):
            for numero, titre in enumerate(colonnes, start=1):
                feuille.cell(1, numero, titre)

        for cellule in feuille[1]:
            cellule.font = Font(bold=True, color="FFFFFF")
            cellule.fill = PatternFill("solid", fgColor="244062")
            cellule.alignment = Alignment(horizontal="center")
        feuille.freeze_panes = "A2"

    classeur.save(FICHIER_EXCEL)


def lire_donnees(classeur):
    articles = {}
    for ligne in classeur["Articles"].iter_rows(min_row=2, values_only=True):
        if not ligne or ligne[0] is None:
            continue
        code = str(ligne[0]).strip()
        if not code:
            continue
        articles[code] = {
            "code": code,
            "designation": str(ligne[1] or ""),
            "unite": str(ligne[2] or "pièce"),
            "initial": nombre(ligne[3]),
            "minimum": nombre(ligne[4]),
        }

    mouvements = []
    totaux = {code: {"entrees": 0.0, "sorties": 0.0} for code in articles}
    for ligne in classeur["Mouvements"].iter_rows(min_row=2, values_only=True):
        if not ligne or ligne[0] is None:
            continue
        mouvement = {
            "date": ligne[0].strftime("%d/%m/%Y %H:%M:%S")
                    if hasattr(ligne[0], "strftime") else str(ligne[0]),
            "type": str(ligne[1] or ""),
            "code": str(ligne[2] or ""),
            "designation": str(ligne[3] or ""),
            "quantite": nombre(ligne[4]),
            "reference": str(ligne[5] or ""),
            "observation": str(ligne[6] or ""),
        }
        mouvements.append(mouvement)
        code = mouvement["code"]
        if code in totaux:
            if mouvement["type"] == "Entrée":
                totaux[code]["entrees"] += mouvement["quantite"]
            elif mouvement["type"] == "Sortie":
                totaux[code]["sorties"] += mouvement["quantite"]

    return articles, mouvements, totaux


def enregistrer_classeur(classeur):
    try:
        classeur.save(FICHIER_EXCEL)
        return True
    except PermissionError:
        messagebox.showerror(
            "Fichier Excel ouvert",
            "Impossible d'enregistrer le fichier.\n\n"
            "Ferme Gestion_Stock_MP.xlsx dans Excel, puis réessaie."
        )
        return False
    except Exception as erreur:
        messagebox.showerror("Erreur Excel", f"L'enregistrement a échoué :\n{erreur}")
        return False


class ApplicationStock:
    def __init__(self, fenetre):
        self.fenetre = fenetre
        self.fenetre.title("Gestion de Stock - Matières premières")
        self.fenetre.geometry("1180x720")
        self.fenetre.minsize(950, 600)
        self.fenetre.configure(bg="#F3F6FA")

        self.articles, self.mouvements, self.totaux, self.stock = {}, [], {}, {}
        self.var_recherche = tk.StringVar()

        self.creer_interface()
        self.var_recherche.trace_add("write", lambda *_: self.afficher_tableaux())
        self.actualiser_donnees()

    def creer_interface(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=27, font=("Arial", 9))
        style.configure("Treeview.Heading", font=("Arial", 9, "bold"))
        style.configure("TButton", font=("Arial", 10), padding=7)
        style.configure("TLabel", font=("Arial", 10))

        haut = tk.Frame(self.fenetre, bg="#17365D", padx=20, pady=16)
        haut.pack(fill="x")
        tk.Label(
            haut, text="GESTION DE STOCK", font=("Arial", 21, "bold"),
            fg="white", bg="#17365D"
        ).pack(anchor="w")
        tk.Label(
            haut, text="Matières premières (MP)  |  Suivi automatique avec Excel",
            font=("Arial", 10), fg="#DCE6F1", bg="#17365D"
        ).pack(anchor="w", pady=(4, 0))

        zone_boutons = tk.Frame(self.fenetre, bg="#F3F6FA", padx=15, pady=12)
        zone_boutons.pack(fill="x")
        ttk.Button(zone_boutons, text="+ Ajouter une MP",
                   command=self.ajouter_article).pack(side="left", padx=4)
        ttk.Button(zone_boutons, text="Entrée / Sortie",
                   command=self.ajouter_mouvement).pack(side="left", padx=4)
        ttk.Button(zone_boutons, text="Actualiser",
                   command=self.actualiser_donnees).pack(side="left", padx=4)
        ttk.Button(zone_boutons, text="Importer N_STOCK_2026.xlsx",
                   command=self.importer_ancien_fichier).pack(side="left", padx=4)
        ttk.Button(zone_boutons, text="Ouvrir le fichier Excel",
                   command=self.ouvrir_excel).pack(side="right", padx=4)

        zone_stats = tk.Frame(self.fenetre, bg="#F3F6FA", padx=15)
        zone_stats.pack(fill="x", pady=(0, 10))
        self.lbl_articles = self.creer_carte(zone_stats, "Références MP", "0", 0)
        self.lbl_alertes = self.creer_carte(zone_stats, "Alertes stock", "0", 1)
        self.lbl_mouvements = self.creer_carte(zone_stats, "Mouvements", "0", 2)

        zone_recherche = tk.Frame(self.fenetre, bg="#F3F6FA", padx=15)
        zone_recherche.pack(fill="x", pady=(0, 6))
        tk.Label(
            zone_recherche, text="Rechercher (code ou désignation) :",
            bg="#F3F6FA", fg="#263445", font=("Arial", 10)
        ).pack(side="left")
        ttk.Entry(zone_recherche, textvariable=self.var_recherche, width=45).pack(
            side="left", padx=8
        )

        notebook = ttk.Notebook(self.fenetre)
        notebook.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        onglet_stock = ttk.Frame(notebook, padding=8)
        onglet_historique = ttk.Frame(notebook, padding=8)
        notebook.add(onglet_stock, text="État de stock")
        notebook.add(onglet_historique, text="Historique des mouvements")

        colonnes_stock = (
            "code", "designation", "unite", "initial", "entrees",
            "sorties", "stock", "minimum", "etat"
        )
        self.table_stock = ttk.Treeview(
            onglet_stock, columns=colonnes_stock, show="headings"
        )
        titres_stock = {
            "code": "Code MP", "designation": "Désignation", "unite": "Unité",
            "initial": "Stock initial", "entrees": "Entrées", "sorties": "Sorties",
            "stock": "Stock actuel", "minimum": "Stock minimum", "etat": "État"
        }
        largeurs_stock = {
            "code": 95, "designation": 320, "unite": 70, "initial": 100,
            "entrees": 90, "sorties": 90, "stock": 100, "minimum": 110, "etat": 100
        }
        for col in colonnes_stock:
            self.table_stock.heading(col, text=titres_stock[col])
            self.table_stock.column(col, width=largeurs_stock[col], anchor="center")
        self.table_stock.column("designation", anchor="w")
        self.table_stock.tag_configure("alerte", background="#FCE4D6")
        self.table_stock.tag_configure("normal", background="#E2F0D9")
        self.ajouter_barre_defilement(onglet_stock, self.table_stock)

        colonnes_hist = (
            "date", "type", "code", "designation", "quantite", "reference", "observation"
        )
        self.table_historique = ttk.Treeview(
            onglet_historique, columns=colonnes_hist, show="headings"
        )
        titres_hist = {
            "date": "Date", "type": "Type", "code": "Code MP",
            "designation": "Désignation", "quantite": "Quantité",
            "reference": "Référence / Origine", "observation": "Observation"
        }
        largeurs_hist = {
            "date": 145, "type": 80, "code": 90, "designation": 280,
            "quantite": 90, "reference": 170, "observation": 220
        }
        for col in colonnes_hist:
            self.table_historique.heading(col, text=titres_hist[col])
            self.table_historique.column(col, width=largeurs_hist[col], anchor="center")
        self.table_historique.column("designation", anchor="w")
        self.table_historique.column("observation", anchor="w")
        self.ajouter_barre_defilement(onglet_historique, self.table_historique)

        tk.Label(
            self.fenetre, text=f"Fichier Excel : {FICHIER_EXCEL}",
            bg="#F3F6FA", fg="#526579", font=("Arial", 9)
        ).pack(anchor="w", padx=18, pady=(0, 7))

    def creer_carte(self, parent, titre, valeur, colonne):
        carte = tk.Frame(
            parent, bg="white", padx=18, pady=10,
            highlightbackground="#D7E0EA", highlightthickness=1
        )
        carte.grid(row=0, column=colonne, sticky="ew", padx=5)
        parent.grid_columnconfigure(colonne, weight=1)
        tk.Label(carte, text=titre, bg="white", fg="#526579",
                 font=("Arial", 9)).pack(anchor="w")
        etiquette = tk.Label(
            carte, text=valeur, bg="white", fg="#17365D",
            font=("Arial", 19, "bold")
        )
        etiquette.pack(anchor="w", pady=(3, 0))
        return etiquette

    def ajouter_barre_defilement(self, parent, tableau):
        cadre = tk.Frame(parent)
        cadre.pack(fill="both", expand=True)
        tableau.pack(in_=cadre, side="left", fill="both", expand=True)
        barre_y = ttk.Scrollbar(cadre, orient="vertical", command=tableau.yview)
        barre_y.pack(side="right", fill="y")
        barre_x = ttk.Scrollbar(parent, orient="horizontal", command=tableau.xview)
        barre_x.pack(fill="x")
        tableau.configure(yscrollcommand=barre_y.set, xscrollcommand=barre_x.set)

    def calculer_stock(self):
        self.stock = {}
        for code, article in self.articles.items():
            mouvements = self.totaux.get(code, {"entrees": 0.0, "sorties": 0.0})
            quantite = article["initial"] + mouvements["entrees"] - mouvements["sorties"]
            self.stock[code] = {
                **article, "entrees": mouvements["entrees"],
                "sorties": mouvements["sorties"], "stock": quantite
            }

    def mettre_a_jour_feuille_stock(self, classeur):
        feuille = classeur["Stock"]
        if feuille.max_row > 1:
            feuille.delete_rows(2, feuille.max_row - 1)

        for code in sorted(self.stock):
            a = self.stock[code]
            etat = "À commander" if est_en_alerte(a["stock"], a["minimum"]) else "Normal"
            feuille.append([
                code, a["designation"], a["unite"], a["initial"],
                a["entrees"], a["sorties"], a["stock"], a["minimum"], etat
            ])

        largeurs = [14, 50, 12, 15, 15, 15, 15, 16, 18]
        for i, largeur in enumerate(largeurs, start=1):
            feuille.column_dimensions[get_column_letter(i)].width = largeur
        feuille.auto_filter.ref = feuille.dimensions
        for ligne in feuille.iter_rows(min_row=2):
            for cellule in ligne:
                cellule.alignment = Alignment(vertical="center")
            if ligne[8].value == "À commander":
                for cellule in ligne:
                    cellule.fill = PatternFill("solid", fgColor="FCE4D6")

    def actualiser_donnees(self):
        try:
            creer_fichier_excel_si_absent()
            classeur = load_workbook(FICHIER_EXCEL)
            self.articles, self.mouvements, self.totaux = lire_donnees(classeur)
            self.calculer_stock()
            self.mettre_a_jour_feuille_stock(classeur)
            if not enregistrer_classeur(classeur):
                return
            self.afficher_tableaux()
        except Exception as erreur:
            messagebox.showerror(
                "Erreur", f"Impossible de lire le fichier Excel :\n{erreur}"
            )

    def afficher_tableaux(self):
        if not hasattr(self, "table_stock"):
            return
        texte = self.var_recherche.get().strip().lower()

        for item in self.table_stock.get_children():
            self.table_stock.delete(item)
        for code in sorted(self.stock):
            a = self.stock[code]
            if texte and texte not in code.lower() and texte not in a["designation"].lower():
                continue
            en_alerte = est_en_alerte(a["stock"], a["minimum"])
            etat = "À commander" if en_alerte else "Normal"
            tag = "alerte" if en_alerte else "normal"
            self.table_stock.insert("", "end", values=(
                code, a["designation"], a["unite"],
                self.formater_nombre(a["initial"]),
                self.formater_nombre(a["entrees"]),
                self.formater_nombre(a["sorties"]),
                self.formater_nombre(a["stock"]),
                self.formater_nombre(a["minimum"]), etat
            ), tags=(tag,))

        for item in self.table_historique.get_children():
            self.table_historique.delete(item)
        for m in reversed(self.mouvements):
            if texte and texte not in m["code"].lower() and texte not in m["designation"].lower():
                continue
            self.table_historique.insert("", "end", values=(
                m["date"], m["type"], m["code"], m["designation"],
                self.formater_nombre(m["quantite"]), m["reference"], m["observation"]
            ))

        nb_alertes = sum(
            1 for a in self.stock.values()
            if est_en_alerte(a["stock"], a["minimum"])
        )
        self.lbl_articles.config(text=str(len(self.articles)))
        self.lbl_alertes.config(text=str(nb_alertes))
        self.lbl_mouvements.config(text=str(len(self.mouvements)))

    @staticmethod
    def formater_nombre(valeur):
        return f"{valeur:g}"

    @staticmethod
    def champ(parent, texte, ligne, valeur=""):
        tk.Label(
            parent, text=texte, bg="white", fg="#263445",
            font=("Arial", 10)
        ).grid(row=ligne, column=0, sticky="w", padx=12, pady=7)
        entree = ttk.Entry(parent, width=34)
        entree.grid(row=ligne, column=1, sticky="ew", padx=12, pady=7)
        if valeur != "":
            entree.insert(0, str(valeur))
        return entree

    def importer_ancien_fichier(self):
        chemin = filedialog.askopenfilename(
            title="Choisir le fichier N_STOCK_2026.xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")]
        )
        if not chemin:
            return

        self.fenetre.config(cursor="watch")
        self.fenetre.update()
        try:
            articles, mouvements, nb_sans_stock = lire_fichier_stock(chemin)
        except Exception as erreur:
            self.fenetre.config(cursor="")
            messagebox.showerror(
                "Import impossible", f"Impossible de lire ce fichier :\n{erreur}"
            )
            return
        finally:
            self.fenetre.config(cursor="")

        if not articles:
            messagebox.showinfo(
                "Rien à importer",
                "Aucun article avec un stock, un inventaire ou un mouvement n'a été trouvé."
            )
            return

        if self.articles or self.mouvements:
            confirmer = messagebox.askyesno(
                "Remplacer les données ?",
                f"Le programme contient déjà {len(self.articles)} article(s) et "
                f"{len(self.mouvements)} mouvement(s).\n\n"
                "L'import va REMPLACER ces données.\n"
                "Une sauvegarde du fichier actuel sera créée avant l'import.\n\n"
                "Continuer ?"
            )
            if not confirmer:
                return

        try:
            creer_fichier_excel_si_absent()
            # Sauvegarde de sécurité avant de remplacer les données existantes.
            if FICHIER_EXCEL.exists():
                sauvegarde = DOSSIER_APP / (
                    "Sauvegarde_Gestion_Stock_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
                )
                import shutil
                shutil.copy2(FICHIER_EXCEL, sauvegarde)

            classeur = load_workbook(FICHIER_EXCEL)
            for nom in ("Articles", "Mouvements"):
                feuille = classeur[nom]
                if feuille.max_row > 1:
                    feuille.delete_rows(2, feuille.max_row - 1)
            for ligne in articles:
                classeur["Articles"].append(ligne)
            for ligne in mouvements:
                classeur["Mouvements"].append(ligne)
            if not enregistrer_classeur(classeur):
                return
        except Exception as erreur:
            messagebox.showerror("Erreur", f"L'import a échoué :\n{erreur}")
            return

        self.actualiser_donnees()
        messagebox.showinfo(
            "Import terminé",
            f"{len(articles)} article(s) et {len(mouvements)} mouvement(s) importés.\n"
            f"{nb_sans_stock} article(s) du catalogue sans quantité ont été ignorés."
        )

    def ajouter_article(self):
        fenetre = tk.Toplevel(self.fenetre)
        fenetre.title("Ajouter une matière première")
        fenetre.geometry("470x350")
        fenetre.resizable(False, False)
        fenetre.configure(bg="white")
        fenetre.transient(self.fenetre)
        fenetre.grab_set()

        tk.Label(
            fenetre, text="NOUVELLE MATIÈRE PREMIÈRE",
            font=("Arial", 14, "bold"), bg="white", fg="#17365D"
        ).grid(row=0, column=0, columnspan=2, pady=16)

        code = self.champ(fenetre, "Code MP :", 1)
        designation = self.champ(fenetre, "Désignation :", 2)
        unite = self.champ(fenetre, "Unité (kg, L, pièce...) :", 3, "pièce")
        initial = self.champ(fenetre, "Stock initial :", 4, "0")
        minimum = self.champ(fenetre, "Stock minimum :", 5, "0")

        def valider():
            c, d = code.get().strip(), designation.get().strip()
            u = unite.get().strip() or "pièce"
            if not c or not d:
                messagebox.showwarning(
                    "Champs obligatoires",
                    "Le code et la désignation sont obligatoires.", parent=fenetre
                )
                return
            if c in self.articles:
                messagebox.showwarning(
                    "Code existant", "Ce code MP existe déjà.", parent=fenetre
                )
                return
            try:
                q_initiale = float(initial.get().replace(",", "."))
                q_min = float(minimum.get().replace(",", "."))
                if q_initiale < 0 or q_min < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(
                    "Quantité invalide",
                    "Saisis des nombres positifs ou égaux à zéro.", parent=fenetre
                )
                return

            try:
                classeur = load_workbook(FICHIER_EXCEL)
                classeur["Articles"].append([c, d, u, q_initiale, q_min])
                if enregistrer_classeur(classeur):
                    fenetre.destroy()
                    self.actualiser_donnees()
                    messagebox.showinfo("Succès", "La matière première a été ajoutée.")
            except Exception as erreur:
                messagebox.showerror("Erreur", str(erreur), parent=fenetre)

        ttk.Button(fenetre, text="Enregistrer", command=valider).grid(
            row=6, column=0, columnspan=2, pady=16
        )

    def ajouter_mouvement(self):
        if not self.articles:
            messagebox.showinfo(
                "Aucune MP",
                "Ajoute d'abord une matière première (ou importe ton fichier)."
            )
            return

        fenetre = tk.Toplevel(self.fenetre)
        fenetre.title("Enregistrer une entrée ou une sortie")
        fenetre.geometry("560x440")
        fenetre.resizable(False, False)
        fenetre.configure(bg="white")
        fenetre.transient(self.fenetre)
        fenetre.grab_set()

        tk.Label(
            fenetre, text="NOUVEAU MOUVEMENT DE STOCK",
            font=("Arial", 14, "bold"), bg="white", fg="#17365D"
        ).grid(row=0, column=0, columnspan=2, pady=16)

        tk.Label(fenetre, text="Type :", bg="white").grid(
            row=1, column=0, sticky="w", padx=12, pady=7
        )
        type_var = tk.StringVar(value="Entrée")
        ttk.Combobox(
            fenetre, textvariable=type_var, values=["Entrée", "Sortie"],
            state="readonly", width=40
        ).grid(row=1, column=1, padx=12, pady=7, sticky="ew")

        tous_les_choix = [
            f"{code} — {a['designation']}"
            for code, a in sorted(self.articles.items())
        ]
        tk.Label(fenetre, text="Rechercher MP :", bg="white").grid(
            row=2, column=0, sticky="w", padx=12, pady=7
        )
        recherche_var = tk.StringVar()
        ttk.Entry(fenetre, textvariable=recherche_var, width=43).grid(
            row=2, column=1, padx=12, pady=7, sticky="ew"
        )

        tk.Label(fenetre, text="Matière première :", bg="white").grid(
            row=3, column=0, sticky="w", padx=12, pady=7
        )
        mp_var = tk.StringVar(value=tous_les_choix[0])
        combo = ttk.Combobox(
            fenetre, textvariable=mp_var, values=tous_les_choix,
            state="readonly", width=40
        )
        combo.grid(row=3, column=1, padx=12, pady=7, sticky="ew")

        def filtrer(*_):
            t = recherche_var.get().strip().lower()
            valeurs = [c for c in tous_les_choix if t in c.lower()]
            combo["values"] = valeurs
            mp_var.set(valeurs[0] if valeurs else "")

        recherche_var.trace_add("write", filtrer)

        quantite = self.champ(fenetre, "Quantité :", 4)
        reference = self.champ(fenetre, "Référence / Origine :", 5)
        observation = self.champ(fenetre, "Observation :", 6)

        def valider():
            if not mp_var.get():
                messagebox.showwarning(
                    "Matière première", "Choisis une matière première.", parent=fenetre
                )
                return
            try:
                q = float(quantite.get().replace(",", "."))
                if q <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(
                    "Quantité invalide",
                    "La quantité doit être supérieure à zéro.", parent=fenetre
                )
                return

            code_mp = mp_var.get().split(" — ", 1)[0]
            article = self.articles.get(code_mp)
            if not article:
                messagebox.showerror(
                    "Erreur", "La matière première sélectionnée est introuvable.",
                    parent=fenetre
                )
                return

            type_mouvement = type_var.get()
            # Recharger les données avant validation pour éviter un stock périmé.
            try:
                classeur = load_workbook(FICHIER_EXCEL)
                articles_actuels, mouvements_actuels, totaux_actuels = lire_donnees(classeur)
                if code_mp not in articles_actuels:
                    messagebox.showerror(
                        "Erreur", "Cette matière première n'existe plus dans le fichier.",
                        parent=fenetre
                    )
                    return
                art_actuel = articles_actuels[code_mp]
                t = totaux_actuels.get(code_mp, {"entrees": 0.0, "sorties": 0.0})
                stock_actuel = art_actuel["initial"] + t["entrees"] - t["sorties"]
                if type_mouvement == "Sortie" and q > stock_actuel:
                    messagebox.showwarning(
                        "Stock insuffisant",
                        f"Stock disponible : {self.formater_nombre(stock_actuel)} "
                        f"{art_actuel['unite']}.", parent=fenetre
                    )
                    return

                classeur["Mouvements"].append([
                    datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    type_mouvement, code_mp, art_actuel["designation"], q,
                    reference.get().strip(), observation.get().strip()
                ])
                if enregistrer_classeur(classeur):
                    fenetre.destroy()
                    self.actualiser_donnees()
                    messagebox.showinfo(
                        "Succès",
                        f"{type_mouvement} enregistrée. Le stock a été recalculé."
                    )
            except Exception as erreur:
                messagebox.showerror(
                    "Erreur", f"Impossible d'enregistrer le mouvement :\n{erreur}",
                    parent=fenetre
                )

        ttk.Button(
            fenetre, text="Enregistrer le mouvement", command=valider
        ).grid(row=7, column=0, columnspan=2, pady=17)

    def ouvrir_excel(self):
        try:
            if not FICHIER_EXCEL.exists():
                creer_fichier_excel_si_absent()
            if sys.platform.startswith("win"):
                os.startfile(str(FICHIER_EXCEL))
            elif sys.platform == "darwin":
                os.system(f'open "{FICHIER_EXCEL}"')
            else:
                os.system(f'xdg-open "{FICHIER_EXCEL}"')
        except Exception as erreur:
            messagebox.showerror(
                "Ouverture impossible", f"Impossible d'ouvrir Excel :\n{erreur}"
            )


if __name__ == "__main__":
    fenetre_principale = tk.Tk()
    try:
        creer_fichier_excel_si_absent()
        ApplicationStock(fenetre_principale)
    except Exception as erreur:
        messagebox.showerror("Erreur au démarrage", str(erreur))
        fenetre_principale.destroy()
        raise SystemExit(1)
    fenetre_principale.mainloop()
