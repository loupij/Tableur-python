try:
    import tkinter as tk
    from tkinter import font, messagebox, filedialog, colorchooser
    # import sqlite3
    import pandas as pd
    import logging
    from PIL import Image, ImageTk
    import random
    import math
    import datetime
    import re
    import traceback
    import sympy as sp
    import platform
    import os
    import psutil
    import subprocess
except ModuleNotFoundError:
    print(f"Veuillez vérifier que tout les modules sont bien installés.\nVous pouvez aussi faire la commande suivante : pip install -r requirements.txt")
    exit()

with open("requirements.txt") as requirements:
    requirements = requirements.readlines()
    req = []
    for i in range(len(requirements)):
        module = requirements[i].replace("\n", "")
        req.append(module)

# paramètres basiques du tableur
NB_LIGNES = 10
NB_COLONNES = 20
COULEUR_EVIDENCE = "yellow"
LOGGING_ENABLED = True
VARIABLE_FONCTION = sp.symbols("x")
POLICE = "Arial" # à refaire
TAILLE_TABLE = NB_LIGNES * NB_COLONNES

# informations logicielles
VERSION = "2.2.2"
REQUIREMENTS = req

# logs
LOG = logging.getLogger()
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("logs.log", encoding="utf-8"),  # Spécifiez l'encodage ici
        logging.StreamHandler()
    ]
)

class Tableur():
    def __init__(self, root):
        """
        Permet la création de l'interface et des données nécéssaire au fonctionnement du tableur.
        """
        self.root = root
        self.root.title("Tableur")
        self.dataframe = pd.DataFrame()

        # barre de formule
        logMessage("[__init__] Création de la barre de formule.")
        self.formulabar = tk.Frame(self.root, bg="lightgray", height=40)
        self.formulabar.pack(fill=tk.X, side=tk.TOP)

        self.cell_content_label = tk.Label(self.formulabar, text="f", bg="lightgray")
        self.cell_content_label.pack(side=tk.LEFT, padx=10)

        self.cell_content_entry = tk.Entry(self.formulabar, width=NB_COLONNES*10)
        self.cell_content_entry.pack(side=tk.LEFT, padx=5)

        # création de la table
        logMessage("[__init__] Création de la table et des cellules.")
        self.table = tk.Frame(self.root)
        self.table.pack(fill=tk.BOTH, expand=True)

        self.cellules = {} # valeurs des cellules
        self.cellules_raw = {} # à mettre d'abord dans create table puis adapter le code
        self.creer_table(NB_COLONNES, NB_LIGNES)
        self.current_cell = (0, 0)

        # barre de menu
        logMessage("[__init__] Création de la barre de menu.")
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)

        self.file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Fichier", menu=self.file_menu)
        self.file_menu.add_command(label="Ouvrir", command=self.ouvrir)
        self.file_menu.add_command(label="Enregistrer", command=self.enregistrer)
        self.save_menu = tk.Menu(self.file_menu, tearoff=0)
        self.file_menu.add_cascade(label="Enregistrer sous", menu=self.save_menu)
        self.save_menu.add_command(label="Enregistrer en CSV", command=self.enregistrer_csv)
        self.save_menu.add_command(label="Enregistrer en XLSX", command=self.enregistrer_excel)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Paramètres", command=self.afficher_parametres)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Fermer", command=self.exit_program)

        self.edit_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Edition", menu=self.edit_menu)
        self.edit_menu.add_command(label="Gras", command=self.gras)
        self.edit_menu.add_command(label="Italique", command=self.italique)
        self.edit_menu.add_command(label="Souligner", command=self.souligner)
        self.edit_menu.add_command(label="Remplissage", command=self.remplissage)
        self.edit_menu.add_command(label="Couleur Police", command=self.couleur_police)
        self.edit_menu.add_command(label="Police", command=self.changer_police)

        self.tools_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Outils", menu=self.tools_menu)
        self.tools_menu.add_command(label="Info cellule", command=self.cell_info)
        self.tools_menu.add_cascade(label="Afficher cellules", command=self.show_cells)
        self.tools_menu.add_cascade(label="Nettoyer", command=self.clear)
        self.help_menu = tk.Menu(self.tools_menu, tearoff=0)
        self.tools_menu.add_cascade(label="Aide", menu=self.help_menu)
        self.help_menu.add_command(label="Formules", command=self.aide_formules)

        # touches
        logMessage("[__init__] Création des touches.")
        self.root.bind("<Key>", self.key_handler)

        # format de celulles pour les formules et autres
        self.formats = {
            # formules
            "SOMME"             : r"^SOMME\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "SOMMEPROD"         : r"^SOMMEPROD\(\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*;\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*\)$",
            "MOYENNE"           : r"^MOYENNE\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "MOYENNE.POND"      : r"^MOYENNE\.POND\(\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*;\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*\)$",
            "RANDINT"           : r"^RANDINT\([A-Z][0-9]+,[A-Z][0-9]+\)$",
            "EXP"               : r"^EXP\([A-Z][0-9]|\d+\)$",
            "ABS"               : r"^ABS\([A-Z][0-9]|\d+\)$",
            "BINOME"            : r"^BINOME\([A-Z][0-9]|\d;[A-Z][0-9]|\d;[A-Z][0-9]|\d\)$",
            "PROD.SCAL"         : r"^PROD\.SCAL\([A-Z][0-9]|\d;[A-Z][0-9]|\d;[A-Z][0-9]|\d\)$",
            "RACINE"            : r"^RACINE\([A-Z][0-9]+\)$",
            "RACINE.DEG"        : r"^RACINE\.DEG\([A-Z][0-9]|\d+;\d\)$",
            "MIN"               : r"^MIN\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "MAX"               : r"^MAX\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "NB"                : r"^NB\([A-Z][0-9]+:[A-Z][0-9]+\)$",
            "NB.SI"             : r"^NB\.SI\([A-Z][0-9]|\d+;\d\)$",
            "ERREUR"            : r"^ERREUR\(\)$",
            "SI"                : r"^SI\((.*?);(.*?)(?:;(.+))?\)$",
            "HEURE"             : r"^HEURE\(\)$",
            "DATE"              : r"^DATE\(\)$",
            "MAINTENANT"        : r"^MAINTENANT\(\)$",
            "UNIX"              : r"^UNIX\(\)$",
            "SIMPLE"            : r"^([A-Z][0-9])+$",
            "LIM"               : r"", # à faire
            "TRIER"             : r"", # à faire
            # formules combinées
            "COMBINEE"          : r"(SOMME|MOYENNE|MIN|MAX|RANDINT|RACINE|MOYENNE\.POND|SI|NB|NB\.SI|SOMMEPROD|EXP|ABS|RACINE\.DEG|BINOME|PROD\.SCAL|LIM|HEURE|DATE|MAINTENANT|UNIX)\([^\)]+\)",
            # autres
            "COULEUR"           : r"^COULEUR\(#[A-Za-z][0-9]+[A-Za-z][0-9]+[A-Za-z][0-9]+\)$",
            "MAIL"              : r"/^([a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})*$/",
            "URL"               : r"/https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#()?&//=]*)/ ",
            "COULEURHEX"        : r"^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$"
        }

        self.erreurs = { # à refaire
            "DIV 0"                     : "",
            "FORMULE INCONNUE"          : "",
            "PLAGES INCORRETES"         : "",
            "PROBABILITE INCORRECTE"    : "",
            "TYPE INCORRECT"            : ""
        }

    def creer_table(self, lignes, cols):
        try:
            """
            Permet de créer la table.
            """
            logMessage("[creer_table] Création de la table.")
            self.dataframe = pd.DataFrame(columns=[self.get_column_label(c) for c in range(cols)])

            for c in range(cols):
                label = tk.Label(self.table, text=self.get_column_label(c), relief="ridge")
                label.grid(row=0, column=c+1, sticky=tk.NSEW)

            for l in range(lignes):
                label = tk.Label(self.table, text=str(l+1), relief="ridge")
                label.grid(row=l+1, column=0, sticky=tk.NSEW)
                for c in range(cols):
                    cellule = tk.Entry(self.table, relief="ridge")
                    cellule.grid(row=l+1, column=c+1, sticky=tk.NSEW)
                    cellule.insert(tk.END, "")
                    cellule.bind("<Button-1>", self.selection_cellule)
                    cellule.bind("<FocusOut>", self.evaluer_cellule)
                    self.cellules[(l, c)] = cellule
                    self.cellules_raw[(l, c)] = cellule
            logMessage("[creer_table] Table créée avec succès")
        except Exception:
            logMessage(f"[creer_table] Une erreur est survenue lors de la création de la table.", typelog="critical")
            logMessage(f"{traceback.format_exc()}", typelog="critical", indent=1)
            self.exit_program()

    def get_column_label(self, nb_col: int):
        """
        Permet d'obtenir une lettre à partir à partir d'un chiffre (par exemple le chiffre 1 correspond à la colonne A)

        Note : je me suis aidé de chatGPT pour cette méthode
        """
        label = ""
        while nb_col >= 0:
            label = chr(nb_col % 26 + 65) + label
            nb_col = nb_col // 26 - 1
        return label

    def selection_cellule(self, event):
        """
        Change la cellule séléctionnée.
        """
        cellule = event.widget
        rows, col = None, None
        for cle, valeur in self.cellules.items():
            if valeur == cellule:
                rows, col = cle
                break
        if rows is not None and col is not None:
            self.current_cell = (rows, col)
            self.cell_selected(event)
            self.selection_cellule_evidence(rows, col)

    def selection_cellule_evidence(self, rows: int, col: int):
        """
        Met la cellule séléctionnée en évidence.
        """
        for cle, cellule in self.cellules.items(): # cellule = valeur
            if cle == (rows, col):
                cellule.config(bg=COULEUR_EVIDENCE)
                cellule.focus_set()
            else:
                cellule.config(bg="white")

    def key_handler(self, event):
        """
        Permet de prendre en charge les touches du clavier pour la navigation.
        """
        rows, col = self.current_cell

        nb_lignes = self.dataframe.shape[0] if self.dataframe.shape[0] > 0 else 0

        if event.keysym == "Up":
            rows = max(0, rows - 1)
        elif event.keysym == "Return":
            rows = max(nb_lignes - 1, rows + 1)

        self.current_cell = (rows, col)
        self.cell_selected(event)
        self.selection_cellule_evidence(rows, col)

    def evaluer_cellule(self, event):
        """
        Lorsque la sélection d'une cellule est perdue, on applique les différentes formules trouvées dans celle-ci.
        """
        cellule = event.widget
        try:
            valeur = cellule.get()
            valeur = valeur.strip()
            valeur = valeur.replace(",", ".")
            self.cellules_raw[self.current_cell] = valeur
            if valeur.startswith("="):
                logMessage(f"[evaluer_cellule] Formule détéctée dans la cellule {self.current_cell} : {valeur}")
                if "+" in valeur or "-" in valeur or "*" in valeur or "/" in valeur:
                    res = self.evaluer_formule_combinee(valeur[1:])
                else:
                    res = self.evaluer_formule(valeur[1:])
                cellule.delete(0, tk.END)
                cellule.insert(tk.END, str(res))
            elif valeur.startswith("'="):
                cellule.delete(0, tk.END)
                cellule.insert(tk.END, valeur[1:])
            else: # pour remplacer les valeurs avec les "," avec des "." pour éviter des soucis avec des nombres décimaux.
                cellule.delete(0, tk.END)
                cellule.insert(tk.END, valeur)

        except Exception as e:
            logMessage(f"[evaluer_cellule] Erreur : {traceback.format_exc()}", typelog="error")
            messagebox.showerror("Erreur", f"Erreur : {traceback.format_exc()}")

    def evaluer_formule(self, valeur: str):
        """
        Cette fonction évalue des formules simples comme SOMME(A1:A3), MOYENNE(A1:B2), etc.
        """
        logMessage(f"[evaluer_formule] Valeur entrée : {valeur}", indent=True)
        try:
            if re.match(self.formats["SOMME"], valeur):
                arguments = valeur[6:-1]
                nombres = self.evaluer_arguments(arguments)
                nb = sum(nombres)
                logMessage(f"[evaluer_formule] Somme de {nombres} calculée : {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["SOMMEPROD"], valeur):
                arguments = valeur[10:-1]
                valeurs1_range, valeurs2_range = arguments.split(";")
                valeurs1 = self.evaluer_arguments(valeurs1_range)
                valeurs2 = self.evaluer_arguments(valeurs2_range)
                if len(valeurs1) != len(valeurs2):
                    return "ERREUR PLAGES INCORRECTES"
                s = 0
                for valeur1 in valeurs1:
                    for valeur2 in valeurs2:
                        s += valeur1 * valeur2
                logMessage(f"[evaluer_formule] Somme produit de {valeurs1, valeurs2} calculé : {s}", indent=True)
                return s
            elif re.fullmatch(self.formats["MOYENNE"], valeur):
                arguments = valeur[8:-1]
                nombres = self.evaluer_arguments(arguments)
                nb = sum(nombres) / len(nombres) if nombres else 0
                logMessage(f"[evaluer_formule] Moyenne de {nombres} calculée : {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["MOYENNE.POND"], valeur):
                arguments = valeur[16:-1]
                valeurs_range, coeffs_range = arguments.split(";")
                valeurs = self.evaluer_arguments(valeurs_range)
                coeffs = self.evaluer_arguments(coeffs_range)
                if len(valeurs) != len(coeffs):
                    return "ERREUR PLAGES INCORRECTES"
                s = 0
                c = 0
                for valeur in valeurs:
                    for coeff in coeffs:
                        s += valeur * coeff
                        c += 1
                nb = s/c if c else 0
                logMessage(f"[evaluer_formule] Moyenne pondéré de {valeurs, coeffs} calculée : {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["MIN"], valeur):
                arguments = valeur[4:-1]
                nombres = self.evaluer_arguments(arguments)
                logMessage(f"[evaluer_formule] Minimum calculé : {min(nombres)}", indent=True)
                return min(nombres)
            elif re.fullmatch(self.formats["MAX"], valeur):
                arguments = valeur[4:-1]
                nombres = self.evaluer_arguments(arguments)
                logMessage(f"[evaluer_formule] Maximum calculé : {max(nombres)}", indent=True)
                return max(nombres)
            elif re.fullmatch(self.formats["NB"], valeur):
                arguments = valeur[3:-1]
                nombres = self.evaluer_arguments(arguments)
                logMessage(f"[evaluer_formule] Nombres de cellules calculé dans la plage {arguments} : {len(nombres)}", indent=True)
                return len(nombres)
            elif re.fullmatch(self.formats["NB.SI"], valeur):
                arguments = valeur[6:-1].split(";")
                nombres = self.evaluer_arguments(arguments[0])
                critere = arguments[1]
                nombres2 = []
                for valeur in nombres:
                    if eval(f"{valeur}{critere}"):
                        nombres.append(valeur)
                logMessage(f"[evaluer_formule] Nombres de cellules calculé dans la plage {arguments} avec la condiion {critere} : {len(nombres2)}", indent=True)
                return len(nombres2)
            elif re.fullmatch(self.formats["RANDINT"], valeur):
                arguments = valeur[8:-1]
                nombres = self.evaluer_arguments(arguments)
                if nombres[0] > nombres[1]:
                    nombres[0], nombres[1] = nombres[1], nombres[0]
                nb = random.randint(nombres[0], nombres[1]) if len(nombres) == 2 else "ERREUR PLAGES"
                logMessage(f"[evaluer_formule] Nombre aléatoire calculé entre {nombres[0]} et {nombres[1]} : {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["EXP"], valeur):
                arguments = valeur[4:-1]
                nombre = self.evaluer_arguments(arguments)
                nb = math.exp(nombre)
                logMessage(f"[evaluer_formule] Exponentielle calculée : {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["ABS"], valeur):
                arguments = valeur[4:-1]
                nombre = self.evaluer_arguments(arguments)
                nb = abs(nombre)
                logMessage(f"[evaluer_formule] Valeur absolue de {nombre} = {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["RACINE"], valeur):
                arguments = valeur[7:-1]
                nombre = self.evaluer_arguments(arguments)
                logMessage(f"[evaluer_formule] Racine calculée : {nombre}", indent=True)
                return float(nombre**5)
            elif re.fullmatch(self.formats["RACINE.DEG"], valeur):
                arguments = valeur[11:-1].split(";")
                nombre = self.evaluer_arguments(arguments[0])
                degree = self.evaluer_arguments(arguments[1])
                nb = nombre**(1/degree) if (type(nombre) == int or float) else "ERREUR TYPE DE CHIFFRE"
                logMessage(f"[evaluer_formule] Racine de degrée {degree} calculée : {nb}", indent=True)
                return nb
            elif re.fullmatch(self.formats["HEURE"], valeur):
                return str(datetime.datetime.now().strftime("%H:%M:%S"))
            elif re.fullmatch(self.formats["DATE"], valeur):
                return str(datetime.datetime.now().strftime("%w/%m/%y"))
            elif re.fullmatch(self.formats["UNIX"], valeur): # à faire
                unix = datetime.datetime.now() - datetime.datetime(1970, 1, 1)
                unix.total_seconds()
                return int(unix)
            elif re.fullmatch(self.formats["MAINTENANT"], valeur):
                return str(datetime.datetime.now().strftime("%w/%m/%y, %H:%M:%S,%f"))
            elif re.fullmatch(self.formats["SI"], valeur):
                match = re.fullmatch(self.formats["SI"], valeur)
                condition = match.group(1)
                alors = match.group(2)
                sinon = match.group(3) if match.group(3) else None
                condition_result = self.evaluer_condition(condition)
                if condition_result:
                    res = self.evaluer_formule(alors)
                elif sinon:
                    res = self.evaluer_formule(sinon)
                else:
                    res = None
                logMessage(f"[evaluer_formule] Condition évaluée : {condition} -> {condition_result} -> {res}", indent=True)
                return res
            elif re.fullmatch(self.formats["ERREUR"], valeur):
                a = "qzd" / "qzd"
            elif re.fullmatch(self.formats["BINOME"], valeur): # à faire
                arguments = valeur[6:-1].split(";")
                n = arguments[0]
                k = arguments[1]
                p = arguments[2]
                if p < 0 or p > 1:
                    return "ERREUR PROBABILITE INCORRECTE"
                res = math.comb(n, k) * p**k * (1-p)**(n-k)
                logMessage(f"[evaluer_formule] X calculé pour loi binomiale de paramètres n = {n}, k = {k} et p = {p} : {res}", indent=True)
                return res
            elif re.fullmatch(self.formats["PROD.SCAL"], valeur):
                arguments = valeur[9:-1].split(";")
                longueur1 = self.evaluer_arguments(arguments[0])
                longueur2 = self.evaluer_arguments(arguments[1])
                angle = self.evaluer_arguments(arguments[2])
                res = longueur1 * longueur2 * math.cos(angle)
                logMessage("[evaluer_formule] Produit sclaire calculé pour les longeurs "+str(longueur1)+" & "+str(longueur2)+" et un angle de "+str(angle)+" : "+str(res), ident=True)
                return res
            elif re.fullmatch(self.formats["LIM"], valeur):
                arguments = valeur[3:-1].split(";")
                fonction = arguments[0]
                cible = arguments[1]
                limite = self.evaluer_limite(fonction,  cible)
                logMessage(f"Limite calculée pour la fonction {fonction} lorsque que {VARIABLE_FONCTION} tend vers {cible} = {limite}", indent=True)
                return limite
            elif re.fullmatch(self.formats["TRIER"]):
                arguments = valeur[5:-1].split(";")
                valeurs1 = self.evaluer_arguments(arguments[0])
                critere = arguments[1]
                if critere == "<" or "décroissant" or "decroissant":
                    res = trier(valeurs1, lambda x, y : x < y)
                elif critere == ">" or "croissant":
                    res = trier(valeurs1, lambda x, y : x > y)
                logMessage(f"[evaluer_formule] La liste {valeurs1} triée selon le critère {critere}: {res}", indent=True)
                return res
            elif re.fullmatch(self.formats["COULEUR"], valeur):
                valeur = valeur[8:-1]
                self.remplissage(valeur)
                return
            elif re.fullmatch(self.formats["SIMPLE"], valeur):
                valeur = valeur[1:]
                valeur = self.evaluer_arguments(valeur)
                return valeur
            else:
                return "#FORMULE?"
        except Exception as e:
            logMessage(f"[evaluer_formule] Erreur : {traceback.format_exc()}", typelog="error")
            messagebox.showerror("Erreur", f"Erreur : {traceback.format_exc()}")

    def evaluer_formule_combinee(self, valeur):
        """
        Cette fonction évalue des formules combinées, comme =SOMME(A1:A3) + MOYENNE(A1:A3).
        Elle remplace chaque fonction par son résultat avant d'évaluer l'expression complète.
        """
        # Rechercher toutes les fonctions dans la formule
        matches = re.finditer(self.formats["COMBINEE"], valeur)

        for match in matches:
            full_match = match.group(0)  # Par exemple SOMME(A1:A3)
            func_name, arguments = full_match.split("(", 1)
            arguments = arguments[:-1]

            if func_name == "SOMME":
                res = self.evaluer_formule("SOMME(" + arguments + ")")
            elif func_name == "SOMMEPROD":
                res = self.evaluer_formule("SOMMEPROD(" + arguments + ")")
            elif func_name == "MOYENNE":
                res = self.evaluer_formule("MOYENNE(" + arguments + ")")
            elif func_name == "MIN":
                res = self.evaluer_formule("MIN(" + arguments + ")")
            elif func_name == "MAX":
                res = self.evaluer_formule("MAX(" + arguments + ")")
            elif func_name == "NB":
                res = self.evaluer_formule("NB(" + arguments + ")")
            elif func_name == "NB.SI":
                res = self.evaluer_formule("NB.SI(" + arguments + ")")
            elif func_name == "RANDINT":
                res = self.evaluer_formule("RANDINT(" + arguments + ")")
            elif func_name == "RACINE":
                res = self.evaluer_formule("RACINE(" + arguments + ")")
            elif func_name == "RACINE.DEG":
                res = self.evaluer_formule("RACINE.DEG(" + arguments + ")")
            elif func_name == "MOYENNE.POND":
                res = self.evaluer_formule("MOYENNE.POND(" + arguments + ")")
            elif func_name == "SI":
                res = self.evaluer_formule("SI(" + arguments + ")")
            elif func_name == "EXP":
                res = self.evaluer_formule("EXP(" + arguments + ")")
            elif func_name == "ABS":
                res = self.evaluer_formule("ABS(" + arguments + ")")
            elif func_name == "BINOME":
                res = self.evaluer_formule("BINOME(" + arguments + ")")
            elif func_name == "PROD.SCAL":
                res = self.evaluer_formule("PROD.SCAL(" + arguments + ")")
            elif func_name == "LIM":
                res = self.evaluer_formule("LIM(" + arguments + ")")
            elif func_name == "UNIX":
                res = self.evaluer_formule("UNIX()")
            elif func_name == "ERREUR":
                res = self.evaluer_formule("ERREUR()")
            valeur = valeur.replace(full_match, str(res))
        valeur = self.replace_cell_references(valeur)
        try:
            return eval(valeur)
        except Exception:
            logMessage(f"[evaluer_formule_combinee] Erreur dans l'évaluation : {traceback.format_exc()}", typelog="error", indent=True)
            messagebox.showerror("Erreur", traceback.format_exc())
            return "ERREUR"

    def evaluer_condition(self, condition):
        """
        Évalue la condition dans une formule SI, par exemple : A1=A2, A1>=A2, A1=2
        """
        condition = condition.strip()
        if "=" in condition:
            if ">" in condition:
                # Si la condition est de type "A1>=A2"
                cell1, cell2 = condition.split(">=")
                return self.evaluer_formule(f"={cell1.strip()}") >= self.evaluer_formule(f"={cell2.strip()}")
            if "<" in condition:
                # Si la condition est de type "A1<=A2"
                cell1, cell2 = condition.split("<=")
                return self.evaluer_formule(f"={cell1.strip()}") <= self.evaluer_formule(f"={cell2.strip()}")
            # Condition de type "A1=A2"
            cell1, cell2 = condition.split("=")
            return self.evaluer_formule(f"={cell1.strip()}") == self.evaluer_formule(f"={cell2.strip()}")

        # Condition de type "A1=2"
        try:
            return float(condition) == self.evaluer_formule(f"={condition.strip()}")
        except:
            return False

    def evaluer_limite(self, fonction, cible):
            """
            Évalue la limite d'une fonction dont la variable converge vers un point ou diverge vers l'infini.
            """
            if cible != "+inf" or "-inf":
                try:
                    point = self.evaluer_arguments(cible)
                    res = sp.limit(fonction, VARIABLE_FONCTION, point)
                except Exception:
                    logMessage(f"[evaluer_formule] Erreur calcul limite : {traceback.format_exc()}", typelog="error", indent=True)
                    res = "ERREUR"
            elif cible == "-inf":
                try:
                    res = sp.limit(fonction, VARIABLE_FONCTION, -sp.oo)
                except Exception:
                    logMessage(f"[evaluer_formule] Erreur calcul limite : {traceback.format_exc()}", typelog="error", indent=True)
                    res = "ERREUR"
            elif cible == "+inf":
                try:
                    res = sp.limit(fonction, VARIABLE_FONCTION, sp.oo)
                except Exception:
                    logMessage(f"[evaluer_formule] Erreur calcul limite : {traceback.format_exc()}", typelog="error", indent=True)
                    res = "ERREUR"
            else:
                res = "ERREUR"
            return res

    def evaluer_arguments(self, arguments):
        """
        Analyse les arguments d'une formule comme (A1:B3, C2, 10) par exemple.
        Retourne une liste de valeur correspondantes.
        """
        res = []
        for arg in re.split(r"[;,]", arguments):
            arg = arg.strip()
            if re.match(r"^[A-Z][0-9]+:[A-Z][0-9]+$", arg):
                start, end = arg.split(":")
                start_index = self.cellule_index(start)
                end_index = self.cellule_index(end)
                for row in range(start_index[0], end_index[0] + 1):
                    for col in range(start_index[1], end_index[1] + 1):
                        valeur = self.cellules.get((row, col), tk.Entry()).get()
                        res.append(float(valeur) if valeur else 0)
                        # logMessage(f"[evaluer_arguments] Cellule {start}:{end} = {valeur}", indent=True)
            elif re.match(r"^[A-Z][0-9]+$", arg):
                row, col = self.cellule_index(arg)
                valeur = self.cellules.get((row, col), tk.Entry()).get()
                res.append(float(valeur) if valeur else 0)
                # logMessage(f"[evaluer_arguments] Cellule {arg} = {valeur}")
            elif re.match(r"^\d+(\.\d+)?$", arg):
                res.append(float(arg))
                # logMessage(f"[evaluer_arguments] Constante {arg}")
        logMessage(f"[evaluer_arguments] {arguments} = {res}", indent=True)
        return res

    def cellule_index(self, cellule_nom):
        """
        Renvoie les coordonnés à partir d'une string de cellule (par exemple cellule_index("A1") renvoie (0,0))
        """
        col = 0
        rows = 0
        for char in cellule_nom:
            if char.isdigit():
                rows = int(cellule_nom[cellule_nom.index(char):]) - 1
                break
            col = col * 26 + (ord(char.upper()) - ord("A"))
        return (rows, col)

    def replace_cell_references(self, expression):
        """
        Remplace les références de cellules (par exemple A1) par leur valeur dans l'expression donnée.
        Par exemple si A1 = 5, "A1*2" renverra "5*2".
        """
        # Trouver toutes les références de cellules dans l'expression (par exemple A1, B2, etc.)
        matches = re.findall(r"([A-Z][0-9]+)", expression)

        for match in matches:
            # Pour chaque référence de cellule, remplacez-la par sa valeur
            row, col = self.cellule_index(match)
            cell_value = self.cellules.get((row, col), tk.Entry()).get() or 0
            expression = expression.replace(match, str(cell_value))  # Remplace la référence par sa valeur

        return expression

    def cell_info(self):
        """
        Renvoie des informations sur la cellule séléctionnée telles que les coordonnés, la valeur, la formule non forrmatée, l'emplacement...
        """
        coos = self.current_cell
        valeur = self.cellules[self.current_cell].get()
        messagebox.showinfo("Information sur la cellule.", f"coordonnés : {coos}\nvaleur = {valeur}\nAutres : {self.cellules[coos]}")

    def show_cells(self):
        """
        Permet d'obtenir des informations sur les cellules (coordonnées et valeurs)
        """
        cellules = []
        cell_text = "(ligne, colonne), valeur" + "\n"
        for cle, valeur in self.cellules.items():

            cellules.append((cle, self.cellules[cle].get()))
        for cell in cellules:
            cell_text += str(cell) + "\n"
        messagebox.showinfo("Cellules", cell_text)

    def clear(self):
        """
        Permet de nettoyer la table.
        """
        self.cellules_temp = self.cellules
        for coos, _ in self.cellules.items():
            self.cellules_temp[coos] = ""
        self.cellules = self.cellules_temp
        self.update_dataframe()

    def ouvrir(self):
        """
        Permet d'ouvrir un fichier.
        """
        file_path = filedialog.askopenfilename(filetypes=[("Fichier CSV", "*.csv")])

        if file_path:
            try:
                self.dataframe = pd.read_csv(file_path)
                if self.dataframe.empty:
                    raise ValueError("Le fichier est vide.")
                self.populate_table()
            except pd.errors.EmptyDataError:
                logMessage(f"[ouvrir] Erreur : {str(e)}", "error")
                messagebox.showerror("Erreur", "Le fichier séléctionné est vide ou non pris en charge.")
            except ValueError as e:
                logMessage(f"[ouvrir] Erreur : {str(e)}", "error")
                messagebox.showerror("Erreur", str(e))
            except Exception as e:
                logMessage(f"[ouvrir] Erreur : {traceback.format_exc()}", "error")
                messagebox.showerror("Erreur", f"Une erreur est survenue {traceback.format_exc()}")

    def enregistrer(self):
        """
        Permet d'enregistrer le fichier aux formats CSV ou XLSX
        """
        types_fichiers = [("fichier CSV", "*.csv"), ("Fichier Excel", "*.xlsx")]
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=types_fichiers)

        if file_path:
            if file_path.endswith(".csv"):
                self.update_dataframe()
                self.dataframe.to_csv(file_path, index=False)
                logMessage(f"[enregistrer] Fichier sauvegardé avec succès : {file_path}")
                messagebox.showinfo("Sauvegarde", "Fichier sauvegardé au format CSV avec succès.")
            elif file_path.endswith(".xlsx"):
                self.update_dataframe()
                self.dataframe.to_excel(file_path, index=False)
                logMessage(f"[enregistrer] Fichier sauvegardé avec succès : {file_path}")
                messagebox.showinfo("Sauvegarde", "Fichier sauvegardé au format XLSX avec succès.")
        else:
            logMessage(f"[enregistrer] Erreur : chemin non valide.", typelog="error")
            logMessage(message=f"{file_path}", typelog="error", indent=1)

    def enregistrer_csv(self):
        """
        Permet d'enregistrer le fichier au format CSV.
        """
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichier CSV", "*.csv")])
        if file_path:
            self.update_dataframe()
            try:
                self.dataframe.to_csv(file_path, index=False)
                logMessage(f"[enregistrer_csv] Fichier sauvegardé avec succès : {file_path}")
                messagebox.showinfo("Sauvegarde", "Fichier CSV Sauvegardé avec succès")
            except Exception as e:
                logMessage(f"[enregistrer_csv] Erreur : {traceback.format_exc()}", typelog="error")
                messagebox.showerror("Erreur", f"Une erreur est survenue lors de la sauvegarde du fichier : {traceback.format_exc()}")

    def enregistrer_excel(self):
        """
        Permet d'enregistrer le fichier au format XLSX.
        """
        messagebox.showinfo("Avertissement", "Attention, ce logiciel ne prend en charge qu'uniquement les fichiers au format CSV. Si vous souhaitez modifier votre fichier plus tard, favorisez plutôt le format CSV.")
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Fichier Excel", "*.xlsx")])
        if file_path:
            self.update_dataframe()
            try:
                self.dataframe.to_excel(file_path, index=False)
                logMessage(f"[enregistrer_xlsx] Fichier sauvegardé avec succès : {file_path}")
                messagebox.showinfo("Sauvegarde", "Fichier CSV sauvegardé avec succès")
            except Exception as e:
                logMessage(f"[enregistrer_xlsx] Erreur : {traceback.format_exc()}", typelog="error")
                messagebox.showerror("Erreur", f"Une erreur est survenue lors de la sauvegarde du fichier : {traceback.format_exc()}")

    def gras(self):
        """
        Met le texte d'une cellule en gras.
        """
        ligne, col = self.current_cell
        current_font = font.Font(font=self.cellules[(ligne, col)].cget("font"))
        if current_font.actual()["weight"] == "normal":
            current_font["weight"] = "bold"
        else:
            current_font["weight"] = "normal"
        self.cellules[(ligne, col)].config(font=current_font)

    def italique(self):
        """
        Met le texte d'une cellule en italique.
        """
        ligne, col = self.current_cell
        current_font = font.Font(font=self.cellules[(ligne, col)].cget("font"))
        if current_font.actual()["slant"] == "roman":
            current_font["slant"] = "italic"
        else:
            current_font["slant"] = "roman"
        self.cellules[(ligne, col)].config(font=current_font)

    def souligner(self): # à refaire
        """
        Souligne le texte.
        """
        ligne, col = self.current_cell
        current_font = font.Font(font=self.cellules[(ligne, col)].cget("font"))
        current_font.configure(underline=True)
        self.cellules[(ligne, col)].config(f=current_font)

    def remplissage(self, couleur=None):
        """
        Change la couleur du texte d'une cellule.
        """
        ligne, col = self.current_cell
        if not couleur:
            couleur = colorchooser.askcolor()[1]
        if re.fullmatch(self.formats["COULEURHEX"], couleur):
            logMessage(f"[remplissage] la couleur de la cellule {self.current_cell} a été changée en {couleur}.")
            self.cellules[(ligne, col)].config(fg=couleur)
        else:
            logMessage(f"[remplissage] Format couleur invalide : {couleur}", typelog="error")

    def couleur_police(self, couleur=None):
        """
        Change la couleur de la police vers la couleur donnée.
        """
        ligne, col = self.current_cell
        if not couleur:
            couleur = colorchooser.askcolor()[1]
        if re.fullmatch(self.formats["COULEURHEX"], couleur):
            logMessage(f"[couleur_police] la couleur de la police de la cellule {self.current_cell} a été changée en {couleur}")
            self.cellules[(ligne, col)].config(fg=couleur)
        else:
            logMessage(f"[couleur_police] Format couleur invalide : {couleur}", typelog="error")

    def changer_police(self, police=POLICE):
        """
        Change la police d'une cellule.
        """
        ligne, col = self.current_cell
        if (ligne, col) in self.cellules():
            cellule = self.cellules[(ligne, col)]
            police2 = font.Font(family=police, size=12)
            cellule.config(font=police2)
        else:
            logMessage("[changer_police] Erreur : la cellule"+str((ligne, col))+"n'existe pas.", typelog="error")

    def populate_table(self):
        """
        Permet d'afficher les données dans la table.
        (la méthode update_dataframe met à jour les données mais ne les affiche pas.)
        """
        ligne, cols = self.dataframe.shape

        for l in range(ligne):
            for c in range(cols):
                valeur = self.dataframe.iat[l, c]

                if pd.isna(valeur):
                    valeur = ""

                self.cellules[(l, c)].delete(0, tk.END)
                self.cellules[(l, c)].insert(tk.END, valeur)

    def update_dataframe(self):
        """
        Met à jour la table avec les nouvelles données.
        (n'affiche pas les données dans la table, c.f.)
        Note : aidé de chatGPT
        """
        lignes = max([cle[0] for cle in self.cellules.keys()]) + 1
        cols = max([cle[1] for cle in self.cellules.keys()]) + 1

        data = []

        vide = True
        for l in range(lignes):
            row_data = []
            for c in range(cols):
                valeur = self.cellules[(l, c)].get()
                row_data.append(valeur)
                if valeur != "":
                    vide = False
            data.append(row_data)

        if vide:
            logMessage("[update_dataframe] La table est vide.", typelog="warning")
            messagebox.showwarning("Avertissement", "La table est vide. Veuillez ajouter des donnés avant de enregistrer.")
            return

        self.dataframe = pd.DataFrame(data, columns=[self.get_column_label(c) for c in range(cols)])

    def update_formulabar(self, valeur):
        """
        Met à jour la barre de formule avec le contenu de la cellule et la formule.
        """
        self.cell_content_entry.delete(0, tk.END)
        self.cell_content_entry.insert(0, valeur)  # Affiche le contenu de la cellule

    def cell_selected(self, event):
        """
        Cette fonction est appelée lorsqu'une cellule est sélectionnée.
        """
        valeur = self.get_cell_value(event)
        self.update_formulabar(valeur)

    def get_cell_value(self, event):
        """
        Retourne la valeur d'une cellule à partir d'un event.
        """
        cellule = event.widget
        valeur = cellule.get()
        return valeur

    def aide_formules(self):
        """
        Affiche de l'aide concernant les formules présentes dans le logiciel.
        """
        texte = """
        LISTE DES FORMULES DISPONIBLES :
        SOMME() -> calcule la somme des valeurs
        SOMMEPROD() -> calcule la somme des produits des valeurs de deux plages
        MOYENNE() -> calcule une moyenne de valeurs
        MOYENNE.POND() -> calcule la moyenne pondérée de deux plages.
            utilisation = MOYENNE.POND(A1:A3;B1;B3)
            A1:A3 représente les valeurs et B1:B3 les coefficiants
        MIN() -> renvoie le minimum de la plage donnée
        MAX() -> renvoie le maximum de la plage donnée
        NB() -> renvoie la longueur de la plage donnée
        SI() -> fonction conditionnelle.
            utilisation : SI(condition;si_vrai;si_faux).
            condition -> par exemple A1=A2 ou A1>=A1.
            si_vrai -> instructions si la condition est remplie. Peut être une valeur simple (A1 ou 50 par exemple) ou une formule.
            si_faux -> instructions si la condition n'est pas remplie. Non obligatoire.
        RANDINT() -> génère un nombre entier aléatoire entre deux valeurs données.
        RACINE() -> renvoie la racine de la valeur donnée.
        RACINE.DEG() -> renvoie la racine de la valeur donnée au degré indiqué.
            utilisation = RACINE.DEG(3, 50) -> racine cubique de 50
        BINOME() -> renvoie le resulstat de la loie binomiale avec les arguments donnés
            utilisation : BINOME(n; k; p) -> BINOME(A1;A2;A3)
        LIM() -> permet de calculer la limite d'une fonction
            utilisation = LIM(fonction, [point ou -inf / +inf])
        PROD.SCAL() -> calcule un produit scalaire à partir de 2 longeurs et un angle
            utilisation = PROD.SCAL(longueur1;longueur2;angle)
        ERREUR() -> provoque une erreur (utilisé pour le développement)
        HEURE() -> renvoie l'heure actuelle
        DATE() -> renvoie la date actuelle
        MAINTENANT() -> renvoie la date et l'heure actuelle
        EXP() -> renvoie la valeur donnée passée par la fonction exponentielle
        ABS() -> renvoie la valeur absolue de la valeur donnée.
        COULEUR() -> remplace la couleur de la cellule spécifiée par la couleur donnée
            utilisation = COULEUR(A1;#FFFFFF)
        TRIER() -> trie une liste de valeur

        EXEMPLES D'OPERATIONS
        =A1+A2
        =5+9
        =A1*2
        =SOMME(A1:B5)
        =SOMME(A1:A2) * MOYENNE(A1:A2)
        =SI(SOMME(A1:A2)>20;SOMME(B1:B2);COULEUR(#FFFFFF))
        """
        messagebox.showinfo("Aide sur les formules", texte)

    def afficher_parametres(self):
        """
        Affiche les paramètres du logiciel.
        """
        texte = f"""
        PARAMETRES & INFORMATIONS (affichage uniquement)

        NB_COLONNES = {NB_COLONNES}
        NB_LIGNES = {NB_LIGNES}
        COULEUR_EVIDENCE = {COULEUR_EVIDENCE}
        LOGGING_ENABLED = {LOGGING_ENABLED}
        VARIABLE_FONCTION = \"{VARIABLE_FONCTION}\"
        POLICE = {POLICE}
        TAILLE_TABLE = {TAILLE_TABLE}
        """
        messagebox.showinfo("Paramètres du tableur", texte)

    def exit_program(self):
        """
        Ferme le logiciel.
        Demande si l'utilisateur souhaite enregistrer son travail avant de quitter.
        """
        reponse = messagebox.askquestion("Enregistrer les modifications ?", "Voulez vous enregistrer les modfications avant de quitter ?")
        if reponse == "yes":
            self.enregistrer()
        logMessage(f"[exit_program] Arrêt.", typelog="critical")
        self.root.quit()
        self.root.destroy()

def fusion(liste1, liste2, critere):
    """
    Fusionne deux listes en utilisant un critère de tri.
    Utilisé par le tri fusion.
    """
    liste = []
    i1, i2 = 0, 0
    while i1 < len(liste1) and i2 < len(liste2):
        if critere(liste1[i1], liste2[i2]):
            liste.append(liste2[i2])
            i2 += 1
        else:
            liste.append(liste1[i1])
            i1 += 1
    return liste + liste1[i1:] + liste2[i2:]

def trier(liste, critere):
    """
    Trie une liste selon un critère de tri à l'aide d'un tri par fusion.
    Le critère par défaut trie en ordre croissant.
    """
    n = len(liste)
    if n < 2: 
        return liste
    else:
        m = n // 2
        gauche = trier(liste[:m], critere)
        droite = trier(liste[m:], critere)
        return fusion(gauche, droite, critere)


def obtenir_specifications_utilisateur():
    """
    Permet d'obtenir des informations sur les spécifications de la machine de l'utilisateur, tel que l'OS, le CPU ou encore la RAM.
    Renvoie les données sous la forme d'une string formatée.
    """
    # Informations sur le système d'exploitation
    os_info = platform.uname()
    systeme = os_info.system
    version = os_info.version
    release = os_info.release
    machine = os_info.machine
    processeur = os_info.processor

    if release == "10":
        win11_warn = "or Windows 11"
    else:
        win11_warn = ""

    # Informations sur la RAM
    ram_totale = round(psutil.virtual_memory().total / (1024**3), 2)

    # Informations sur le CPU
    cpu_count = os.cpu_count()
    frequence_cpu = psutil.cpu_freq().max  # En MHz

    gpu_info = obtenir_infos_gpu()

    result = f"""
    [INITIALISATION] Spécifications de la machine de l'utilisateur
    Système d'exploitation : {systeme} {release} (Version: {version}) {win11_warn}
    Architecture machine   : {machine}
    Processeur             : {processeur}
    RAM Totale (Go)        : {ram_totale}
    Nombre de CPU          : {cpu_count}
    Fréquence Max CPU (MHz): {frequence_cpu}
    GPU                    : {gpu_info if isinstance(gpu_info, str) else ', '.join([f"{gpu['Nom']} ({gpu['Mémoire Totale (Go)']} Go, Charge: {gpu['Charge (%)']}%)" for gpu in gpu_info])}
    """

    return result.strip()

def obtenir_infos_gpu():
    """
    Tente de récupérer les informations sur les GPU disponibles.
    Fonctionne uniquement si l'outil `nvidia-smi` est disponible.
    """
    try:
        # nvidia-smi
        resultat = subprocess.run(
            ["nvidia-smi", "--query-gpu=name,memory.total,utilization.gpu", "--format=csv,noheader,nounits"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        if resultat.returncode == 0:
            gpu_infos = []
            for ligne in resultat.stdout.strip().split("\n"):
                nom, memoire, utilisation = ligne.split(", ")
                gpu_infos.append({
                    "Nom": nom.strip(),
                    "Mémoire Totale (Go)": round(float(memoire) / 1024, 2),  # Converti en Go
                    "Charge (%)": float(utilisation)
                })
            return gpu_infos
        else:
            return "nvidia-smi non disponible ou aucun GPU NVIDIA détecté."
    except FileNotFoundError:
        return "L'outil nvidia-smi n'est pas installé sur ce système."

def obtenir_infos_logiciel():
    """
    Permet d'obtenir des informations relatives au logiciel.
    """
    texte = f"""
    [INITIALISATION] Spécifications du logiciel
    Version =  {VERSION}
    Modules = {REQUIREMENTS}
    """
    return texte.strip()

def copylist(liste):
    """
    Permet de copier une liste sans avoir le problème de copie de liste.
    """
    liste2 = []
    for element in liste:
        liste2.append(element)
    return liste2

def decomp(valeur):
    """
    Décompose une string en liste de string et de int
    """
    valeur_temp = []
    for c in range(len(valeur)):
        try:
            var = int(valeur[c])
        except ValueError:
            var = valeur[c]
        valeur_temp.append(var)
    return valeur_temp

def logMessage(message: str, typelog="info", indent=False, force=False):
    """
    Renvoie un message de log dans le fichier log.
    Envoie également le message dans le terminal.
    """
    if LOGGING_ENABLED or force:
        if indent:
            message = "    " + message
        # message = datetime.datetime.now().strftime("%w/%m/%y, %H:%M:%S,%f") + " : " + message
        if typelog == "info":
            LOG.info(message)
        elif typelog == "warning":
            LOG.warning(message)
        elif typelog == "error":
            LOG.error(message)
        elif typelog == "debug":
            LOG.debug(message)
        elif typelog == "critical":
            LOG.critical(message)
        else:
            logMessage(f"[logMessage] Erreur : le typelog de LOG n'est pas reconnu : typelog={typelog}, type(typelog)={type(typelog)}, indent={indent}", typelog="error")

if __name__ == "__main__":
    logMessage("[INITIALISATION] Lancement en cours.", force=True)
    infos_config  = obtenir_specifications_utilisateur()
    infos_logiciel = obtenir_infos_logiciel()
    logMessage(infos_config, force=True)
    logMessage(infos_logiciel, force=True)

    root = tk.Tk()
    app = Tableur(root)
    root.mainloop()
