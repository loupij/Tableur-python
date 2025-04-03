try:
    import tkinter as tk
    from tkinter import font, messagebox, filedialog, colorchooser
    # import sqlite3
    import pandas as pd
    import logging
    # from PIL import Image, ImageTk # sera utilisé plus tard quand on voudra mettre des icones pour certaines fonctionalités (gras, italique et changement de couleur notamment)
    import random
    import math
    import datetime
    import re
    import traceback
    import sympy as sp
except ModuleNotFoundError:
    with open("requirements.txt") as requirements:
        print("Veuillez vérifier que les modules suivants sont bien installés : "+str(requirements[7:])+"\nVous pouvez aussi faire la commande suivante : pip install requirements.txt")
    exit()

# lettres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# logs
LOG = logging.getLogger()
logging.basicConfig(filename="log.txt", level=logging.DEBUG) # fichier log

# paramètres basiques du tableur
NB_LIGNES = 10
NB_COLONNES = 20
COULEUR_EVIDENCE = "yellow"
LOGGING_ENABLED = True # ne fonctionne pas avec EduPython
VARIABLE_FONCTION = sp.symbols("x")
POLICE = "Arial" # à refaire

class Tableur():
    def __init__(self, root):
        self.root = root
        self.root.title("Tableur")
        self.dataframe = pd.DataFrame()

        # barre de formule
        self.cellbar = tk.Frame(self.root, bg="lightgray", height=40)
        self.cellbar.pack(fill=tk.X, side=tk.TOP)

        self.cell_content_label = tk.Label(self.cellbar, text="f", bg="lightgray")
        self.cell_content_label.pack(side=tk.LEFT, padx=10)

        self.cell_content_entry = tk.Entry(self.cellbar, width=NB_COLONNES*10)
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
        logMessage("[__init__] Création de la cellbar de menu.")
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
            "COULEURHEX"        : r"#[A-Za-z][0-9]+[A-Za-z][0-9]+[A-Za-z][0-9]+",
            "COMBINEE"          : r"(SOMME|MOYENNE|MIN|MAX|RANDINT|RACINE|MOYENNE\.POND|SI|LEN|SOMMEPROD|EXP|ABS|RACINE\.DEG|BINOME|PROD\.SCAL|LIM)\([^\)]+\)",
            # formules
            "SOMME"             : r"^SOMME\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "SOMMEPROD"         : r"^SOMMEPROD\(\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*;\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*\)$",
            "MOYENNE"           : r"^MOYENNE\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "MOYENNE.POND"      : r"^MOYENNE\.POND\(\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*;\s*([A-Z][0-9]+:[A-Z][0-9]+)\s*\)$",
            "RANDINT"           : r"^RANDINT\([A-Z][0-9]+,[A-Z][0-9]+\)$",
            "EXP"               : r"^EXP\([A-Z][0-9]|\d+\)$",
            "ABS"               : r"^ABS\([A-Z][0-9]|\d+\)$",
            "BINOME"            : r"^BINOME\([A-Z][0-9]|\d;[A-Z][0-9]|\d;[A-Z][0-9]|\d\)$",
            "PROD.SCAL"         : r"^PROD\.SCAL((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)",
            "RACINE"            : r"^RACINE\([A-Z][0-9]+\)$",
            "RACINE.DEG"        : r"^RACINE\.DEG\([A-Z][0-9]|\d+;\d\)$",
            "MIN"               : r"^MIN\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "MAX"               : r"^MAX\((([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)([;,] ?))*([A-Z][0-9]+:[A-Z][0-9]+|[A-Z][0-9]+|\d+)\)$",
            "LEN"               : r"^LEN\([A-Z][0-9]+:[A-Z][0-9]+\)$",
            "ERREUR"            : r"^ERREUR\(\)$",
            "SI"                : r"^SI\((.*?);(.*?)(?:;(.+))?\)$",
            "HEURE"             : r"^HEURE\(\)$",
            "DATE"              : r"^DATE\(\)$",
            "MAINTENANT"        : r"^MAINTENANT\(\)$",
            "UNIX"              : r"^UNIX\(\)$",
            "SIMPLE"            : r"^([A-Z]|[0-9])+$",
            "LIM"               : r"", # à faire
            # autres
            "COULEUR"           : r"^COULEUR\(#[A-Za-z][0-9]+[A-Za-z][0-9]+[A-Za-z][0-9]+\)$",
            "MAIL"              : r"/^([a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})*$/",
            "URL"               : r"/https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#()?&//=]*)/ "
        }

        self.erreurs = { # à faire
            "DIV 0"                     : "",
            "FORMULE INCONNUE"          : "",
            "PLAGES INCORRETES"         : "",
            "PROBABILITE INCORRECTE"    : "",
            "TYPE INCORRECT"            : ""
        }

    def creer_table(self, lignes: int, cols: int):
        try:
            """
            Permet de créer la table.
            """
            # logMessage("[creer_table] Création de la table.")
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
        except Exception as e:
            logMessage("[creer_table] Une erreur est survenue lors de la création de la table.", typelog="critical")
            logMessage("[creer_table] "+str(traceback.format_exc()), typelog="critical", indent=1)
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
            self.cellules_raw[self.current_cell] = valeur
            if valeur.startswith("="):
                valeur.replace(",", ".") # on remplace les "," par des "." pour éviter des problèmes avec les nombres décimaux
                logMessage("[evaluer_cellule] Formule détéctée dans la cellule " + str(self.current_cell) + " : " + str(valeur))
                if "+" in valeur or "-" in valeur or "*" in valeur or "/" in valeur:
                    res = self.evaluer_formule_combinee(valeur[1:])
                else:
                    res = self.evaluer_formule(valeur[1:])
                cellule.delete(0, tk.END)
                cellule.insert(tk.END, str(res))
            elif valeur.startswith("'="):
                cellule.delete(0, tk.END)
                cellule.insert(tk.END, valeur[1:])
        except Exception as e:
            logMessage("[evaluer_cellule] Erreur : " + str(e), typelog="error")
            messagebox.showerror("Erreur", str(traceback.format_exc()))

    def evaluer_formule(self, valeur: str):
        """
        Cette fonction évalue des formules simples comme SOMME(A1:A3), MOYENNE(A1:B2), etc.
        """
        logMessage("[evaluer_formule] Valeur entrée : " + valeur)
        try:
            if re.match(self.formats["SOMME"], valeur):
                arguments = valeur[6:-1]
                nombres = self.parse_formula_arguments(arguments)
                nb = sum(nombres)
                logMessage("[evaluer_formule] Somme de " + str(nombres) + " calculée : " + str(nb), indent=True)
                return nb
            elif re.fullmatch(self.formats["SOMMEPROD"], valeur):
                arguments = valeur[10:-1]
                valeurs1_range, valeurs2_range = arguments.split(";")
                valeurs1 = self.parse_formula_arguments(valeurs1_range)
                valeurs2 = self.parse_formula_arguments(valeurs2_range)
                if len(valeurs1) != len(valeurs2):
                    return "ERREUR PLAGES INCORRECTES"
                s = 0
                for valeur1 in valeurs1:
                    for valeur2 in valeurs2:
                        s += valeur1 * valeur2
                logMessage("[evaluer_formule] Somme produit de " + str(valeurs1)+" et " + str(valeurs2) + " calculé : " + str(s), indent=True)
                return s
            elif re.fullmatch(self.formats["MOYENNE"], valeur):
                arguments = valeur[8:-1]
                nombres = self.parse_formula_arguments(arguments)
                nb = sum(nombres) / len(nombres) if nombres else 0
                logMessage("[evaluer_formule] Moyenne de " + str(nombres) + " calculée : " + str(nb), indent=True)
                return nb
            elif re.fullmatch(self.formats["MOYENNE.POND"], valeur):
                arguments = valeur[16:-1]
                valeurs_range, coeffs_range = arguments.split(";")
                valeurs = self.parse_formula_arguments(valeurs_range)
                coeffs = self.parse_formula_arguments(coeffs_range)
                if len(valeurs) != len(coeffs):
                    return "ERREUR PLAGES INCORRECTES"
                s = 0
                c = 0
                for valeur in valeurs:
                    for coeff in coeffs:
                        s += valeur * coeff
                        c += 1
                nb = s/c if c else 0
                logMessage("[evaluer_formule] Moyenne pondéré de " + str(valeurs) + "avec les coeffs" +  str(coeffs) + " calculée : " + str(nb), indent=True)
                return nb
            elif re.fullmatch(self.formats["MIN"], valeur):
                arguments = valeur[4:-1]
                nombres = self.parse_formula_arguments(arguments)
                logMessage("[evaluer_formule] Minimum calculé : " + str(min(nombres)), indent=True)
                return min(nombres)
            elif re.fullmatch(self.formats["MAX"], valeur):
                arguments = valeur[4:-1]
                nombres = self.parse_formula_arguments(arguments)
                logMessage("[evaluer_formule] Maximum calculé : " + str(max(nombres)), indent=True)
                return max(nombres)
            elif re.fullmatch(self.formats["LEN"], valeur):
                arguments = valeur[4:-1]
                nombres = self.parse_formula_arguments(arguments)
                logMessage("[evaluer_formule] Longueur calculé : " + str(len(nombres)), indent=True)
                return len(nombres)
            elif re.fullmatch(self.formats["RANDINT"], valeur):
                arguments = valeur[8:-1]
                nombres = self.parse_formula_arguments(arguments)
                if nombres[0] > nombres[1]:
                    nombres[0], nombres[1] = nombres[1], nombres[0]
                nb = random.randint(nombres[0], nombres[1]) if len(nombres) == 2 else "ERREUR PLAGES"
                logMessage("[evaluer_formule] Nombre aléatoire calculé entre " + str(nombres[0]) + " et " + str(nombres[1]) + " : " + str(nb), indent=True)
                return nb
            elif re.fullmatch(self.formats["EXP"], valeur):
                arguments = valeur[4:-1]
                nombre = self.parse_formula_arguments(arguments)
                nb = math.exp(nombre)
                logMessage("[evaluer_formule] Exponentielle calculée : " + str(nb), indent=True)
                return nb
            elif re.fullmatch(self.formats["ABS"], valeur):
                arguments = valeur[4:-1]
                nombre = self.parse_formula_arguments(arguments)
                nb = abs(nombre)
                logMessage("[evaluer_formule] Valeur absolue de " + str(nombre) + " = " + str(nb), indent=True)
                return nb
            elif re.fullmatch(self.formats["RACINE"], valeur):
                arguments = valeur[7:-1]
                nombre = self.parse_formula_arguments(arguments)
                logMessage("[evaluer_formule] Racine calculée : " + str(nombre))
                return float(nombre**5)
            elif re.fullmatch(self.formats["RACINE.DEG"], valeur):
                arguments = valeur[11:-1].split(";")
                nombre = self.parse_formula_arguments(arguments[0])
                degree = self.parse_formula_arguments(arguments[1])
                nb = nombre**(1/degree) if (type(nombre) == int or float) else "ERREUR TYPE DE CHIFFRE"
                logMessage("[evaluer_formule] Racine de degrée " + str(degree) + "calculée : " + str(nb), indent=True)
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
                return res
            elif re.fullmatch(self.formats["ERREUR"], valeur):
                a = "qzd" / "qzd"
            elif re.fullmatch(self.formats["BINOME"], valeur):
                arguments = valeur[6:-1].split(";")
                n = arguments[0]
                k = arguments[1]
                p = arguments[2]
                if p < 0 or p > 1:
                    return "ERREUR PROBABILITE INCORRECTE"
                res = math.comb(n, k) * p**k * (1-p)**(n-k)
                logMessage("[evaluer_formule] P(X) calculée pour loi binomiale de paramètres n = "+str(n)+", "+str(k)+" = "+str(k)+" et p = "+str(p)+" : "+str(res), indent=True)
                return res
            elif re.fullmatch(self.formats["PROD.SCAL"], valeur):
                arguments = valeur[9:-1].split(";")
                longueur1 = self.parse_formula_arguments(arguments[0])
                longueur2 = self.parse_formula_arguments(arguments[1])
                angle = self.parse_formula_arguments(arguments[2])
                res = longueur1 * longueur2 * math.cos(angle)
                logMessage("[evaluer_formule] Produit sclaire calculé pour les longeurs "+str(longueur1)+" & "+str(longueur2)+" et un angle de "+str(angle)+" : "+str(res), ident=True)
                return res
            elif re.fullmatch(self.formats["LIM"], valeur):
                arguments = valeur[3:-1].split(";")
                fonction = arguments[0]
                cible = arguments[1]
                limite = self.evaluer_limite(fonction, cible)
                logMessage("Limite calculée pour la fonction"+str(fonction)+" lorsque que "+str(VARIABLE_FONCTION)+" tend vers "+str(cible)+" : "+str(limite), indent=True)
                return limite
            elif re.fullmatch(self.formats["COULEUR"], valeur):
                valeur = valeur[8:-1]
                self.remplissage(valeur)
                return ""
            elif re.match(self.formats["SIMPLE"], valeur):
                row, col = self.cellule_index(valeur)
                logMessage(str((row, col)))
                return self.cellules.get((row, col), tk.Entry()).get() or 0
            else:
                return "#FORMULE?"
        except Exception as e:
            logMessage("[evaluer_formule] Erreur : " + str(e) + "\n" + str(traceback.format_exc()), typelog="error")
            messagebox.showerror("Erreur", "[evaluer_formule] Erreur : \n" + str(traceback.format_exc()))

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
            elif func_name == "LEN":
                res = self.evaluer_formule("LEN(" + arguments + ")")
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
        except Exception as e:
            logMessage("[evaluer_formule_combinee] Erreur dans l'évaluation : " + str(e), typelog="error")
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
                return self.evaluer_formule("="+cell1.strip()) >= self.evaluer_formule("="+cell2.strip())
            if "<" in condition:
                # Si la condition est de type "A1<=A2"
                cell1, cell2 = condition.split("<=")
                return self.evaluer_formule("="+cell1.strip()) <= self.evaluer_formule("="+cell2.strip())
            # Condition de type "A1=A2"
            cell1, cell2 = condition.split("=")
            return self.evaluer_formule("="+cell1.strip()) == self.evaluer_formule("="+cell2.strip())

        # Condition de type "A1=2"
        try:
            return float(condition) == self.evaluer_formule("="+condition.strip())
        except:
            return False

    def evaluer_limite(self, fonction, cible):
            """
            Évalue la limite d'une fonction dont la variable converge vers un point ou diverge vers l'infini.
            """
            if cible != "+inf" or "-inf":
                try:
                    point = self.parse_formula_arguments(cible)
                    res = sp.limit(fonction, VARIABLE_FONCTION, point)
                except Exception:
                    logMessage(f"[evaluer_limite] Erreur calcul limite : {traceback.format_exc()}")
                    res = "ERREUR"
            elif cible == "-inf":
                try:
                    res = sp.limit(fonction, VARIABLE_FONCTION, -sp.oo)
                except Exception:
                    logMessage(f"[evaluer_limite] Erreur calcul limite : {traceback.format_exc()}")
                    res = "ERREUR"
            elif cible == "+inf":
                try:
                    res = sp.limit(fonction, VARIABLE_FONCTION, sp.oo)
                except Exception:
                    logMessage(f"[evaluer_limite] Erreur calcul limite : {traceback.format_exc()}")
                    res = "ERREUR"
            else:
                res = "ERREUR"
            return res
    
    def parse_formula_arguments(self, arguments):
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
                        # logMessage("[parse_formula_arguments] Cellule "+str(start)+":"+str(end)+" = "+str(valeur), indent=True)
            elif re.match(r"^[A-Z][0-9]+$", arg):
                row, col = self.cellule_index(arg)
                valeur = self.cellules.get((row, col), tk.Entry()).get()
                res.append(float(valeur) if valeur else 0)
                # logMessage("[parse_formula_arguments] Cellule arg = "+str(valeur))
            elif re.match(r"^\d+(\.\d+)?$", arg):
                res.append(float(arg))
                # logMessage("[parse_formula_arguments] Constante arg")
        # logMessage("[parse_formula_arguments] " + arguments + " = " + res)
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
        coos = self.current_cell
        valeur = self.cellules[self.current_cell].get()
        messagebox.showinfo("Information sur la cellule.", "coordonnés : " + coos +"\nvaleur = " + valeur + "\nAutres : " + self.cellules[coos])

    def show_cells(self):
        cellules = []
        cell_text = "(ligne, colonne), valeur" + "\n"
        for cle, valeur in self.cellules.items():

            cellules.append((cle, self.cellules[cle].get()))
        for cell in cellules:
            cell_text += str(cell) + "\n"
        messagebox.showinfo("Cellules", cell_text)

    def clear(self):
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
                logMessage("[ouvrir] Erreur : " + str(e), "error")
                messagebox.showerror("Erreur", "Le fichier séléctionné est vide ou non pris en charge.")
            except ValueError as e:
                logMessage("[ouvrir] Erreur : " + str(e), "error")
                messagebox.showerror("Erreur", str(e))
            except Exception as e:
                logMessage("[ouvrir] Erreur : " + str(e), "error")
                messagebox.showerror("Erreur", "Une erreur est survenue " + str(traceback.format_exc()))

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
                logMessage("[enregistrer] Fichier sauvegardé avec succès : " + str(file_path))
                messagebox.showinfo("Sauvegarde", "Fichier sauvegardé au format CSV avec succès.")
            elif file_path.endswith(".xlsx"):
                self.update_dataframe()
                self.dataframe.to_excel(file_path, index=False)
                logMessage("[enregistrer] Fichier sauvegardé avec succès : " + str(file_path))
                messagebox.showinfo("Sauvegarde", "Fichier sauvegardé au format XLSX avec succès.")
        else:
            logMessage("[enregistrer] Erreur : chemin non valide.", typelog="error")
            logMessage(str(file_path), typelog="error", indent=1)
            pass

    def enregistrer_csv(self):
        """
        Permet d'enregistrer le fichier au format CSV.
        """
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichier CSV", "*.csv")])
        if file_path:
            self.update_dataframe()
            try:
                self.dataframe.to_csv(file_path, index=False)
                logMessage("[enregistrer_csv] Fichier sauvegardé avec succès : " + str(file_path))
                messagebox.showinfo("Sauvegarde", "Fichier CSV Sauvegardé avec succès")
            except Exception as e:
                logMessage("[enregistrer_csv] Erreur : " + str(traceback.format_exc()), typelog="error")
                messagebox.showerror("Erreur", "Une erreur est survenue lors de la sauvegarde du fichier : " +  str(traceback.format_exc()))

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
                logMessage("[enregistrer_xlsx] Fichier sauvegardé avec succès : " + str(file_path))
                messagebox.showinfo("Sauvegarde", "Fichier CSV sauvegardé avec succès")
            except Exception as e:
                logMessage("[enregistrer_xlsx] Erreur : " + str(traceback.format_exc()), typelog="error")
                messagebox.showerror("Erreur", "Une erreur est survenue lors de la sauvegarde du fichier : " + str(traceback.format_exc()))

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
            logMessage("[remplissage] la couleur de la cellule " + str(self.current_cell) + " a été changée en " +  str(couleur))
            self.cellules[(ligne, col)].config(fg=couleur)
        else:
            logMessage("[remplissage] Format couleur invalide : " + str(couleur), typelog="error")

    def couleur_police(self, couleur=None):
        """
        Change la couleur de la police vers la couleur donnée.
        """
        ligne, col = self.current_cell
        if not couleur:
            couleur = colorchooser.askcolor()[1]
        if re.fullmatch(self.formats["COULEURHEX"], couleur):
            logMessage("[couleur_police] la couleur de la police de la cellule "  + str(self.current_cell) + " a été changée en " + str(couleur))
            self.cellules[(ligne, col)].config(fg=couleur)
        else:
            logMessage("[couleur_police] Format couleur invalide : " + str(couleur), typelog="error")

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
            logMessage("[changer_police]Erreur : la cellule"+str((ligne, col))+"n'existe pas.")

    def populate_table(self):
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
            logMessage("La table est vide.", typelog="warning")
            messagebox.showwarning("Avertissement", "La table est vide. Veuillez ajouter des donnés avant de enregistrer.")
            return

        self.dataframe = pd.DataFrame(data, columns=[self.get_column_label(c) for c in range(cols)])

    def exit_program(self):
        reponse = messagebox.askquestion("Enregistrer les modifications ?", "Voulez vous enregistrer les modfications avant de quitter ?")
        if reponse == "yes":
            self.enregistrer()
        elif reponse == "no":
            logMessage("[exit_program] Arrêt.", typelog="critical")
            self.root.quit()
            self.root.destroy()

    def update_toolbar(self, valeur):
        """
        Met à jour la barre d'outils avec le contenu de la cellule et la formule.
        """
        self.cell_content_entry.delete(0, tk.END)
        self.cell_content_entry.insert(0, valeur)  # Affiche le contenu de la cellule

    def cell_selected(self, event):
        """
        Cette fonction est appelée lorsqu'une cellule est sélectionnée.
        """
        valeur = self.get_cell_value(event)
        self.update_toolbar(valeur)

    def get_cell_value(self, event):
        """
        Retourne la valeur d'une cellule. (Exemple simple ici).
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
        LEN() -> renvoie la longueur de la plage donnée
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
            utilisation = LIM(fonction;[point ou -inf / +inf])
        PROD.SCAL() -> calcule un produit scalaire à partir de 2 longueurs et un angle
            utilisation = PROD.SCAL(longueur1;longueur2;angle)
        ERREUR() -> provoque une erreur (utilisé pour le développement)
        HEURE() -> renvoie l'heure actuelle
        DATE() -> renvoie la date actuelle
        MAINTENANT() -> renvoie la date et l'heure actuelle
        EXP() -> renvoie la valeur donnée passée par la fonction exponentielle
        ABS() -> renvoie la valeur absolue de la valeur donnée.
        COULEUR() -> remplace la couleur de la cellule spécifiée par la couleur donnée
            utilisation = COULEUR(A1;#FFFFFF)

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
        texte = """
        PARAMETRES (affichage uniquement)

        NB_COLONNES = """+str(NB_COLONNES)+"""
        NB_LIGNES = """+str(NB_LIGNES)+"""
        COULEUR_EVIDENCE = """+str(COULEUR_EVIDENCE)+"""
        LOGGING_ENABLED = """+str(LOGGING_ENABLED)+"""
        VARIABLE_FONCTION = """+str(VARIABLE_FONCTION)+"""
        POLICE = """+str(POLICE)

        messagebox.showinfo("Paramètres du tableur", texte)

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

def logMessage(message: str, typelog="info", indent=False):
    """
    Renvoie un message de log dans le fichier log.
    Envoie également le message dans le terminal.
    """
    if LOGGING_ENABLED:
        if indent:
            message = "    " + message
        message = datetime.datetime.now().strftime("%w/%m/%y, %H:%M:%S,%f") + " : " + message
        if typelog == "info":
            LOG.info(message)
            print(message)
        elif typelog == "warning":
            LOG.warning(message)
            print(message)
        elif typelog == "error":
            LOG.error(message)
            print(message)
        elif typelog == "debug":
            LOG.debug(message)
            print(message)
        elif typelog == "critical":
            LOG.critical(message)
            print(message)
        else:
            logMessage("[# logMessage] Erreur : le typelog de log n'est pas reconnu : typelog="+str(typelog)+", type(typelog)="+type(typelog)+", indent="+str(indent), typelog="error")
            print("[# logMessage] Erreur : le typelog de log n'est pas reconnu : typelog="+typelog+", type(typelog)="+type(typelog)+", indent="+indent)

if __name__ == "__main__":
    root = tk.Tk()
    app = Tableur(root)
    root.mainloop()