import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import pandas as pd



# fonction pour supprimer une ligne
def delete_row():
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        id_to_delete = item['values'][0]
        cursor.execute("DELETE FROM data WHERE id=?", (id_to_delete,))
        conn.commit()
        load_data()
    else:
        messagebox.showwarning("Sélectionner une ligne", "Veuillez sélectionner une ligne à supprimer.")

# fonction export excel
def export_to_excel():
    df = pd.read_sql_query("SELECT * FROM data", conn)
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Exportation réussie", f"Les données ont été exportées vers {file_path}")

# fonction somme (voir somme)
def calculate_sum():
    selected_items = tree.selection()
    if selected_items:
        total = 0
        for item in selected_items:
            values = tree.item(item, 'values')
            total += sum([float(v) for v in values[1:] if v])
        messagebox.showinfo("Somme", f"La somme des valeurs sélectionnées est: {total}")
    else:
        messagebox.showwarning("Sélectionner des lignes", "Veuillez sélectionner des lignes pour calculer la somme.")

# fonction moyenne (voir moyenne)
def calculate_avg():
    selected_items = tree.selection()
    if selected_items:
        total = 0
        count = 0
        for item in selected_items:
            values = tree.item(item, 'values')
            numeric_values = [float(v) for v in values[1:] if v]
            total += sum(numeric_values)
            count += len(numeric_values)
        if count > 0:
            average = total / count
            messagebox.showinfo("Moyenne", f"La moyenne des valeurs sélectionnées est: {average}")
        else:
            messagebox.showwarning("Aucune valeur numérique", "Aucune valeur numérique sélectionnée.")
    else:
        messagebox.showwarning("Sélectionner des lignes", "Veuillez sélectionner des lignes pour calculer la moyenne.")

# selectionner une base données (voir ouvir un fichier)
def select_database():
    global conn, cursor
    db_path = filedialog.askopenfilename(filetypes=[("SQLite Database Files", "*.db"), ("All Files", "*.*")])
    if db_path:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        load_data()

# fonction pour charger les données
def load_data():
    for row in tree.get_children():
        tree.delete(row)
    cursor.execute("SELECT * FROM data")
    rows = cursor.fetchall()
    for row in rows:
        tree.insert("", tk.END, values=row)

# fonction pour ajouter des données
def add_data():
    cursor.execute("INSERT INTO data (col1, col2, col3) VALUES (?, ?, ?)",
                   (entry1.get(), entry2.get(), entry3.get()))
    conn.commit()
    load_data()

# créer la base de données sql
def create_database():
    global conn, cursor
    db_path = filedialog.asksaveasfilename(defaultextension=".db",
                                           filetypes=[("SQLite Database Files", "*.db"), ("All Files", "*.*")])
    if db_path:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            col1 REAL,
            col2 REAL,
            col3 REAL
        )
        ''')
        conn.commit()
        load_data()
        messagebox.showinfo("Base de données créée", f"La nouvelle base de données a été créée à {db_path}")

# fenetre tkinter
root = tk.Tk()
root.title("Tableur")

# bouton pour ouvrir un fichier
select_db_button = tk.Button(root, text="Sélectionner la base de données", command=select_database)
select_db_button.grid(row=0, column=0, columnspan=2)

# bouton pour créer un nouveau fichier
create_db_button = tk.Button(root, text="Créer une nouvelle base de données", command=create_database)
create_db_button.grid(row=0, column=2, columnspan=2)

# entrées
entry1 = tk.Entry(root)
entry2 = tk.Entry(root)
entry3 = tk.Entry(root)
entry1.grid(row=1, column=1)
entry2.grid(row=2, column=1)
entry3.grid(row=3, column=1)

# colonnes
tk.Label(root, text="Colonne 1").grid(row=1, column=0)
tk.Label(root, text="Colonne 2").grid(row=2, column=0)
tk.Label(root, text="Colonne 3").grid(row=3, column=0)

# ajouter des données
add_button = tk.Button(root, text="Ajouter", command=add_data)
add_button.grid(row=4, column=1)

# suppr ligne
delete_row_button = tk.Button(root, text="Supprimer la ligne", command=delete_row)
delete_row_button.grid(row=4, column=2)

# export excel
export_button = tk.Button(root, text="Exporter vers Excel", command=export_to_excel)
export_button.grid(row=4, column=3)

# somme
sum_button = tk.Button(root, text="Calculer la somme", command=calculate_sum)
sum_button.grid(row=4, column=4)

# moyenne
avg_button = tk.Button(root, text="Calculer la moyenne", command=calculate_avg)
avg_button.grid(row=4, column=5)

# afficher les données
columns = ("id", "col1", "col2", "col3")
tree = ttk.Treeview(root, columns=columns, show="headings")
tree.heading("id", text="ID")
tree.heading("col1", text="Colonne 1")
tree.heading("col2", text="Colonne 2")
tree.heading("col3", text="Colonne 3")
tree.grid(row=5, column=0, columnspan=5)

# boucle principale
root.mainloop()

# 
if 'conn' in globals():
    conn.close()
