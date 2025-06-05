import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import pyodbc

SERVER = r'.\SQLEXPRESS' # nom/chemin du serveur à changer
DATABASE = 'db_stloup-exceptions' # nom de la bdd à changer 


def connect_to_sqlserver():
    try:
        conn_str = (
            "Driver={ODBC Driver 17 for SQL Server};"
            f"Server={SERVER};"
            f"Database={DATABASE};"
            "Trusted_Connection=yes;"
            "autocommit=True"
        )
        return pyodbc.connect(conn_str)
    except pyodbc.Error as e:
        messagebox.showerror("Erreur connexion SQL", f"Impossible de se connecter à la base : {e}")
        return None



def create_tables():
    conn = connect_to_sqlserver()
    cursor = conn.cursor()
    cursor.execute("""
    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'saint_loup')
    CREATE TABLE saint_loup (
        id INT IDENTITY PRIMARY KEY,
        code VARCHAR(13) UNIQUE NOT NULL
    )
    """)
    cursor.execute("""
    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'exceptions')
    CREATE TABLE exceptions (
        id INT IDENTITY PRIMARY KEY,
        code VARCHAR(13) UNIQUE NOT NULL
    )
    """)
    conn.commit()
    conn.close()


def fetch_all_codes(table_name):
    conn = connect_to_sqlserver()
    cursor = conn.cursor()
    cursor.execute(f"SELECT id, code FROM {table_name} ORDER BY code")
    rows = cursor.fetchall()
    conn.close()
    return rows


def insert_codes_with_validation(lines, progress_callback=None):
    inserted = 0
    duplicates = 0
    errors = []
    total = len(lines)
    conn = connect_to_sqlserver()
    cursor = conn.cursor()

    for i, line in enumerate(lines):
        code = line.strip()

        # Ignorer les lignes vides
        if not code:
            if progress_callback:
                progress_callback((i+1) * 100 / total)
            continue

        if len(code) < 13:
            errors.append(f"Ligne {i+1}: code trop court '{code}'")
            if progress_callback:
                progress_callback((i+1) * 100 / total)
            continue
        code = code[:13]
        if not code.isdigit():
            errors.append(f"Ligne {i+1}: code non numérique '{code}'")
            if progress_callback:
                progress_callback((i+1) * 100 / total)
            continue

        try:
            table = "saint_loup" if code.startswith("348094") else "exceptions"
            cursor.execute(f"INSERT INTO {table} (code) VALUES (?)", code)
            inserted += 1
        except pyodbc.IntegrityError as e:
            if "UNIQUE" in str(e).upper():
                duplicates += 1
            else:
                errors.append(f"Ligne {i+1}: {str(e)}")
        except Exception as e:
            errors.append(f"Ligne {i+1}: {str(e)}")

        if progress_callback:
            progress_callback((i+1) * 100 / total)

    conn.commit()
    conn.close()
    return inserted, duplicates, errors



class App(TkinterDnD.Tk):
    ITEMS_PER_PAGE = 50

    def __init__(self):
        super().__init__()
        self.title("Import codes vers SQL Server")
        self.geometry("1280x720")
        self.configure(bg="#f0f4f8")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"), background="#1E90FF", foreground="white")
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)

        btn_frame = tk.Frame(self, bg="#f0f4f8")
        btn_frame.grid(row=0, column=0, sticky="ew", pady=8, padx=10)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        self.import_btn = tk.Button(btn_frame, text="Importer fichier texte", command=self.load_file,
                                    font=("Segoe UI", 14, "bold"), bg="#1E90FF", fg="white",
                                    activebackground="#104E8B", activeforeground="white",
                                    relief="flat", padx=15, pady=8)
        self.import_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.export_btn = tk.Button(btn_frame, text="Exporter les codes", command=self.export_to_file,
                                    font=("Segoe UI", 14, "bold"), bg="#32CD32", fg="white",
                                    activebackground="#228B22", activeforeground="white",
                                    relief="flat", padx=15, pady=8)
        self.export_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.tabs = ttk.Notebook(self)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.tab_saintloup = ttk.Frame(self.tabs)
        self.tab_exceptions = ttk.Frame(self.tabs)
        self.tabs.add(self.tab_saintloup, text="Saint-Loup (348094...)")
        self.tabs.add(self.tab_exceptions, text="Exceptions")

        self.tree_sl = self.create_treeview(self.tab_saintloup, "saint_loup")
        self.tree_ex = self.create_treeview(self.tab_exceptions, "exceptions")

        nav_frame_sl = tk.Frame(self.tab_saintloup)
        nav_frame_sl.pack(fill='x', pady=5)
        self.prev_btn_sl = tk.Button(nav_frame_sl, text="← Précédent", command=lambda: self.change_page("sl", -1))
        self.prev_btn_sl.pack(side='left', padx=5)
        self.page_label_sl = tk.Label(nav_frame_sl, text="Page 1")
        self.page_label_sl.pack(side='left', padx=5)

        self.page_entry_sl = tk.Entry(nav_frame_sl, width=5)
        self.page_entry_sl.pack(side='left', padx=(10, 0))
        self.go_btn_sl = tk.Button(nav_frame_sl, text="Aller", command=lambda: self.goto_page("sl"))
        self.go_btn_sl.pack(side='left', padx=5)

        self.next_btn_sl = tk.Button(nav_frame_sl, text="Suivant →", command=lambda: self.change_page("sl", 1))
        self.next_btn_sl.pack(side='left', padx=5)

        nav_frame_ex = tk.Frame(self.tab_exceptions)
        nav_frame_ex.pack(fill='x', pady=5)
        self.prev_btn_ex = tk.Button(nav_frame_ex, text="← Précédent", command=lambda: self.change_page("ex", -1))
        self.prev_btn_ex.pack(side='left', padx=5)
        self.page_label_ex = tk.Label(nav_frame_ex, text="Page 1")
        self.page_label_ex.pack(side='left', padx=5)

        self.page_entry_ex = tk.Entry(nav_frame_ex, width=5)
        self.page_entry_ex.pack(side='left', padx=(10, 0))
        self.go_btn_ex = tk.Button(nav_frame_ex, text="Aller", command=lambda: self.goto_page("ex"))
        self.go_btn_ex.pack(side='left', padx=5)

        self.next_btn_ex = tk.Button(nav_frame_ex, text="Suivant →", command=lambda: self.change_page("ex", 1))
        self.next_btn_ex.pack(side='left', padx=5)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate")
        self.progress.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))

        self.status_label = tk.Label(self, text="", anchor="w", font=("Segoe UI", 10), bg="#f0f4f8")
        self.status_label.grid(row=3, column=0, sticky="ew", padx=10)

        self.log_text = tk.Text(self, height=8, font=("Consolas", 9), state="disabled",
                                bg="#ffffff", fg="#333333", relief="sunken", borderwidth=1)
        self.log_text.grid(row=4, column=0, sticky="nsew", padx=10, pady=(5, 10))
        self.grid_rowconfigure(4, weight=1)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.drop_event)

        self.current_file_path = None


        self.current_page_sl = 0
        self.current_page_ex = 0
        self.all_codes_sl = []
        self.all_codes_ex = []

        create_tables()
        self.refresh_tables()

    def create_treeview(self, parent, table_name):
        cols = ("id", "code")
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        tree.heading("id", text="ID")
        tree.heading("code", text="Code")
        tree.column("id", width=50, anchor='center')
        tree.column("code", width=300, anchor='center')
        tree.pack(expand=True, fill='both')

        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Modifier", command=lambda t=table_name, tr=tree: self.edit_code(t, tr))
        menu.add_command(label="Supprimer", command=lambda t=table_name, tr=tree: self.delete_code(t, tr))

        def popup(event):
            if tree.selection():
                menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Button-3>", popup)
        return tree

    def refresh_tables(self):
        self.all_codes_sl = fetch_all_codes("saint_loup")
        self.all_codes_ex = fetch_all_codes("exceptions")
        self.current_page_sl = 0
        self.current_page_ex = 0
        self.update_treeview("sl")
        self.update_treeview("ex")

    def update_treeview(self, table_key):
        if table_key == "sl":
            tree = self.tree_sl
            all_codes = self.all_codes_sl
            current_page = self.current_page_sl
            page_label = self.page_label_sl
            prev_btn = self.prev_btn_sl
            next_btn = self.next_btn_sl
        else:
            tree = self.tree_ex
            all_codes = self.all_codes_ex
            current_page = self.current_page_ex
            page_label = self.page_label_ex
            prev_btn = self.prev_btn_ex
            next_btn = self.next_btn_ex

        tree.delete(*tree.get_children())

        start = current_page * self.ITEMS_PER_PAGE
        end = start + self.ITEMS_PER_PAGE
        page_codes = all_codes[start:end]

        for item in page_codes:
            tree.insert("", "end", values=(item[0], item[1]))

        total_pages = max(1, (len(all_codes) + self.ITEMS_PER_PAGE - 1) // self.ITEMS_PER_PAGE)
        page_label.config(text=f"Page {current_page + 1} / {total_pages}")

        prev_btn.config(state="normal" if current_page > 0 else "disabled")
        next_btn.config(state="normal" if current_page < total_pages - 1 else "disabled")

    def change_page(self, table_key, delta):
        if table_key == "sl":
            new_page = self.current_page_sl + delta
            max_page = (len(self.all_codes_sl) - 1) // self.ITEMS_PER_PAGE
            if 0 <= new_page <= max_page:
                self.current_page_sl = new_page
                self.update_treeview("sl")
        else:
            new_page = self.current_page_ex + delta
            max_page = (len(self.all_codes_ex) - 1) // self.ITEMS_PER_PAGE
            if 0 <= new_page <= max_page:
                self.current_page_ex = new_page
                self.update_treeview("ex")

    def goto_page(self, table_key):
        if table_key == "sl":
            try:
                page = int(self.page_entry_sl.get()) - 1
            except ValueError:
                messagebox.showerror("Erreur", "Entrez un numéro de page valide.")
                return
            max_page = (len(self.all_codes_sl) - 1) // self.ITEMS_PER_PAGE
            if 0 <= page <= max_page:
                self.current_page_sl = page
                self.update_treeview("sl")
            else:
                messagebox.showerror("Erreur", "Numéro de page hors limites.")
        else:
            try:
                page = int(self.page_entry_ex.get()) - 1
            except ValueError:
                messagebox.showerror("Erreur", "Entrez un numéro de page valide.")
                return
            max_page = (len(self.all_codes_ex) - 1) // self.ITEMS_PER_PAGE
            if 0 <= page <= max_page:
                self.current_page_ex = page
                self.update_treeview("ex")
            else:
                messagebox.showerror("Erreur", "Numéro de page hors limites.")

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner un fichier texte",
            filetypes=[("Fichiers texte", "*.txt"), ("Tous les fichiers", "*.*")]
        )
        if filepath:
            self.current_file_path = filepath
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            answer = messagebox.askyesno("Confirmation", f"Importer {len(lines)} codes depuis '{filepath}' ?")
            if answer:
                self.import_btn.config(state="disabled")
                self.status_label.config(text="Import en cours...")
                threading.Thread(target=self.import_codes_thread, args=(lines,)).start()

    def drop_event(self, event):
        files = self.split_drop_files(event.data)
        if files:
            filepath = files[0]
            if filepath.lower().endswith(".txt"):
                self.current_file_path = filepath
                with open(filepath, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                answer = messagebox.askyesno("Confirmation", f"Importer {len(lines)} codes depuis '{filepath}' ?")
                if answer:
                    self.import_btn.config(state="disabled")
                    self.status_label.config(text="Import en cours...")
                    threading.Thread(target=self.import_codes_thread, args=(lines,)).start()
            else:
                messagebox.showerror("Erreur", "Seuls les fichiers .txt sont acceptés.")

    @staticmethod
    def split_drop_files(data):
        if data.startswith('{') and data.endswith('}'):
            data = data[1:-1]
        return data.split()

    def import_codes_thread(self, lines):
        def progress_callback(percent):
            self.progress['value'] = percent

        inserted, duplicates, errors = insert_codes_with_validation(lines, progress_callback)
        self.progress['value'] = 0
        self.import_btn.config(state="normal")

        self.status_label.config(text=f"Import terminé: {inserted} insérés, {duplicates} doublons.")
        self.log_text.config(state="normal")
        if errors:
            self.log_text.insert("end", "Erreurs détectées :\n" + "\n".join(errors) + "\n\n")
        else:
            self.log_text.insert("end", "Aucune erreur détectée.\n\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

        self.refresh_tables()

    def edit_code(self, table_name, tree):
        selected = tree.selection()
        if not selected:
            return
        item_id = selected[0]
        current_values = tree.item(item_id, "values")
        current_code = current_values[1]

        new_code = simpledialog.askstring("Modifier le code", "Entrez le nouveau code 13 chiffres :", initialvalue=current_code)
        if new_code is None:
            return
        new_code = new_code.strip()
        if len(new_code) != 13 or not new_code.isdigit():
            messagebox.showerror("Erreur", "Le code doit contenir exactement 13 chiffres.")
            return

        conn = connect_to_sqlserver()
        cursor = conn.cursor()
        try:
            cursor.execute(f"UPDATE {table_name} SET code=? WHERE id=?", new_code, current_values[0])
            conn.commit()
            self.status_label.config(text=f"Code modifié : {current_code} → {new_code}")
        except pyodbc.IntegrityError:
            messagebox.showerror("Erreur", "Le code existe déjà dans la base.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur SQL : {str(e)}")
        finally:
            conn.close()
        self.refresh_tables()

    def delete_code(self, table_name, tree):
        selected = tree.selection()
        if not selected:
            return
        item_id = selected[0]
        current_values = tree.item(item_id, "values")
        code = current_values[1]
        answer = messagebox.askyesno("Confirmation", f"Voulez-vous vraiment supprimer le code {code} ?")
        if answer:
            conn = connect_to_sqlserver()
            cursor = conn.cursor()
            try:
                cursor.execute(f"DELETE FROM {table_name} WHERE id=?", current_values[0])
                conn.commit()
                self.status_label.config(text=f"Code supprimé : {code}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur SQL : {str(e)}")
            finally:
                conn.close()
            self.refresh_tables()

    def export_to_file(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Fichiers texte", "*.txt"), ("Tous les fichiers", "*.*")],
            title="Enregistrer les codes exportés"
        )
        if not filepath:
            return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                for _, code in self.all_codes_sl:
                    f.write(code + '\n')
                for _, code in self.all_codes_ex:
                    f.write(code + '\n')
            self.status_label.config(text=f"Export réussi vers {filepath}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'export : {str(e)}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
