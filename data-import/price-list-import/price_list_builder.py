"""
Aggiorna i prezzi di un listino esportato da ARCA Evolution con quelli di un
file Excel (codice articolo, prezzo) e produce il file pronto da reimportare.
"""
import pandas as pd
import os
import subprocess
import sys
import openpyxl
import tempfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import re
import numpy as np

class ListiniArcaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Creazione Listini Arca")
        self.root.geometry("700x500")
        self.root.configure(bg='#f0f0f0')
        
        self.file1_path = ""
        self.temp_folder = ""
        self.temp_file_path = ""
        
        self.setup_ui()
        self.show_welcome_message()
    
    def setup_ui(self):
        # Frame principale
        main_frame = tk.Frame(self.root, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Selezione primo file
        tk.Label(main_frame, text="Seleziona il Listino originale estratto da Arca:", 
                bg='#f0f0f0', font=('Arial', 10)).pack(anchor='w', pady=(0, 5))
        
        file1_frame = tk.Frame(main_frame, bg='#f0f0f0')
        file1_frame.pack(fill='x', pady=(0, 15))
        
        self.file1_entry = tk.Entry(file1_frame, font=('Arial', 10))
        self.file1_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        tk.Button(file1_frame, text="Sfoglia...", command=self.browse_file1,
                 bg='#e1e1e1', font=('Arial', 9)).pack(side='right')
        
        # Selezione secondo file
        tk.Label(main_frame, text="Inserisci il Listino creato su Excel:", 
                bg='#f0f0f0', font=('Arial', 10)).pack(anchor='w', pady=(0, 5))
        
        file2_frame = tk.Frame(main_frame, bg='#f0f0f0')
        file2_frame.pack(fill='x', pady=(0, 15))
        
        self.file2_entry = tk.Entry(file2_frame, font=('Arial', 10))
        self.file2_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        tk.Button(file2_frame, text="Sfoglia...", command=self.browse_file2,
                 bg='#e1e1e1', font=('Arial', 9)).pack(side='right')
        
        # Bottoni azione
        buttons_frame = tk.Frame(main_frame, bg='#f0f0f0')
        buttons_frame.pack(fill='x', pady=(0, 15))
        
        tk.Button(buttons_frame, text="Visualizza Anteprima", command=self.show_preview,
                 bg='#4CAF50', fg='white', font=('Arial', 10, 'bold'),
                 padx=15, pady=5).pack(side='left', padx=(0, 10))
        
        tk.Button(buttons_frame, text="Esegui la modifica", command=self.execute_modification,
                 bg='#2196F3', fg='white', font=('Arial', 10, 'bold'),
                 padx=15, pady=5).pack(side='left', padx=(0, 10))
        
        tk.Button(buttons_frame, text="Esci dal programma", command=self.exit_program,
                 bg='#f44336', fg='white', font=('Arial', 10, 'bold'),
                 padx=15, pady=5).pack(side='right')
        
        # Area di anteprima
        tk.Label(main_frame, text="Anteprima modifiche:", 
                bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor='w', pady=(0, 5))
        
        self.preview_text = scrolledtext.ScrolledText(main_frame, height=15, 
                                                     font=('Consolas', 9),
                                                     state='disabled')
        self.preview_text.pack(fill='both', expand=True)
    
    def show_welcome_message(self):
        welcome_msg = """Benvenuto in Creazione Listini Arca

Modalità d'uso dello script:

"Generare file Originale"
Avviare Arca, recarsi sul pannello di sinistra, cercare la voce 'Listini' -> 'Listini Standard' -> 'Manutenzione'.
Adesso premere sul pannello di sinistra 'Esporta Listino', premere sulla freccia destra in alto, selezionare il listino da esportare (qualora il listino non fosse presente bisogna prima crearlo!), adesso premere semplicemente su 'Esporta' in basso a destra.

"Copia nella directory principale dello script"
Adesso bisogna che sia il listino creato su excel che il file estratto con Arca siano entrambi nella stessa directory dello script.

"Creazione Listino"
Adesso basterà caricare il primo foglio originale estratto da Arca ed il secondo foglio creato su excel.
Premere su visualizza anteprima, controllare le modifiche e poi successivamente si potrà creare il foglio corretto."""
        
        messagebox.showinfo("Creazione Listini Arca", welcome_msg)
    
    def browse_file1(self):
        filename = filedialog.askopenfilename(
            title="Seleziona il Listino originale estratto da Arca",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")]
        )
        if filename:
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, filename)
    
    def browse_file2(self):
        filename = filedialog.askopenfilename(
            title="Seleziona il Listino creato su Excel",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")]
        )
        if filename:
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, filename)
    
    def show_preview(self):
        file1_path = self.file1_entry.get()
        file2_path = self.file2_entry.get()
        
        if not file1_path or not file2_path:
            messagebox.showerror("Errore", "Seleziona entrambi i file prima di continuare.")
            return
        
        if not os.path.exists(file1_path) or not os.path.exists(file2_path):
            messagebox.showerror("Errore", "Uno o entrambi i file selezionati non esistono.")
            return
        
        try:
            self.file1_path = file1_path
            self.temp_folder = os.path.dirname(file1_path)
            
            df1 = pd.read_excel(file1_path)  # Listino estratto da ARCA
            df2 = pd.read_excel(file2_path, header=None)  # Listino creato su Excel (SENZA header)
            
            wb = openpyxl.load_workbook(file1_path)
            ws = wb.active
            
            preview_text = ""
            
            for index, row in df1.iterrows():
                valore_colonna_df1 = row[df1.columns[0]]
                
                # Pulizia del valore (rimuove spazi)
                valore_pulito = str(valore_colonna_df1).strip()
                
                # Cerca con strip per gestire spazi nascosti
                matching_row_df2 = df2[df2[df2.columns[0]].astype(str).str.strip() == valore_pulito]
                
                if not matching_row_df2.empty:
                    prezzo_stringa = matching_row_df2.iloc[0][df2.columns[1]]
                    
                    is_string = isinstance(prezzo_stringa, str)
                    if not is_string:
                        prezzo_stringa = f'€ {round(float(prezzo_stringa), 2)}'
                    elif is_string:
                        prezzo_stringa = float(prezzo_stringa.replace('€', '').strip().replace(',', '.'))
                        prezzo_stringa = f'€ {round(float(prezzo_stringa), 2)}'
                    
                    match = re.search(r'\€\s*([\d.,]+)', prezzo_stringa)
                    if match:
                        prezzo_stringa = prezzo_stringa.replace("€ ", "").replace(",", ".")
                        prezzo_corrispondente = float(prezzo_stringa)
                        ws.cell(row=index+2, column=5, value=round(prezzo_corrispondente, 2))
                        preview_text += f"Riga {index+2}: Prezzo modificato a {round(prezzo_corrispondente, 2)}\n"
            
            self.temp_file_path = os.path.join(self.temp_folder, "temporaneo.xlsx")
            
            wb.save(self.temp_file_path)
            wb.close()
            
            # Aggiorna l'area di anteprima
            self.preview_text.config(state='normal')
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, preview_text)
            self.preview_text.config(state='disabled')
            
            # Apri il file temporaneo con il programma predefinito
            try:
                if sys.platform.startswith("win"):
                    os.startfile(self.temp_file_path)
                elif sys.platform.startswith("darwin"):  # macOS
                    subprocess.run(["open", self.temp_file_path], check=False)
                else:  # Linux
                    subprocess.run(["xdg-open", self.temp_file_path], check=False)
            except Exception:
                messagebox.showwarning(
                    "Attenzione",
                    "Impossibile aprire automaticamente il file. Aprilo manualmente."
                )
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'elaborazione: {str(e)}")
    
    def execute_modification(self):
        if not self.temp_file_path:
            messagebox.showerror("Errore", "Devi prima visualizzare l'anteprima.")
            return
        
        try:
            shutil.copy(self.temp_file_path, os.path.join(self.temp_folder, "Listino_pronto.xlsx"))
            os.remove(self.temp_file_path)  # Rimuovi il file temporaneo creato
            self.temp_file_path = ""  # Resetta il percorso del file temporaneo
            messagebox.showinfo("Successo", "Modifiche effettuate con successo.")
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante il salvataggio: {str(e)}")
    
    def exit_program(self):
        # Pulisci eventuali file temporanei prima di uscire
        if self.temp_file_path and os.path.exists(self.temp_file_path):
            try:
                os.remove(self.temp_file_path)
            except:
                pass
        
        self.root.quit()

def main():
    root = tk.Tk()
    app = ListiniArcaApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()