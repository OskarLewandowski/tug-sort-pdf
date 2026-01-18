import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz  # PyMuPDF
import csv
import re
import os
import threading

class PDFSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sorter Dokumentów PDF + Podmiana")
        self.root.geometry("500x600")
        
        # Zmienne
        self.pdf_path = tk.StringVar()
        self.csv_path = tk.StringVar()
        self.regex_pattern = tk.StringVar()
        self.replace_placeholder = tk.StringVar()
        self.sort_order = tk.StringVar(value="ascending")
        self.pages_per_doc = tk.IntVar(value=1)
        
        self.create_widgets()

    def create_widgets(self):
        pwr = {'padx': 10, 'pady': 5, 'sticky': 'w'}
        
        tk.Label(self.root, text="1. Wybierz plik PDF:").grid(row=0, column=0, **pwr)
        tk.Entry(self.root, textvariable=self.pdf_path, width=30).grid(row=0, column=1, **pwr)
        tk.Button(self.root, text="...", command=self.browse_pdf).grid(row=0, column=2, **pwr)
        
        tk.Label(self.root, text="2. Wybierz plik CSV:").grid(row=1, column=0, **pwr)
        tk.Entry(self.root, textvariable=self.csv_path, width=30).grid(row=1, column=1, **pwr)
        tk.Button(self.root, text="...", command=self.browse_csv).grid(row=1, column=2, **pwr)
        
        tk.Label(self.root, text="3. Sortuj po nazwie (Regex):").grid(row=2, column=0, **pwr)
        tk.Entry(self.root, textvariable=self.regex_pattern, width=30).grid(row=2, column=1, columnspan=2, **pwr)
        
        tk.Label(self.root, text="4. Zastąp (Placeholder w PDF):").grid(row=3, column=0, **pwr)
        tk.Entry(self.root, textvariable=self.replace_placeholder, width=30).grid(row=3, column=1, columnspan=2, **pwr)
        
        tk.Label(self.root, text="   (np.: NR_SPRAWY)").grid(row=4, column=1, sticky='w', padx=10)

        tk.Label(self.root, text="5. Sortowanie:").grid(row=5, column=0, **pwr)
        frame_sort = tk.Frame(self.root)
        frame_sort.grid(row=5, column=1, columnspan=2, sticky='w')
        tk.Radiobutton(frame_sort, text="Rosnąco", variable=self.sort_order, value="ascending").pack(side='left')
        tk.Radiobutton(frame_sort, text="Malejąco", variable=self.sort_order, value="descending").pack(side='left')
        
        tk.Label(self.root, text="6. Ilość stron (jeden dok.):").grid(row=6, column=0, **pwr)
        tk.Entry(self.root, textvariable=self.pages_per_doc, width=10).grid(row=6, column=1, **pwr)
        
        self.btn_start = tk.Button(self.root, text="START", font=('Arial', 12, 'bold'), command=self.start_thread, height=2, width=20)
        self.btn_start.grid(row=7, column=0, columnspan=3, pady=20)
        
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=450, mode="determinate")
        self.progress.grid(row=8, column=0, columnspan=3, padx=20, pady=10)
        
        self.status_label = tk.Label(self.root, text="Gotowy")
        self.status_label.grid(row=9, column=0, columnspan=3)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename: self.pdf_path.set(filename)

    def browse_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt")])
        if filename: self.csv_path.set(filename)

    def start_thread(self):
        threading.Thread(target=self.process_documents).start()

    def process_documents(self):
        pdf_file = self.pdf_path.get()
        csv_file = self.csv_path.get()
        pattern = self.regex_pattern.get()
        placeholder = self.replace_placeholder.get() # np. NUMER_GO123
        pages_count = self.pages_per_doc.get()
        
        if not pdf_file or not os.path.exists(pdf_file):
            messagebox.showerror("Błąd", "Wybierz poprawny plik PDF.")
            return
            
        try:
            self.btn_start.config(state='disabled')
            self.status_label.config(text="Analizowanie PDF...")
            
            doc = fitz.open(pdf_file)
            total_pages = len(doc)
            
            # --- Etap 1: Analiza i Sortowanie ---
            documents_metadata = []
            total_docs = (total_pages + pages_count - 1) // pages_count
            self.progress["maximum"] = total_docs
            self.progress["value"] = 0

            for i in range(total_docs):
                start_page = i * pages_count
                end_page = min(start_page + pages_count, total_pages)
                
                # Pobranie tekstu do sortowania
                page_text = doc[start_page].get_text()
                sort_key = ""
                
                if pattern:
                    match = re.search(pattern, page_text)
                    if match:
                        sort_key = match.group(0)
                    else:
                        sort_key = "ZZZ_BRAK" # Na koniec listy
                
                documents_metadata.append({
                    'index': i,
                    'start_page': start_page,
                    'end_page': end_page,
                    'sort_key': sort_key
                })
                self.progress["value"] += 0.5

            # Wykonaj sortowanie
            reverse_sort = (self.sort_order.get() == "descending")
            documents_metadata.sort(key=lambda x: x['sort_key'], reverse=reverse_sort)
            
            # --- Etap 2: Wczytanie danych z CSV ---
            replacements = []
            if csv_file and os.path.exists(csv_file):
                try:
                    # Próba UTF-8 (standard), potem CP1250 (Excel Windows)
                    try:
                        with open(csv_file, newline='', encoding='utf-8') as f:
                            reader = csv.reader(f)
                            replacements = [row[0] for row in reader if row]
                    except UnicodeDecodeError:
                        with open(csv_file, newline='', encoding='cp1250') as f:
                            reader = csv.reader(f)
                            replacements = [row[0] for row in reader if row]
                except Exception as e:
                    print(f"Błąd czytania CSV: {e}")

            # --- Etap 3: Generowanie PDF i Podmiana ---
            self.status_label.config(text="Generowanie i podmienianie danych...")
            output_doc = fitz.open()
            
            for idx, meta in enumerate(documents_metadata):
                # Kopiuj strony
                output_doc.insert_pdf(doc, from_page=meta['start_page'], to_page=meta['end_page']-1)
                
                # Jeśli mamy wartość do podmiany dla tego dokumentu
                if placeholder and idx < len(replacements):
                    new_value = str(replacements[idx]) # np. SA/101/2024
                    
                    # Obliczamy zakres stron w NOWYM pliku, które przed chwilą dodaliśmy
                    current_len = len(output_doc)
                    doc_len = meta['end_page'] - meta['start_page']
                    start_new_idx = current_len - doc_len
                    
                    # Szukamy placeholdera na stronach tego konkretnego dokumentu
                    for p_num in range(start_new_idx, current_len):
                        page = output_doc[p_num]
                        
                        # Znajdź wszystkie wystąpienia NUMER_GO123
                        hits = page.search_for(placeholder)
                        
                        for rect in hits:
                            # 1. Zasłoń stary tekst białym prostokątem
                            page.add_redact_annot(rect, fill=(1, 1, 1))
                            page.apply_redactions()
                            
                            # 2. Wstaw nowy tekst
                            # Używamy rect.x0 (lewa krawędź) i rect.y1 (dolna krawędź - linia pisma)
                            # Odejmujemy 2 punkty od wysokości, żeby tekst ładnie "siedział" na linii
                            insert_point = fitz.Point(rect.x0, rect.y1 - 2)
                            
                            page.insert_text(
                                insert_point, 
                                new_value, 
                                fontsize=12,    # Rozmiar czcionki (dostosuj jeśli za mała/duża)
                                color=(0, 0, 0) # Czarny kolor
                            )

                self.progress["value"] = (total_docs / 2) + ((idx + 1) / total_docs * (total_docs / 2))

            # --- Zapis ---
            output_filename = pdf_file.replace(".pdf", "_gotowy.pdf")
            output_doc.save(output_filename)
            output_doc.close()
            doc.close()
            
            self.status_label.config(text=f"Sukces!")
            messagebox.showinfo("Gotowe", f"Plik zapisany jako:\n{output_filename}")
            
        except Exception as e:
            messagebox.showerror("Błąd", str(e))
            print(e)
        finally:
            self.btn_start.config(state='normal')
            self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFSorterApp(root)
    root.mainloop()