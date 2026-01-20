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
       self.root.title("PDF Sorter")
       self.root.geometry("500x400")
       self.root.minsize(500, 400)
       self.root.maxsize(500, 400)


       # Zmienne do wyświetlania w GUI (tylko nazwa pliku)
       self.pdf_display = tk.StringVar()
       self.csv_display = tk.StringVar()


       # Zmienne logiczne (pełne ścieżki do plików, ukryte przed wzrokiem)
       self.full_pdf_path = ""
       self.full_csv_path = ""
       # --- ZMIANA: Zmienna na ścieżkę zapisu ---
       self.full_save_path = ""


       self.regex_pattern = tk.StringVar()
       self.replace_placeholder = tk.StringVar()
       self.sort_order = tk.StringVar(value="ascending")
       self.pages_per_doc = tk.IntVar(value=1)


       self.create_widgets()


   def create_widgets(self):
       pwr = {'padx': 10, 'pady': 5, 'sticky': 'w'}


       tk.Label(self.root, text="1. Wybierz plik PDF:").grid(row=0, column=0, **pwr)
       tk.Entry(self.root, textvariable=self.pdf_display, width=30, state='readonly').grid(row=0, column=1, **pwr)
       tk.Button(self.root, text="Wybierz plik...", command=self.browse_pdf).grid(row=0, column=2, **pwr)


       tk.Label(self.root, text="2. Wybierz plik CSV:").grid(row=1, column=0, **pwr)
       tk.Entry(self.root, textvariable=self.csv_display, width=30, state='readonly').grid(row=1, column=1, **pwr)
       tk.Button(self.root, text="Wybierz plik...", command=self.browse_csv).grid(row=1, column=2, **pwr)


       tk.Label(self.root, text="3. Sortuj po nazwie (Regex):").grid(row=2, column=0, **pwr)
       tk.Entry(self.root, textvariable=self.regex_pattern, width=30).grid(row=2, column=1, columnspan=2, **pwr)


       tk.Label(self.root, text="4. Zastąp (Placeholder):").grid(row=3, column=0, **pwr)
       tk.Entry(self.root, textvariable=self.replace_placeholder, width=30).grid(row=3, column=1, columnspan=2, **pwr)


       tk.Label(self.root, text="5. Sortowanie:").grid(row=4, column=0, **pwr)
       frame_sort = tk.Frame(self.root)
       frame_sort.grid(row=4, column=1, columnspan=2, sticky='w')
       tk.Radiobutton(frame_sort, text="Rosnąco", variable=self.sort_order, value="ascending").pack(side='left')
       tk.Radiobutton(frame_sort, text="Malejąco", variable=self.sort_order, value="descending").pack(side='left')


       tk.Label(self.root, text="6. Ilość stron (jeden dok.):").grid(row=5, column=0, **pwr)
       tk.Entry(self.root, textvariable=self.pages_per_doc, width=10).grid(row=5, column=1, **pwr)


       self.btn_start = tk.Button(self.root, text="START", font=('Arial', 12, 'bold'), command=self.start_thread,
                                  height=2, width=20)
       self.btn_start.grid(row=7, column=0, columnspan=3, pady=20)


       self.progress = ttk.Progressbar(self.root, orient="horizontal", length=450, mode="determinate")
       self.progress.grid(row=8, column=0, columnspan=3, padx=20, pady=10)


       self.status_label = tk.Label(self.root, text="Gotowy")
       self.status_label.grid(row=9, column=0, columnspan=3)


   def browse_pdf(self):
       filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
       if filename:
           self.full_pdf_path = filename
           self.pdf_display.set(os.path.basename(filename))


   def browse_csv(self):
       filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt")])
       if filename:
           self.full_csv_path = filename
           self.csv_display.set(os.path.basename(filename))


   def start_thread(self):
       # --- ZMIANA: Pytamy o plik wyjściowy PRZED uruchomieniem wątku ---
       if not self.full_pdf_path:
           messagebox.showerror("Błąd", "Najpierw wybierz plik PDF!")
           return


       # Sugerowana nazwa pliku
       default_name = os.path.basename(self.full_pdf_path).replace(".pdf", "-Posortowane.pdf")


       # Otwieramy okno "Zapisz jako"
       save_path = filedialog.asksaveasfilename(
           defaultextension=".pdf",
           filetypes=[("PDF Files", "*.pdf")],
           initialfile=default_name,
           title="Wybierz miejsce zapisu i nazwę pliku"
       )


       # Jeśli użytkownik kliknie "Anuluj", przerywamy
       if not save_path:
           return


       # Zapisujemy wybraną ścieżkę do zmiennej klasy
       self.full_save_path = save_path


       # Dopiero teraz uruchamiamy wątek przetwarzania
       threading.Thread(target=self.process_documents).start()


   def process_documents(self):
       pdf_file = self.full_pdf_path
       csv_file = self.full_csv_path


       # --- ZMIANA: Pobieramy ścieżkę zapisu ustaloną w start_thread ---
       output_path = self.full_save_path


       pattern = self.regex_pattern.get()
       placeholder = self.replace_placeholder.get()
       pages_count = self.pages_per_doc.get()


       if not pdf_file or not os.path.exists(pdf_file):
           messagebox.showerror("Błąd", "Wybierz poprawny plik PDF.")
           return


       try:
           self.btn_start.config(state='disabled')
           self.status_label.config(text="Analizowanie PDF...")


           doc = fitz.open(pdf_file)
           total_pages = len(doc)


           documents_metadata = []
           total_docs = (total_pages + pages_count - 1) // pages_count
           self.progress["maximum"] = total_docs
           self.progress["value"] = 0


           # --- ETAP 1: SKANOWANIE DOKUMENTÓW ---
           for i in range(total_docs):
               start_page = i * pages_count
               end_page = min(start_page + pages_count, total_pages)


               # Pobranie tekstu z pierwszej strony dokumentu
               page_text = doc[start_page].get_text()


               # Zmienne do sortowania
               raw_match = ""
               sort_value = 0  # Domyślnie 0


               if pattern:
                   match = re.search(pattern, page_text)
                   if match:
                       raw_match = match.group(0)


                       # Wyciąganie liczby
                       numbers = re.findall(r'\d+', raw_match)
                       if numbers:
                           sort_value = int(numbers[0])
                       else:
                           sort_value = -1
                   else:
                       sort_value = -1  # Brak dopasowania regexa


               documents_metadata.append({
                   'index': i,
                   'start_page': start_page,
                   'end_page': end_page,
                   'sort_value': sort_value,
                   'debug_text': raw_match
               })
               self.progress["value"] += 0.5


           # --- ETAP 2: SORTOWANIE LICZBOWE ---
           reverse_sort = (self.sort_order.get() == "descending")
           documents_metadata.sort(key=lambda x: x['sort_value'], reverse=reverse_sort)


           # --- ETAP 3: WCZYTANIE CSV ---
           replacements = []
           if csv_file and os.path.exists(csv_file):
               try:
                   try:
                       with open(csv_file, newline='', encoding='utf-8') as f:
                           reader = csv.reader(f)
                           replacements = [row[0] for row in reader if row]
                   except UnicodeDecodeError:
                       with open(csv_file, newline='', encoding='cp1250') as f:
                           reader = csv.reader(f)
                           replacements = [row[0] for row in reader if row]
               except Exception as e:
                   print(f"Błąd CSV: {e}")


           # --- ETAP 4: TWORZENIE PDF ---
           self.status_label.config(text="Generowanie posortowanego pliku...")
           output_doc = fitz.open()


           for idx, meta in enumerate(documents_metadata):
               output_doc.insert_pdf(doc, from_page=meta['start_page'], to_page=meta['end_page'] - 1)


               if placeholder and idx < len(replacements):
                   new_value = str(replacements[idx])


                   current_len = len(output_doc)
                   doc_len = meta['end_page'] - meta['start_page']
                   start_new_idx = current_len - doc_len


                   for p_num in range(start_new_idx, current_len):
                       page = output_doc[p_num]
                       hits = page.search_for(placeholder)
                       for rect in hits:
                           page.add_redact_annot(rect, fill=(1, 1, 1))
                           page.apply_redactions()
                           insert_point = fitz.Point(rect.x0, rect.y1 - 2)
                           page.insert_text(insert_point, new_value, fontsize=10, fontname="Helvetica-Bold",
                                            color=(0, 0, 0))


               self.progress["value"] = (total_docs / 2) + ((idx + 1) / total_docs * (total_docs / 2))


           # --- Zapis do pliku wybranego w oknie dialogowym ---
           output_doc.save(output_path)
           output_doc.close()
           doc.close()


           self.status_label.config(text="Zakończono sukcesem!")
           messagebox.showinfo("Gotowe", f"Plik zapisany jako:\n{os.path.basename(output_path)}")


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