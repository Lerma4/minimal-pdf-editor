import customtkinter as ctk
from tkinter import filedialog, Listbox, END
from pypdf import PdfReader, PdfWriter
import os

class PDFEditor(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Editor")
        self.geometry("700x500")

        ctk.set_appearance_mode("dark")

        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(padx=20, pady=20, fill="both", expand=True)

        self.split_tab = self.tab_view.add("Dividi PDF")
        self.merge_tab = self.tab_view.add("Unisci PDF")

        self.setup_split_tab()
        self.setup_merge_tab()

    def setup_split_tab(self):
        self.split_file_path = ""
        self.split_output_dir = ""

        # --- Widgets ---
        self.select_split_button = ctk.CTkButton(self.split_tab, text="Seleziona PDF", command=self.select_split_file)
        self.select_split_button.pack(pady=10)

        self.split_file_label = ctk.CTkLabel(self.split_tab, text="Nessun file selezionato")
        self.split_file_label.pack(pady=5)

        self.select_dir_button = ctk.CTkButton(self.split_tab, text="Seleziona Cartella di Destinazione", command=self.select_output_dir)
        self.select_dir_button.pack(pady=10)

        self.split_dir_label = ctk.CTkLabel(self.split_tab, text="Nessuna cartella selezionata")
        self.split_dir_label.pack(pady=5)

        self.ranges_label = ctk.CTkLabel(self.split_tab, text="Intervalli di pagine (es. 1-3, 5, 8-10):")
        self.ranges_label.pack(pady=5)

        self.ranges_entry = ctk.CTkEntry(self.split_tab, width=300)
        self.ranges_entry.pack(pady=5)

        self.split_button = ctk.CTkButton(self.split_tab, text="Dividi e Salva", command=self.split_pdf)
        self.split_button.pack(pady=20)

        self.split_status_label = ctk.CTkLabel(self.split_tab, text="")
        self.split_status_label.pack(pady=5)

    def setup_merge_tab(self):
        self.merge_file_paths = []

        # --- Widgets ---
        self.select_merge_button = ctk.CTkButton(self.merge_tab, text="Seleziona PDF da Unire", command=self.select_merge_files)
        self.select_merge_button.pack(pady=10)

        self.merge_listbox = Listbox(self.merge_tab, selectmode="extended", bg="#2B2B2B", fg="white", borderwidth=0, highlightthickness=0)
        self.merge_listbox.pack(pady=10, fill="both", expand=True)

        self.reorder_frame = ctk.CTkFrame(self.merge_tab)
        self.reorder_frame.pack(pady=5)

        self.up_button = ctk.CTkButton(self.reorder_frame, text="Su", command=self.move_up)
        self.up_button.pack(side="left", padx=5)

        self.down_button = ctk.CTkButton(self.reorder_frame, text="Giù", command=self.move_down)
        self.down_button.pack(side="left", padx=5)

        self.merge_button = ctk.CTkButton(self.merge_tab, text="Unisci e Salva", command=self.merge_pdfs)
        self.merge_button.pack(pady=20)

        self.merge_status_label = ctk.CTkLabel(self.merge_tab, text="")
        self.merge_status_label.pack(pady=5)

    def select_split_file(self):
        self.split_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.split_file_path:
            self.split_file_label.configure(text=os.path.basename(self.split_file_path))
            self.split_status_label.configure(text="")

    def split_pdf(self):
        if not self.split_file_path:
            self.split_status_label.configure(text="Errore: Seleziona un file PDF.", text_color="red")
            return

        ranges_str = self.ranges_entry.get()
        if not ranges_str:
            self.split_status_label.configure(text="Errore: Inserisci gli intervalli di pagine.", text_color="red")
            return

        if not self.split_output_dir:
            self.split_status_label.configure(text="Errore: Seleziona una cartella di destinazione.", text_color="red")
            return

        try:
            reader = PdfReader(self.split_file_path)
            total_pages = len(reader.pages)
            parsed_ranges = self.parse_ranges(ranges_str, total_pages)

            base_name = os.path.splitext(os.path.basename(self.split_file_path))[0]

            for r in parsed_ranges:
                writer = PdfWriter()
                for page_num in r:
                    writer.add_page(reader.pages[page_num - 1])

                range_str = f"p{r[0]}-{r[-1]}" if len(r) > 1 else f"p{r[0]}"
                output_filename = os.path.join(self.split_output_dir, f"{base_name}_{range_str}.pdf")
                with open(output_filename, "wb") as f:
                    writer.write(f)

            self.split_status_label.configure(text="Divisione completata con successo!", text_color="green")

        except Exception as e:
            self.split_status_label.configure(text=f"Errore: {e}", text_color="red")

    def select_output_dir(self):
        self.split_output_dir = filedialog.askdirectory()
        if self.split_output_dir:
            self.split_dir_label.configure(text=self.split_output_dir)
            self.split_status_label.configure(text="")

    def parse_ranges(self, ranges_str, total_pages):
        ranges = []
        parts = ranges_str.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                if start < 1 or end > total_pages or start > end:
                    raise ValueError("Intervallo di pagine non valido")
                ranges.append(list(range(start, end + 1)))
            else:
                page = int(part)
                if page < 1 or page > total_pages:
                    raise ValueError("Numero di pagina non valido")
                ranges.append([page])
        return ranges

    def select_merge_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if files:
            self.merge_file_paths.extend(files)
            self.update_merge_listbox()
            self.merge_status_label.configure(text="")

    def update_merge_listbox(self):
        self.merge_listbox.delete(0, END)
        for f in self.merge_file_paths:
            self.merge_listbox.insert(END, os.path.basename(f))

    def move_up(self):
        selected_indices = self.merge_listbox.curselection()
        for i in selected_indices:
            if i > 0:
                self.merge_file_paths[i], self.merge_file_paths[i-1] = self.merge_file_paths[i-1], self.merge_file_paths[i]
        self.update_merge_listbox()

    def move_down(self):
        selected_indices = self.merge_listbox.curselection()
        for i in reversed(selected_indices):
            if i < len(self.merge_file_paths) - 1:
                self.merge_file_paths[i], self.merge_file_paths[i+1] = self.merge_file_paths[i+1], self.merge_file_paths[i]
        self.update_merge_listbox()

    def merge_pdfs(self):
        if len(self.merge_file_paths) < 2:
            self.merge_status_label.configure(text="Seleziona almeno due file per l'unione.", text_color="red")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not save_path:
            return

        try:
            writer = PdfWriter()
            for pdf_file in self.merge_file_paths:
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    writer.add_page(page)

            with open(save_path, "wb") as f:
                writer.write(f)

            self.merge_status_label.configure(text="Unione completata!", text_color="green")
            self.merge_file_paths = []
            self.update_merge_listbox()

        except Exception as e:
            self.merge_status_label.configure(text=f"Errore: {e}", text_color="red")


if __name__ == "__main__":
    app = PDFEditor()
    app.mainloop()