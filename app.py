import customtkinter as ctk
from tkinter import filedialog, Listbox, END
from pypdf import PdfReader, PdfWriter
import os
import fitz  # PyMuPDF
from PIL import Image, ImageTk

class PDFEditor(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Editor")
        self.geometry("1000x700") # Increased size for previews

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

        # --- Main Frame ---
        main_frame = ctk.CTkFrame(self.split_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Controls Frame ---
        controls_frame = ctk.CTkFrame(main_frame)
        controls_frame.pack(side="left", fill="y", padx=10, pady=10)

        # --- Preview Frame ---
        preview_frame = ctk.CTkFrame(main_frame)
        preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.split_preview_label = ctk.CTkLabel(preview_frame, text="Anteprima PDF")
        self.split_preview_label.pack(expand=True)

        # --- Widgets ---
        self.select_split_button = ctk.CTkButton(controls_frame, text="Seleziona PDF", command=self.select_split_file)
        self.select_split_button.pack(pady=10, padx=10, fill="x")

        self.split_file_label = ctk.CTkLabel(controls_frame, text="Nessun file selezionato")
        self.split_file_label.pack(pady=5, padx=10)

        self.select_dir_button = ctk.CTkButton(controls_frame, text="Seleziona Cartella di Destinazione", command=self.select_output_dir)
        self.select_dir_button.pack(pady=10, padx=10, fill="x")

        self.split_dir_label = ctk.CTkLabel(controls_frame, text="Nessuna cartella selezionata")
        self.split_dir_label.pack(pady=5, padx=10)

        self.ranges_label = ctk.CTkLabel(controls_frame, text="Intervalli di pagine (es. 1-3, 5, 8-10):")
        self.ranges_label.pack(pady=5, padx=10)

        self.ranges_entry = ctk.CTkEntry(controls_frame, width=250)
        self.ranges_entry.pack(pady=5, padx=10)
        self.ranges_entry.bind("<KeyRelease>", self.update_split_preview)

        self.split_button = ctk.CTkButton(controls_frame, text="Dividi e Salva", command=self.split_pdf)
        self.split_button.pack(pady=20, padx=10, fill="x")

        self.split_status_label = ctk.CTkLabel(controls_frame, text="")
        self.split_status_label.pack(pady=5, padx=10)

    def setup_merge_tab(self):
        self.merge_file_paths = []
        self.merge_output_dir = ""

        # --- Main Frame ---
        main_frame = ctk.CTkFrame(self.merge_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Controls Frame ---
        controls_frame = ctk.CTkFrame(main_frame)
        controls_frame.pack(side="left", fill="y", padx=10, pady=10)

        # --- Preview Frame ---
        preview_frame = ctk.CTkFrame(main_frame)
        preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.merge_preview_label = ctk.CTkLabel(preview_frame, text="Anteprima PDF")
        self.merge_preview_label.pack(expand=True)

        # --- Widgets ---
        self.select_merge_button = ctk.CTkButton(controls_frame, text="Seleziona PDF da Unire", command=self.select_merge_files)
        self.select_merge_button.pack(pady=10, padx=10, fill="x")

        self.merge_listbox = Listbox(controls_frame, selectmode="extended", bg="#2B2B2B", fg="white", borderwidth=0, highlightthickness=0, height=10)
        self.merge_listbox.pack(pady=10, padx=10, fill="x")
        self.merge_listbox.bind("<<ListboxSelect>>", self.update_merge_preview)

        self.reorder_frame = ctk.CTkFrame(controls_frame)
        self.reorder_frame.pack(pady=5)

        self.up_button = ctk.CTkButton(self.reorder_frame, text="Su", command=self.move_up)
        self.up_button.pack(side="left", padx=5)

        self.down_button = ctk.CTkButton(self.reorder_frame, text="Giù", command=self.move_down)
        self.down_button.pack(side="left", padx=5)

        self.select_merge_dir_button = ctk.CTkButton(controls_frame, text="Seleziona Cartella di Destinazione", command=self.select_merge_output_dir)
        self.select_merge_dir_button.pack(pady=10, padx=10, fill="x")

        self.merge_dir_label = ctk.CTkLabel(controls_frame, text="Nessuna cartella selezionata")
        self.merge_dir_label.pack(pady=5, padx=10)

        self.output_filename_label = ctk.CTkLabel(controls_frame, text="Nome file unito:")
        self.output_filename_label.pack(pady=5, padx=10)

        self.output_filename_entry = ctk.CTkEntry(controls_frame, placeholder_text="merged.pdf")
        self.output_filename_entry.pack(pady=5, padx=10, fill="x")

        self.merge_button = ctk.CTkButton(controls_frame, text="Unisci e Salva", command=self.merge_pdfs)
        self.merge_button.pack(pady=20, padx=10, fill="x")

        self.merge_status_label = ctk.CTkLabel(controls_frame, text="")
        self.merge_status_label.pack(pady=5, padx=10)

    def select_split_file(self):
        self.split_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.split_file_path:
            self.split_file_label.configure(text=os.path.basename(self.split_file_path))
            self.split_status_label.configure(text="")
            self.update_split_preview()

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
            self.update_merge_preview()

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
        self.update_merge_preview()

    def move_down(self):
        selected_indices = self.merge_listbox.curselection()
        for i in reversed(selected_indices):
            if i < len(self.merge_file_paths) - 1:
                self.merge_file_paths[i], self.merge_file_paths[i+1] = self.merge_file_paths[i+1], self.merge_file_paths[i]
        self.update_merge_listbox()
        self.update_merge_preview()

    def select_merge_output_dir(self):
        self.merge_output_dir = filedialog.askdirectory()
        if self.merge_output_dir:
            self.merge_dir_label.configure(text=self.merge_output_dir)
            self.merge_status_label.configure(text="")

    def merge_pdfs(self):
        if len(self.merge_file_paths) < 2:
            self.merge_status_label.configure(text="Errore: Seleziona almeno due file.", text_color="red")
            return

        if not self.merge_output_dir:
            self.merge_status_label.configure(text="Errore: Seleziona una cartella di destinazione.", text_color="red")
            return

        output_filename = self.output_filename_entry.get()
        if not output_filename:
            output_filename = "merged.pdf"
        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"

        save_path = os.path.join(self.merge_output_dir, output_filename)

        try:
            writer = PdfWriter()
            for pdf_file in self.merge_file_paths:
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    writer.add_page(page)

            with open(save_path, "wb") as f:
                writer.write(f)

            self.merge_status_label.configure(text=f"Unione completata!\nSalvataggio in {output_filename}", text_color="green")
            self.merge_file_paths = []
            self.update_merge_listbox()
            self.merge_preview_label.configure(image=None, text="Anteprima PDF")
            self.merge_preview_label.image = None


        except Exception as e:
            self.merge_status_label.configure(text=f"Errore: {e}", text_color="red")

    def update_split_preview(self, *args):
        if not self.split_file_path:
            return

        page_num = 1
        try:
            ranges_str = self.ranges_entry.get()
            if ranges_str:
                first_part = ranges_str.split(',')[0].strip()
                if '-' in first_part:
                    page_num = int(first_part.split('-')[0])
                else:
                    page_num = int(first_part)
        except (ValueError, IndexError):
            page_num = 1 # Default to first page on invalid input

        self.show_pdf_preview(self.split_file_path, page_num, self.split_preview_label)

    def update_merge_preview(self, event=None):
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices:
            # You might want to clear the preview or show a default message
            self.merge_preview_label.configure(image=None, text="Anteprima PDF")
            self.merge_preview_label.image = None
            return

        selected_index = selected_indices[0]
        if 0 <= selected_index < len(self.merge_file_paths):
            filepath = self.merge_file_paths[selected_index]
            self.show_pdf_preview(filepath, 1, self.merge_preview_label)

    def show_pdf_preview(self, filepath, page_num, preview_label):
        try:
            doc = fitz.open(filepath)
            page_count = doc.page_count

            if not (1 <= page_num <= page_count):
                preview_label.configure(text=f"Pagina {page_num} non valida.", image=None)
                preview_label.image = None
                return

            page = doc.load_page(page_num - 1)  # 0-indexed

            # Render page to a pixmap
            pix = page.get_pixmap()

            # Convert to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Resize if necessary to fit the preview panel
            max_size = (400, 550)
            img.thumbnail(max_size, Image.Resampling.LANCZOS)

            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

            preview_label.configure(image=ctk_img, text="")
            preview_label.image = ctk_img # Keep a reference

        except Exception as e:
            preview_label.configure(text=f"Errore anteprima:\n{e}", image=None)
            preview_label.image = None


if __name__ == "__main__":
    app = PDFEditor()
    app.mainloop()