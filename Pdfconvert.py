import os
import subprocess
import customtkinter as ctk
from tkinter import filedialog
from docx2pdf import convert

# -------------------- App Setup --------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class WordToPDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Word to PDF Converter")
        self.geometry("550x380")
        self.resizable(False, False)

        self.selected_files = []

        self.build_ui()

    def build_ui(self):
        # Title
        self.title_label = ctk.CTkLabel(
            self,
            text="Word ➜ PDF Converter",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.title_label.pack(pady=20)

        # Select Files Button
        self.select_button = ctk.CTkButton(
            self,
            text="Select Word Files",
            width=220,
            height=40,
            command=self.select_files
        )
        self.select_button.pack(pady=10)

        # Files Display Box
        self.files_box = ctk.CTkTextbox(
            self,
            width=480,
            height=140,
            state="disabled"
        )
        self.files_box.pack(pady=10)

        # Convert Button
        self.convert_button = ctk.CTkButton(
            self,
            text="Convert to PDF",
            width=220,
            height=40,
            command=self.convert_files
        )
        self.convert_button.pack(pady=10)

        # Status Label
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=10)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Word Documents",
            filetypes=[("Word Documents", "*.docx")]
        )

        if not files:
            return

        for file in files:
            if file not in self.selected_files:
                self.selected_files.append(file)

        self.update_files_box()

    def update_files_box(self):
        self.files_box.configure(state="normal")
        self.files_box.delete("1.0", "end")

        for file in self.selected_files:
            self.files_box.insert("end", os.path.basename(file) + "\n")

        self.files_box.configure(state="disabled")

    def convert_files(self):
        if not self.selected_files:
            self.status_label.configure(
                text="❌ No files selected",
                text_color="red"
            )
            return

        try:
            self.status_label.configure(text="⏳ Converting files...")
            self.update()

            last_pdf_path = None

            for file_path in self.selected_files:
                convert(file_path)
                last_pdf_path = file_path.replace(".docx", ".pdf")

            self.status_label.configure(
                text=f"Converted {len(self.selected_files)} file(s) successfully",
                text_color="green"
            )

            # Open file location (Windows)
            if last_pdf_path and os.name == "nt":
                subprocess.run(
                    ["explorer", "/select,", os.path.normpath(last_pdf_path)]
                )

        except Exception as e:
            self.status_label.configure(
                text=f"Error:\n{e}",
                text_color="red"
            )

# -------------------- Run App --------------------
if __name__ == "__main__":
    app = WordToPDFApp()
    app.mainloop()
