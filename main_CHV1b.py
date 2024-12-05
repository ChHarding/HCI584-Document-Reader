import tkinter as tk
from tkinter import PhotoImage
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from collections import Counter
import pandas as pd
import os
import re
from docx import Document  # For reading .docx files
import PyPDF2  # For reading .pdf files

class DocumentReader(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Document Reader")
        self.geometry("500x700")  # Set a larger window size for better spacing
        self.config(bg="#0d2d2f")  # Light background

        # Font and color settings
        self.font_style = ("Helvetica", 12)
        self.font_color_white = "#FFFFFF"  # White text for labels
        self.font_color_black = "#000000"  # Black text for input and output areas
        self.bg_color = "#0d2d2f"  # Light background color
        self.entry_width = 50  # Set a consistent width for entries
        
        # Create GUI components
        self.create_widgets()

    def create_widgets(self):
        # Load the image and store it as an attribute to avoid garbage collection
        self.image1 = PhotoImage(file="ReadFileButton.png")
        self.image2 = PhotoImage(file="CountButton.png")
        self.image3 = PhotoImage(file="DownButton.png")

        # GUI "Read File" button
        self.open_button = tk.Button(self, image=self.image1, command=self.open_file)
        self.open_button.grid(row=0, column=0, pady=10, padx=20, sticky="w")

        # GUI text input field (Black text)
        self.text_area = ScrolledText(self, wrap='word', height=8, width=50, font=self.font_style, bg="#E0F2F1", fg=self.font_color_black)  # Black text
        self.text_area.grid(row=1, column=0, pady=10, padx=20)

        # GUI "Include Terms" label (White text)
        self.include_terms_label = tk.Label(self, text="Enter Terms to Include in Your Search (comma separated):", font=self.font_style, bg=self.bg_color, fg=self.font_color_white)
        self.include_terms_label.grid(row=2, column=0, pady=5, padx=20, sticky="w")
        self.include_terms_entry = tk.Entry(self, width=self.entry_width, font=self.font_style, fg=self.font_color_black)
        self.include_terms_entry.grid(row=3, column=0, pady=5, padx=20, sticky="w")

        # GUI "Exclude Terms" label (White text)
        self.exclude_terms_label = tk.Label(self, text="Enter Terms to Exclude from Your Search (comma separated):", font=self.font_style, bg=self.bg_color, fg=self.font_color_white)
        self.exclude_terms_label.grid(row=4, column=0, pady=5, padx=20, sticky="w")
        self.exclude_terms_entry = tk.Entry(self, width=self.entry_width, font=self.font_style, fg=self.font_color_black)
        self.exclude_terms_entry.grid(row=5, column=0, pady=5, padx=20, sticky="w")

        # GUI "Count Words" button
        self.count_button = tk.Button(self, image=self.image2, command=self.count_words)
        self.count_button.grid(row=6, column=0, pady=10, padx=20, sticky="w")

        # GUI "Export to Excel" button
        self.export_button = tk.Button(self, image=self.image3, command=self.export_to_excel)
        self.export_button.grid(row=7, column=0, pady=10, padx=20, sticky="w")

        # GUI text output field (Black text)
        self.text_area_out = ScrolledText(self, wrap='word', height=10, width=50, font=self.font_style, bg="#E0F2F1", fg=self.font_color_black)  # Black text
        self.text_area_out.grid(row=8, column=0, pady=10, padx=20)

        # Store terms to include/exclude
        self.word_count = None
        self.include_terms = set()
        self.exclude_terms = set()

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Text Files", "*.txt"),
            ("Word Documents", "*.docx"),
            ("PDF Files", "*.pdf"),
            ("All Files", "*.*")
        ])

        if not file_path:
            return

        text = ""
        try:
            if file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as file:
                    text = file.read()
            elif file_path.endswith('.docx'):
                doc = Document(file_path)
                for paragraph in doc.paragraphs:
                    text += paragraph.text + "\n"
            elif file_path.endswith('.pdf'):
                with open(file_path, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    for page in reader.pages:
                        text += page.extract_text() + "\n"
            else:
                messagebox.showerror("Error", "Unsupported file type.")
                return

            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")

    def count_words(self):
        content = self.text_area.get("1.0", tk.END).strip().lower()

        # Remove symbols, punctuation, and numbers
        content = re.sub(r'[^a-z\s]', '', content)

        include_terms_input = self.include_terms_entry.get().strip().lower()
        self.include_terms = set(include_terms_input.split(',')) if include_terms_input else set()

        exclude_terms_input = self.exclude_terms_entry.get().strip().lower()
        self.exclude_terms = set(exclude_terms_input.split(',')) if exclude_terms_input else set()

        words = content.split()

        if not self.include_terms:
            filtered_words = [word for word in words if word not in self.exclude_terms]
        else:
            filtered_words = [
                word for word in words if word in self.include_terms and word not in self.exclude_terms
            ]

        self.word_count = Counter(filtered_words)
        self.display_word_count()

    def display_word_count(self):
        if self.word_count:
            self.text_area_out.delete(1.0, tk.END)
            self.text_area_out.insert(tk.END, "Word Counts (Filtered by Include and Exclude Terms):\n\n")

            sorted_word_count = sorted(self.word_count.items())
            for word, count in sorted_word_count:
                self.text_area_out.insert(tk.END, f"{word}: {count}\n")

    def export_to_excel(self):
        if not self.word_count:
            messagebox.showwarning("Warning", "Count words before exporting.")
            return

        word_list = list(self.word_count.items())
        df = pd.DataFrame(word_list, columns=["Word", "Count"])

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Exported to {os.path.basename(file_path)}")


if __name__ == "__main__":
    app = DocumentReader()
    app.mainloop()
