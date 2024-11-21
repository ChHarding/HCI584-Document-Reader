import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from collections import Counter
import pandas as pd
import os
import re
from docx import Document  # For reading .docx files
import PyPDF2            # For reading .pdf files


class DocumentReader(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Document Reader")
        self.geometry("500x700")  # Set a larger window size for better spacing
        self.config(bg="#B2EBF2")  # Light background

        # Font and color settings
        self.font_style = ("Helvetica", 12)
        self.button_font = ("Helvetica", 10, "bold")
        self.bg_color = "#E0F2F1"  # Light background color
        self.button_color = "#004D40"  # Dark Green button color
        self.button_fg = "#E0F2F1" # Aqua button font
        self.entry_width = 50  # Set a consistent width for entries
        
        # Create GUI components
        self.create_widgets()

    def create_widgets(self):
        # GUI "Read File" button
        self.open_button = tk.Button(self, text="Read File", command=self.open_file,
                                     font=self.button_font, bg=self.button_color, fg=self.button_fg,
                                     relief="flat", width=20, height=2)
        self.open_button.grid(row=0, column=0, pady=10, padx=20, sticky="ew")

        # GUI text input field
        self.text_area = ScrolledText(self, wrap='word', height=8, width=50, font=self.font_style, bg="#E0F2F1", fg="#000000")
        self.text_area.grid(row=1, column=0, pady=10, padx=20)

        # GUI "Include Terms" label and input field
        self.include_terms_label = tk.Label(self, text="Enter Terms to Include in Your Search (comma separated):", font=self.font_style, bg=self.bg_color)
        self.include_terms_label.grid(row=2, column=0, pady=5, padx=20, sticky="w")
        self.include_terms_entry = tk.Entry(self, width=self.entry_width, font=self.font_style)
        self.include_terms_entry.grid(row=3, column=0, pady=5, padx=20, sticky="w")

        # GUI "Exclude Terms" label and input field
        self.exclude_terms_label = tk.Label(self, text="Enter Terms to Exclude from Your Search (comma separated):", font=self.font_style, bg=self.bg_color)
        self.exclude_terms_label.grid(row=4, column=0, pady=5, padx=20, sticky="w")
        self.exclude_terms_entry = tk.Entry(self, width=self.entry_width, font=self.font_style)
        self.exclude_terms_entry.grid(row=5, column=0, pady=5, padx=20, sticky="w")

        # GUI "Count Words" button
        self.count_button = tk.Button(self, text="Count Words", command=self.count_words,
                                      font=self.button_font, bg=self.button_color, fg=self.button_fg,
                                      relief="flat", width=20, height=2)
        self.count_button.grid(row=6, column=0, pady=10, padx=20, sticky="ew")

        # GUI "Export to Excel" button
        self.export_button = tk.Button(self, text="Export to Excel", command=self.export_to_excel,
                                       font=self.button_font, bg=self.button_color, fg=self.button_fg,
                                       relief="flat", width=20, height=2)
        self.export_button.grid(row=7, column=0, pady=10, padx=20, sticky="ew")

        # GUI text output field
        self.text_area_out = ScrolledText(self, wrap='word', height=10, width=50, font=self.font_style, bg="#E0F2F1", fg="#000000")
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
            
            self.text_area.delete(1.0, tk.END)  # Clear the text area
            self.text_area.insert(tk.END, text)  # Insert new text
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")    

    def count_words(self):
        content = self.text_area.get("1.0", tk.END).strip().lower()  # Convert to lowercase

        # Remove symbols, punctuation, and numbers
        content = re.sub(r'[^a-z\s]', '', content)

        # Get the user input from the Include Terms field and split into a set of terms
        include_terms_input = self.include_terms_entry.get().strip().lower()
        self.include_terms = set(include_terms_input.split(',')) if include_terms_input else set()

        # Get the user input from the Exclude Terms field and split into a set of terms
        exclude_terms_input = self.exclude_terms_entry.get().strip().lower()
        self.exclude_terms = set(exclude_terms_input.split(',')) if exclude_terms_input else set()

        words = content.split()

        # If no include terms are specified, include all words
        if not self.include_terms:
            filtered_words = [word for word in words if word not in self.exclude_terms]  # Exclude only specified terms
        else:
            # Filter out terms based on Include and Exclude lists
            filtered_words = [
                word for word in words if word in self.include_terms and word not in self.exclude_terms
            ]

        self.word_count = Counter(filtered_words)

        self.display_word_count()

    def display_word_count(self):
        if self.word_count:
            self.text_area_out.delete(1.0, tk.END)  # Clear current text in the output area
            self.text_area_out.insert(tk.END, "Word Counts (Filtered by Include and Exclude Terms):\n\n")  # Add header

            # Sort word counts alphabetically
            sorted_word_count = sorted(self.word_count.items())
            for word, count in sorted_word_count:
                self.text_area_out.insert(tk.END, f"{word}: {count}\n") 

    def export_to_excel(self):
        if not self.word_count:
            messagebox.showwarning("Warning", "Count words before exporting.")
            return

        word_list = list(self.word_count.items())
        df = pd.DataFrame(word_list, columns=["Word", "Count"])

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Exported to {os.path.basename(file_path)}")

if __name__ == "__main__":
    app = DocumentReader()
    app.mainloop()