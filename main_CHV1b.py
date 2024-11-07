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
        
        # GUI "Read File" button
        self.open_button = tk.Button(self, text="Read File", command=self.open_file)
        self.open_button.pack(pady=5)
        
        # GUI text input field
        self.text_area = ScrolledText(self, wrap='word', height=10, width=50)
        self.text_area.pack(pady=5)
        
        # GUI "Include Terms" label and input field
        self.include_terms_label = tk.Label(self, text="Enter Terms to Include (comma separated):")
        self.include_terms_label.pack(pady=5)
        self.include_terms_entry = tk.Entry(self, width=50)
        self.include_terms_entry.pack(pady=5)

        # GUI "Exclude Terms" label and input field
        self.exclude_terms_label = tk.Label(self, text="Enter Terms to Exclude (comma separated):")
        self.exclude_terms_label.pack(pady=5)
        self.exclude_terms_entry = tk.Entry(self, width=50)
        self.exclude_terms_entry.pack(pady=5)

        # GUI "Count Words" button
        self.count_button = tk.Button(self, text="Count Words", command=self.count_words)
        self.count_button.pack(pady=5)

        # GUI "Export to Excel" button
        self.export_button = tk.Button(self, text="Export to Excel", command=self.export_to_excel)
        self.export_button.pack(pady=5)

        # GUI text output field
        self.text_area_out = ScrolledText(self, wrap='word', height=10, width=50)
        self.text_area_out.pack(pady=5)

        self.word_count = None
        self.include_terms = set()  # Store the terms to include
        self.exclude_terms = set()  # Store the terms to exclude

        # Set GUI window size
        self.geometry("500x600") 

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