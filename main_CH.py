import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from collections import Counter
import pandas as pd
import os

class DocumentReader(tk.Tk):  # CH
    def __init__(self): # CH
        super().__init__()  #CH

        self.title("Document Reader")
        
        # GUI "Read File" button
        self.open_button = tk.Button(self, text="Read File", command=self.open_file)
        self.open_button.pack(pady=5)
        
        # GUI text input field
        self.text_area = tk.Text(self, wrap='word', height=20, width=50)
        self.text_area.pack(pady=10)
        
        # GUI "Count Words" button
        self.count_button = tk.Button(self, text="Count Words", command=self.count_words)
        self.count_button.pack(pady=5)

        # GUI "Export to Excel" button
        self.export_button = tk.Button(self, text="Export to Excel", command=self.export_to_excel)
        self.export_button.pack(pady=5)

        self.word_count = None

        # Set GUI window size
        self.geometry("500x700") 

    def open_file(self):
        file_path = filedialog.askopenfilename()
        with open(file_path, 'r') as file:
            text = file.read()
            self.text_area.insert(tk.END, text)
            #count_words(text)    

    def count_words(self):
        content = self.text_area.get("1.0", tk.END).strip()
        words = content.split()
        self.word_count = Counter(words)

        self.display_word_count()

    def display_word_count(self):
        if self.word_count:
            display_window = tk.Toplevel(self)
            display_window.title("Word Counts")
            for word, count in self.word_count.items():
                label = tk.Label(display_window, text=f"{word}: {count}")
                label.pack()

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