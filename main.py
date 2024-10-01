import tkinter as tk
from tkinter import filedialog, messagebox
from collections import Counter
import pandas as pd
import os

class DocumentReader:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Reader")
        
        self.open_button = tk.Button(root, text="Read File", command=self.open_file)
        self.open_button.pack(pady=5)
        
        self.text_area = tk.Text(root, wrap='word', height=20, width=50)
        self.text_area.pack(pady=10)
        
        self.count_button = tk.Button(root, text="Count Words", command=self.count_words)
        self.count_button.pack(pady=5)

        self.export_button = tk.Button(root, text="Export to Excel", command=self.export_to_excel)
        self.export_button.pack(pady=5)

        self.word_count = None

    def open_file(self):
        file_path = filedialog.askopenfilename()
        with open(file_path, 'r') as file:
            text = file.read()
            #text_widget.insert(tk.END, text)
            #count_words(text)    

    def count_words(self):
        content = self.text_area.get("1.0", tk.END).strip()
        words = content.split()
        self.word_count = Counter(words)

        self.display_word_count()

    def display_word_count(self):
        if self.word_count:
            display_window = tk.Toplevel(self.root)
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
    root = tk.Tk()
    app = DocumentReader(root)
    root.mainloop()
