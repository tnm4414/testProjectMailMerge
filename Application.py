import tkinter as tk
import os
from tkinter import filedialog, messagebox
import main
import pandas as pd

# Suppress deprecation warnings
os.environ['TK_SILENCE_DEPRECATION'] = '1'


def select_and_process_file():
    template_path = 'Header_Template.csv'  # Static file path for the template
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        try:
            os.replace(file_path, template_path)
            main.process_csv_and_generate_doc(template_path, 'Output.docx')
            messagebox.showinfo("Success", "Word document created successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


def main():
    root = tk.Tk()
    root.title("CSV to Word Mail Merge")

    btn_upload = tk.Button(root, text="Upload and Process CSV", command=select_and_process_file)
    btn_upload.pack(pady=20)

    root.mainloop()


if __name__ == "__main__":
    main()
