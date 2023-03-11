import tkinter as tk
from tkinter import filedialog
import tkinter.font as font
import os
from pdf2docx import Converter
from termcolor import colored

root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("700x500")
root.configure(bg="#263D42")

def select_file():
    file_path = filedialog.askopenfilename(initialdir="/", title="Select a PDF file",
                                           filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
    if file_path:
        convert_file(file_path)

def convert_file(file_path):
    save_location = os.path.join(os.getcwd(), "Document.docx")
    cv = Converter(file_path)
    cv.convert(save_location, start=0, end=None)
    cv.close()
    success_text = tk.Label(frame, text=colored("PDF has been successfully converted to DOCX!", "green"),
                            font=("Courier", 16), bg="#263D42", fg="#fff")
    success_text.pack(pady=(0, 0))

def clear_frame():
    for widget in frame.winfo_children():
        widget.destroy()

def about():
    about_text = """PDF to DOCX Converter v1.0
A simple tool to convert PDF files to DOCX format.
Made by Aniket kaloo """
    about_label = tk.Label(frame, text=about_text, font=("Courier", 14), bg="#263D42", fg="#fff")
    about_label.pack(pady=(50, 0))

# Header
header_label = tk.Label(root, text="PDF to DOCX Converter", font=("Courier", 24), bg="#263D42", fg="#fff")
header_label.pack(pady=(20, 0))

# Frame
frame = tk.Frame(root, bg="#c2d6d6")
frame.place(relwidth=0.9, relheight=0.7, relx=0.05, rely=0.2)

# Buttons
button_font = font.Font(size=20)

select_button = tk.Button(frame, text="Select PDF file", bg="#ff2200", fg="#fff",
                          command=select_file, font=button_font)
select_button.pack(padx=20, pady=(50, 10), fill="x")

clear_button = tk.Button(frame, text="Clear", bg="#c2d6d6", fg="#000",
                         command=clear_frame, font=button_font)
clear_button.pack(padx=20, pady=10, fill="x")

about_button = tk.Button(frame, text="About", bg="#c2d6d6", fg="#000",
                         command=about, font=button_font)
about_button.pack(padx=20, pady=10, fill="x")

exit_button = tk.Button(root, text="Exit", bg="#ff2200", fg="#fff",
                        command=root.destroy, font=button_font)
exit_button.pack(pady=(20, 0), fill="x")

root.mainloop()
