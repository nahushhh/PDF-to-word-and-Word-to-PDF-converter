import tkinter as tk
import os
from docx2pdf import convert
from pdf2docx import parse,Converter
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfile, askdirectory
from tkinter.messagebox import showinfo

win = tk.Tk()
win.title("Word to PDF and PDF to Word Converter")


def doc_to_pdf_openfile():
    file = askopenfile(filetypes=[("Word Files", "*.docx")])
    convert(file.name)
    showinfo("Done", "File Sucessfully Converted")

def doc_to_pdf_openfolder():
    folder = askdirectory()
    print(folder)
    convert(folder)
    showinfo("Done", "Folder Files Sucessfully Converted")

def pdf_to_doc_openfile():
    file = askopenfile(filetypes=[("PDF", "*.pdf")])
    parse(file.name)
    showinfo("Done", "File Sucessfully Converted")

def pdf_to_doc_openfolder():
    folder = askdirectory()
    path = os.listdir(folder)
    for i in path:
        if i[-4:]=='.pdf':
            length = len(i)
            parse(f"D:/temp/test/pdf/{i[0:length-4]}.pdf",f"D:/temp/test/pdf/{i[0:length-4]}.docx",start=0,end=None)

    showinfo("Done", "Folder Files Sucessfully Converted")


"""Word to PDF"""
title_label_1 = tk.Label(win, text="Word to PDF Converter ")
title_label_1.grid(row=0, column=0, padx=5, pady=5)

label = tk.Label(win, text="Choose File/Folder : ")
label.grid(row=1, column=0, padx=5, pady=5)

button1 = ttk.Button(win, text="Select File", width=30, command=doc_to_pdf_openfile)
button1.grid(row=2, column=0, padx=5, pady=5)

button2 = ttk.Button(win, text="Select Folder", width=30, command=doc_to_pdf_openfolder)
button2.grid(row=3, column=0, padx=5, pady=5)


"""PDF to Word"""
title_label_2 = tk.Label(win, text="PDF to Word Converter ")
title_label_2.grid(row=5, column=0, padx=5, pady=5)

label1 = tk.Label(win, text="Choose File/Folder : ")
label1.grid(row=6, column=0, padx=5, pady=5)

button3 = ttk.Button(win, text="Select File", width=30, command=pdf_to_doc_openfile)
button3.grid(row=7, column=0, padx=5, pady=5)

button4 = ttk.Button(win, text="Select Folder", width=30, command=pdf_to_doc_openfolder)
button4.grid(row=8, column=0, padx=5, pady=5)

helpText = """
    How to Use -
    Use the application to either convert single file or a folder containing multiple doc files!
    Simply select your file/folder and the app will take care of the rest!
"""

helpLabel = tk.Label(win, text=helpText)
helpLabel.grid(row=9, column=0, padx=5, pady=5)

win.mainloop()