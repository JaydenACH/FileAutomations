import os
import tkinter as tk
import tkinter.messagebox as msgbox
from tkinter.filedialog import askopenfiles, asksaveasfile
import customtkinter as ctk
from PyPDF2 import PdfFileMerger, PdfFileReader

"""
This application provide a graphical user interface to merge pdfs using the libary PyPDF2. 
There are 2 methods for the user to choose from, depending on the circumstances they face.

Method 1 is having a folder containing all pdfs to be merged. Paste the link into application and provide a
new file name to merge.

Method 2 is to browser for various pdfs across different directory and prompt a save-as window to select
save file location and new file name.

Room for improvement will be sorting pages and delete pdfs which were wrongly loaded in method 2. Also another 
function to be added in this application can be spliting 1 pdfs into multiple page.

Welcome all contributors for my simple application.
"""

# initialize the tkinter window
app = ctk.CTk()
app.geometry("550x750")
app.resizable(False, False)
app.title("PDF Merger")

# gather pdf files that user provide
list_of_pdfs = []

# merge with method 1
def mergepdf():
    try:
        merged_object = PdfFileMerger()
        filename = new_pdf.get() # get new file name from user
        pdfpath = pdf_path.get() # get folder path which consists of all pdfs to be merged
        os.chdir(pdfpath)
        for item in os.listdir(pdfpath):
            if item.endswith('.pdf') or item.endswith('.PDF'):
                list_of_pdfs.append(item)
        for pdf in list_of_pdfs:
            merged_object.append(PdfFileReader(pdf, 'rb'))
        merged_object.write(f"{filename}.pdf") # save the merged pdfs into 1 file under user provided name
    except FileNotFoundError:
        msgbox.showerror("Warning", "No path given!")

# merge with method 2
def load_pdfs():
    global list_of_pdfs, pdfs

    # let user browse for various files across their drive directory
    path_pdf = askopenfiles(parent=app, title="Browse For PDFs..",
                            mode='r', filetypes=[("PDF", "*.pdf")])
    for item in path_pdf:
        list_of_pdfs.append(item.name)

    # display the files that been inputted into the application for merging
    for i in path_pdf:
        cur = pdfs.get()
        nex = cur + '\n' + i.name
        pdfs.set(nex)

# merge with method 2, taking files gathered from load_pdfs() function to be merged
def saveas():
    global list_of_pdfs
    combinedpdf = PdfFileMerger()
    for pdf in list_of_pdfs:
        combinedpdf.append(PdfFileReader(pdf, 'rb'))
    # prompt user to select save file location and file name
    savepath = asksaveasfile(mode='w', defaultextension='.pdf', filetypes=[("PDF", "*.pdf")])
    combinedpdf.write(savepath.name)

# exit function to bind with Esc button on keyboard
def quit_app():
    app.destroy()

# GUI Layout of the application
cright = u"\u00A9" "Created by Jayden Ang"
pdfs = tk.StringVar()

lbl_method1 = ctk.CTkLabel(master=app, text="Method 1 :", text_font=('Arial Bold',))
lbl_method2 = ctk.CTkLabel(master=app, text="Method 2 :", text_font=('Arial Bold',))
lbl_pdf_path = ctk.CTkLabel(master=app, text="Put your folder link here : ", anchor='w')
lbl_new_pdf = ctk.CTkLabel(master=app, text="New PDF Name : ", anchor='w')
lbl_copyright = ctk.CTkLabel(master=app, text=cright, anchor='w')
lbl_lblcap = ctk.CTkLabel(master=app, text="PDFs loaded are as below:")
lbl_listofpdf = ctk.CTkLabel(master=app, textvariable=pdfs, bg_color='#003366',
                             width=450, height=250, justify='left', wrap=450)

divider1 = ctk.CTkCanvas(master=app, width=500, height=0.1)
divider2 = ctk.CTkCanvas(master=app, width=500, height=0.1)

pdf_path = ctk.CTkEntry(master=app, width=300)
new_pdf = ctk.CTkEntry(master=app, width=300)

export_button = ctk.CTkButton(master=app, text="Export", command=mergepdf)
load_button = ctk.CTkButton(master=app, text="Load", command=load_pdfs)
save_button = ctk.CTkButton(master=app, text="Save As", command=saveas)

lbl_copyright.grid(row=0, column=0, columnspan=2, padx=5, pady=10)
divider1.grid(row=1, column=0, columnspan=2, pady=5, sticky='nsew')
lbl_method1.grid(row=2, column=0, pady=10, sticky='w')
lbl_pdf_path.grid(row=3, column=0, padx=25, pady=10, sticky='w')
pdf_path.grid(row=3, column=1, padx=25, pady=10)
lbl_new_pdf.grid(row=4, column=0, padx=25, pady=10, sticky='w')
new_pdf.grid(row=4, column=1, padx=25, pady=10)
export_button.grid(row=5, column=0, columnspan=2, padx=25, pady=10, sticky='ew')
divider2.grid(row=6, column=0, columnspan=2, pady=30, sticky='nsew')
lbl_method2.grid(row=7, column=0, pady=10, sticky='w')
lbl_lblcap.grid(row=8, column=0, padx=20, pady=10, sticky='w')
lbl_listofpdf.grid(row=9, column=0, columnspan=2, padx=25, pady=10)
load_button.grid(row=10, column=0, padx=25, pady=10)
save_button.grid(row=10, column=1, padx=25, pady=10, sticky='e')

app.bind('<Escape>', lambda e: quit_app())

app.mainloop()
