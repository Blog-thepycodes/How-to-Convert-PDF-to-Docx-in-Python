import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from pdf2docx import Converter
import os
from docx import Document
import PyPDF2
 
 
def create_widgets(root, input_file, output_file, status_label):
   # Button to select the input PDF file
   input_button = tk.Button(root, text="Select PDF File", command=lambda: load_input_file(input_file, status_label))
   input_button.pack(pady=20)
 
 
   # Button to select the output DOCX file location
   output_button = tk.Button(root, text="Select Output DOCX File", command=lambda: select_output_file(output_file, status_label))
   output_button.pack(pady=20)
 
 
   # Button to start the conversion process
   convert_button = tk.Button(root, text="Convert", command=lambda: convert_pdf2docx(input_file, output_file, status_label))
   convert_button.pack(pady=20)
 
 
   # Label to display the status of the conversion
   status_label.pack(pady=10)
 
 
def load_input_file(input_file, status_label):
   input_file['path'] = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
   if input_file['path']:
       status_label.config(text=f"Selected PDF: {input_file['path']}")
 
 
def select_output_file(output_file, status_label):
   output_file['path'] = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("DOCX files", "*.docx")])
   if output_file['path']:
       status_label.config(text=f"Output will be saved as: {output_file['path']}")
 
 
def convert_pdf2docx(input_file, output_file, status_label):
   if input_file['path'] and output_file['path']:
       try:
           with open(input_file['path'], 'rb') as pdf_file:
               pdf_reader = PyPDF2.PdfReader(pdf_file)
               if pdf_reader.is_encrypted:
                   messagebox.showinfo("Info", "The selected PDF file is encrypted. Please enter the password.")
                   password = simpledialog.askstring("Password", "Enter Password:", show='*')
                   if password:
                       pdf_reader.decrypt_password(password)
                   else:
                       raise Exception("Password not provided.")
 
 
           cv = Converter(input_file['path'])
           cv.convert(output_file['path'], start=0, end=None)
           cv.close()
 
 
           if os.path.exists(output_file['path']) and os.path.getsize(output_file['path']) > 0:
               status_label.config(text="Conversion Completed Successfully")
               messagebox.showinfo("Success", "PDF successfully converted to DOCX!")
               check_docx_file(output_file['path'])
           else:
               raise Exception("The file appears to be empty or missing.")
       except Exception as e:
           messagebox.showerror("Error", f"Failed to convert PDF: {e}")
           status_label.config(text="Conversion Failed")
 
 
def check_docx_file(path):
   try:
       doc = Document(path)
       messagebox.showinfo("File Check", f"Successfully opened the DOCX file. It contains {len(doc.paragraphs)} paragraphs.")
   except Exception as e:
       messagebox.showerror("File Check Error", f"Failed to open the DOCX file: {e}")
 
 
def main():
   root = tk.Tk()
   root.title("PDF to DOCX Converter - The Pycodes")
   root.geometry("400x200")
   input_file = {'path': None}
   output_file = {'path': None}
   status_label = tk.Label(root, text="", fg="green")
   create_widgets(root, input_file, output_file, status_label)
   root.mainloop()
 
 
if __name__ == "__main__":
   main()
