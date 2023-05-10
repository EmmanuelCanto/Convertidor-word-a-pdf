import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from docx2pdf import convert
import os

root = tk.Tk()

def select_word_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    word_file_label.config(text=file_path)


def convert_file():
    word_file = word_file_label.cget("text")
    pdf_file = word_file[:-4] + ".pdf"

    # Convertir el archivo Word a PDF
    convert(word_file, pdf_file, keep_active=False)

    # Mostrar mensaje de éxito
    messagebox.showinfo("Éxito", f"El archivo {pdf_file} se ha creado con éxito.")

def handle_drop(event): 
    try:
        file_path = event.widget.selection_get().strip()
        file_path = event.widget.text()
        # Resto del código para convertir el archivo
    except tk.TclError:
        # Manejar el caso en que la selección primaria no existe
        pass

root.geometry("300x150")
root.title("Convertidor")

# Crear etiqueta para mostrar la ruta del archivo Word
word_file_label = tk.Label(root, text="Selecciona un archivo Word")
word_file_label.pack(pady=10)

# Crear botón para seleccionar archivo Word
select_word_file_button = tk.Button(root, text="Seleccionar archivo Word", command=select_word_file)
select_word_file_button.pack(pady=10)

# Crear botón para convertir a PDF
convert_button = tk.Button(root, text="Convertir a PDF", command=convert_file)
convert_button.pack(pady=10)

# Permitir que se pueda soltar el archivo Word sobre la ventana
root.bind("<ButtonRelease-1>", handle_drop)
root.bind("<Enter>", lambda e: e.widget.focus_set())

root.mainloop()
