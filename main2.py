import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from docx2pdf import convert
import os

root = tk.Tk()

def select_word_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    word_file_canvas.itemconfigure(word_file_text, text=file_path)
    word_file_canvas.itemconfigure(word_file_icon, state=tk.HIDDEN)

def convert_file():
    word_file = word_file_canvas.itemcget(word_file_text, "text")
    pdf_file = word_file[:-4] + ".pdf"

    # Convertir el archivo Word a PDF
    convert(word_file, pdf_file, keep_active=False)

    # Mostrar mensaje de éxito
    messagebox.showinfo("Éxito", f"El archivo {pdf_file} se ha creado con éxito.")
    
    # Reiniciar la zona de arrastre
    word_file_canvas.itemconfigure(word_file_text, text="Arrastra aquí un archivo Word")
    word_file_canvas.itemconfigure(word_file_icon, state=tk.NORMAL)

def handle_drop(event):
    # Obtener la ruta del archivo
    file_path = event.data["text"].strip()

    # Verificar que sea un archivo Word
    if file_path.endswith(".docx"):
        # Actualizar la zona de arrastre con la ruta del archivo
        word_file_canvas.itemconfigure(word_file_text, text=file_path)
        word_file_canvas.itemconfigure(word_file_icon, state=tk.HIDDEN)
    else:
        messagebox.showwarning("Error", "Solo se permiten archivos Word (*.docx)")

root.geometry("300x200")
root.title("Convertidor")

# Crear canvas para la zona de arrastre
word_file_canvas = tk.Canvas(root, width=200, height=100, bg="white", highlightthickness=2, highlightbackground="black")
word_file_canvas.pack(pady=10)

# Agregar texto e icono a la zona de arrastre
word_file_text = word_file_canvas.create_text(100, 50, text="Arrastra aquí un archivo Word", font=("Arial", 12), fill="black")
word_file_icon = word_file_canvas.create_image(100, 50, image=tk.PhotoImage(file="word_icon.png"), anchor=tk.CENTER)

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
