import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def add_headers_to_excel(excel_file_path, output_file_path):
    try:
        # Leer el archivo Excel sin encabezados.
        df = pd.read_excel(excel_file_path, header=None)
        
        # Definir los nuevos encabezados en el orden especificado.
        headers = [
            "IdPrestamo",
            "RFC",
            "Nombre",
            "QnaInicial",
            "QnaFinal",
            "CapitalPrestado",
            "InteresDelPrestamo",
            "FondoDeGarantia",
            "DescuentoQnal",
            "Total",
            "Plazo",
            "Estatus",
            "ColumnaVacia",
            "TotalRecuperado",
            "CapitalRecuperado",
            "InteresRecuperado",
            "FondoRecuperado",
            "QnasPagadas",
            "QnasTranscurridas",
            "QnasAtrasadas",
            "EstatusDePagos",
            "EstatusMaestro",
            "NombrePrestador",
            "Regiones",
            "Bancos"
        ]
        
        # Asignar los nuevos encabezados al DataFrame.
        df.columns = headers
        
        # Guardar el DataFrame con los nuevos encabezados en un nuevo archivo Excel.
        df.to_excel(output_file_path, index=False)
        
        messagebox.showinfo("Éxito", f"Encabezados agregados y archivo guardado como {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error al procesar el archivo: {e}")

def select_file():
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_file_path:
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_file_path:
            add_headers_to_excel(excel_file_path, output_file_path)

# Crear la ventana principal de tkinter.
root = tk.Tk()
root.title("Agregar Encabezados a Excel")

# Crear un botón para seleccionar el archivo Excel.
select_button = tk.Button(root, text="Seleccionar archivo Excel", command=select_file)
select_button.pack(pady=20)

# Ejecutar el bucle principal de tkinter.
root.mainloop()