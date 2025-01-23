import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter

def convert_txt_to_excel(txt_file_path, excel_file_path):
    # Crear un nuevo libro de Excel y agregar una hoja de trabajo.
    workbook = xlsxwriter.Workbook(excel_file_path)
    worksheet = workbook.add_worksheet()

    # Escribir encabezados en el archivo Excel.
    worksheet.write('G1', 'RFC')
    worksheet.write('C1', 'NOMBRE')
    worksheet.write('L1', 'IMPORTE')  

    # Leer el archivo .txt y escribir su contenido en el archivo Excel.
    with open(txt_file_path, 'r') as txt_file:
        for row_num, line in enumerate(txt_file, start=1):
            rfc = line[0:14].strip()
            nombre = line[14:50].strip()
            importe = line[77:90].strip()
            
            worksheet.write(row_num, 6, rfc)  # Columna G (índice 6)
            worksheet.write(row_num, 2, nombre)  # Columna C (índice 2)
            worksheet.write(row_num, 11, importe)  # Columna L (índice 11)

    # Cerrar el libro de Excel.
    workbook.close()

def select_file():
    txt_file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if txt_file_path:
        excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_file_path:
            convert_txt_to_excel(txt_file_path, excel_file_path)
            messagebox.showinfo("Success", f"File converted and saved as {excel_file_path}")

# Crear la ventana principal de tkinter.
root = tk.Tk()
root.title("Convertidor de TXT a Excel")

# Crear un botón para seleccionar el archivo .txt.
select_button = tk.Button(root, text="Seleccionar archivo TXT", command=select_file)
select_button.pack(pady=20)

# Ejecutar el bucle principal de tkinter.
root.mainloop()