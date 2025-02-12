import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def convert_txt_to_excel(txt_file_path, excel_file_path):
    # Crear un DataFrame vacío para almacenar los datos.
    data = {
        'RFC': [],
        'IMPORTE': [],
        'RESPUESTA': []
    }

    # Leer el archivo .txt y extraer los datos, omitiendo la primera y la última línea.
    with open(txt_file_path, 'r') as txt_file:
        lines = txt_file.readlines()[1:-1]  # Omitir la primera y la última línea
        for line in lines:
            rfc = line[135:150].strip()
            importe_str = line[22:29].strip()
            importe = float(importe_str[:-2] + '.' + importe_str[-2:])  # Convertir a formato decimal
            respuesta = line[277:279].strip()
            
            data['RFC'].append(rfc)
            data['IMPORTE'].append(importe)
            data['RESPUESTA'].append(respuesta)

    # Crear un DataFrame con los datos extraídos.
    df = pd.DataFrame(data)

    # Crear un archivo Excel con encabezados azules y filtros.
    writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Obtener el objeto workbook y worksheet.
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Formatear los encabezados.
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#DDEBF7',
        'border': 1
    })

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Agregar autofiltros.
    worksheet.autofilter(0, 0, 0, len(df.columns) - 1)

    # Guardar y cerrar el archivo Excel.
    writer.close()

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