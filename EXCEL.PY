import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re

def process_excel_file(excel_file_path, output_file_path):
    # Leer el archivo Excel.
    df = pd.read_excel(excel_file_path)
    
    # Imprimir los nombres de las columnas para verificar.
    print("Columnas en el archivo Excel:", df.columns.tolist())
    
    # Eliminar todos los caracteres excepto números y puntos decimales en la columna IMPORTE.
    df['IMPORTE'] = df['IMPORTE'].astype(str).apply(lambda x: re.sub(r'[^0-9.]', '', x))
    
    # Convertir la columna IMPORTE a tipo numérico.
    df['IMPORTE'] = pd.to_numeric(df['IMPORTE'], errors='coerce').fillna(0)
    
    # Agrupar por RFC y sumar los importes.
    try:
        df_grouped = df.groupby('RFC', as_index=False).agg({
            'ENTIDAD': 'first',
            'PROCESO_DE_NOMINA': 'first',
            'NOMBRE': 'first',
            'PRIMER_APELLIDO': 'first',
            'SEGUNDO_APELLIDO': 'first',
            'CURP': 'first',
            'RFC': 'first',
            'CTA_INTERBANCARIA': 'first',
            'CLC': 'first',
            'CVE_CONCEPTO': 'first',
            'DESCRIPCION': 'first',
            'IMPORTE': 'sum'
        })
    except KeyError as e:
        messagebox.showerror("Error", f"Column(s) {e} do not exist in the file.")
        return

    # Calcular la suma total de los importes.
    total_importe = df_grouped['IMPORTE'].sum()

    # Agregar una fila adicional con la suma total.
    total_row = pd.DataFrame([{
        'ENTIDAD': '',
        'PROCESO_DE_NOMINA': '',
        'NOMBRE': '',
        'PRIMER_APELLIDO': '',
        'SEGUNDO_APELLIDO': '',
        'CURP': '',
        'RFC': 'TOTAL',
        'CTA_INTERBANCARIA': '',
        'CLC': '',
        'CVE_CONCEPTO': '',
        'DESCRIPCION': '',
        'IMPORTE': total_importe
    }])
    df_grouped = pd.concat([df_grouped, total_row], ignore_index=True)

    # Formatear los valores de la columna IMPORTE para que tengan dos decimales.
    df_grouped['IMPORTE'] = df_grouped['IMPORTE'].apply(lambda x: f"{x:.2f}")

    # Guardar el resultado en un nuevo archivo Excel.
    df_grouped.to_excel(output_file_path, index=False)

def select_file():
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_file_path:
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_file_path:
            process_excel_file(excel_file_path, output_file_path)
            messagebox.showinfo("Success", f"File processed and saved as {output_file_path}")

# Crear la ventana principal de tkinter.
root = tk.Tk()
root.title("Procesador de Excel")

# Crear un botón para seleccionar el archivo Excel.
select_button = tk.Button(root, text="Seleccionar archivo Excel", command=select_file)
select_button.pack(pady=20)

# Ejecutar el bucle principal de tkinter.
root.mainloop()