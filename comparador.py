import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

def compare_excel_files(file1_path, file2_path):
    try:
        # Leer los archivos Excel sin encabezados y especificando las columnas por índice.
        df1 = pd.read_excel(file1_path, usecols=[4, 6], header=None, names=['RFC', 'Descuento'], skiprows=1)
        df2 = pd.read_excel(file2_path, usecols=[3, 5], header=None, names=['RFC', 'Descuento'], skiprows=1)
        
        # Convertir la columna Descuento a tipo numérico, manejando los diferentes formatos.
        df1['Descuento'] = pd.to_numeric(df1['Descuento'].replace({',': ''}, regex=True), errors='coerce').fillna(0)
        df2['Descuento'] = pd.to_numeric(df2['Descuento'], errors='coerce').fillna(0)
        
        # Calcular la suma de los descuentos para cada archivo.
        total_descuento_file1 = df1['Descuento'].sum()
        total_descuento_file2 = df2['Descuento'].sum()
        
        # Comparar los RFCs.
        merged_df = pd.merge(df1, df2, on='RFC', how='outer', indicator=True)
        
        # Identificar los RFCs faltantes en cada archivo.
        missing_in_file1 = merged_df[merged_df['_merge'] == 'right_only']
        missing_in_file2 = merged_df[merged_df['_merge'] == 'left_only']
        
        # Crear una nueva ventana para mostrar los resultados.
        result_window = tk.Toplevel(root)
        result_window.title("Diferencias de RFCs")
        
        # Agregar un título en la parte superior.
        title_label = tk.Label(result_window, text="202401", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)
        
        # Mostrar las sumas de los descuentos.
        total_label_file1 = tk.Label(result_window, text=f"Total Descuento Archivo 1: {total_descuento_file1:.2f}", font=("Helvetica", 12))
        total_label_file1.pack(pady=5)
        total_label_file2 = tk.Label(result_window, text=f"Total Descuento Archivo 2: {total_descuento_file2:.2f}", font=("Helvetica", 12))
        total_label_file2.pack(pady=5)
        
        # Crear un Treeview para mostrar los resultados.
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))
        style.configure("Treeview", font=("Helvetica", 10), rowheight=25)
        style.map("Treeview", background=[("selected", "lightblue")], foreground=[("selected", "black")])
        
        tree = ttk.Treeview(result_window, columns=("RFC", "Archivo"), show='headings', selectmode="extended")
        tree.heading("RFC", text="RFC")
        tree.heading("Archivo", text="Archivo")
        
        # Insertar los RFCs faltantes en el primer archivo.
        for index, row in missing_in_file1.iterrows():
            tree.insert("", "end", values=(row['RFC'], "Faltante en Archivo 1"))
        
        # Insertar los RFCs faltantes en el segundo archivo.
        for index, row in missing_in_file2.iterrows():
            tree.insert("", "end", values=(row['RFC'], "Faltante en Archivo 2"))
        
        tree.pack(expand=True, fill='both')
        
        # Añadir líneas de separación
        tree.tag_configure('oddrow', background='lightgrey')
        tree.tag_configure('evenrow', background='white')
        
        for i, item in enumerate(tree.get_children()):
            if i % 2 == 0:
                tree.item(item, tags=('evenrow',))
            else:
                tree.item(item, tags=('oddrow',))
        
        # Función para copiar los valores seleccionados al portapapeles.
        def copy_to_clipboard():
            selected_items = tree.selection()
            if selected_items:
                values = [tree.item(item, "values")[0] for item in selected_items]
                clipboard_text = "\n".join(values)
                result_window.clipboard_clear()
                result_window.clipboard_append(clipboard_text)
                messagebox.showinfo("Copiado", "RFCs copiados al portapapeles.")
            else:
                messagebox.showwarning("Advertencia", "No hay elementos seleccionados.")
        
        # Botón para copiar los RFCs seleccionados al portapapeles.
        copy_button = tk.Button(result_window, text="Copiar RFCs seleccionados", command=copy_to_clipboard)
        copy_button.pack(pady=10)
        
    except FileNotFoundError:
        messagebox.showerror("Error", "Uno o ambos archivos no se encontraron.")
    except KeyError:
        messagebox.showerror("Error", "Las columnas requeridas no están presentes en los archivos.")
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error al procesar los archivos: {e}")

def select_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="Seleccionar el primer archivo Excel")
    if file1_path:
        file1_label.config(text=f"Archivo 1: {file1_path}")

def select_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="Seleccionar el segundo archivo Excel")
    if file2_path:
        file2_label.config(text=f"Archivo 2: {file2_path}")

def compare_files():
    if file1_path and file2_path:
        compare_excel_files(file1_path, file2_path)
    else:
        messagebox.showerror("Error", "Por favor, selecciona ambos archivos Excel.")

# Crear la ventana principal de tkinter.
root = tk.Tk()
root.title("Comparador de RFC y Descuentos")

# Variables para almacenar las rutas de los archivos.
file1_path = ""
file2_path = ""

# Botones y etiquetas para seleccionar los archivos Excel.
file1_button = tk.Button(root, text="Seleccionar el primer archivo Excel", command=select_file1)
file1_button.pack(pady=10)
file1_label = tk.Label(root, text="Archivo 1: No seleccionado")
file1_label.pack(pady=5)

file2_button = tk.Button(root, text="Seleccionar el segundo archivo Excel", command=select_file2)
file2_button.pack(pady=10)
file2_label = tk.Label(root, text="Archivo 2: No seleccionado")
file2_label.pack(pady=5)

compare_button = tk.Button(root, text="Comparar archivos Excel", command=compare_files)
compare_button.pack(pady=20)

# Botón para salir de la aplicación.
exit_button = tk.Button(root, text="Salir", command=root.quit)
exit_button.pack(pady=10)

# Ejecutar el bucle principal de tkinter.
root.mainloop()