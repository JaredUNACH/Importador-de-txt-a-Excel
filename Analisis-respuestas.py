import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

def compare_excel_files(file1_path, file2_path):
    try:
        # Leer los archivos Excel, solo la columna A (RFC).
        df1 = pd.read_excel(file1_path, usecols=[0], header=None, names=['RFC'])
        df2 = pd.read_excel(file2_path, usecols=[0], header=None, names=['RFC'])
        
        # Convertir las columnas RFC a conjuntos para facilitar la comparación.
        rfc_set1 = set(df1['RFC'])
        rfc_set2 = set(df2['RFC'])
        
        # Encontrar los RFCs faltantes en cada archivo.
        missing_in_file1 = rfc_set2 - rfc_set1
        missing_in_file2 = rfc_set1 - rfc_set2
        
        # Crear un DataFrame para mostrar los resultados.
        result_data = {
            'RFC Faltante en Archivo 1': list(missing_in_file1),
            'RFC Faltante en Archivo 2': list(missing_in_file2)
        }
        result_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in result_data.items()]))
        
        # Crear una nueva ventana para mostrar los resultados.
        result_window = tk.Toplevel(root)
        result_window.title("Diferencias de RFCs")
        
        # Crear un Canvas con barras de desplazamiento.
        canvas = tk.Canvas(result_window)
        scrollbar_y = tk.Scrollbar(result_window, orient="vertical", command=canvas.yview)
        scrollbar_x = tk.Scrollbar(result_window, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Crear un Frame para los botones.
        button_frame = ttk.Frame(result_window)
        
        # Agregar un título en la parte superior.
        title_label = tk.Label(scrollable_frame, text="Comparación de RFCs", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)
        
        # Crear un Treeview para mostrar los RFCs faltantes.
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))
        style.configure("Treeview", font=("Helvetica", 10), rowheight=25)
        style.map("Treeview", background=[("selected", "lightblue")], foreground=[("selected", "black")])
        
        tree_missing = ttk.Treeview(scrollable_frame, columns=("RFC", "Archivo"), show='headings', selectmode="extended")
        tree_missing.heading("RFC", text="RFC")
        tree_missing.heading("Archivo", text="Archivo")
        
        # Insertar los RFCs faltantes en el primer archivo.
        for rfc in missing_in_file1:
            tree_missing.insert("", "end", values=(rfc, "Faltante en Archivo 1"))
        
        # Insertar los RFCs faltantes en el segundo archivo.
        for rfc in missing_in_file2:
            tree_missing.insert("", "end", values=(rfc, "Faltante en Archivo 2"))
        
        tree_missing.pack(expand=True, fill='both')
        
        # Añadir líneas de separación
        tree_missing.tag_configure('oddrow', background='lightgrey')
        tree_missing.tag_configure('evenrow', background='white')
        
        for i, item in enumerate(tree_missing.get_children()):
            if i % 2 == 0:
                tree_missing.item(item, tags=('evenrow',))
            else:
                tree_missing.item(item, tags=('oddrow',))
        
        # Función para copiar los valores seleccionados al portapapeles.
        def copy_to_clipboard(tree):
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
        copy_button_missing = tk.Button(button_frame, text="Copiar RFCs seleccionados", command=lambda: copy_to_clipboard(tree_missing))
        copy_button_missing.pack(pady=10)
        
        # Función para exportar la comparación de RFCs a un archivo Excel.
        def export_rfc_comparison():
            output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if output_file_path:
                result_df.to_excel(output_file_path, index=False)
                messagebox.showinfo("Éxito", f"Comparación de RFCs exportada a {output_file_path}")
        
        # Botón para exportar la comparación de RFCs a un archivo Excel.
        export_rfc_button = tk.Button(button_frame, text="Exportar comparación de RFCs a Excel", command=export_rfc_comparison)
        export_rfc_button.pack(pady=10)
        
        # Empaquetar el Canvas y las barras de desplazamiento.
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        
        # Empaquetar el Frame de los botones.
        button_frame.pack(side="right", fill="y")
        
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
root.title("Comparador de RFCs en Excel")

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