import pandas as pd
import pyodbc
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from decimal import Decimal
import threading

# Función para procesar el archivo Excel
def process_file(file_path, connection_string, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo):
    try:
        df = pd.read_excel(file_path, usecols=[6, 11], header=None, names=['RFC', 'Descuento'], skiprows=1)

        df['Descuento'] = df['Descuento'].astype(str).str.replace("'", "", regex=False).str.replace(",", "", regex=False).replace({"": "0"})
        df['Descuento'] = pd.to_numeric(df['Descuento'], errors='coerce').fillna(0)

        show_preview(df, connection_string, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo)
    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar el archivo: {e}")

# Función para mostrar la previsualización
def show_preview(df, connection_string, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo):
    preview_window = tk.Toplevel()
    preview_window.title("Previsualización de Datos")
    preview_window.geometry("600x400")

    tree = ttk.Treeview(preview_window, columns=("RFC", "Descuento"), show="headings")
    tree.heading("RFC", text="RFC")
    tree.heading("Descuento", text="Descuento")
    tree.column("RFC", width=200, anchor="center")
    tree.column("Descuento", width=200, anchor="center")
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    for _, row in df.iterrows():
        tree.insert("", "end", values=(row['RFC'], row['Descuento']))

    status_label = tk.Label(preview_window, text="")
    status_label.pack(pady=10)

    progress_bar = ttk.Progressbar(preview_window, orient="horizontal", length=400, mode="determinate")
    progress_bar.pack(pady=10)

    tk.Button(preview_window, text="Subir Datos", command=lambda: upload_data(df, connection_string, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo, preview_window, status_label, progress_bar)).pack(pady=10)
    tk.Button(preview_window, text="Cancelar", command=preview_window.destroy).pack(pady=10)

# Función para subir los datos a la base de datos
def upload_data(df, connection_string, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo, preview_window, status_label, progress_bar):
    def run():
        try:
            conn = pyodbc.connect(connection_string)
            cursor = conn.cursor()

            total_rows = len(df)
            progress_bar["maximum"] = total_rows

            for i, (_, row) in enumerate(df.iterrows(), start=1):
                cursor.execute("""
                    EXEC dbo.usp_RespuestasDomiciliacion
                    @RFC = ?, @Monto = ?, @CodigoRespuesta = ?, @FechaImportacion = ?, 
                    @IdTipoFondo = ?, @QnaAplicacion = ?, @QnaInicial = ?, @QnaFinal = ?, 
                    @QnaProceso = ?, @NumPagos = ?
                """, row['RFC'], Decimal(str(row['Descuento'])), '00', pd.Timestamp.now(),
                   tipo_fondo, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number)

                progress_bar["value"] = i
                status_label.config(text=f"Subiendo datos... {i}/{total_rows}")
                preview_window.update_idletasks()

            conn.commit()
            cursor.close()
            conn.close()

            messagebox.showinfo("Éxito", "Datos guardados correctamente.")
            preview_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Error al subir los datos: {e}")

    threading.Thread(target=run).start()  # Evita que la GUI se congele

# Función para seleccionar el archivo Excel
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            qna_aplicacion = simpledialog.askinteger("Qna Aplicación", "Ingrese Qna Aplicación:")
            qna_inicial = simpledialog.askinteger("Qna Inicial", "Ingrese Qna Inicial:")
            qna_final = simpledialog.askinteger("Qna Final", "Ingrese Qna Final:")
            qna_proceso = simpledialog.askinteger("Qna Proceso", "Ingrese Qna Proceso:")
            payment_number = simpledialog.askinteger("Número de Cobro", "Ingrese Número de Cobro:")
            tipo_fondo = simpledialog.askinteger("Tipo de Fondo", "Ingrese Tipo de Fondo (1 para Fabes, 2 para Jubilados):")

            if None in [qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo]:
                messagebox.showwarning("Advertencia", "Debe ingresar todos los valores.")
                return
            
            connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=76.74.150.73;DATABASE=FabesBD;UID=sa;PWD=tQs73Z39XcGi;TrustServerCertificate=Yes;"
            process_file(file_path, connection_string, qna_aplicacion, qna_inicial, qna_final, qna_proceso, payment_number, tipo_fondo)
        except Exception as e:
            messagebox.showerror("Error", f"Error en los parámetros ingresados: {e}")

# Crear ventana principal
root = tk.Tk()
root.title("Subir Domiciliación")

tk.Button(root, text="Seleccionar archivo Excel", command=select_file).pack(pady=20)

root.mainloop()