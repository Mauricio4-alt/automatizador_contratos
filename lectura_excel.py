import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from contratos import crear_contratos

def buscar_excel():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Excel", "*.xlsx")]
    )
    if ruta:
        excel_path.set(ruta)

def generar():
    try:
        if not excel_path.get():
            raise ValueError("Seleccione un archivo Excel")

        df = pd.read_excel(excel_path.get(), sheet_name="Contratados")
        crear_contratos(df)

        messagebox.showinfo("Éxito", "Contratos generados correctamente")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Generador de Contratos")
root.resizable(False, False)

excel_path = tk.StringVar()

tk.Label(root, text="Archivo Excel:").grid(row=0, column=0, padx=10, pady=10)

tk.Entry(root, textvariable=excel_path, width=40, state="readonly")\
    .grid(row=0, column=1, padx=5)

tk.Button(root, text="Buscar", command=buscar_excel)\
    .grid(row=0, column=2, padx=5)

tk.Button(root, text="Generar contratos", command=generar, width=25)\
    .grid(row=1, column=1, pady=20)

root.mainloop()
