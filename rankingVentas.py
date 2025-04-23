import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def procesar_excel():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if not archivo:
        return

    try:
        df = pd.read_excel(archivo)
        df['FECHA_EMI'] = pd.to_datetime(df['FECHA_EMI'])
        df['MES'] = df['FECHA_EMI'].dt.strftime('%B %Y')
        df['TOTAL_SIN_IVA'] = df['IMP_EXENTO'].fillna(0) + df['IMP_GRAVAD'].fillna(0)

        resumen = df.groupby(['RAZON_SOC', 'MES'])['TOTAL_SIN_IVA'].sum().reset_index()
        tabla_final = resumen.pivot(index='RAZON_SOC', columns='MES', values='TOTAL_SIN_IVA').fillna(0)
        tabla_final = tabla_final.loc[:, sorted(tabla_final.columns, key=lambda x: pd.to_datetime(x, format='%B %Y'))]

        salida = archivo.replace('.xlsx', '_resumen.xlsx')
        tabla_final.to_excel(salida)
        messagebox.showinfo("Proceso finalizado", f"Resumen guardado en:\n{salida}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

root = tk.Tk()
root.withdraw()
procesar_excel()
