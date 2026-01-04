import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import time

def agregar_a_excel(ruta_archivo, fila):
    # Leer el Excel con header=None
    df_actual = pd.read_excel(ruta_archivo, header=None)
    
    # Si está vacío, definimos columnas
    if df_actual.empty:
        df_actual = pd.DataFrame(columns=range(len(fila)))
    
    # Agregamos la nueva fila
    df_actual.loc[len(df_actual)] = fila
    
    # Guardamos de nuevo con openpyxl (más seguro que xlsxwriter si no está instalado)
    with pd.ExcelWriter(ruta_archivo, engine="openpyxl", mode="w") as writer:
        df_actual.to_excel(writer, index=False, header=False)


def cargar_excel():
    archivo = filedialog.askopenfilename(
        title="Selecciona un archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")],
    )
    if archivo:
        try:
            # Leer hoja llamada Data
            df = pd.read_excel(archivo, sheet_name="Data", header=None)
            ruta_script = os.path.dirname(os.path.abspath(__file__))
            ruta_carpeta = os.path.join(ruta_script, "formateados")
            os.makedirs(ruta_carpeta, exist_ok=True)
            nombre_archivo = entrada_nombre.get().strip()
            if not nombre_archivo:
                messagebox.showwarning("Atención", "Debes ingresar un nombre para el nuevo Excel")
                return
            ruta_salida = os.path.join(ruta_carpeta, f"{nombre_archivo}.xlsx")
            
            # Crear DataFrame vacío
            nuevo_excel = pd.DataFrame(columns=[0,1,2,3,4])  
            with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode="w") as writer:
                nuevo_excel.to_excel(writer, index=False, header=False)
       
            time.sleep(1)

            # Recorremos fila por fila
            for index, fila in df.iterrows():
                cambio = ""
                # Si es la fila 0 → se agrega tal cual sin formatear
                if index == 0:
                    inicio = str(fila[2])
                    fin = str(fila[3])

                    # Intentar formatear solo si son horas válidas
                    # try:
                    #     inicio_dt = datetime.strptime(inicio, "%H:%M:%S")
                    #     inicio_fmt = inicio_dt.strftime("%H:%M:%S")
                    # except:
                    #     inicio_fmt = inicio

                    # try:
                    #     fin_dt = datetime.strptime(fin, "%H:%M:%S")
                    #     fin_fmt = fin_dt.strftime("%H:%M:%S")
                    # except:
                    #     fin_fmt = fin

                    add = [
                        fila[0],
                        fila[1],
                        inicio,
                        fin,
                        fila[4]
                    ]
                    agregar_a_excel(ruta_salida, add)
                    continue

                fecha = fila[0]    
                codigo = fila[1]
                inicio = str(fila[2])
                fin = str(fila[3])
                texto = fila[4]
            
                # Convertir ambos a datetime
                inicio_dt = datetime.strptime(inicio, "%H:%M:%S")
                fin_dt = datetime.strptime(fin, "%H:%M:%S")

                # Por defecto: guardar en formato AM/PM
                add = [
                    fecha,
                    codigo,
                    inicio_dt.strftime("%H:%M:%S"),
                    fin_dt.strftime("%H:%M:%S"),
                    texto
                ]

                # Validar con la fila anterior
                final_anterior = df.iloc[index - 1, 3]
                final_anterior_hora = str(final_anterior)[:5]
                codigo_anterior = df.iloc[index - 1, 1]
                if inicio[:5] == final_anterior_hora and codigo == codigo_anterior:
                    hora_dt_mas_1 = inicio_dt + pd.Timedelta(minutes=1)

                    # Si la nueva hora de inicio es igual a la hora final → sumar 1 minuto a la final
                    if hora_dt_mas_1.strftime("%H:%M") == fin_dt.strftime("%H:%M"):
                        fin_dt += pd.Timedelta(minutes=1)
                        print("la fila" + str(index) + " se ha modificado para evitar horas iguales")
                    # Ahora guardamos la fila actualizada
                    add = [
                        fecha,
                        codigo,
                        hora_dt_mas_1.strftime("%H:%M:%S"),
                        fin_dt.strftime("%H:%M:%S"),
                        texto
                    ]

                # ✅ Verificar que inicio y fin en "add" no sean iguales (sin segundos)
                if add[2][:5] == add[3][:5]:
                    fin_dt = datetime.strptime(add[3], "%H:%M:%S") + pd.Timedelta(minutes=1)
                    add[3] = fin_dt.strftime("%H:%M:%S" \
                    "") 
                    print("la fila" + str(index) + " se ha modificado para evitar horas iguales")
                    
                print(index)
                agregar_a_excel(ruta_salida, add)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")


# Ventana principal
ventana = tk.Tk()
ventana.title("formateador")

tk.Label(ventana, text="Nombre del nuevo Excel:").pack(pady=5,padx=10)
entrada_nombre = tk.Entry(ventana, width=30)
entrada_nombre.pack(pady=5, padx=10)

boton = tk.Button(ventana, text="Cargar Excel", command=cargar_excel, width=20, height=2)
boton.pack(pady=20)

ventana.mainloop()
