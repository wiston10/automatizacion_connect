import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry   # 👈 Importar DateEntry
import pandas as pd
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By  
from selenium.webdriver.common.keys import Keys
import os
from datetime import datetime

perfilepath = "/Users/wistonnavarroaroca/Library/Application Support/Firefox/Profiles/asu90e8f.default-release"


url = "https://dish-cc.my.connect.aws/historical-metrics/agent-activity-audit"

import time
from datetime import datetime

def get_time(time_str):
    """
    Convierte 'Nov 19, 2025 07:13:18 AM'
    a '07:13:18' en 24 horas.
    """
    try:
        fecha_hora = datetime.strptime(time_str.strip(), "%b %d, %Y %I:%M:%S %p")
        return fecha_hora.strftime("%H:%M:%S")
    except Exception as e:
        print("⚠️ Error parseando fecha:", time_str, e)
        return ""


def agregar_a_excel(ruta_archivo, fila):
    # Leer el Excel con header=None
    df_actual = pd.read_excel(ruta_archivo, header=None)
    
    # Si está vacío, definimos columnas
    if df_actual.empty:
        df_actual = pd.DataFrame(columns=range(len(fila)))
    
    # Agregamos la nueva fila
    df_actual.loc[len(df_actual)] = fila
    
    # Guardamos de nuevo
    df_actual.to_excel(ruta_archivo, index=False, header=False)
 

def consultar_web(driver, correo, fecha, zona, oracle_id, wait,rutasalida,esprimero):
    """
    Usa el navegador existente para consultar los datos
    """
    print(f"Consultando datos para {correo}, {fecha}, {zona}, {oracle_id}")
    try:
        # selector por correo de usuario 
        # Esperar que el overlay desaparezca
        element = driver.find_element(By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/button/span[2]")

        # Si existe, hacer click inmediatamente
        if element:
            try:
                element.click()
                
            except Exception:
                 # Si falla el click normal, usar JavaScript
                driver.execute_script("arguments[0].click();", element)
            
        time.sleep(1)


        caja = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div[1]/div[3]/div[2]/div/div/div[1]/input"))
            )
        caja.click()
        time.sleep(1)
        # agregar el correo de usuario a el input 
        caja.send_keys(correo)
        time.sleep(3)
        opciones = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.awsui_selectable-item_15o6u_2imeb_145"))
        )
        #dar click al primer elemento que sale en la lista si esta el usuario deberia ser el primero 
        opciones[0].click()
    
        if esprimero:
        # click botton open
            boton_open_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/button/span[2]'))
            )
            boton_open_input.click()
            time.sleep(2)
        
            time_zone = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div[1]/div[3]/div[2]/div/div/div/input"))
            )
            time_zone.click()
            time_zone.send_keys(zona)
            time.sleep(2)
            primer_opcion = wait.until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    'ul[id^="option-list"] > li.awsui_selectable-item_15o6u_2imeb_145:nth-child(1)'
            ))
            )
            primer_opcion.click()

            ## siguente fecha 
            input_date = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='YYYY/MM/DD']"))
            )

            input_date.click()
            input_date.clear()
            input_date.send_keys(fecha)

        
    
        # buscar elementos
        button_search = wait.until(
                EC.element_to_be_clickable((By.XPATH,'/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[2]/button'))
        )
        #buscar 
        button_search.click()
        time.sleep(2)
    
        # esperamos los resultados
        # --- Checar si hay datos ---
         
        timeout = 10  # segundos
        intervalo = 1  # cada cuánto revisar
        tiempo_transcurrido = 0 
         
        while tiempo_transcurrido < timeout:
            # Verificamos si aparece el mensaje de "no hay datos"
            # no_data_elements = driver.find_elements(By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr/td/div/div")
            # if no_data_elements and no_data_elements[0].is_displayed():
            #     print("⚠️ No hay datos para esta consulta")
            #     time.sleep(1)
            #     break

            # Verificamos si aparece la tabla y es visible
            tabla_elements = driver.find_elements(By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/table")

            if tabla_elements and tabla_elements[0].is_displayed():
                tabla = tabla_elements[0]

                # Obtener encabezados del thead
                encabezados = [th.text.split("\n")[0] for th in tabla.find_elements(By.XPATH, ".//thead/tr/th")]

                datos = []
                no_records = tabla.find_elements(
                By.XPATH,
                ".//tbody/tr/td/div/div[contains(@class,'awsui_empty')]"
                )

                if no_records:
                    print("⚠️ No hay datos para esta consulta: " ,correo)
                    break
                # Obtener todas las filas del tbody
                for fila in tabla.find_elements(By.XPATH, ".//tbody/tr"):
                    columnas = fila.find_elements(By.TAG_NAME, "td")
                    fila_dict = {encabezados[i]: columnas[i].text for i in range(len(columnas))}
                    datos.append(fila_dict)

                # 1️⃣ GUARDAR EN TEMPORAL: Extraer todos los rangos primero
                temporal_data = []
                new_data = []
                estado_actual = ""

                for i, fila in enumerate(datos):
                    time_str = fila["Time"]

                    if i == 0:
                        new_data = []
                        estado_actual = "Offline"

                    # Si pasa de Available → Offline: se cierra rango
                    if fila["Status"] == "Offline" and estado_actual == "Available":
                        salida = get_time(time_str)
                        new_data.append(salida)
                        new_data.append("Central Standard Time")
                        estado_actual = "Offline"

                    # Si pasa de Offline → Available: se abre rango
                    elif estado_actual == "Offline" and fila["Status"] != "Offline":
                        if len(new_data) >= 3:
                            temporal_data.append(new_data.copy())

                        time_data = get_time(time_str)
                        new_data = [fecha, oracle_id, time_data]
                        estado_actual = "Available"

                    # Si ya junta 5 elementos, guardar en temporal
                    if len(new_data) >= 5:
                        temporal_data.append(new_data.copy())
                        new_data = []

                # 2️⃣ REVISAR HORA Y MINUTO IGUALES: Comparar con fila anterior
                for idx, fila_temp in enumerate(temporal_data):
                    if idx > 0:
                        # Comparar hora:minuto de inicio actual con hora:minuto de fin anterior
                        hora_inicio_actual = fila_temp[2][:5]  # HH:MM
                        hora_fin_anterior = temporal_data[idx-1][3][:5]  # HH:MM
                        oracle_id_anterior = temporal_data[idx-1][1]
                        
                        # Si son iguales y mismo Oracle ID, ajustar
                        if hora_inicio_actual == hora_fin_anterior and fila_temp[1] == oracle_id_anterior:
                            try:
                                inicio_dt = datetime.strptime(fila_temp[2], "%H:%M:%S")
                                fin_dt = datetime.strptime(fila_temp[3], "%H:%M:%S")
                                
                                # Sumar 1 minuto al inicio
                                inicio_dt += pd.Timedelta(minutes=1)
                                
                                # Si inicio queda igual al fin, sumar 1 minuto al fin también
                                if inicio_dt.strftime("%H:%M") == fin_dt.strftime("%H:%M"):
                                    fin_dt += pd.Timedelta(minutes=1)
                                
                                fila_temp[2] = inicio_dt.strftime("%H:%M:%S")
                                fila_temp[3] = fin_dt.strftime("%H:%M:%S")
                                print(f"⚠️ Ajustado rango {idx} para {oracle_id} - hora igual detectada")
                            except Exception as e:
                                print(f"⚠️ Error al ajustar hora: {e}")
                    
                    # Verificar que inicio y fin no sean iguales en la misma fila
                    if fila_temp[2][:5] == fila_temp[3][:5]:
                        try:
                            fin_dt = datetime.strptime(fila_temp[3], "%H:%M:%S")
                            fin_dt += pd.Timedelta(minutes=1)
                            fila_temp[3] = fin_dt.strftime("%H:%M:%S")
                            print(f"⚠️ Ajustado fin de rango {idx} - inicio/fin eran iguales")
                        except Exception as e:
                            print(f"⚠️ Error al ajustar fin: {e}")
                
                # 3️⃣ ENVIAR AL EXCEL: Guardar todas las filas procesadas
                for fila_final in temporal_data:
                    agregar_a_excel(rutasalida, fila_final)

                             
                break
            # Si no apareció ninguno, esperar un segundo y volver a revisar
            time.sleep(intervalo)
            tiempo_transcurrido += intervalo
        else:
            # Si se cumple el timeout sin encontrar ninguno
            print("⏱ Timeout: no apareció ni el mensaje de 'no hay datos' ni la tabla")
        
        
       
        
 
     

    except Exception as e:
        print("❌ Error en la consulta:", e)

def abrir_firefox():
    # Cargar el perfil que usas normalmente (default-release)
    profile = FirefoxProfile(perfilepath)

    options = Options()
    options.profile = profile  # asignar el perfil

    # Opciones para intentar ocultar Selenium
    options.set_preference("dom.webdriver.enabled", False)
    options.set_preference("useAutomationExtension", False)
    options.set_preference("general.useragent.override", 
                           "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:141.0) Gecko/20100101 Firefox/141.0")

    # Iniciar Firefox con ese perfil
    driver = webdriver.Firefox(options=options)
    driver.maximize_window()
    driver.get(url)
    return driver


def cargar_excel():
    archivo = filedialog.askopenfilename(
        title="Selecciona un archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        try:
            #excel con la hoja llamada Boots
            df = pd.read_excel(archivo,sheet_name='Boost')

            # Tomamos la fecha y la opción seleccionada
            fecha = entry_fecha.get()  # con DateEntry también usas .get()
            opcion = combo.get()
    
            # validar datos necesarios para el procedimiento si no se tienen mostrar error
            if not fecha:
                messagebox.showerror("Error", "Debe seleccionar una fecha antes de continuar.")
                return
            if not opcion:
                messagebox.showerror("Error", "Debe seleccionar una zona horaria antes de continuar.")
                return

            # Crear carpeta 'output' junto al script
            ruta_script = os.path.dirname(os.path.abspath(__file__))
            ruta_carpeta = os.path.join(ruta_script, "output")
            os.makedirs(ruta_carpeta, exist_ok=True)
            # Nombre del nuevo archivo según la fecha
            ruta_salida = os.path.join(ruta_carpeta, f"{fecha}.xlsx")
            # Crear DataFrame vacío para ir llenando filas
            nuevo_excel = pd.DataFrame(columns=[0,1,2,3,4])  # solo 5 columnas, sin títulos
            # Guardar inicialmente el Excel vacío
            nuevo_excel.to_excel(ruta_salida, index=False, header=False)
            driver = abrir_firefox()
            wait = WebDriverWait(driver, 60)
            time.sleep(2)
            # Recorremos fila por fila y usamos el correo
            for index, fila in df.iterrows():
                es_primero = index == df.index[0]
                correo = fila["AWS Name"]    
                oracle_id = fila["Oracle or Kronos ID"]
                consultar_web(driver, correo, fecha, opcion, oracle_id,wait,ruta_salida,es_primero)

            driver.quit() 

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")




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


def formatear_excel():
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
                    
                    elif fin_dt.strftime("%H:%M")== hora_dt_mas_1.strftime("%H:%M"):
                        hora_dt_mas_1 += pd.Timedelta(minutes=1)
                        print("la fila" + str(index) + "se ha modificado para evitar horas iguales")

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


# ═══════════════════════════════════════════════════════════════════════════
# 🎨 INTERFAZ GRÁFICA PROFESIONAL
# ═══════════════════════════════════════════════════════════════════════════

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Sistema de Gestión de Datos | TTEC")
ventana.geometry("600x450")
ventana.configure(bg="#f0f4f8")
ventana.resizable(False, False)

# Estilos personalizados
style = ttk.Style()
style.theme_use('clam')

# Estilo para el Combobox
style.configure("Custom.TCombobox", 
                fieldbackground="#ffffff",
                background="#2c3e50",
                foreground="#2c3e50",
                arrowcolor="#2c3e50",
                bordercolor="#3498db",
                lightcolor="#ffffff",
                darkcolor="#2c3e50")

# ═══════════════════════════════════════════════════════════════════════════
# 📋 FRAME PRINCIPAL CON TÍTULO
# ═══════════════════════════════════════════════════════════════════════════

# Frame del título
frame_titulo = tk.Frame(ventana, bg="#2c3e50", height=80)
frame_titulo.pack(fill="x", pady=0)

titulo_label = tk.Label(
    frame_titulo,
    text="📊 PROCESADOR DE DATOS",
    font=("Helvetica Neue", 22, "bold"),
    bg="#2c3e50",
    fg="#ffffff"
)
titulo_label.pack(pady=20)

subtitulo_label = tk.Label(
    frame_titulo,
    text="Extracción y formateo automático de información",
    font=("Helvetica Neue", 10),
    bg="#2c3e50",
    fg="#bdc3c7"
)
subtitulo_label.pack()

# ═══════════════════════════════════════════════════════════════════════════
# 📝 FRAME DE CONTENIDO
# ═══════════════════════════════════════════════════════════════════════════

frame_contenido = tk.Frame(ventana, bg="#f0f4f8")
frame_contenido.pack(fill="both", expand=True, padx=40, pady=30)

# ─────────────────────────────────────────────────────────────────────────
# 📅 SECCIÓN DE FECHA
# ─────────────────────────────────────────────────────────────────────────

frame_fecha = tk.Frame(frame_contenido, bg="#f0f4f8")
frame_fecha.pack(fill="x", pady=10)

label_fecha = tk.Label(
    frame_fecha,
    text="📅 Fecha de consulta:",
    font=("Helvetica Neue", 12, "bold"),
    bg="#f0f4f8",
    fg="#2c3e50"
)
label_fecha.pack(anchor="w", pady=(0, 5))

entry_fecha = DateEntry(
    frame_fecha,
    width=25,
    background="#3498db",
    foreground="white",
    borderwidth=2,
    date_pattern='yyyy-mm-dd',
    font=("Helvetica Neue", 11)
)
entry_fecha.pack(fill="x", ipady=5)

# ─────────────────────────────────────────────────────────────────────────
# 🌍 SECCIÓN DE ZONA HORARIA
# ─────────────────────────────────────────────────────────────────────────

frame_zona = tk.Frame(frame_contenido, bg="#f0f4f8")
frame_zona.pack(fill="x", pady=10)

label_zona = tk.Label(
    frame_zona,
    text="🌍 Zona horaria:",
    font=("Helvetica Neue", 12, "bold"),
    bg="#f0f4f8",
    fg="#f0f4f8"
)
label_zona.pack(anchor="w", pady=(0, 5))

combo = ttk.Combobox(
    frame_zona,
    values=["America/Mexico_City", "America/Argentina/Buenos_Aires", "America/Lima"],
    font=("Helvetica Neue", 11),
    state="readonly",
    style="Custom.TCombobox"
)
combo.current(0)
combo.pack(fill="x", ipady=5)

# ─────────────────────────────────────────────────────────────────────────
# 📁 SECCIÓN DE BOTONES
# ─────────────────────────────────────────────────────────────────────────

frame_botones = tk.Frame(frame_contenido, bg="#98bee3")
frame_botones.pack(fill="x", pady=20)

# Botón principal - Cargar Excel
boton_cargar = tk.Button(
    frame_botones,
    text="📂 CARGAR Y PROCESAR EXCEL",
    command=cargar_excel,
    font=("Helvetica Neue", 12, "bold"),
    bg="#3498db",
    fg="black",
    activebackground="#2980b9",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    padx=20,
    pady=15
)
boton_cargar.pack(fill="x")

# Efecto hover para el botón
def on_enter_cargar(e):
    boton_cargar.config(bg="#2980b9")

def on_leave_cargar(e):
    boton_cargar.config(bg="#3498db")

boton_cargar.bind("<Enter>", on_enter_cargar)
boton_cargar.bind("<Leave>", on_leave_cargar)

# ═══════════════════════════════════════════════════════════════════════════
# 📌 FOOTER
# ═══════════════════════════════════════════════════════════════════════════

frame_footer = tk.Frame(ventana, bg="#ecf0f1", height=40)
frame_footer.pack(side="bottom", fill="x")

footer_label = tk.Label(
    frame_footer,
    text="© 2025 | Sistema automatizado de gestión de datos",
    font=("Helvetica Neue", 9),
    bg="#ecf0f1",
    fg="#7f8c8d"
)
footer_label.pack(pady=10)

# ═══════════════════════════════════════════════════════════════════════════
# 🚀 INICIAR APLICACIÓN
# ═══════════════════════════════════════════════════════════════════════════

ventana.mainloop()
