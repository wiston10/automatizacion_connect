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
        time.sleep(9)
        element = driver.find_element(By.XPATH, "/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div")

        # Si existe, hacer click inmediatamente
        if element:
            try:
                element.click()
                
            except Exception:
                 # Si falla el click normal, usar JavaScript
                driver.execute_script("arguments[0].click();", element)
        time.sleep(1)


        caja = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div[1]"))
            )

        caja.click()
        time.sleep(3)
        # agregar el correo de usuario a el input 
        agregar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div[1]/div[3]/div[2]/div/div/div[1]/input"))
        )
        agregar.send_keys(correo)
        # caja.send_keys(correo)
        time.sleep(6)
        opciones = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul[id^='option-list'] > li"))
        )
        #dar click al primer elemento que sale en la lista si esta el usuario deberia ser el primero 
        opciones[0].click()
    
        if esprimero:
        # click botton open
            boton_open_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/button/span[2]'))
            )
            boton_open_input.click()
            time.sleep(3)
        
            time_zone = wait.until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div/div/div/div'))
            )
        
            time_zone.clear()
            time_zone.send_keys("America/Mexico_City")  # Cambia por la zona que necesites
           
            time_zone.send_keys(zona)
            time.sleep(5)
            primer_opcion = wait.until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    'ul[id^="option-list"] > li.awsui_selectable-item_15o6u_2imeb_145:nth-child(1)'
            ))
            )
            primer_opcion.click()

            ## siguente fecha 
            input_date = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input.awsui_input_2rhyz_uz5yt_149"))
            )

            input_date.click()
            input_date.clear()
            input_date.send_keys(fecha)

        
    
        # buscar elementos
        print("Buscando datos...")
        button_search = wait.until(
                EC.element_to_be_clickable((By.XPATH,'/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div[1]/div/div/div/div/div/div[2]'))
        )
        #buscar 
        button_search.click()
        time.sleep(3)
    
        # esperamos los resultados
        # --- Checar si hay datos ---
         
        timeout = 12  # segundos
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

                # Lógica original para procesar datos
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
                            agregar_a_excel(rutasalida, new_data)

                        time_data = get_time(time_str)
                        new_data = [fecha, oracle_id, time_data]
                        estado_actual = "Available"

                    # Si ya junta 5 elementos, guardar
                    if len(new_data) >= 5:
                        agregar_a_excel(rutasalida, new_data)
                        new_data = []

                             
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





# Ventana principal
ventana = tk.Tk()
ventana.title("Lector de Excel")

# Input fecha con calendario
tk.Label(ventana, text="Seleccione una fecha:").pack()
entry_fecha = DateEntry(ventana, width=20, background='darkblue',
                        foreground='white', date_pattern='yyyy-mm-dd')  # 👈 Aquí va el calendario
entry_fecha.pack(pady=5)

# Select con opciones predeterminadas (combobox)
tk.Label(ventana, text="Seleccione una zona horaria:").pack()
combo = ttk.Combobox(ventana, values=["America/Mexico_City", "America/Argentina/Buenos_Aires", "America/Lima"])
combo.current(0)  # Selecciona la primera opción por defecto
combo.pack(pady=5)

# Botón para cargar Excel
boton = tk.Button(ventana, text="Cargar Excel", command=cargar_excel, width=20, height=2)
boton.pack(pady=20)

ventana.mainloop()
