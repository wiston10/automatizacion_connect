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
def get_time(time_str):
    """
    Convierte un string con formato '30 sept 2025, 12:00:06'
    y devuelve solo la hora en formato 24h, ejemplo '12:00:06'
    """
    meses = {
        "ene": "Jan", "feb": "Feb", "mar": "Mar", "abr": "Apr",
        "may": "May", "jun": "Jun", "jul": "Jul", "ago": "Aug",
        "sept": "Sep", "oct": "Oct", "nov": "Nov", "dic": "Dec"
    }

    # Reemplazar el mes en español por inglés
    for mes_es, mes_en in meses.items():
        if mes_es in time_str.lower():
            time_str_en = time_str.lower().replace(mes_es, mes_en)
            break
    else:
        time_str_en = time_str

    # Convertir a datetime
    fecha_hora = datetime.strptime(time_str_en, "%d %b %Y, %H:%M:%S")

    # Devolver solo la hora en 24h
    return fecha_hora.strftime("%H:%M:%S")

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
 

def consultar_web(driver, correo, fecha, zona, oracle_id, wait,rutasalida):
    """
    Usa el navegador existente para consultar los datos
    """
    print(f"Consultando datos para {correo}, {fecha}, {zona}, {oracle_id}")
    try:
        # Esperar que el overlay desaparezca
        elements = driver.find_elements(By.CSS_SELECTOR, "#s2id_autogen2 > a:nth-child(1) > abbr:nth-child(2)")

        # Si existe, hacer click inmediatamente
        if elements:
            try:
                elements[0].click()
                
            except Exception:
                 # Si falla el click normal, usar JavaScript
                driver.execute_script("arguments[0].click();", elements[0])


        caja = wait.until(
            EC.element_to_be_clickable((By.ID, "s2id_autogen2"))  # el div visible
        )
        caja.click()
        
         # 2. Una vez abierto, ubica el input de búsqueda de select2
        campo_busqueda = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#select2-drop > div:nth-child(1) > input:nth-child(1)"))
        )
        
        # 3. Escribe el correo en el input
        campo_busqueda.clear()
        campo_busqueda.send_keys(correo)
        time.sleep(2)  # dar un respiro para que carguen las coincidencias
        
        campo_busqueda.send_keys(Keys.ENTER)  # presiona Enter para seleccionar la primera coincidencia
        
        #rellenar lo que es ahora la fecha 
        fecha_input = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="startDate"]'))
        )
        fecha_input.clear() 
        fecha_input.send_keys(fecha)
    
        #rellenar campos de zona horaria 
        zona_input = wait.until(
            EC.element_to_be_clickable((By.ID, 's2id_timezoneSelect'))
        )
        zona_input.click()
        # una vez abierto ubicar el input para escribir la zona 
        zona_busqueda = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#select2-drop > div:nth-child(1) > input:nth-child(1)"))
        )
        zona_busqueda.clear()
        zona_busqueda.send_keys(zona)
        time.sleep(3)   
        zona_busqueda.send_keys(Keys.ENTER)  # presiona Enter para seleccionar la primera coincidencia

        # una vez los campos llenos procedemos a buscar los elementos 
        buscar_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div[2]/div[2]/div/form/div/div/div[3]/div[1]/button"))
        )
        buscar_btn.click()
         
        # esperamos los resultados
         # --- Checar si hay datos ---
         
        timeout = 12  # segundos
        intervalo = 1  # cada cuánto revisar
        tiempo_transcurrido = 0 
         
        while tiempo_transcurrido < timeout:
        # Verificamos si aparece el mensaje de "no hay datos"
            no_data_elements = driver.find_elements(By.ID, "emptyTraceAlert")
            if no_data_elements and no_data_elements[0].is_displayed():
                print("⚠️ No hay datos para esta consulta")
                time.sleep(1)
                break

            # Verificamos si aparece la tabla y es visible
            tabla_elements = driver.find_elements(By.XPATH, "/html/body/lily-ui-app-layout/div/div/div[2]/main/div/div/div[2]/div[2]/div/div[3]/div/table")
            if tabla_elements and tabla_elements[0].is_displayed():
                tabla = tabla_elements[0]
                   # Obtener encabezados (thead)
                   # Obtener encabezados (thead)
                encabezados = [th.text.split("\n")[0] for th in tabla.find_elements(By.XPATH, ".//thead/tr/th")]

                datos = []
                for fila in tabla.find_elements(By.XPATH, ".//tbody/tr"):
                    columnas = fila.find_elements(By.TAG_NAME, "td")
                    fila_dict = {encabezados[i]: columnas[i].text for i in range(len(columnas))}
                    datos.append(fila_dict)
                new_data = []
                estado_actual = ""
                for i, fila in enumerate(datos):
                    time_str = fila["desperdiciado"]
                    if i == 0:
                        new_data=[]
                        estado_actual = "Offline"
                    

                    if fila["Estado"] == "Offline" and estado_actual == "Available":
                        salida = get_time(time_str)
                        print(salida)
                        new_data.append(salida)
                        new_data.append("Central Standard Time")
                        estado_actual = "Offline"

                    elif estado_actual == "Offline" and fila["Estado"] != "Offline":
                        # Guardar la fila anterior si estaba completa
                        if len(new_data) >= 3:
                            agregar_a_excel(rutasalida, new_data)
                        # Empezar nueva fila
                        time_data = get_time(time_str)
                        new_data = [fecha, oracle_id, time_data]
                        estado_actual = "Available"
                    
                    # Guardar si ya tenemos 5 elementos
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
                correo = fila["AWS Name"]    
                oracle_id = fila["Oracle or Kronos ID"]
                consultar_web(driver, correo, fecha, opcion, oracle_id,wait,ruta_salida)

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
