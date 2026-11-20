def consultar_web(driver, correo, fecha, zona, oracle_id, wait,rutasalida,esprimero):
    """
    Usa el navegador existente para consultar los datos
    """
    print(f"Consultando datos para {correo}, {fecha}, {zona}, {oracle_id}")
    try:

     
        caja = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div[1]"))
            )
        #tiempo para que cargue las opciones puede cambiarse dependiendo la conexion
        caja.click()
        time.sleep(2)
        # agregar el correo de usuario a el input 
        agregar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div/div/div/div/div/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div/div/div[1]/div[3]/div[2]/div/div/div[1]/input"))
        )
        agregar.send_keys(correo)
        # caja.send_keys(correo)
        #tiempo para que cargue las opciones puede cambiarse dependiendo la conexion
        time.sleep(2)
        opciones = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul[id^='option-list'] > li"))
        )
        #dar click al primer elemento que sale en la lista si esta el usuario deberia ser el primero 
        #tiempo para que cargue las opciones puede cambiarse dependiendo la conexion
        opciones[0].click()
        time.sleep(2)
        print("que viva el guarapo la lechona y el tamal hp")
        boton_buscar = wait.until(
            EC.element_to_be_clickable((By.XPATH,'/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div[1]/div/div/div/div/div/div[2]/button'))
        )
        boton_buscar.click()
        # --- Checar si hay datos ---
        # modificar para evitar errores en recargos de tablas 
        # 
        timeout = 10  # segundos
        intervalo = 3  # cada cuánto revisar
        tiempo_transcurrido = 0 
         
        while tiempo_transcurrido < timeout:

            # Verificamos si aparece la tabla y es visible
            tabla_elements = driver.find_elements(By.XPATH, "/html/body/lily-ui-app-layout/div/div/div/div[2]/main/lily-historical-metrics/div/main/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/table")

            if tabla_elements and tabla_elements[0].is_displayed():
                print("✅ Tabla encontrada, procesando datos...")
                tabla = tabla_elements[0]

                # Obtener encabezados del thead
                encabezados = [th.text.split("\n")[0] for th in tabla.find_elements(By.XPATH, ".//thead/tr/th")]
                # print("Encabezados:", encabezados)
                datos = []
                no_records = tabla.find_elements(
                By.XPATH,
                ".//tbody/tr/td/div/div[contains(@class,'awsui_empty')]"
                )
              
                if no_records:
                    print("⚠️ No hay datos para esta consulta: " ,correo)
                    break
                # Esperar un momento para que la tabla termine de cargar
                time.sleep(2)
                # Obtener todas las filas del tbody
                filas_tabla = tabla.find_elements(By.XPATH, ".//tbody/tr")
                print(f"Filas encontradas: {len(filas_tabla)}")
                if not filas_tabla or len(filas_tabla) == 1:
                    print(f"❌ No existen datos en la tabla, no hay tamal ni lechona para: {correo}")
                    break
                for fila in filas_tabla:
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
