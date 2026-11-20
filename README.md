# Automatización Python 🤖

Herramienta de automatización desarrollada en Python para facilitar tareas repetitivas con Selenium, Pandas y GUI tkinter. Este proyecto permite automatizar flujos de trabajo web, procesamiento de datos y generación de reportes.

## 📋 Descripción

Este proyecto es una colección de scripts de automatización que utiliza:
- **Selenium**: Para automatización de navegador web y scraping
- **Pandas**: Para procesamiento y manipulación de datos en Excel
- **tkinter**: Para crear interfaces gráficas de usuario (GUI)
- **python-docx**: Para generación de documentos Word

## 🎯 Funcionalidades Principales

- Automatización de tareas web mediante Selenium (WebDriver Firefox)
- Lectura y escritura de archivos Excel (.xlsx)
- Interfaz gráfica interactiva con tkinter
- Selector de fechas con calendario integrado
- Procesamiento de datos con Pandas
- Generación de reportes automatizados

## 📦 Requisitos Previos

- Python 3.8 o superior
- Firefox instalado (para Selenium WebDriver)
- pip (gestor de paquetes de Python)
- Virtualenv (opcional pero recomendado)

## 🚀 Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/wiston10/automatizacion_connect.git
cd Automatizacion-py
```

### 2. Crear un entorno virtual (recomendado)
```bash
# En macOS/Linux
python3 -m venv .venv
source .venv/bin/activate

# En Windows
python -m venv .venv
.venv\Scripts\activate
```

### 3. Instalar las dependencias
```bash
pip install -r requirements.txt
```

## 📚 Dependencias

```
python-docx==1.1.0      # Generación de documentos Word
pandas==2.2.2            # Procesamiento de datos
selenium==4.23.1         # Automatización web
tkcalendar==1.6.1        # Widget de calendario para tkinter
```

## 📂 Estructura del Proyecto

```
Automatizacion-py/
├── README.md                 # Este archivo
├── requirements.txt          # Dependencias del proyecto
├── .gitignore               # Archivo de exclusiones de git
│
├── cool.py                  # Script principal de automatización
├── actualizado.py           # Versión actualizada
├── index.py                 # Script de índice/utilidades
├── index1.py                # Variante del índice
├── ttec.py                  # Script específico para TTEC
├── scripts.py               # Colección de scripts auxiliares
├── formateador.py           # Formateador de datos
├── formateador1.py          # Variante del formateador
│
├── output/                  # Archivos de salida generados
├── formateados/             # Datos procesados/formateados
│
└── [archivos de datos]      # Archivos Excel y documentos de trabajo
    ├── BD_Sara.xlsx
    ├── BD_WINSTON.xlsx
    ├── usuarios.xlsx
    └── ...
```

## 🔧 Uso

### Ejecutar el script principal
```bash
python cool.py
```

### Ejecutar otros scripts
```bash
python actualizado.py
python ttec.py
python scripts.py
```

## 🌐 Configuración

Algunos scripts requieren configuración previa:

- **Firefox Profile Path**: Actualiza la ruta del perfil de Firefox en el script
- **URLs destino**: Modifica las URLs según tus necesidades
- **Rutas de archivos**: Ajusta las rutas de entrada/salida de datos

## 📊 Ejemplos de Uso

### Procesar datos Excel
```python
import pandas as pd

# Leer archivo Excel
df = pd.read_excel('usuarios.xlsx')

# Procesar datos
df_procesado = df.drop_duplicates()

# Guardar resultado
df_procesado.to_excel('usuarios_procesado.xlsx', index=False)
```

### Automatizar navegación web
```python
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Firefox()
driver.get('https://ejemplo.com')
# ... más acciones de automatización
driver.quit()
```

## ⚠️ Notas Importantes

- Los scripts pueden requerir credenciales para acceder a sitios específicos
- Algunos procesos pueden tomar tiempo dependiendo del volumen de datos
- Se recomienda ejecutar primero en entorno de prueba
- Revisa los logs y archivos de salida para verificar el correcto funcionamiento

## 🐛 Troubleshooting

### "WebDriver not found"
Instala geckodriver o asegúrate que Firefox está instalado correctamente

### "Module not found"
Verifica que el entorno virtual está activado y las dependencias están instaladas:
```bash
pip install -r requirements.txt
```

### Problemas con rutas de archivos
Usa rutas absolutas o relativas consistentes en tus scripts

## 📝 Notas de Desarrollo

- Mantén los scripts modulares y reutilizables
- Comenta el código para facilitar el entendimiento
- Crea branches para nuevas funcionalidades
- Realiza commits con mensajes descriptivos

## 👨‍💻 Autor

Wiston Navarro Aroca

## 📄 Licencia

Este proyecto está disponible bajo la licencia MIT.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:
1. Crea un fork del proyecto
2. Crea una rama para tu feature (`git checkout -b feature/NuevaFuncionalidad`)
3. Realiza commit de tus cambios (`git commit -am 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/NuevaFuncionalidad`)
5. Abre un Pull Request

## 📞 Contacto

Para preguntas o sugerencias, abre un issue en el repositorio de GitHub.
