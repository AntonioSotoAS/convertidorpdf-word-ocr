# Convertidor de PDF a Word con OCR (Tesseract)

Aplicación de escritorio en Python con interfaz gráfica (Tkinter) para convertir documentos PDF (escaneados o basados en texto) a archivos editables de Microsoft Word (`.docx`) manteniendo la estructura y aplicando reconocimiento óptico de caracteres (OCR) mediante **Tesseract OCR** y **OpenCV**.

---

## 📋 Características

* **OCR Integrado**: Convierte PDFs escaneados o imágenes a texto editable en Word.
* **Detección Automática de Tesseract**: Localiza la instalación de Tesseract en las rutas habituales de Windows.
* **Interfaz Gráfica Fácil de Usar**: Construida con Tkinter, incluye barra de progreso y opción de cancelación.
* **Limpieza y Formateo de Texto**: Detección básica de títulos, eliminación de caracteres nulos y saltos de línea innecesarios.

---

## 🛠️ Requisitos Previos

1. **Python 3.8 o superior** instalado en el sistema.
2. **Tesseract OCR Engine** instalado en Windows (instrucciones detalladas abajo).

---

## ⚙️ Paso a Paso: Instalación y Configuración de Tesseract OCR

Para que la conversión por OCR funcione correctamente, **Tesseract OCR debe estar instalado** en tu equipo y accesible.

### 1. Descargar e Instalar Tesseract

* Puedes usar el instalador ejecutable que viene incluido en la raíz de este proyecto:  
  `tesseract-ocr-w64-setup-5.5.3.20260724.exe`
* O descargar la última versión para Windows desde el repositorio oficial de UB-Mannheim:  
  [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)

> ⚠️ **¡IMPORTANTE DURANTE LA INSTALACIÓN!**  
> Al ejecutar el instalador, en la ventana donde se seleccionan los componentes (**Choose Components**):
> 1. Despliega la opción **Additional language data (download)**.
> 2. Marca la casilla de **Spanish (`spa`)** para instalar los datos de idioma en español (o el idioma que necesites). Sin esto, el OCR en español fallará o dará errores.

---

### 2. Verificar la Ruta de Instalación (`tesseract.exe`)

Por defecto en Windows, Tesseract suele instalarse en una de las siguientes ubicaciones:

* **Para todos los usuarios:**  
  `C:\Program Files\Tesseract-OCR\tesseract.exe`
* **Para el usuario actual:**  
  `C:\Users\<TuUsuario>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe`
* **Instalación mediante PDF24 u otras herramientas:**  
  `C:\Program Files\PDF24\tesseract\tesseract.exe`

Asegúrate de abrir el Explorador de archivos de Windows y **comprobar en qué carpeta exacta quedó guardado `tesseract.exe`**.

---

### 3. Agregar Tesseract al PATH de Windows (Recomendado)

Agregar Tesseract a las Variables de Entorno permite que cualquier script o terminal reconozca el comando `tesseract` globalmente:

1. Presiona `Win + R`, escribe `sysdm.cpl` y pulsa **Enter**.
2. Ve a la pestaña **Opciones avanzadas** y haz clic en **Variables de entorno...**.
3. En la sección **Variables del sistema** (o variables de usuario), selecciona la variable `Path` y haz clic en **Editar...**.
4. Haz clic en **Nuevo** y añade la ruta de la carpeta donde instalaste Tesseract (por ejemplo: `C:\Program Files\Tesseract-OCR`).
5. Haz clic en **Aceptar** en todas las ventanas.
6. **Verificación**: Abre un nuevo `cmd` o `PowerShell` y escribe:
   ```cmd
   tesseract --version
   ```
   Si la consola muestra la versión instalada, ¡el PATH está configurado correctamente!

---

### 4. Cómo Verifica y Configura la Ruta `main.py`

El archivo `main.py` incluye una función llamada `configurar_tesseract()` que busca automáticamente la ruta de `tesseract.exe` en este orden:

1. Verifica si está registrado en el **PATH del sistema** (`shutil.which("tesseract")`).
2. Revisa las rutas de instalación estándar de Windows:
   * `%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe`
   * `%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe`
   * `C:\Program Files\Tesseract-OCR\tesseract.exe`
   * `C:\Program Files\PDF24\tesseract\tesseract.exe`
   * `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`

#### 🔴 ¿Qué hacer si instalaste Tesseract en una ruta distinta?
Si instalaste Tesseract en una carpeta personalizada (por ejemplo `D:\Herramientas\Tesseract-OCR\tesseract.exe`), abre [main.py](file:///c:/Users/amontejo/Desktop/depdfaword/main.py#L43-L57) y agrega tu ruta en la lista `posibles_rutas`:

```python
posibles_rutas = [
    r"D:\TuRutaPersonalizada\Tesseract-OCR\tesseract.exe",  # <-- Agrega tu ruta aquí
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    # ...
]
```

---

## 🚀 Instalación y Ejecución del Proyecto

### 1. Clonar o descargar el repositorio
```bash
git clone https://github.com/AntonioSotoAS/convertidorpdf-word-ocr.git
cd convertidorpdf-word-ocr
```

### 2. Crear y activar entorno virtual
```bash
# En Windows (PowerShell / CMD)
python -m venv venv

# Activar en PowerShell:
.\venv\Scripts\Activate.ps1

# O activar en CMD:
.\venv\Scripts\activate.bat
```

### 3. Instalar dependencias de Python
```bash
pip install -r requirements.txt
```

Las librerías requeridas en `requirements.txt` son:
* `PyMuPDF` (fitz) - Extracción e inspección rápida de PDFs.
* `pytesseract` - Wrapper de Python para Tesseract OCR.
* `opencv-python` & `numpy` - Procesamiento y mejora de imágenes antes del OCR.
* `Pillow` - Manipulación de imágenes.
* `python-docx` - Creación y edición de archivos Word (`.docx`).

### 4. Ejecutar la Aplicación
```bash
python main.py
```

---

## 📂 Estructura del Proyecto

```text
convertidorpdf-word-ocr/
│
├── main.py                                      # Código principal de la aplicación GUI y lógica de conversión
├── requirements.txt                             # Librerías de Python requeridas
├── README.md                                    # Documentación del proyecto
└── tesseract-ocr-w64-setup-5.5.3.20260724.exe  # Instalador de Tesseract OCR para Windows (opcional/incluido)
```

---

## ❓ Solución de Problemas Frecuentes

* **Error `pytesseract.TesseractNotFoundError`**:
  * Tesseract no está instalado o `main.py` no encuentra la ruta ejecutable. Revisa el **Paso 2 y 4** para asegurarte de que `tesseract.exe` exista en una de las rutas escaneadas o agrégalo al PATH del sistema.

* **Error de idioma `TesseractError: (1, 'Error opening data file ... spa.traineddata')`**:
  * No seleccionaste el paquete de idioma en español durante la instalación de Tesseract. Vuelve a ejecutar el instalador `tesseract-ocr-w64-setup-5.5.3.20260724.exe`, selecciona "Modify" o reinstala asegurándote de marcar **Spanish (`spa`)** dentro de *Additional language data*.

---

## 📄 Créditos y Licencia

Desarrollado por **Antonio Soto Developer**.

Este proyecto ha sido creado por Antonio Soto Developer. Todos los derechos sobre la autoría y el código pertenecen a su desarrollador. Siéntete libre de utilizarlo, personalizarlo y adaptarlo a tus necesidades.

