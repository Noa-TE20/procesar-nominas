# Gestor de Nóminas PDF (Renombrado y OCR)

Esta herramienta automatizada está diseñada para la gestión de archivos PDF de nóminas. El script permite extraer de forma inteligente el nombre del trabajador, el mes y el año para renombrar los archivos localmente o directamente en **SharePoint**. Además, incluye un módulo de **OCR** para hacer que los PDFs escaneados sean legibles y permitan la búsqueda de texto.

---

## 🚀 Funcionalidades Principales

* **Renombrado Inteligente:** Utiliza `pdfplumber` y `Tesseract OCR` para localizar el nombre del empleado, el mes y el año dentro del documento.
* **Integración con SharePoint:** Conexión con sitios remotos de SharePoint mediante Client ID y Client Secret para renombrar archivos en la nube.
* **OCR por Lotes:** Convierte PDFs que son solo imágenes (no buscables) en documentos de alta calidad con capa de texto usando `OCRmyPDF`.
* **Gestión de Duplicados:** Añade contadores automáticamente (ej. `(2)`) si existen múltiples nóminas para la misma persona en el mismo periodo.

---

## 🛠 Requisitos Previos

Antes de ejecutar el script, es necesario tener instaladas las siguientes dependencias del sistema:

1.  **Tesseract OCR:** [Descargar aquí](https://github.com/UB-Mannheim/tesseract/wiki) (Necesario para la extracción de texto desde imágenes).
2.  **Poppler:** (Necesario para `pdf2image` para convertir páginas de PDF a imágenes).
3.  **OCRmyPDF & Unpaper:** (Necesario para el módulo de procesamiento OCR).

### Librerías de Python
Instala todas las dependencias de Python usando el archivo `requirements.txt` incluido:
```bash
pip install -r requirements.txt
```

---

## 📖 Uso

Ejecuta el script principal y sigue el menú interactivo en la terminal:

```bash
python procesar_nominas.py
```

### Opciones del Menú:
1.  **Renombrar Nóminas (Local):** Procesa los archivos dentro de una carpeta en tu ordenador.
2.  **Renombrar Nóminas (SharePoint):** Requiere la URL del sitio, Client ID y Client Secret.
3.  **Procesamiento OCR:** Selecciona una carpeta para aplicar OCR a todos los PDFs y hacerlos buscables.

---

## 📂 Estructura del Proyecto

| Archivo | Descripción |
| :--- | :--- |
| `procesar_nominas.py` | Script principal con la lógica de renombrado y OCR. |
| `requirements.txt` | Lista de librerías de Python necesarias. |
| `README.md` | Documentación del proyecto (este archivo). |
| `.gitignore` | Configuración para excluir carpetas de prueba y entornos virtuales de Git. |

---

## ⚠️ Nota sobre Configuración
El script utiliza patrones de expresiones regulares (Regex) específicos para encontrar los nombres:
* **Patrón:** `Trabajador[:.]?\s*([A-Z...])`
* **Meses:** Configurado para castellano (`ENERO`, `FEBRERO`, etc.).

---

### Seguridad
El script utiliza `getpass.getpass()` para la entrada del Client Secret de SharePoint, asegurando que las credenciales no se muestren en pantalla ni se guarden accidentalmente en el historial de la terminal.
