import os
import re
import getpass
import tempfile
import subprocess
from pathlib import Path

# External Libraries
import pytesseract
from pdf2image import convert_from_path
import pdfplumber
import pikepdf

# SharePoint Imports
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# ==========================================
# CONFIGURATION & PATTERNS (RENAME MODULE)
# ==========================================

MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

NAME_PATTERN = r"Trabajador[:.]?\s*([A-ZÁÉÍÓÚÑñ\s]+?)(?=\s+(?:NIF|DNI|NIE|NUM|N\.I\.F|No|Num|Afili|$)|\n|$)"
YEAR_PATTERN = r"\b(20[2-3]\d)\b"

# ==========================================
# CORE EXTRACTION LOGIC (RENAME MODULE)
# ==========================================

def extract_date_info(text):
    found_month, found_year = "MES_DESCONOCIDO", "AÑO_DESCONOCIDO"
    text_up = text.upper()
    
    for mes in MESES:
        if mes in text_up:
            found_month = mes
            break
            
    year_match = re.search(YEAR_PATTERN, text)
    if year_match:
        found_year = year_match.group(1)
        
    return found_month, found_year

def format_worker_name(raw_name):
    clean = re.sub(r'[^A-ZÁÉÍÓÚÑñ\s]', '', raw_name, flags=re.IGNORECASE)
    parts = " ".join(clean.split()).split()
    
    while parts and len(parts[-1]) <= 1:
        parts.pop()
        
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1]} {' '.join(parts[2:])}".upper()
    return " ".join(parts).upper()

def extract_content(pdf_path, dpi=300):
    images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=dpi)
    if images:
        return pytesseract.image_to_string(images[0].convert('L'), lang='spa')
    return ""

def analyze_pdf_for_rename_data(pdf_path):
    """Returns (month, year, worker_name) or None if not found."""
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    if not re.search(NAME_PATTERN, text, re.IGNORECASE):
        text = extract_content(pdf_path, dpi=300)

    if not re.search(NAME_PATTERN, text, re.IGNORECASE):
        text = extract_content(pdf_path, dpi=600)

    name_match = re.search(NAME_PATTERN, text, re.IGNORECASE)
    if name_match:
        worker_name = format_worker_name(name_match.group(1))
        month, year = extract_date_info(text)
        return month, year, worker_name
    return None

# ==========================================
# PROCESSING LOGIC (RENAME MODULE)
# ==========================================

def process_local_rename(folder_path):
    folder = Path(folder_path)
    if not folder.exists():
        print(f"  [ERROR] La carpeta local {folder_path} no existe.")
        return

    payroll_counts = {}
    for pdf_file in folder.glob('*.pdf'):
        if "NOMINA" in pdf_file.name.upper():
            continue

        print(f"Procesando: {pdf_file.name}...")
        result = analyze_pdf_for_rename_data(pdf_file)
        
        if result:
            month, year, worker_name = result
            key = (month, year, worker_name)
            count = payroll_counts.get(key, 0) + 1
            payroll_counts[key] = count
            
            nomina_label = "NOMINA" if count == 1 else f"NOMINA {count}"
            new_filename = f"{nomina_label} {month} {year} {worker_name}.pdf"
            
            new_path = folder / new_filename
            final_counter = 2
            while new_path.exists():
                new_path = folder / f"{nomina_label} {month} {year} {worker_name} ({final_counter}).pdf"
                final_counter += 1
            
            pdf_file.rename(new_path)
            print(f"  [EXITO] -> {new_path.name}")
        else:
            print(f"  [AVISO] No se encontró nombre en {pdf_file.name}")

def process_sharepoint_rename(site_url, client_id, client_secret, target_folder):
    try:
        ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
        folder = ctx.web.get_folder_by_server_relative_url(target_folder)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()
    except Exception as e:
        print(f"  [ERROR] No se pudo conectar a SharePoint: {e}")
        return

    payroll_counts = {}
    
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        for sp_file in files:
            if "NOMINA" in sp_file.name.upper() or not sp_file.name.lower().endswith(".pdf"):
                continue
                
            print(f"Procesando (SharePoint): {sp_file.name}...")
            local_temp_file = temp_path / sp_file.name
            
            with open(local_temp_file, "wb") as local_file:
                sp_file.download(local_file)
                ctx.execute_query()

            result = analyze_pdf_for_rename_data(local_temp_file)
            
            if result:
                month, year, worker_name = result
                key = (month, year, worker_name)
                count = payroll_counts.get(key, 0) + 1
                payroll_counts[key] = count
                
                nomina_label = "NOMINA" if count == 1 else f"NOMINA {count}"
                new_filename = f"{nomina_label} {month} {year} {worker_name}.pdf"
                new_relative_url = f"{target_folder.rstrip('/')}/{new_filename}"
                
                try:
                    sp_file.moveto(new_relative_url, 1)
                    ctx.execute_query()
                    print(f"  [EXITO] Renombrado en SharePoint -> {new_filename}")
                except Exception as e:
                    print(f"  [ERROR] Falló al renombrar en SharePoint: {e}")
            else:
                print(f"  [AVISO] No se encontró nombre en {sp_file.name}")

# ==========================================
# OCR & SEARCHABLE PDF MODULE
# ==========================================

def is_already_ocrd(pdf_path):
    """Detect if the PDF was already processed by OCRmyPDF."""
    try:
        with pikepdf.Pdf.open(pdf_path) as pdf:
            return "OCRmyPDF" in str(pdf.docinfo)
    except Exception:
        return False

def process_single_ocr(input_path):
    print(f"\n--- Procesando OCR: {input_path.name} ---")

    if is_already_ocrd(input_path):
        print("  [i] Archivo ya procesado previamente. Omitiendo.")
        return

    temp_final = input_path.with_suffix(".ocr.tmp.pdf")

    cmd = [
        "ocrmypdf",
        "--rotate-pages",
        "--rotate-pages-threshold", "1",
        "--deskew",
        "--clean", # NOTE: Requires 'unpaper' on Windows. Remove if it causes Exit Status 3.
        "--force-ocr",
        "--language", "spa",
        str(input_path),
        str(temp_final)
    ]

    try:
        subprocess.run(cmd, check=True, capture_output=True)
        temp_final.replace(input_path)
        print(f"  [✓] Finalizado con éxito: {input_path.name}")

    except subprocess.CalledProcessError as e:
        print(f"  [X] Error de OCR en {input_path.name}")
        print(f"      Stderr: {e.stderr.decode() if e.stderr else 'Error desconocido'}")

    finally:
        if temp_final.exists():
            temp_final.unlink()

def process_batch_ocr(folder_path):
    folder = Path(folder_path)
    if not folder.exists():
        print(f"  [ERROR] El directorio {folder} no fue encontrado.")
        return

    pdf_files = list(folder.glob("*.pdf"))
    if not pdf_files:
        print("  [AVISO] No se encontraron archivos PDF en la carpeta.")
        return

    for pdf in pdf_files:
        process_single_ocr(pdf)

# ==========================================
# INTERACTIVE CLI MENU
# ==========================================

def main():
    while True:
        print("\n" + "=" * 50)
        print("  HERRAMIENTAS DE GESTIÓN DE DOCUMENTOS (PDF)  ")
        print("=" * 50)
        print("¿Qué tarea deseas realizar?")
        print(" 1. Renombrar Nóminas (Extracción de Trabajador/Mes)")
        print(" 2. Hacer PDFs buscables y legibles (Procesamiento OCR)")
        print(" 3. Salir")
        
        tarea = input("\nElige una opción (1, 2 o 3): ").strip()
        
        if tarea == "3":
            print("Saliendo del programa...")
            break
            
        elif tarea == "1":
            print("\n--- RENOMBRAR NÓMINAS ---")
            print(" 1. Carpeta Local")
            print(" 2. SharePoint Remoto")
            opcion_renombre = input("Elige origen (1 o 2): ").strip()
            
            if opcion_renombre == "1":
                folder_path = input("Introduce la ruta de la carpeta local (ej. ./nominas): ").strip()
                print("\n[CONFIRMACIÓN] Modo: Renombrar | Origen: Local | Ruta:", folder_path)
                if input("¿Continuar? (s/n): ").strip().lower() in ['s', 'y']:
                    print("\nIniciando proceso local...\n")
                    process_local_rename(folder_path)
                    
            elif opcion_renombre == "2":
                site_url = input("URL de SharePoint (ej. https://dominio.sharepoint.com/sites/RRHH): ").strip()
                target_folder = input("Ruta de la carpeta (ej. /sites/RRHH/Documentos/Nominas): ").strip()
                client_id = input("Client ID: ").strip()
                client_secret = getpass.getpass("Client Secret (oculto): ").strip()
                
                print("\n[CONFIRMACIÓN] Modo: Renombrar | Origen: SharePoint Remoto")
                print(f"Sitio: {site_url} \nCarpeta: {target_folder}")
                if input("¿Continuar? (s/n): ").strip().lower() in ['s', 'y']:
                    print("\nIniciando proceso en SharePoint...\n")
                    process_sharepoint_rename(site_url, client_id, client_secret, target_folder)
            else:
                print("Opción no válida.")

        elif tarea == "2":
            print("\n--- PROCESAMIENTO OCR (PDFs Buscables) ---")
            folder_path = input("Introduce la ruta de la carpeta local con los PDFs (ej. ./lotes): ").strip()
            
            print("\n[CONFIRMACIÓN] Modo: OCR Batch | Ruta:", folder_path)
            if input("¿Continuar? (s/n): ").strip().lower() in ['s', 'y']:
                print("\nIniciando procesamiento OCR...\n")
                process_batch_ocr(folder_path)
                
        else:
            print("Opción no válida. Por favor, elige 1, 2 o 3.")

if __name__ == "__main__":
    main()