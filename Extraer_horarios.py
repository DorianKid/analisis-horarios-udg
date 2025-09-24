from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import win32com.client as win32
from datetime import date
from constantes import (CLAVES_CARRERAS, CENTROS, CICLOS)
import os, json, time, re

# -----------------------------
# Configuración general
# -----------------------------
SIIAU_URL = "http://consulta.siiau.udg.mx/wco/sspseca.forma_consulta"
DIAS_ORDEN = ['L', 'M', 'I', 'J', 'V', 'S']  # Lunes a Sábado
DEFAULT_WAIT = 5
AÑO = date.today().year

# -----------------------------
# Funciones Selenium
# -----------------------------
def get_chromedriver_service():
    """
    Pide al usuario la RUTA COMPLETA al chromedriver y valida que:
    - Sea un archivo existente
    - Se llame exactamente 'chromedriver.exe' (Windows)
    """
    while True:
        ruta = input('Ruta completa a "chromedriver.exe": ').strip()

        # quitar comillas si las pegas
        if (ruta.startswith('"') and ruta.endswith('"')) or (ruta.startswith("'") and ruta.endswith("'")):
            ruta = ruta[1:-1]

        p = Path(os.path.expanduser(os.path.expandvars(ruta))).resolve()

        if not p.is_file():
            print("La ruta no apunta a un archivo existente. Intenta nuevamente.")
            continue

        if p.name.lower() != "chromedriver.exe":
            print(f'El archivo debe llamarse exactamente "chromedriver.exe" (recibido: "{p.name}").')
            continue

        return Service(executable_path=str(p))

def make_driver(headless: bool = True, window_size: str = "1366,768") -> webdriver.Chrome:
    """
    Crea un webdriver.Chrome con opciones sensatas para scraping.
    - headless=True para ejecutar sin ventana (puedes poner False para depurar).
    - window_size para evitar layouts móviles.
    """
    service = get_chromedriver_service()
    opts = Options()
    if headless:
        # 'new' evita warnings con Chrome moderno
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument(f"--window-size={window_size}")
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(30)
    return driver

def open_siiau(driver: webdriver.Chrome, url: str = SIIAU_URL, wait_seconds: int = 3) -> WebDriverWait:
    """
    Abre el formulario de consulta de SIIAU y devuelve un WebDriverWait listo.
    """
    driver.get(url)
    time.sleep(2)  # pequeño respiro para assets iniciales
    return WebDriverWait(driver, wait_seconds)

def select_filters(driver, wait, ciclo: str, centro: str, clave_carrera: str) -> None:
    """Elige ciclo, centro, escribe clave de carrera, sube el límite y consulta."""
    Select(driver.find_element(By.ID, "cicloID")).select_by_value(str(ciclo))
    Select(driver.find_element(By.NAME, "cup")).select_by_value(str(centro))

    # Clave de carrera (majrp)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[name="majrp"]'))).send_keys(clave_carrera)

    # Subir la cantidad de horarios
    cantidad = driver.find_element(By.XPATH, "/html/body/font/font/form/table/tbody/tr[9]/td[1]/input")
    driver.execute_script("arguments[0].setAttribute('value','1000');", cantidad) # Obtener toda la tabla de horarios

    # Consultar
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='idConsultar']"))).click()

# -----------------------------
# Helpers
# -----------------------------
def AutoFit_columns_width(ruta_archivo):
    # Crear una instancia de la aplicación de Excel
    xlApp = win32.Dispatch('Excel.Application')

    # Abrir el libro de Excel
    wb = xlApp.Workbooks.Open(ruta_archivo)

    # Iterar sobre todas las hojas del libro
    for ws in wb.Worksheets:
        # Ajustar automáticamente el ancho de todas las columnas en la hoja actual
        ws.Columns.AutoFit()

    # Guardar los cambios y cerrar el libro
    wb.Save()
    wb.Close()

    # Cerrar la aplicación de Excel
    xlApp.Quit()

def create_dir(path: str) -> str:
    """
    Crea (si no existe) y devuelve la ruta absoluta de un directorio.
    """
    abs_path = os.path.abspath(path)
    os.makedirs(abs_path, exist_ok=True)
    return abs_path

def validate_input(prompt: str, valids: list):
    """
    Pide un valor al usuario y valida que esté en la lista válida.
    """

    value = input(prompt).strip().upper()
        
    while value not in valids:
        print("Valor inválido. Intenta nuevamente.")
        value = input(prompt).strip().upper()

    return value

def save_excel(df: pd.DataFrame, out_xlsx: str) -> None:
    base_cols = ['Clave','Hora inicio','Hora fin','Día','Salón','Materia','Profesor','Alumnos']
    cols = [c for c in base_cols if (df is not None and c in getattr(df,'columns',[]))]
    if df is None or df.empty:
        df = pd.DataFrame(columns=cols)

    out_dir = Path(out_xlsx).parent
    out_dir.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        for dia in DIAS_ORDEN:
            dfd = df[df['Día'] == dia] if 'Día' in df.columns else pd.DataFrame(columns=cols)
            if dfd.empty:
                pd.DataFrame(columns=cols).to_excel(writer, sheet_name=dia, index=False)
            else:
                dfd[cols].to_excel(writer, sheet_name=dia, index=False)
            
def save_txt(txt_path, content):
    """
    Escribe (sobrescribe) un .txt en la ruta dada.
    Acepta str, list, dict (los serializa a JSON).
    """
    if not isinstance(content, str):
        content = json.dumps(content, ensure_ascii=False)
    # OJO: aquí NO creamos carpetas; asegúrate de que el directorio exista antes.
    with open(txt_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)
        if content and not content.endswith("\n"):
            f.write("\n")
    return txt_path

def load_datos_from_txt(txt_path: str):
    with open(txt_path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    if isinstance(payload, dict) and "data" in payload:
        datos = payload["data"]
    else:
        datos = payload
    if not isinstance(datos, list) or len(datos) < 7:
        raise ValueError("El TXT no contiene la estructura esperada de 'datos'.")
    return datos  # lista: [CLAVE, MATERIA, SECCION, CUP, DIS, HORARIO, PROFESOR]

# -----------------------------
# Extracción de datos
# -----------------------------
def parse_table_raw(driver):
    """
    Lee la tabla y devuelve los datos
    """
    time.sleep(2)
    filas = driver.find_elements(By.XPATH, '/html/body/table[1]/tbody/tr')
    cols  = driver.find_elements(By.XPATH, './/th')
    n_f, n_c = len(filas), len(cols)
    if n_f == 0 or n_c == 0:
        raise RuntimeError("No se detectó tabla de horarios.")

    datos = []
    for i in range(1, n_c):
        col_vals = []
        for j in range(3, n_f + 1):
            celda = driver.find_element(By.XPATH, f'/html/body/table[1]/tbody/tr[{j}]/td[{i}]').text
            col_vals.append(celda)
        datos.append(col_vals)

    # El layout típico requiere quitar 2 columnas de ruido
    del datos[0]
    del datos[3]

    return datos

# -----------------------------
# Limpieza de datos
# -----------------------------

def transform_datos(datos) -> pd.DataFrame:
    # Desempaquetar columnas
    CLAVE, MATERIA, SECCION, CUP, DIS, HORARIO, PROFESOR = [datos[i] for i in range(7)]

    # Limpieza de profesor (como el original)
    PROFESOR = [str(p).replace("01 ", "").strip() for p in PROFESOR]

    # 1) Buscar rango de fechas en cualquier elemento de HORARIO
    patron = re.compile(
        r'(?P<desde>(?:0[1-9]|[12]\d|3[01])/(?:0[1-9]|1[0-2])/\d{2})\s*-\s*'
        r'(?P<hasta>(?:0[1-9]|[12]\d|3[01])/(?:0[1-9]|1[0-2])/\d{2})'
    )
    fecha_rango = None
    for h in HORARIO:
        m = patron.search(str(h))
        if m:
            fecha_rango = f"{m.group('desde')} - {m.group('hasta')}"
            break  # con uno basta

    # 2) Limpieza de HORARIO (como el original) + quitar rango si se encontró
    HORARIO_LIMPIO = []
    for h in HORARIO:
        s = str(h)
        if fecha_rango:
            s = s.replace(fecha_rango, "")
        s = (
            s.replace("CS", "")
             .replace("-", " ")
             .replace("LFS0", "LFS-")
             .replace("001 ", "1")
             .replace("01 ", "")
             .replace(".", "")
             .replace(" A00", "-")
             .replace("DED", "")
             .replace(" A0", "-")
             .replace(" A", "-")
             .replace("V LC0", "V-")
             .replace("DUCT1 LC0", "T1-")
             .strip()
        )
        HORARIO_LIMPIO.append(s)

    # 3) ALUMNOS = CUP - DIS (misma tolerancia a errores del original)
    ALUMNOS = []
    for cup, dis in zip(CUP, DIS):
        try:
            alumnos = int(str(cup).strip()) - int(str(dis).strip())
        except Exception:
            alumnos = ""
        ALUMNOS.append(alumnos)

    # 4) Expandir por día (4/5/6 tokens), respetando multilínea
    datos_org = []
    for clave, mat, alumnos, horario, profesor in zip(CLAVE, MATERIA, ALUMNOS, HORARIO_LIMPIO, PROFESOR):
        if not horario:
            continue
        lineas = horario.split("\n") if "\n" in horario else [horario]
        for linea in lineas:
            toks = linea.split()
            if len(toks) == 4:
                hi, hf, d1, salon = toks
                dias = [d1]
            elif len(toks) == 5:
                hi, hf, d1, d2, salon = toks
                dias = [d1, d2]
            elif len(toks) >= 6:
                hi, hf, d1, d2, d3, salon = toks[:6]
                dias = [d1, d2, d3]
            else:
                continue

            for d in dias:
                dd = str(d).strip().upper()
                if dd in DIAS_ORDEN:
                    # Nota: aquí generamos el mismo orden de campos que usabas originalmente
                    datos_org.append([clave, mat, alumnos, hi, hf, dd, salon, profesor])

    # 5) DataFrame final como el original
    columns = ['Clave', 'Materia', 'Alumnos', 'Hora inicio', 'Hora fin', 'Día', 'Salón', 'Profesor']
    df = pd.DataFrame(datos_org, columns=columns)

    # Filtro DES (mantener filas donde Salón NO contiene DES)
    if not df.empty:
        df = df[~df['Salón'].astype(str).str.contains('DES', na=False)]

    # Reordenar columnas y ordenar por día/hora (igual que el original)
    df = df[['Clave', 'Hora inicio', 'Hora fin', 'Día', 'Salón', 'Materia', 'Profesor', 'Alumnos']]
    if not df.empty:
        df['Día'] = pd.Categorical(df['Día'], categories=DIAS_ORDEN, ordered=True)
        df = df.sort_values(['Día', 'Hora inicio'], ignore_index=True)

    return df

#Preguntar al usuario las claves, ciclo y centro a buscar carrreras.
clave_carrera = validate_input('\nEscribe la clave de tu carrera: ', CLAVES_CARRERAS)
ciclo = validate_input('\nEscribe el ciclo Ej: 2025B -> "202520" (Ciclo A -> "10" o Ciclo B -> "20"): ', CICLOS)

print('\nLista de centros: ')
for key, value in CENTROS.items():
    print(value, ' -> ', key)
    time.sleep(.1)

centro = validate_input('\nEscribe la clave del centro de esa carrera Ej: CUCEI -> "D": ', list(CENTROS.keys()))

# Preguntar donde crear directorio, nombre archivo final
ruta_guardado = input('\nIngresa la ruta absoluta donde se guardará el archivo excel: ').strip()

# -----------------------------
# Scrapping
# -----------------------------
scrape = validate_input('¿Necesitas scrapear o ya tienes un txt con la inormación cruda? (S/N)\n', ["S", "N"])

filename = "_".join(["HORARIOS", clave_carrera, CENTROS[centro], ciclo])
ruta_excel = ruta_guardado + "\\" + filename + ".xlsx"
ruta_txt = ruta_guardado + "\\" + filename + ".txt"

# --- EXPORTACIÓN ---
base_dir = Path(os.path.expanduser(ruta_guardado)).resolve()
base_dir.mkdir(parents=True, exist_ok=True)

filename_base = "_".join(["HORARIOS", clave_carrera, CENTROS[centro], ciclo])
ruta_excel = base_dir / f"{filename_base}.xlsx"
ruta_txt   = base_dir / f"{filename_base}.txt"

if scrape == "S":
    driver = make_driver(headless=False)
    wait   = open_siiau(driver)

    try:
        select_filters(driver, wait, ciclo=ciclo, centro=centro, clave_carrera=clave_carrera)
        raw_data = parse_table_raw(driver)  # o pasa "DD/MM/AA - DD/MM/AA" si lo quieres limpiar
    finally:
        driver.quit()

    # Serializa y guarda RAW en TXT (usa meta útil + datos crudos)
    payload_raw = {
        "ciclo": ciclo,
        "centro": centro,
        "carrera": clave_carrera,
        "data": raw_data  # <- 'datos' que regresaste desde parse_table_raw
    }
    save_txt(str(ruta_txt), payload_raw)
    print("TXT:  ", os.path.abspath(str(ruta_txt)))

else:
    raw_data = load_datos_from_txt(str(ruta_txt))

df = transform_datos(raw_data)

# Guarda EXCEL
save_excel(df, str(ruta_excel))
print("Excel:", os.path.abspath(str(ruta_excel)))

# (Opcional) AutoFit con COM de Excel si estás en Windows con Office
try:
    AutoFit_columns_width(ruta_excel)
except Exception:
    pass
