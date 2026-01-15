import os
import pandas as pd
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    return webdriver.Chrome(options=options)

def run_scraper():
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    LIMITE_BATCH = 1000 
    
    driver = get_driver()
    wait = WebDriverWait(driver, 20)

    try:
        logging.info("Iniciando sesión...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        wait.until(EC.visibility_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR')).click();")
        wait.until(EC.url_contains("/inicio"))
        logging.info("✅ Login exitoso.")

        # CARGA DE EXCEL
        file_input = "Reclamos.xlsx"
        df = pd.read_excel(file_input, engine='openpyxl') # Quitamos dtype=str aquí para manejarlo manual
        col_name = "Seguimiento_Extraido"
        
        # Asegurar columna DG (110)
        if len(df.columns) <= 110:
            while len(df.columns) <= 110:
                df[f"Col_Extra_{len(df.columns)}"] = ""
        df.columns.values[110] = col_name

        # Convertir columna de NURC (índice 5) a string limpio sin .0
        def clean_nurc(val):
            if pd.isna(val): return ""
            s = str(val).strip()
            if s.endswith('.0'): s = s[:-2]
            return s

        # Identificar pendientes
        df[col_name] = df[col_name].fillna("")
        mask_pendientes = (df[col_name] == "")
        indices_pendientes = df.index[mask_pendientes].tolist()
        
        total_a_procesar = min(len(indices_pendientes), LIMITE_BATCH)
        logging.info(f"Registros pendientes totales: {len(indices_pendientes)}")
        
        contador = 0
        for idx in indices_pendientes:
            if contador >= LIMITE_BATCH: break
            
            # Obtener NURC de la columna 6 (índice 5)
            pqr_nurc = clean_nurc(df.iloc[idx, 5])
            
            if not pqr_nurc or pqr_nurc == "":
                logging.warning(f"Fila {idx}: NURC vacío, saltando.")
                continue
            
            logging.info(f"Procesando [{contador+1}/{total_a_procesar}] - NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            resultado = "Sin registros"
            try:
                # Espera inteligente corta
                time.sleep(5) 
                script_js = """
                let table = document.querySelector('app-list-seguimientos table');
                if (!table) return "TABLA_NO_ENCONTRADA";
                let rows = table.querySelectorAll('tbody tr');
                if (rows.length > 0 && !rows[0].innerText.includes('No hay datos')) {
                    return Array.from(rows).map(r => {
                        let c = r.querySelectorAll('td');
                        let d = c[3] ? (c[3].querySelector('div') ? c[3].querySelector('div').innerText : c[3].innerText) : "";
                        return `[${c[0].innerText.trim()}]: ${d.trim()}`;
                    }).join('\\n-----\\n');
                }
                return "SIN_SEGUIMIENTO";
                """
                res = driver.execute_script(script_js)
                if res: resultado = res
            except Exception as e:
                resultado = f"Error: {str(e)[:30]}"

            df.at[idx, col_name] = resultado
            contador += 1
            
            # Guardado preventivo cada 20 registros por si falla la conexión
            if contador % 20 == 0:
                df.to_excel("Reclamos.xlsx", index=False)

        # Guardado final
        df.to_excel("Reclamos.xlsx", index=False)
        df.to_excel("Reclamos_scraping.xlsx", index=False)
        logging.info(f"✅ Proceso terminado. Procesados: {contador}")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
