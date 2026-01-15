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

        # --- CARGA CRÍTICA: Forzamos STRING para evitar redondeos ---
        file_input = "Reclamos.xlsx"
        # Leemos TODO como string para que las columnas F, J, L, M, W, Y, Z, BH, CN no se alteren
        df = pd.read_excel(file_input, engine='openpyxl', dtype=str) 
        
        col_name = "Seguimiento_Extraido"
        
        # Asegurar columna DG (110)
        if len(df.columns) <= 110:
            while len(df.columns) <= 110:
                df[f"Col_Extra_{len(df.columns)}"] = ""
        df.columns.values[110] = col_name

        # Limpiar la columna de seguimiento para identificar pendientes
        df[col_name] = df[col_name].fillna("").strip() if hasattr(df[col_name], 'str') else df[col_name].fillna("")
        indices_pendientes = df.index[df[col_name] == ""].tolist()
        
        total_a_procesar = min(len(indices_pendientes), LIMITE_BATCH)
        logging.info(f"Registros pendientes: {len(indices_pendientes)}")
        
        contador = 0
        for idx in indices_pendientes:
            if contador >= LIMITE_BATCH: break
            
            # Obtener NURC (Columna F -> Índice 5) de forma segura
            pqr_nurc = str(df.iloc[idx, 5]).strip().split('.')[0]
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == "":
                continue
            
            logging.info(f"[{contador+1}/{total_a_procesar}] Procesando NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            resultado = "Sin registros"
            try:
                # Espera de renderizado
                time.sleep(6) 
                script_js = """
                let table = document.querySelector('app-list-seguimientos table');
                if (!table) return null;
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
            except:
                resultado = "Error en extracción"

            df.at[idx, col_name] = resultado
            contador += 1
            
            # Guardado preventivo
            if contador % 25 == 0:
                df.to_excel("Reclamos.xlsx", index=False)

        # Guardado final preservando el formato texto
        df.to_excel("Reclamos.xlsx", index=False)
        df.to_excel("Reclamos_scraping.xlsx", index=False)
        logging.info(f"✅ Proceso terminado. Total: {contador}")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
