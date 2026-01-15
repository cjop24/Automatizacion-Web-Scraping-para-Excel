import os
import pandas as pd
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración de Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    # No cargar imágenes para máxima velocidad
    options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    return webdriver.Chrome(options=options)

def run_scraper():
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    # LÍMITE DE PROCESAMIENTO POR SESIÓN
    LIMITE_BATCH = 1000 
    
    driver = get_driver()
    wait = WebDriverWait(driver, 20)

    try:
        # --- LOGIN ---
        logging.info("Iniciando sesión...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        wait.until(EC.visibility_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR')).click();")
        wait.until(EC.url_contains("/inicio"))
        logging.info("✅ Login exitoso.")

        # --- CARGA DE EXCEL ---
        file_input = "Reclamos.xlsx"
        df = pd.read_excel(file_input, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111): df[f"C_{i}"] = ""
        df.columns.values[110] = col_name

        # --- FILTRADO: Solo procesar lo que esté vacío ---
        # Esto permite retomar el trabajo si se corta o si procesamos por lotes
        pendientes = df[df[col_name].isna() | (df[col_name] == "")]
        total_a_procesar = min(len(pendientes), LIMITE_BATCH)
        
        logging.info(f"Registros pendientes totales: {len(pendientes)}")
        logging.info(f"Procesando lote de: {total_a_procesar} registros")

        contador = 0
        for index, row in pendientes.iterrows():
            if contador >= LIMITE_BATCH: break
            
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': continue
            
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # Extracción atómica por JS (Paciencia de 8 seg)
                script_js = """
                return (function() {
                    let table = document.querySelector('app-list-seguimientos table');
                    if (!table) return null;
                    let rows = table.querySelectorAll('tbody tr');
                    if (rows.length > 0 && !rows[0].innerText.includes('No hay datos')) {
                        return Array.from(rows).map(r => {
                            let c = r.querySelectorAll('td');
                            let d = c[3].querySelector('div') ? c[3].querySelector('div').innerText : c[3].innerText;
                            return `[${c[0].innerText.trim()}]: ${d.trim()}`;
                        }).join('\\n---\\n');
                    }
                    return "SIN_REGISTROS";
                })();
                """
                
                resultado = "Sin registros"
                for _ in range(4): # 4 reintentos de 2 segundos cada uno
                    res = driver.execute_script(script_js)
                    if res and res != "SIN_REGISTROS":
                        resultado = res
                        break
                    time.sleep(2)

                df.at[index, col_name] = resultado
                contador += 1
                if contador % 50 == 0:
                    logging.info(f"Progreso: {contador}/{total_a_procesar}")

            except Exception:
                df.at[index, col_name] = "Error de carga"

        # Guardar el mismo archivo para "recordar" el progreso
        df.to_excel("Reclamos.xlsx", index=False)
        # También guardar una copia con el nombre de salida para GitHub Artifacts
        df.to_excel("Reclamos_scraping.xlsx", index=False)
        logging.info(f"✅ Lote completado. {contador} registros procesados.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
