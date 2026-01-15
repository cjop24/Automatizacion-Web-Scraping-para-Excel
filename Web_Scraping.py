import os
import pandas as pd
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- Configuración de Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# --- Configuración de Chrome (Basada en tu script funcional) ---
options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

def run_scraper():
    # Credenciales desde Secrets
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    def safe_click(locator, timeout=25):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
        el.click()
        return el

    try:
        logging.info("--- Iniciando Sesión (Lógica Verificada) ---")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # Proceso de Login exacto al que te funciona
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        
        safe_click((By.XPATH, "//button[contains(., 'INGRESAR')]"))
        
        # Espera crítica para confirmar entrada
        wait.until(EC.url_contains("/inicio"))
        logging.info("✅ Login exitoso detectado.")

        # --- Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        
        # Asegurar columna DG (índice 110)
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111):
                df[f"Col_Temp_{i}"] = ""
        df.columns.values[110] = col_name

        # --- Bucle de Extracción ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': continue
                
            logging.info(f"Procesando NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # Esperar al componente visual identificado en el OuterHTML
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-list-seguimientos")))
                time.sleep(12) # Tiempo para que Angular cargue los datos internos
                
                # Extracción quirúrgica de la tabla y los divs de descripción
                script_js = """
                let filas = document.querySelectorAll('app-list-seguimientos table tbody tr');
                let logs = [];
                filas.forEach(f => {
                    let c = f.querySelectorAll('td');
                    if (c.length >= 4 && !f.innerText.includes('No hay datos')) {
                        let fecha = c[0].innerText.trim();
                        let divDesc = c[3].querySelector('div');
                        let textoDesc = divDesc ? divDesc.innerText.trim() : c[3].innerText.trim();
                        logs.push(`[${fecha}]: ${textoDesc}`);
                    }
                });
                return logs.join('\\n---\\n');
                """
                
                resultado = driver.execute_script(script_js)
                df.at[index, col_name] = resultado if resultado else "Sin seguimientos"
                logging.info(f"-> Capturado con éxito para {pqr_nurc}")

            except Exception as e:
                logging.warning(f"-> No se encontró tabla en {pqr_nurc}")
                df.at[index, col_name] = "Error: Tabla no cargó"
                driver.save_screenshot(f"error_{pqr_nurc}.png")
            
            time.sleep(2)

        # Guardado final
        df.to_excel("Reclamos_scraping.xlsx", index=False)
        logging.info("--- Proceso Finalizado con Éxito ---")

    except Exception as e:
        logging.error(f"Falla crítica: {e}")
        driver.save_screenshot("DEBUG_FINAL.png")
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
