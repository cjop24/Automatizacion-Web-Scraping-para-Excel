import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Configuración de Chrome optimizada
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 40) # Aumentamos espera a 40 segundos

    try:
        # --- PASO 1: Login ---
        print("Abriendo página de login...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        print("Intentando click en INGRESAR...")
        boton_js = """
        var botones = document.querySelectorAll('button');
        for (var i = 0; i < botones.length; i++) {
            if (botones[i].textContent.includes('INGRESAR')) {
                botones[i].click();
                return true;
            }
        }
        return false;
        """
        driver.execute_script(boton_js)
        time.sleep(10)

        # --- PASO 2: Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl')

        # SOLUCIÓN AL INDEX ERROR: 
        # Si el archivo no tiene 111 columnas, creamos la columna DG (índice 110)
        while df.shape[1] <= 110:
            df[f"Nueva_Columna_{df.shape[1]}"] = ""
        
        # Le ponemos nombre a la columna de destino si es necesario
        target_col = 110 # Columna DG

        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip()
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Procesando NURC: {pqr_nurc}")
            
            driver.get(url_reclamo)
            
            try:
                # PASO 2.3: Extraer Seguimiento
                # Usamos un selector más genérico por si ID falla
                seguimiento_elem = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#main_table_wrapper, .dataTables_wrapper")))
                df.iat[index, target_col] = seguimiento_elem.text
                print(f"Datos extraídos para {pqr_nurc}")
                
            except TimeoutException:
                print(f"Aviso: No se encontró tabla para {pqr_nurc}")
                df.iat[index, target_col] = "Sin datos de seguimiento o error de carga"
            
            time.sleep(3)

        # Guardar archivo final
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Proceso finalizado con éxito.")

    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        driver.save_screenshot("debug_error.png")
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
