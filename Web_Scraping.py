import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
    wait = WebDriverWait(driver, 45)

    try:
        # --- PASO 1: Login ---
        print("Accediendo a SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        # Click en INGRESAR
        boton_js = "var btn = document.querySelector('button.mat-flat-button'); if(btn) btn.click();"
        driver.execute_script(boton_js)
        time.sleep(10)

        # --- PASO 2: Lectura de Excel (Solución Redondeo) ---
        file_path = "Reclamos.xlsx"
        # Forzamos lectura como string para evitar el redondeo de los NURC
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        # Asegurar columna DG (índice 110)
        while df.shape[1] <= 110:
            df[f"Columna_Seguimiento_{df.shape[1]}"] = ""
        
        target_col = 110

        # --- PASO 3: Bucle de Extracción ---
        for index, row in df.iterrows():
            # Limpieza de NURC para evitar notación científica y decimales
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Procesando NURC: {pqr_nurc}")
            driver.get(url_reclamo)
            
            try:
                # NUEVO SELECTOR: Basado en tu hallazgo del HTML
                # Esperamos específicamente al componente <app-follow>
                selector_follow = "app-follow"
                wait.until(EC.visibility_of_element_located((By.TAG_NAME, selector_follow)))
                
                # Pausa extra para que Angular termine de llenar la tabla dentro del componente
                time.sleep(7)
                
                # Intentamos extraer del componente app-follow
                elemento_seguimiento = driver.find_element(By.TAG_NAME, selector_follow)
                texto_extraido = elemento_seguimiento.text.strip()

                # Si app-follow está vacío, intentamos con el selector jerárquico que encontraste
                if len(texto_extraido) < 10:
                    selector_css = "#contenido > div > div:nth-child(1) > div > mat-card:nth-child(3) > app-follow"
                    texto_extraido = driver.find_element(By.CSS_SELECTOR, selector_css).text.strip()

                df.iat[index, target_col] = texto_extraido
                print(f"-> EXITO: {pqr_nurc} (Datos capturados)")
                
            except Exception:
                print(f"-> AVISO: No se encontro contenido en app-follow para {pqr_nurc}")
                driver.save_screenshot(f"debug_{pqr_nurc}.png")
                df.iat[index, target_col] = "Sin informacion encontrada en el componente de seguimiento"
            
            time.sleep(2)

        # Guardar archivo final
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Proceso finalizado con exito.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
