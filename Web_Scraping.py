import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración de Chrome optimizada para GitHub Actions
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
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("document.querySelector('button.mat-flat-button').click();")
        time.sleep(10)

        # --- PASO 2: Lectura de Excel ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        # Asegurar columna DG (índice 110)
        while df.shape[1] <= 110:
            df[f"Columna_Seguimiento_{df.shape[1]}"] = ""
        
        target_col = 110

        # --- PASO 3: Bucle de Extracción ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Abriendo NURC: {pqr_nurc}")
            driver.get(url_reclamo)
            
            try:
                # SELECTOR PROFUNDO PROPORCIONADO
                selector_css = "#contenido > div > div:nth-child(1) > div > mat-card:nth-child(3) > app-follow > mat-card-content"
                
                # Esperar a que el contenedor esté presente
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector_css)))
                
                # Pausa para renderizado de Angular (importante para que el texto aparezca)
                time.sleep(8)
                
                # Intentamos extraer el texto directamente
                elemento = driver.find_element(By.CSS_SELECTOR, selector_css)
                texto_extraido = elemento.text.strip()

                # Si el texto directo falla, buscamos tablas o divs internos (capas más profundas)
                if len(texto_extraido) < 5:
                    print(f"Buscando capas más profundas para {pqr_nurc}...")
                    # Buscamos cualquier tabla o contenido de texto dentro del mat-card-content
                    inner_content = driver.execute_script(f"return document.querySelector('{selector_css}').innerText;")
                    texto_extraido = inner_content.strip() if inner_content else ""

                df.iat[index, target_col] = texto_extraido
                
                if len(texto_extraido) > 10:
                    print(f"-> ÉXITO: {pqr_nurc} ({len(texto_extraido)} caracteres extraídos)")
                else:
                    print(f"-> AVISO: {pqr_nurc} extraído pero parece vacío.")

            except Exception as e:
                print(f"-> ERROR en NURC {pqr_nurc}: {e}")
                driver.save_screenshot(f"debug_{pqr_nurc}.png")
                df.iat[index, target_col] = "Error de localización de elemento"
            
            time.sleep(2)

        # Guardar resultados
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Archivo Reclamos_scraping.xlsx generado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
