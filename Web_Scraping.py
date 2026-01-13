import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # Ingreso de credenciales
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        # Login Robusto: Intentar click por texto si el selector de clase falla
        try:
            btn_xpath = "//button[contains(., 'INGRESAR')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath))).click()
            print("Click exitoso.")
        except:
            print("Fallo click normal, intentando via Script...")
            driver.execute_script("document.querySelectorAll('button').forEach(b => { if(b.innerText.includes('INGRESAR')) b.click(); });")
        
        time.sleep(12)

        # Cargar Excel
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        while df.shape[1] <= 110:
            df[f"Col_Extra_{df.shape[1]}"] = ""
        
        target_col = 110

        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Abriendo NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # El selector profundo que proporcionaste
                selector = "#contenido > div > div:nth-child(1) > div > mat-card:nth-child(3) > app-follow > mat-card-content"
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                time.sleep(10)
                
                # Extraer texto usando InnerText para asegurar captura de datos dinámicos
                texto = driver.execute_script(f"return document.querySelector('{selector}').innerText;")
                df.iat[index, target_col] = texto.strip() if texto else "Contenedor vacío"
                print(f"-> Capturado: {pqr_nurc}")
                
            except Exception:
                print(f"-> No se encontró contenido para {pqr_nurc}")
                df.iat[index, target_col] = "Sin seguimiento"
                driver.save_screenshot(f"error_nurc_{pqr_nurc}.png")
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)

    except Exception as e:
        print(f"Error fatal: {e}")
        driver.save_screenshot("debug_error.png")
        # Creamos un archivo vacío para que GitHub Actions no falle al buscar artifacts
        with open("Reclamos_scraping.xlsx", "w") as f: f.write("error")
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
