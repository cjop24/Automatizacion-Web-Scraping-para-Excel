import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración extrema para estabilidad en GitHub Actions
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    # Uso de Secrets para seguridad como solicitaste
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 50) # Aumentamos tiempo de espera

    try:
        # --- LOGIN ---
        print("Accediendo a la plataforma...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("document.querySelector('button.mat-flat-button').click();")
        time.sleep(15) # Tiempo de carga inicial

        # --- EXCEL ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        
        # Asegurar columna DG (índice 110)
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111):
                df[f"Col_{i}"] = ""
        df.columns.values[110] = col_name

        # --- EXTRACCIÓN ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            url = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Abriendo: {pqr_nurc}")
            driver.get(url)
            
            try:
                # 1. Espera forzada al componente principal que vimos en tu foto
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-follow")))
                
                # 2. Scroll para asegurar que Angular 'despierte' el componente
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
                time.sleep(15) # Espera larga para que la base de datos responda

                # 3. SCRIPT JS DE EXTRACCIÓN TOTAL (Busca la tabla o cualquier texto en app-follow)
                js_extract = """
                let container = document.querySelector('app-list-seguimientos');
                if (!container) container = document.querySelector('app-follow');
                
                if (container) {
                    let rows = container.querySelectorAll('tr');
                    let data = [];
                    rows.forEach(r => {
                        if (r.innerText.trim().length > 10 && !r.innerText.includes('Fecha')) {
                            data.push(r.innerText.replace(/\\t/g, ' | ').trim());
                        }
                    });
                    return data.length > 0 ? data.join('\\n---\\n') : container.innerText;
                }
                return "CONTENEDOR_NO_HALLADO";
                """
                
                texto = driver.execute_script(js_extract)
                df.at[index, col_name] = texto
                print(f"-> EXITO: {pqr_nurc}")

            except Exception:
                # Si falla, tomamos TODO lo que haya en la pantalla como último recurso
                print(f"-> FALLO SELECTOR: Intentando captura de emergencia en {pqr_nurc}")
                try:
                    emergencia = driver.execute_script("return document.body.innerText;")
                    df.at[index, col_name] = "CAPTURA_EMERGENCIA: " + emergencia[:500]
                except:
                    df.at[index, col_name] = "ERROR_TOTAL_DE_CARGA"
                
                driver.save_screenshot(f"error_{pqr_nurc}.png")
            
            time.sleep(3)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Fin del proceso.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
