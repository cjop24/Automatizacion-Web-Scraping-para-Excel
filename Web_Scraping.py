import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración optimizada
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
        # --- LOGIN ---
        print("Iniciando sesión...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.execute_script("document.querySelectorAll('button').forEach(b => { if(b.innerText.includes('INGRESAR')) b.click(); });")
        time.sleep(12)

        # --- PREPARACIÓN EXCEL ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111): df[f"Col_Aux_{i}"] = ""
        df.columns.values[110] = col_name

        # --- BUCLE DE EXTRACCIÓN ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
            
            print(f"Procesando NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # 1. Esperar a que el componente padre esté en el HTML
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-list-seguimientos")))
                
                # 2. Bajar el scroll para activar la carga de la tabla
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                # 3. ESPERA ACTIVA: Reintentar la extracción hasta que la tabla tenga filas reales
                resultado = "Sin datos"
                for intento in range(15): # Reintenta durante 15 segundos
                    script_js = """
                    let filas = document.querySelectorAll('app-list-seguimientos table tbody tr');
                    let logs = [];
                    filas.forEach(f => {
                        let c = f.querySelectorAll('td');
                        if(c.length >= 4 && !f.innerText.includes('No hay datos') && f.innerText.trim() !== "") {
                            logs.push(`[${c[0].innerText.trim()}] ${c[2].innerText.trim()}: ${c[3].innerText.trim()}`);
                        }
                    });
                    return logs.join('\\n---\\n');
                    """
                    resultado = driver.execute_script(script_js)
                    if resultado and len(resultado) > 10:
                        break # Si encontró datos, sale del bucle de reintentos
                    time.sleep(1) # Si no, espera un segundo y vuelve a intentar

                df.at[index, col_name] = resultado if resultado else "Tabla vacía tras espera"
                print(f"-> Resultado: {'Capturado' if len(resultado)>10 else 'Vacío'}")

            except Exception as e:
                print(f"-> Error en carga: {pqr_nurc}")
                driver.save_screenshot(f"error_{pqr_nurc}.png")
                df.at[index, col_name] = "Error de localización de tabla"
            
            time.sleep(2)

        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Finalizado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
