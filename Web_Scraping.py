import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuración optimizada para evitar bloqueos en GitHub Actions
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

def run_scraper():
    # Uso de Secrets configurados en GitHub para proteger tus credenciales
    user = os.getenv("PQRD_USER")
    password = os.getenv("PQRD_PASS")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 60)

    try:
        # --- PASO 1: LOGIN ---
        print("Iniciando sesión en SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        # Esperar y llenar campos
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)
        
        # Click robusto por JavaScript para evitar que el botón sea 'tapado' por otros elementos
        driver.execute_script("""
            let btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('INGRESAR'));
            if(btn) btn.click();
        """)
        time.sleep(15) # Tiempo de carga para el Dashboard principal

        # --- PASO 2: PREPARACIÓN DEL EXCEL ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        col_name = "Seguimiento_Extraido"
        
        # Solución al IndexError: Asegurar que exista la columna DG (índice 110)
        if len(df.columns) <= 110:
            for i in range(len(df.columns), 111):
                df[f"Col_Temp_{i}"] = ""
        df.columns.values[110] = col_name

        # --- PASO 3: EXTRACCIÓN CON NOVEDAD DE OUTERHTML ---
        for index, row in df.iterrows():
            # Extraer NURC (Columna 6 del Excel original)
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            if not pqr_nurc or pqr_nurc == 'nan': break
                
            print(f"Procesando NURC: {pqr_nurc}")
            driver.get(f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}")
            
            try:
                # 1. Esperar al componente de seguimientos identificado en tu foto
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-list-seguimientos")))
                
                # 2. Bajar scroll a media página para forzar renderizado de Angular
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
                
                # 3. SCRIPT JS DE EXTRACCIÓN QUIRÚRGICA
                # Extrae de cada fila la Fecha y el texto dentro del DIV de Descripción
                script_js = """
                let filas = document.querySelectorAll('app-list-seguimientos table tbody tr');
                let logs = [];
                
                filas.forEach(fila => {
                    let celdas = fila.querySelectorAll('td');
                    if (celdas.length >= 4) {
                        let fecha = celdas[0].innerText.trim();
                        // Buscamos el div con scroll que identificamos en el OuterHTML
                        let divDesc = celdas[3].querySelector('div');
                        let descripcion = divDesc ? divDesc.innerText.trim() : celdas[3].innerText.trim();
                        
                        if (descripcion && !descripcion.includes('No hay datos')) {
                            logs.push(`[${fecha}]: ${descripcion}`);
                        }
                    }
                });
                return logs.join('\\n---\\n');
                """
                
                # Espera activa (Retry) de hasta 15 segundos para que la tabla se llene
                resultado_final = ""
                for intento in range(15):
                    resultado_final = driver.execute_script(script_js)
                    if resultado_final: break
                    time.sleep(1)
                
                df.at[index, col_name] = resultado_final if resultado_final else "Sin seguimiento disponible"
                print(f"-> EXITO: {pqr_nurc} capturado.")

            except Exception as e:
                print(f"-> AVISO: No se cargó tabla para {pqr_nurc}")
                df.at[index, col_name] = "Componente no localizado"
                driver.save_screenshot(f"error_{pqr_nurc}.png")
            
            time.sleep(2) # Pausa entre cada NURC para no saturar el servidor

        # --- PASO 4: GUARDADO ---
        df.to_excel("Reclamos_scraping.xlsx", index=False)
        print("Proceso completado. Archivo Reclamos_scraping.xlsx generado.")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
