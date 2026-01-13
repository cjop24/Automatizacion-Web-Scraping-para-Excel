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
    # Recuperar credenciales de Secrets
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

        # Click en INGRESAR mediante JS para evitar bloqueos de renderizado
        boton_js = "var btn = document.querySelector('button.mat-flat-button'); if(btn) btn.click();"
        driver.execute_script(boton_js)
        time.sleep(10) # Tiempo para que cargue el sistema

        # --- PASO 2: Lectura de Excel (EVITAR REDONDEO) ---
        file_path = "Reclamos.xlsx"
        # Leemos todo como string para que no altere los NURC largos
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)

        # Asegurar que la columna DG (índice 110) exista
        while df.shape[1] <= 110:
            df[f"Columna_Seguimiento_{df.shape[1]}"] = ""
        
        target_col = 110

        # --- PASO 3: Bucle de Extracción ---
        for index, row in df.iterrows():
            # Limpieza profunda del NURC (columna 6 / índice 5)
            # Eliminamos espacios, posibles .0 de conversiones previas y lo tratamos como texto puro
            pqr_nurc = str(row.iloc[5]).strip().split('.')[0]
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                print(f"Llegamos al final de los datos en fila {index + 1}")
                break
                
            url_reclamo = f"https://pqrdsuperargo.supersalud.gov.co/gestion/supervisar/{pqr_nurc}"
            print(f"Extrayendo NURC: {pqr_nurc}")
            
            driver.get(url_reclamo)
            
            try:
                # Esperamos a que la tabla de seguimiento (main_table) aparezca
                # Basado en Fotografía 2: #main_table_wrapper es el ID correcto
                wait.until(EC.presence_of_element_located((By.ID, "main_table_wrapper")))
                
                # Scroll y pausa para que Angular pueble la tabla
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(6) 
                
                # Capturamos el texto de la tabla
                seguimiento_elem = driver.find_element(By.ID, "main_table_wrapper")
                texto_final = seguimiento_elem.text.strip()

                # Si el contenedor principal trae poco texto, intentamos con el tag table directamente
                if len(texto_final) < 20:
                    tablas = driver.find_elements(By.TAG_NAME, "table")
                    if tablas:
                        texto_final = tablas[-1].text.strip() # Usualmente la última tabla es la de seguimiento

                df.iat[index, target_col] = texto_final
                print(f"-> EXITO: {pqr_nurc}")
                
            except Exception as e:
                print(f"-> AVISO: No se capturaron datos para {pqr_nurc}")
                # Guardamos captura para verificar qué falló visualmente
                driver.save_screenshot(f"debug_{pqr_nurc}.png")
                df.iat[index, target_col] = "Sin datos (Revisar captura de pantalla)"
            
            # Pausa para no saturar el servidor
            time.sleep(3)

        # Guardar archivo final
        df.to_excel("Reclamos_scraping.xlsx", index=False, engine='openpyxl')
        print("Proceso completado. Archivo generado: Reclamos_scraping.xlsx")

    except Exception as e:
        print(f"Error general: {e}")
        driver.save_screenshot("error_fatal.png")
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    run_scraper()
