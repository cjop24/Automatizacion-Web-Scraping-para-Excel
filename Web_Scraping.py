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
    wait = WebDriverWait(driver, 45) # Mayor tolerancia para carga de datos

    try:
        # --- PASO 1: Login ---
        print("Accediendo a SuperArgo...")
        driver.get("https://pqrdsuperargo.supersalud.gov.co/login")
        
        wait.until(EC.presence_of_element_located((By.ID, "user"))).send_keys(user)
        driver.find_element(By.ID, "password").send_keys(password)

        print("Enviando formulario de acceso...")
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
        time.sleep(10) # Espera post-login

        # --- PASO 2: Preparación de Excel ---
        file_path = "Reclamos.xlsx"
        df = pd.read_excel(file_path, engine='openpyxl')

        # Asegurar que la columna DG (índice 110) exista
        while df.shape[1] <= 110:
            df[f"Columna_Aux_{df.shape[1]}"] = ""
        
        target_col = 110 # Columna DG

        # --- PASO 3: Bucle de Extracción ---
        for index, row in df.iterrows():
            pqr_nurc = str(row.iloc[5]).strip()
            
            if not pqr_nurc or pqr_nurc == 'nan' or pqr_nurc == '':
                print(f"Fin de registros detectado en fila
