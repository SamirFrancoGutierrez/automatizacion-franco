from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
import time

# Crear opciones del navegador Chrome
chrome_options = webdriver.ChromeOptions()

# Inicializar el controlador de Chrome con las opciones
driver = webdriver.Chrome(options=chrome_options)

# Leer el archivo Excel
df = pd.read_excel("C:/Users/U22244871/Desktop/automatizacion-franco/transportes-data-pendientes.xlsx")
print(df)

chofer_dict = {
    'BRITO CELESTINO ARAUJO': '//*[@id="i5"]/div[3]/div',
    'MARTIN SANCHEZ JARAMILLO': '//*[@id="i8"]/div[3]/div',
    'MARCO ANTON CORONADO': '//*[@id="i11"]/div[3]/div',
    'GIAMPEAR PURIZACA MARTINEZ': '//*[@id="i14"]/div[3]/div'
    
}

plantel_dict = {
    '110': '//*[@id="i37"]/div[3]/div',
    '111': '//*[@id="i40"]/div[3]/div',
    '112': '//*[@id="i43"]/div[3]/div',
    '113': '//*[@id="i46"]/div[3]/div',
    '114': '//*[@id="i49"]/div[3]/div'
}

for row, datos in df.iterrows():
    Chofer = datos['Chofer']
    GuiaFranco = datos['GuiaFranco']
    GuiaChimu = datos['GuiaChimu']
    PuntoPartida = datos['PuntoPartida']
    PuntoDestino = datos['PuntoDestino']
    Servicio = datos['Servicio']
    Plantel = datos['Plantel']
    Cantidad= datos['Cantidad']
    # Abrir formulario
    driver.get("https://forms.gle/P9DYSy8eeiM9dbWk9")
    time.sleep(2)
    # Chofer
    time.sleep(2)
    driver.find_element(By.XPATH, chofer_dict[Chofer]).click()
    time.sleep(1)
    # G.R.R TRANSPORTES FRANCO
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
    input_element.send_keys(GuiaFranco)
    # G.R.R CHIMU
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input')
    input_element.send_keys(GuiaChimu)
    # Servicio
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input')
    input_element.send_keys(Servicio)
    # Punto de Partida
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input')
    input_element.send_keys(PuntoPartida) 
     # Punto de Destino
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div/div[1]/input')
    input_element.send_keys(PuntoDestino) 
    #PL
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div/div[1]/input' )
    input_element.send_keys(Plantel)
    #CANTIDAD M3
    input_element = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[8]/div/div/div[2]/div/div[1]/div/div[1]/input' )
    input_element.send_keys(Cantidad)
    time.sleep(2)
     #Enviar Formulario
    driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span' ).click()
# Cerrar la página después de que todas las iteraciones se completen

driver(quit)