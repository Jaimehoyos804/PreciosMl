from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
import time
import xlwings as xw
import tkinter as tk
from PIL import Image, ImageTk
from threading import Thread
import tkinter.ttk as ttk
import os
import re
def abrir_navegador():
    ruta_driver = os.path.join(os.getcwd(), "Recursos", "msedgedriver.exe")
    service = Service(executable_path=ruta_driver)
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Edge(service=service, options=options)
    return driver

#Abrir archivo
archivo_excel = "Insumo.xlsm"
app = xw.App(visible=False)
libro = app.books.open(archivo_excel, update_links=False, read_only=False)
hoja = libro.sheets["DATOS"]


# Leer productos desde columna A
productos = []
ultima_fila = hoja.range("A" + str(hoja.cells.last_cell.row)).end("up").row

for fila in range(2, ultima_fila + 1):
    producto = hoja.range(f"A{fila}").value
    if producto:
        productos.append((fila, producto))

# Variables globales para mantener el navegador abierto entre funciones
driver_exito = None
driver_Alix = None
driver = abrir_navegador()
driver.get("https://www.mercadolibre.com.co/")
time.sleep(3)
def buscar_en_mercado_libre():
    
    for fila, nombres in productos:
        try:
            box = driver.find_element(By.NAME, "as_word")
            box.clear()
            box.send_keys(nombres)
            box.send_keys(Keys.ENTER)
            time.sleep(5)
            driver.execute_script("window.scrollBy(0, 250)")
            elemento = driver.find_element(By.XPATH, '//*[@id="root-app"]/div/div[2]/section/div[5]/ol/li[1]/div/div/div/div[2]/h3/a')
            elemento.click()
            time.sleep(5)
            nombre = driver.find_element(By.CLASS_NAME, "ui-pdp-title").text.strip()
            precio = driver.find_element(By.CLASS_NAME, "andes-money-amount__fraction").text.strip()
            PrecioLimpio = precio.replace("$", "").replace(".", "")
            PrecioInt = int(PrecioLimpio)
            hoja.range(f"B{fila}").value = nombre
            hoja.range(f"C{fila}").value = PrecioInt
        except Exception as e:
            hoja.range(f"B{fila}").value = "No encontrado"
            print(f"[ML] {nombres}: error -> {e}")
    driver.quit()
    global driver_exito
    driver_exito = abrir_navegador()
    driver_exito.get("https://www.exito.com/")
    print("🟡 Éxito abierto, esperando clic en botón EXITO...")

def buscar_exito():
    global driver_exito
    if not driver_exito:
        print("❌ El navegador de Éxito no está abierto.")
        return
    ultima_fila = hoja.range("A" + str(hoja.cells.last_cell.row)).end("up").row
    for fila in range(2, ultima_fila + 1):
        nombre_Exito = hoja[f"A{fila}"].value
        if not  nombre_Exito:
            break 
        productos.append((fila, producto))
        try:
            buscador = driver_exito.find_element(By.XPATH, "//*[@id='header-page']/section/div/div[1]/div[2]/form/input")
            buscador.clear()
            buscador.send_keys(nombre_Exito)
            buscador.send_keys(Keys.ENTER)
            time.sleep(10)
            driver_exito.execute_script("window.scrollBy(0, 250)")
            primer_producto = driver_exito.find_element(By.CLASS_NAME, 'productCard_productLinkInfo__It3J2')
            primer_producto.click()
            time.sleep(3)
            nombre = driver_exito.find_element(By.CLASS_NAME, "product-title_product-title__heading___mpLA").text.strip()
            precio = driver_exito.find_element(By.CLASS_NAME, "ProductPrice_container__JKbri").text.strip()
            time.sleep(1)
            precio_numerico = re.findall(r"\d+", precio) # encuentra todos los grupos de números
            if precio_numerico:
                PrecioInt = int("".join(precio_numerico))  # junta los grupos, por ejemplo ['2','299','900'] -> 2299900
            else:
                PrecioInt = 0
            hoja.range(f"D{fila}").value = nombre
            hoja.range(f"E{fila}").value = PrecioInt
        except Exception as e:
            hoja.range(f"D{fila}").value = "No encontrado"
            hoja.range(f"E{fila}").value = "-"
            print(f"[EXITO] {nombre_Exito}: error -> {e}")

    driver_exito.quit()
    print("✅ Éxito: productos guardados.")
    ultima_fila = hoja.range("C" + str(hoja.cells.last_cell.row)).end("up").row
    for fila in range (2, ultima_fila + 1):
        PrecioMl = hoja.range(f"C{fila}").value
        PrecioEx = hoja.range(f"E{fila}").value
        if PrecioMl < PrecioEx:
            hoja.range(f"F{fila}").value = "Mercado Libre"
        elif PrecioEx == PrecioMl:
            hoja.range(f"F{fila}").value = "Iguales"
        else:
            hoja.range(f"F{fila}").value = "Exito" 
    libro.save()
def ejecutar_busqueda_ml():
    buscar_en_mercado_libre()
    libro.save()
    print("🟢 Mercado Libre: productos guardados.")

def ejecutar_busqueda_exito():
    buscar_exito()

def cerrar_excel():
    try:
        libro.save()
        libro.close()
        app.quit()
        print("✅ Excel cerrado correctamente.")
    except Exception as e:
        print(f"⚠️ Error al cerrar Excel: {e}")


def mostrar_boton():
    ventana = tk.Tk()
    ventana.title("BUSCAR PRODUCTOS")
    ventana.configure(bg="#f0f0f0")

    # Imagen
    imagen = Image.open("robot.png")
    imagen = imagen.resize((150, 200), Image.Resampling.LANCZOS)
    imagen_tk = ImageTk.PhotoImage(imagen)
    label_imagen = tk.Label(ventana, image=imagen_tk, bg="#f0f0f0")
    label_imagen.image = imagen_tk
    label_imagen.pack(pady=10)

    # Crear estilo
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("ML.TButton",
        font=("Segoe UI", 12),
        padding=10,
        background="#008000",  # verde oscuro
        foreground="white"
    )
    style.map("ML.TButton",
        background=[("active", "#006400")]
    )

    style.configure("Exito.TButton",
        font=("Segoe UI", 12),
        padding=10,
        background="#003366",  # azul oscuro
        foreground="white"
    )
    style.map("Exito.TButton",
        background=[("active", "#001f4d")]
    )

    # Botón Mercado Libre
    ttk.Button(
        ventana,
        text="🛒 MERCADO LIBRE",
        style="ML.TButton",
        command=ejecutar_busqueda_ml
    ).pack(pady=8, ipadx=10, ipady=5)

    # Botón Éxito
    ttk.Button(
        ventana,
        text="🏬 ÉXITO",
        style="Exito.TButton",
        command=ejecutar_busqueda_exito
    ).pack(pady=8, ipadx=10, ipady=5)   
    
    def al_cerrar():
        cerrar_excel()
        ventana.destroy()
    ventana.protocol("WM_DELETE_WINDOW", al_cerrar)

    ventana.mainloop()
mostrar_boton()
