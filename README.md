# Comparador de precios
# 📊 Comparador de Precios Automático: Mercado Libre vs Éxito

Este proyecto es una herramienta automatizada desarrollada en Python que permite buscar productos en línea y comparar precios entre **Mercado Libre** y **Éxito**. Está diseñado para facilitar la toma de decisiones de compra o análisis de precios en entornos administrativos y de compras.

## 🚀 Características

- Búsqueda automática de productos desde un archivo Excel (`Insumo.xlsm`).
- Scraping en vivo de:
  - [Mercado Libre Colombia](https://www.mercadolibre.com.co)
  - [Éxito](https://www.exito.com/)
- Registro del nombre del producto encontrado y su precio.
- Comparación automática entre los precios y clasificación del más económico.
- Interfaz gráfica con `Tkinter` para facilitar la interacción.
- Resultados almacenados directamente en el archivo Excel.

## 📁 Estructura del Proyecto
/PRACTICA/
│
├── ProyectoV2.2.py # Script principal
├── Insumo.xlsm # Archivo Excel de entrada/salida
├── msedgedriver.exe # WebDriver para Microsoft Edge
├── robot.png # Imagen decorativa para la interfaz


## 🛠️ Requisitos

- Python 3.11+
- Microsoft Edge instalado
- Microsoft Excel (para lectura y escritura del .xlsm)
- Librerías Python:
  - `selenium`
  - `xlwings`
  - `openpyxl`
  - `tkinter`
  - `Pillow` (para imágenes en la GUI)

Instala dependencias con:

```bash
pip install selenium xlwings openpyxl pillow
📓 Cómo usar
Abre el archivo Insumo.xlsm y llena la columna A con los nombres de los productos a buscar.

Ejecuta ProyectoV2.2.py.

Desde la interfaz, selecciona:

🛒 MERCADO LIBRE para iniciar la búsqueda y scraping en ML.

🏬 ÉXITO para consultar los mismos productos en Éxito.

El sistema completará automáticamente las columnas B-E con nombres y precios.

La columna F indicará automáticamente cuál tienda tiene el precio más bajo
⚙️ Consideraciones técnicas
El scraping requiere buena conexión a Internet.

Si Excel está abierto mientras se ejecuta el script, puede bloquear la escritura. Asegúrate de cerrarlo o permitir acceso.

Asegúrate de tener msedgedriver.exe en la ruta esperada o dentro de la carpeta /recursos.

🔐 Seguridad
Este script solo automatiza búsquedas públicas y no requiere credenciales.

No recoge, almacena ni transmite información privada.

No genera tráfico malicioso ni masivo.

👨‍💻 Autor
Jaime Hoyos
Estudiante de Ingeniería de Software

