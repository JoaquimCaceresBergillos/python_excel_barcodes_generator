import time
inicio = time.time()  # guardar tiempo inicial

# ---------- CONFIGURACIÓN ------------------------------------------------------------------------
ARCHIVO_ENTRADA = "articulos.xlsx" # Archivo original
ARCHIVO_SALIDA = "articulos_con_barcodes.xlsx" # Archivo resultante con códigos de barras
CARPETA_IMAGENES = "barcodes" # carpeta para guardar las imágenes de los códigos de barras

COLUMNA_CODIGO = "codigo" # columna con el número original

# AJUSTES CODE128
GENERAR_CODE128 = True
OPTIONS_CODE128 = {
    "module_width": 0.2,    # ancho de cada barra
    "module_height": 8,    # altura de las barras
    "font_size": 8,        # tamaño del número debajo
    "text_distance": 3.5,   # distancia entre barra y texto
    "quiet_zone": 4,        # margen en blanco a los lados
    "background": "white",  # color de fondo
    "foreground": "black",  # color de las barras
    "write_text": True,     # mostrar número debajo
    "dpi": 300              # resolución
}
COLUMNA_CODE128 = "code128" # columna donde se generará el code128
OPTIONS_CODE128_ANCHO_IMAGEN = 150 # Ancho de la imagen del code128 en Excel
OPTIONS_CODE128_ALTO_IMAGEN = 45 # Alto de la imagen del code128 en Excel


# AJUSTES EAN13
GENERAR_EAN13 = True
OPTIONS_EAN13 = {
    "module_width": 0.2,    # ancho de cada barra
    "module_height": 8,    # altura de las barras
    "font_size": 6,        # tamaño del número debajo
    "text_distance": 2.5,   # distancia entre barra y texto
    "quiet_zone": 4,        # margen en blanco a los lados
    "background": "white",  # color de fondo
    "foreground": "black",  # color de las barras
    "write_text": True,     # mostrar número debajo
    "dpi": 300              # resolución
}
COLUMNA_EAN13 = "ean13"    # columna donde se generará el EAN13
OPTIONS_EAN13_ANCHO_IMAGEN = 150  # Ancho de la imagen del ean13 en Excel
OPTIONS_EAN13_ALTO_IMAGEN = 45 # Alto de la imagen del ean13 en Excel

# ----------  FIN CONFIGURACIÓN -------------------------------------------------------------------











import os
import shutil
import pandas as pd
from barcode import Code128, EAN13
from barcode.writer import ImageWriter
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# Borra el contenido de las carpetas si existen
if os.path.exists(CARPETA_IMAGENES):
    shutil.rmtree(CARPETA_IMAGENES)

# Crear carpeta para imágenes si no existe
os.makedirs(f"{CARPETA_IMAGENES}/code128", exist_ok=True)
os.makedirs(f"{CARPETA_IMAGENES}/ean13", exist_ok=True)

# Leer Excel
df = pd.read_excel(ARCHIVO_ENTRADA)

# Crear columnas vacías según flags
if GENERAR_CODE128:
    df[COLUMNA_CODE128] = ""
if GENERAR_EAN13:
    df[COLUMNA_EAN13] = ""

# Listas de rutas
rutas_code128 = []
rutas_ean13  = []

for i, codigo in enumerate(df[COLUMNA_CODIGO]):
    codigo_str = str(int(codigo))
    
    # ---------- Code128 ----------
    if GENERAR_CODE128:
        ruta_c128 = f"{CARPETA_IMAGENES}/code128/{codigo_str}_code128_{i}.png"
        barcode_c128 = Code128(codigo_str, writer=ImageWriter())
        barcode_c128.save(ruta_c128.replace(".png",""), options=OPTIONS_CODE128)
        rutas_code128.append(ruta_c128)
    
    # ---------- EAN13 ----------
    if GENERAR_EAN13:
        ean13_str = codigo_str.zfill(13)  # rellenar a 13 dígitos
        ruta_ean13 = f"{CARPETA_IMAGENES}/ean13/{ean13_str}_ean13_{i}.png"
        barcode_ean13 = EAN13(ean13_str, writer=ImageWriter())
        barcode_ean13.save(ruta_ean13.replace(".png",""), options=OPTIONS_EAN13)
        rutas_ean13.append(ruta_ean13)

# Guardar Excel temporal
df.to_excel("temp.xlsx", index=False)

# Abrir Excel con openpyxl
wb = load_workbook("temp.xlsx")
ws = wb.active

# Función para encontrar índice de columna por nombre
def indice_columna(nombre_columna):
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value == nombre_columna:
            return idx
    return None

idx_c128 = indice_columna(COLUMNA_CODE128) if GENERAR_CODE128 else None
idx_ean13 = indice_columna(COLUMNA_EAN13) if GENERAR_EAN13 else None

# Insertar imágenes según flags
for i in range(2, ws.max_row + 1):
    if GENERAR_CODE128:
        img_c128 = Image(rutas_code128[i-2])
        img_c128.width = OPTIONS_CODE128_ANCHO_IMAGEN
        img_c128.height = OPTIONS_CODE128_ALTO_IMAGEN
        celda_c128 = f"{ws.cell(row=1, column=idx_c128).column_letter}{i}"
        ws.add_image(img_c128, celda_c128)
    
    if GENERAR_EAN13:
        img_ean13 = Image(rutas_ean13[i-2])
        img_ean13.width = OPTIONS_EAN13_ANCHO_IMAGEN
        img_ean13.height = OPTIONS_EAN13_ALTO_IMAGEN
        celda_ean13 = f"{ws.cell(row=1, column=idx_ean13).column_letter}{i}"
        ws.add_image(img_ean13, celda_ean13)
    
    ws.row_dimensions[i].height = 50

# Guardar Excel final
wb.save(ARCHIVO_SALIDA)

# Limpiar archivo temporal
os.remove("temp.xlsx")

print(f"✅ Excel generado: '{ARCHIVO_SALIDA}'")

fin = time.time()  # guardar tiempo final
print(f"⏳ Tiempo de ejecución: {fin - inicio:.2f} segundos")