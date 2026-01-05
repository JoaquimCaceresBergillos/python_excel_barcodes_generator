# Generador de Códigos de Barras para Excel v2

Este script permite generar códigos de barras **Code128** y **EAN13** a partir de archivos Excel, insertando las imágenes de los códigos directamente en nuevas hojas de Excel.

---

## 1. Requisitos e Instalación

### 1.1. Instalar Python

1. Descarga Python desde la página oficial: [https://www.python.org/downloads/](https://www.python.org/downloads/)  
2. Durante la instalación, asegúrate de **marcar la opción "Add Python to PATH"**.
3. Verifica la instalación abriendo la terminal o CMD y ejecutando:

```bash
python --version
```

Deberías ver algo como:

```
Python 3.11.4
```

---

### 1.2 Instalar dependencias

El script requiere las siguientes librerías de Python:

- pandas
- openpyxl
- python-barcode
- Pillow (para manipulación de imágenes)

Instálalas ejecutando:

```bash
pip install pandas openpyxl python-barcode Pillow
```

---

## 2. Uso del programa

### 2.1. Preparar los archivos de entrada

- El script espera archivos **Excel (.xlsx)**.
- Cada archivo debe contener una columna con los códigos de barras a generar (por ejemplo: `cod_barras`).
- Los valores de los códigos deben ser **numéricos**. Para EAN13, el script completará con ceros a la izquierda si es necesario.

Ejemplo de archivo Excel de entrada:

| cod_barras | nombre_producto |
|------------|----------------|
| 1234567890 | Producto A     |
| 9876543210 | Producto B     |

---

### 2.2. Ejecutar el programa

1. Abre la terminal o CMD.
2. Navega hasta la carpeta donde se encuentra el script.
3. Ejecuta el script:

```bash
python nombre_del_script.py
```

El script te pedirá:

1. **Directorio que contiene los archivos Excel**.
2. **Nombre de la columna que contiene los códigos de barras**.
3. **Cantidad de filas por archivo de salida** (opcional, para dividir archivos grandes).

---

### 2.3. Archivos generados

Por cada archivo Excel de entrada, se generará:

1. Carpeta `Exportación` dentro del directorio de entrada.
2. Subcarpeta por cada archivo procesado con la siguiente estructura:

```
Exportación/
├─ archivo1/
│  ├─ barcodes/
│  │  ├─ bloque_1/
│  │  │  ├─ code128/   (si está activado)
│  │  │  ├─ ean13/     (si está activado)
│  │  │  └─ temp.xlsx
│  │  └─ bloque_2/
│  │     └─ ...
│  ├─ archivo1_barcodes_1.xlsx
│  └─ archivo1_barcodes_2.xlsx (si se dividió en bloques)
└─ archivo2/
   └─ ...
```

- Cada archivo Excel final contendrá las imágenes de los códigos de barras insertadas en las columnas correspondientes.
- Se ajusta automáticamente la altura de las filas para que las imágenes se vean correctamente.

---

### 2.4. Personalización extra

Dentro del script puedes modificar:

- `GENERAR_CODE128` y `GENERAR_EAN13` para habilitar o deshabilitar la generación de cada tipo de código.
- `OPTIONS_CODE128` y `OPTIONS_EAN13` para cambiar el tamaño, color y estilo de los códigos de barras.
- `OPTIONS_CODE128_ANCHO_IMAGEN`, `OPTIONS_CODE128_ALTO_IMAGEN` y equivalentes de EAN13 para ajustar la visualización en Excel.

---

### 3. Notas adicionales

- Si un código es inválido o vacío, se saltará automáticamente.
- El script imprime en consola mensajes sobre el progreso.
- Compatible con Windows, Linux y macOS.

---

**Autor:** JoaquimCB  
**Versión:** 2.0

