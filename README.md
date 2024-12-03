# Módulo de Procesamiento de Facturas Electrónicas (FEL) para SICOFI

## Descripción General

Este módulo de Python procesa **archivos XML de facturas electrónicas** (FEL) y los convierte en archivos CSV y Excel con el formato necesario para ser importados al sistema **SICOFI**. Soporta facturas de **ventas** y **compras**, automatizando la extracción, transformación y validación de datos.

---

## Características

- **Conversión de Archivos**:
  - Convierte archivos `.xls` a `.xlsx` para garantizar la compatibilidad.
  - Extrae datos relevantes de archivos XML y Excel.
- **Validación de Datos**:
  - Valida la estructura y el contenido de las facturas.
  - Identifica y omite archivos inválidos o incompletos.
- **Formatos de Salida**:
  - Genera archivos CSV listos para ser importados a SICOFI.
  - Crea reportes detallados en Excel para análisis adicional.
- **Interfaz de Usuario**:
  - Interfaz gráfica (GUI) implementada con `Tkinter`.
  - Botones para seleccionar archivos, generar reportes y abrir carpetas.

---

## Requisitos del Sistema

1. **Python 3.8+**
2. Dependencias:
   - `tkinter`
   - `pandas`
   - `openpyxl`
   - `xlrd`
   - `xml.etree.ElementTree`
   - `re`, `os`, `math`

---

## Estructura de Archivos

1. **Entradas**:
   - Archivos XML de facturas electrónicas.
   - Archivos `.xls` o `.xlsx` de datos adicionales.
2. **Salidas**:
   - **CSV**: Carpeta `Documentos/FEL-A-SICOFI/COMPRAS` para compras y `Documentos/FEL-A-SICOFI/VENTAS` para ventas.
   - **Excel**: Archivos detallados con formato profesional.

---

## Funcionalidades Principales

### 1. Procesamiento de Ventas (XML VENTAS a SICOFI)
Extrae información de los archivos XML de ventas, validando los datos y generando:
- Archivo CSV para importación a SICOFI.
- Archivo Excel con un reporte detallado de inventario.

### 2. Procesamiento de Compras (XML COMPRAS a SICOFI)
Procesa archivos XML de compras, mostrando:
- Una ventana previa con detalles de cada factura.
- Opción para aceptar o rechazar la factura antes de incluirla en el archivo final.
- Genera dos archivos:
  - CSV con facturas procesadas.
  - CSV con facturas no procesadas para manejo manual.

### 3. Conversión de Archivos Excel
Convierte archivos `.xls` a `.xlsx`, corrigiendo errores comunes en codificaciones de caracteres.

### 4. Validación de Archivos
Detecta y registra errores en:
- Archivos nulos (sin contenido útil).
- Facturas con problemas de formato descargadas desde SAT.

---

## Uso del Módulo

1. **Inicio**:
   - Ejecuta el script para abrir la aplicación GUI.
   - Selecciona la opción deseada en el menú.

2. **Procesamiento de Ventas**:
   - Haz clic en el botón **XML VENTAS a INVENTARIO**.
   - Selecciona los archivos XML desde la ventana emergente.
   - El módulo generará los archivos CSV y Excel.

3. **Procesamiento de Compras**:
   - Haz clic en el botón **XML COMPRAS a SICOFI**.
   - Selecciona los archivos XML desde la ventana emergente.
   - Se mostrarán los detalles de cada factura para revisión.
   - Al finalizar, los archivos procesados se guardan en la carpeta correspondiente.

4. **Abrir Carpeta de Archivos**:
   - Haz clic en **ABRIR CARPETA ARCHIVOS** para acceder a los archivos generados.

5. **Salir**:
   - Usa el menú **Salir** para cerrar la aplicación.

---

## Directorios de Salida

- **Documentos/FEL-A-SICOFI**:
  - Carpeta raíz para todos los archivos generados.
- **Documentos/FEL-A-SICOFI/COMPRAS**:
  - CSV y log de compras procesadas y no procesadas.
- **Documentos/FEL-A-SICOFI/VENTAS**:
  - CSV y reportes detallados de ventas.

---

## Log de Errores

El módulo genera un archivo log en caso de errores:
- Archivos nulos.
- Archivos descargados incorrectamente del portal SAT.

---

## Créditos

Desarrollado por **Jonatan Gonzalez**. Este módulo simplifica el procesamiento de facturas electrónicas para integrarlas al sistema SICOFI, ahorrando tiempo y reduciendo errores.
