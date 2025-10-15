import camelot
import pandas as pd
import os
import re

# --- FUNCIÓN PARA APLICAR FORMATO NUMÉRICO EN EXCEL ---
def aplicar_formato_numerico(worksheet, df):
    # Formato para Excel: #,##0.00 -> separador de miles con punto, dos decimales
    FORMATO_NUMERICO_CHILE = '#,##0.00' 

    # Itera sobre cada columna del DataFrame
    for col_idx, dtype in enumerate(df.dtypes, 1):
        # Si la columna es de tipo numérico...
        if pd.api.types.is_numeric_dtype(dtype):
            # ...itera sobre cada celda de esa columna en la hoja de Excel
            for row_idx in range(2, worksheet.max_row + 1):
                celda = worksheet.cell(row=row_idx, column=col_idx)
                # Evita formatear celdas vacías
                if celda.value is not None:
                    celda.number_format = FORMATO_NUMERICO_CHILE

# --- FUNCIÓN PARA CONVERTIR STRINGS A NÚMEROS ---
def convertir_a_numerico(valor):
    # Si el valor ya es un número o no es un texto, lo devolvemos tal cual.
    if not isinstance(valor, str):
        return valor

    # 1. Quitar símbolos que no son parte del número ($, espacios, saltos de línea).
    valor_limpio = re.sub(r'[$\s\n]', '', valor)

    # 2. Quitar los separadores de miles (el punto '.').
    sin_separador_miles = valor_limpio.replace('.', '')

    # 3. Reemplazar la coma decimal por un punto decimal.
    formato_python = sin_separador_miles.replace(',', '.')

    # 4. Intentar convertir el string ya limpio a un número flotante.
    try:
        return float(formato_python)
    except (ValueError, TypeError):
        # Si no es un número válido, devolver el valor original.
        return valor

# --- FUNCIÓN PRINCIPAL DE EXTRACCIÓN ---
def extract_all_tables_to_excel(pdf_path, output_excel_path):
    """
    Lee, extrae, limpia y guarda todas las tablas de un PDF a un Excel con formato numérico.
    """
    print(f"--- Iniciando el procesamiento del archivo: {pdf_path} ---")

    if not os.path.exists(pdf_path):
        print(f"Error: El archivo PDF no se encontró en la ruta: {pdf_path}")
        return
    
    try:
        tables = camelot.read_pdf(
            pdf_path,
            flavor='stream',
            pages='all',
            edge_tol=500
        )
    except Exception as e:
        print(f"Ocurrió un error al leer el PDF con Camelot: {e}")
        return

    if tables.n == 0:
        print("No se encontraron tablas en el documento.")
        return

    print(f"¡Éxito! Se encontraron {tables.n} tablas en total.")
    print(f"Limpiando y guardando las tablas en: {output_excel_path}")

    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            
            # 1. Limpiar los datos para convertirlos a número en Pandas
            df_procesado = table.df.map(convertir_a_numerico)

            # Usamos un nombre de hoja único para cada tabla
            sheet_name = f"Pagina_{table.page}_Tabla_{i+1}"
            sheet_name = sheet_name[:31] # El nombre de la hoja no puede superar 31 caracteres

            # 2. Guardar el DataFrame en la hoja
            df_procesado.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
            
            # 3. Aplicar el formato de visualización en Excel
            worksheet = writer.sheets[sheet_name]
            aplicar_formato_numerico(worksheet, df_procesado)
            
            print(f"  -> Guardando y formateando la hoja '{sheet_name}'")

    print(f"--- Proceso completado. El archivo '{output_excel_path}' ha sido creado. ---")


if __name__ == "__main__":
    
    pdf_file = "06 LV Cta 1.pdf" 
    
    # Se genera un nuevo archivo Excel para no sobreescribir el anterior
    output_excel_file = "reporte_larrainvial_cta1.xlsx"
    
    extract_all_tables_to_excel(pdf_file, output_excel_file)