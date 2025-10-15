from camelot import read_pdf
import pandas as pd
import os
import re
import sys
import argparse


def aplicar_formato_numerico(worksheet, df):
    
    FORMATO_NUMERICO_CHILE = '#,##0.00' 

    for col_idx, dtype in enumerate(df.dtypes, 1):
        if pd.api.types.is_numeric_dtype(dtype):
            for row_idx in range(2, worksheet.max_row + 1):
                celda = worksheet.cell(row=row_idx, column=col_idx)
                if celda.value is not None:
                    try:
                        celda.number_format = FORMATO_NUMERICO_CHILE
                    except Exception as e:
                        print(f"No se pudo aplicar formato a la celda en Fila: {row_idx}, Columna: {col_idx} - {e}")


def convertir_a_numerico(valor):
    
    if not isinstance(valor, str):
        return valor

    valor_limpio = re.sub(r'[$\s\n]', '', valor)
    
    if not valor_limpio:
        return None

    sin_separador_miles = valor_limpio.replace('.', '')
    formato_python = sin_separador_miles.replace(',', '.')

    try:
        return float(formato_python)
    except (ValueError, TypeError):
        return valor 
    
def extract_and_transform_security_tables(pdf_path, output_excel_path=None):

    # Si no se proporciona ruta de salida, generarla automáticamente
    if output_excel_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.dirname(pdf_path) if os.path.dirname(pdf_path) else '.'
        output_excel_path = os.path.join(output_dir, f"{base_name}.xlsx")

    if not os.path.exists(pdf_path):
        print(f"Error: El archivo PDF no se encontró en la ruta: {pdf_path}")
        return None
    
    try:
        tables = read_pdf(
            pdf_path,
            flavor='stream',
            pages='all',
            edge_tol=500
        )
    except Exception as e:
        print(f"Ocurrió un error al leer el PDF con Camelot: {e}")
        return None

    if tables.n == 0:
        print("No se encontraron tablas en el documento.")
        return None

    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            df_original = table.df
            
            df_original.columns = df_original.iloc[0]
            df_procesado = df_original.iloc[1:].copy()

            df_procesado = df_procesado.applymap(convertir_a_numerico)

            sheet_name = f"Pag_{table.page}_Tabla_{i + 1}"
            sheet_name = sheet_name[:31] 

            df_procesado.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
            
            worksheet = writer.sheets[sheet_name]
            aplicar_formato_numerico(worksheet, df_procesado)
            
            print(f"  -> Guardando y formateando la hoja '{sheet_name}'")

    print(f" Proceso completado. El archivo '{output_excel_path}' ha sido creado.")
    return output_excel_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Procesador de cartolas de inversiones de Security (PDF a Excel)'
    )
    parser.add_argument(
        'pdf_file',
        nargs='?',
        default="Agosto Security.pdf",
        help='Ruta al archivo PDF de la cartola de Security'
    )
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        default=None,
        help='Ruta del archivo Excel de salida (opcional, se genera automáticamente si no se especifica)'
    )

    args = parser.parse_args()

    # Verificar que el archivo existe
    if not os.path.exists(args.pdf_file):
        print(f"\nError: El archivo '{args.pdf_file}' no existe.")
        print("Por favor verifica la ruta del archivo.\n")
        sys.exit(1)

    # Procesar la cartola
    print(f"\nProcesando cartola: {args.pdf_file}")
    print("="*60)

    output_file = extract_and_transform_security_tables(args.pdf_file, args.output_file)

    if output_file:
        print("="*60)
        print(f"Proceso exitoso!")
        print(f"Archivo generado: {output_file}\n")
    else:
        print("\nError: No se pudo procesar la cartola.\n")
        sys.exit(1)