Extracción de Tablas PDF a Excel
Este proyecto es un script de Python diseñado para automatizar la extracción y transformación de tablas de documentos PDF, con un enfoque particular en los estados de cuenta de inversiones de Security. El script lee las tablas del PDF, limpia y convierte los datos numéricos y exporta cada tabla a una hoja separada de un archivo de Excel.

Características
Extracción de Tablas: Utiliza la librería camelot para identificar y extraer tablas de cualquier página de un PDF.

Limpieza de Datos: Las columnas numéricas se limpian de caracteres especiales (como $, ., , y espacios) y se convierten al formato numérico adecuado.

Formato de Excel: Aplica un formato numérico con separadores de miles y decimales al estilo chileno (#,##0.00) para una mejor visualización en el archivo de Excel de salida.

Generación de Hojas: Cada tabla extraída del PDF se guarda en una hoja individual dentro del archivo de Excel.

Interfaz de Línea de Comandos: Permite procesar archivos PDF de manera sencilla a través de la terminal, con opciones para especificar la ruta del archivo de entrada y la de salida.

Requisitos
Asegúrate de tener Python 3.x instalado. Puedes instalar las librerías necesarias con el siguiente comando:

Bash

pip install pandas openpyxl camelot-py
pip install "camelot-py[cv]"
Nota: Para que Camelot funcione correctamente, es necesario tener instalado Ghostscript.

Uso
El script se puede ejecutar directamente desde la línea de comandos.

Uso Básico:
Copia el archivo PDF que deseas procesar en la misma carpeta que el script y ejecuta:

Bash

python main.py
Por defecto, el script buscará un archivo llamado Agosto Security.pdf. Se creará un archivo de Excel con el mismo nombre en la misma carpeta.

Especificar un Archivo PDF:
Puedes especificar el archivo de entrada como un argumento:

Bash

python main.py "ruta/a/tu/archivo.pdf"
Especificar Archivo de Salida:
Usa la opción -o o --output para definir la ruta y el nombre del archivo de Excel de salida.

Bash

python main.py "ruta/a/tu/archivo.pdf" -o "ruta/de/salida/resultados.xlsx"
Ejemplo
Al ejecutar el script, verás la siguiente salida en tu terminal, que muestra el progreso del proceso:

Bash

Procesando cartola: Agosto Security.pdf
============================================================
  -> Guardando y formateando la hoja 'Pag_1_Tabla_1'
  -> Guardando y formateando la hoja 'Pag_2_Tabla_2'
Proceso completado. El archivo 'Agosto Security.xlsx' ha sido creado.
============================================================
Proceso exitoso!
Archivo generado: Agosto Security.xlsx