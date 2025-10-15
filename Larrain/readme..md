Extracción de Tablas de PDF de LarrainVial a Excel
Este proyecto es un script de Python que automatiza la extracción, limpieza y transformación de datos de tablas presentes en los estados de cuenta de LarrainVial en formato PDF. El objetivo es convertir de manera eficiente las tablas del PDF a un archivo de Excel (.xlsx), facilitando el análisis y la manipulación de los datos financieros.

Características Principales
Extracción de Tablas: Utiliza la potente librería camelot para identificar y extraer todas las tablas de un PDF, incluso de documentos con estructuras complejas.

Limpieza de Datos: Las funciones de procesamiento de datos eliminan símbolos de moneda, separadores de miles y espacios, y convierten las cadenas de texto que representan valores numéricos a un formato de número flotante.

Formato de Excel: Aplica un formato numérico con separadores de miles y dos decimales (#,##0.00) para que los datos sean legibles y estén listos para ser utilizados en cálculos de Excel.

Múltiples Hojas de Trabajo: Cada tabla extraída del PDF se guarda en una hoja separada dentro del mismo archivo de Excel, organizando los datos de forma lógica.

Automatización Sencilla: El script está configurado para ejecutarse directamente, sin necesidad de argumentos de línea de comandos, apuntando a un archivo PDF específico y generando un archivo de salida predeterminado.

Requisitos
Asegúrate de tener Python 3.x y las siguientes librerías instaladas. Puedes instalarlas con pip:

Bash

pip install pandas openpyxl camelot-py
pip install "camelot-py[cv]"
Nota: La librería camelot-py depende de Ghostscript, el cual debe estar instalado en tu sistema para que el programa funcione. Puedes descargarlo desde aquí.

Uso
Asegúrate de que el archivo PDF con el que quieres trabajar (06 LV Cta 1.pdf por defecto) esté en el mismo directorio que el script de Python.

Ejecuta el script desde tu terminal:

Bash

python tu_script.py
(Reemplaza tu_script.py con el nombre de tu archivo Python).

Flujo del Proceso
El script sigue estos pasos:

Verifica la existencia del archivo PDF de entrada.

Utiliza camelot para leer y extraer todas las tablas del PDF.

Itera sobre cada tabla, la convierte a un DataFrame de pandas y limpia los datos.

Exporta el DataFrame a un archivo de Excel.

Aplica el formato numérico deseado a las celdas correspondientes en Excel.

Confirma que el proceso se ha completado y que el archivo de salida (reporte_larrainvial_cta1.xlsx por defecto) se ha generado.