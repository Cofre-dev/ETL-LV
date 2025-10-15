Este proyecto es una herramienta de Python poderosa y flexible, diseñada para automatizar la extracción, limpieza y transformación de datos tabulares desde archivos PDF (como estados de cuenta y cartolas de inversión) directamente a hojas de cálculo de Excel (.xlsx).

Aunque está optimizada para manejar formatos comunes de instituciones financieras chilenas (como LarrainVial y Security), su núcleo puede adaptarse a cualquier documento PDF estructurado con tablas.

Características Principales
Extracción de Tablas Robustas: Utiliza la avanzada librería camelot para identificar y extraer todas las tablas presentes en cualquier página del PDF, manejando eficazmente estructuras complejas y texto espaciado.

Limpieza de Datos Inteligente:

Elimina caracteres no numéricos como símbolos de moneda ($), espacios y saltos de línea.

Convierte automáticamente los valores numéricos con formato regional (separador de miles con punto, decimal con coma) al formato estándar de Python (float).

Formato de Excel Profesional: Aplica un formato numérico localizado (#,##0.00) en la hoja de Excel, asegurando que los datos sean legibles y estén listos para análisis con el formato financiero adecuado (punto como separador de miles y coma como separador decimal).

Organización Detallada: Cada tabla extraída del PDF se exporta y se guarda en una hoja separada dentro del mismo archivo de Excel, manteniendo la estructura organizada.

Múltiples Modos de Ejecución:

Modo Interactivo (LarrainVial): Configurado para una ejecución simple sin argumentos, ideal para un flujo de trabajo fijo con un nombre de archivo de entrada y salida predeterminados.

Modo de Línea de Comandos (Security): Admite argumentos para especificar dinámicamente el archivo PDF de entrada y la ruta/nombre del archivo Excel de salida.

Requisitos e Instalación
Asegúrate de tener Python 3.x instalado. Luego, instala las librerías esenciales usando pip:

Bash

pip install pandas openpyxl camelot-py
pip install "camelot-py[cv]"
Dependencia Crítica (Ghostscript):

Para que camelot pueda procesar archivos PDF, es obligatorio tener instalado Ghostscript en tu sistema. Asegúrate de descargarlo e instalarlo desde su sitio web oficial.

Uso del Script
El script permite dos métodos de ejecución principales, dependiendo de cómo esté configurado tu archivo principal (main.py o tu_script.py).

1. Modo de Línea de Comandos (Flexible)
Ideal para procesar diferentes archivos PDF y especificar nombres de salida.

Comando	Descripción
Uso Básico	Procesa un archivo predeterminado (ej: Agosto Security.pdf).
python main.py	
Especificar PDF	Procesa el archivo en la ruta/nombre especificado.
python main.py "ruta/a/mi/cartola.pdf"	
Especificar Salida	Define la ruta y el nombre del archivo Excel de salida.
python main.py "input.pdf" -o "ruta/resultados.xlsx"	

Exportar a Hojas de cálculo
2. Modo Fijo (Ejecución Directa)
Ideal si solo necesitas procesar un archivo con un nombre fijo (ej. el estado de cuenta de LarrainVial).

Asegúrate de que el archivo PDF (ej: 06 LV Cta 1.pdf) esté en el mismo directorio.

Ejecuta el script sin argumentos:

Bash

python tu_script.py
El archivo Excel de salida (ej: reporte_larrainvial_cta1.xlsx) se generará automáticamente en el mismo directorio.

Flujo del Proceso Detallado
Verificación: Confirma la existencia del archivo PDF de entrada.

Extracción: Utiliza camelot.read_pdf para identificar y extraer todas las tablas.

Limpieza: Recorre el DataFrame de cada tabla y aplica la función de convertir_a_numerico para estandarizar los valores.

Exportación: Escribe el DataFrame limpio en una nueva hoja del archivo Excel de salida.

Formato: Aplica la función aplicar_formato_numerico para asegurar que la visualización final en Excel sea correcta para los datos financieros.

Confirmación: Muestra un mensaje en la terminal con la ruta final del archivo generado.
