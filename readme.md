# Generador de Diplomas

Este proyecto contiene dos scripts en Python que permiten generar diplomas personalizados a partir de un listado de alumnos en un archivo Excel. 
Se utiliza una plantilla de imagen para el formato del diploma, a la cual se añade el nombre del alumno y la fecha en que completó el curso, exportándolo en formato PDF.
En la carpeta *plantilla* incluyo un archivo Figma con tres opciones de plantillas para el diploma.

### Funcionalidades:
- Lee un archivo Excel con los nombres de los alumnos y la fecha de finalización del curso.
- Usa una plantilla de imagen para generar los diplomas.
- Inserta el nombre del alumno y la fecha de finalización en la plantilla.
- Exporta el diploma en formato PDF para cada alumno.

### Versiones:
1. **diplomas_batch.py**:
   - Usa una plantilla en formato PNG.
   - Puede presentar problemas de resolución si la imagen no es lo suficientemente grande.

2. **diplomas_batch_svg.py**:
   - Usa una plantilla en formato SVG para evitar problemas de resolución.
   - Recomendado para plantillas con alta calidad.

### Requisitos:
- Python 3.6+
- Las siguientes dependencias deben instalarse:
  - `Pillow`: Para la manipulación de imágenes.
  - `cairosvg`: Para manipular y exportar SVG a PDF (en diplomas_batch_svg.py).
  - `pandas`: Para leer el archivo Excel.
  - `openpyxl`: Soporte para archivos Excel.

Puedes instalar todas las dependencias ejecutando el siguiente comando:

```bash
pip install -r requirements.txt
