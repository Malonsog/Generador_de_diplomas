import os
import pandas as pd
from PIL import ImageFont, ImageDraw, Image
import cairosvg

# Cargar los datos desde Excel o CSV
def load_data(file_path):
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    else:
        raise ValueError("Archivo no soportado. Usa CSV o Excel.")
    
# Función para calcular el tamaño del texto usando Pillow
def get_text_size(text, font_path, font_size):
    # Cargar la fuente con el tamaño deseado
    font = ImageFont.truetype(font_path, font_size)
    
    # Crear una imagen temporal para medir el texto
    image = Image.new('RGB', (1, 1))
    draw = ImageDraw.Draw(image)
    
    # Calcular el tamaño del texto (ancho y alto) usando textbbox()
    bbox = draw.textbbox((0, 0), text, font=font)
    
    # El tamaño del texto será el ancho y alto del bounding box
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    return text_width, text_height

# Función para agregar texto al archivo SVG
def add_text_to_svg(input_svg, output_svg, nombre, fecha, svg_width):
    # Leer el SVG original
    with open(input_svg, 'r', encoding='utf-8') as f:
        svg_content = f.read()

    # Obtener el tamaño del texto para el nombre
    font_path_name = "C:/Windows/Fonts/segoeui.ttf"
    font_size_name = 28
    name_width, name_height = get_text_size(nombre, font_path_name, font_size_name)

    # Calcular posición centrada para el nombre
    pos_x = (svg_width - name_width) // 2

    # Añadir texto del nombre y la fecha dentro del SVG
    text_to_add = f'''
    <text x="{pos_x}" y="255" font-size="28" fill="#000000" font-family="Segoe UI Semibold">{nombre}</text>
    <text x="460" y="453" font-size="16" fill="#696969" font-family="Segoe UI Light">{fecha}</text>
    '''
    
    # Añadir el texto antes de la etiqueta de cierre </svg>
    svg_content = svg_content.replace('</svg>', f'{text_to_add}\n</svg>')

    # Guardar el archivo SVG modificado
    with open(output_svg, 'w', encoding='utf-8') as f:
        f.write(svg_content)

# Función para convertir el archivo SVG modificado a PDF
def convert_svg_to_pdf(input_svg, output_pdf):
    # Convertir el archivo SVG a PDF
    cairosvg.svg2pdf(url=input_svg, write_to=output_pdf)

# Función para generar un certificado
def generate_certificate(nombre, fecha, grupo, plantilla_svg, directorio_salida, svg_width):
    # Asegurarse de que la carpeta de salida exista
    if not os.path.exists(directorio_salida):
        os.makedirs(directorio_salida)

    # Ruta para el SVG modificado
    output_svg_path = f"{directorio_salida}/{nombre.replace(' ', '_')}.svg"

    # Añadir texto al SVG (nombre y fecha)
    add_text_to_svg(plantilla_svg, output_svg_path, nombre, fecha, svg_width)

    # Ruta para el PDF final
    output_pdf_path = f"{directorio_salida}/{grupo.replace(' ', '_')}-{nombre.replace(' ', '_')}.pdf"

    # Convertir el archivo SVG modificado a PDF
    convert_svg_to_pdf(output_svg_path, output_pdf_path)

    # Eliminar el SVG si no es necesario
    os.remove(output_svg_path)

# Función principal
def main():
    # Ruta al archivo de datos (Excel o CSV)
    directorio_datos = os.path.join(os.path.dirname(__file__), 'datos')
    archivo_datos = os.path.join(directorio_datos, 'alumnos.xlsx')
    directorio_plantilla = os.path.join(os.path.dirname(__file__), 'plantilla')
    plantilla_svg = os.path.join(directorio_plantilla, 'plantilla_diploma.svg')
    directorio_salida = 'certificados'
    svg_width = 842

    # Cargar los datos
    data = load_data(archivo_datos)

    # Crear los certificados para cada nombre en la tabla
    for index, row in data.iterrows():
        nombre = row['Nombre']
        fecha = row['Fecha'].strftime('%d/%m/%Y')
        grupo = row['Grupo']
        generate_certificate(nombre, fecha, grupo, plantilla_svg, directorio_salida, svg_width)

    print("Certificados generados con éxito.")

if __name__ == "__main__":
    main()
