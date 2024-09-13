import os
import pandas as pd
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont

# Cargar los datos desde Excel o CSV
def load_data(file_path):
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    else:
        raise ValueError("Archivo no soportado. Usa CSV o Excel.")

# Insertar texto centrado
def draw_centered_text(draw, text, position, font, image_width, color):
    # Obtener el tamaño del texto usando textbbox()
    bbox = draw.textbbox((0, 0), text, font=font)  # (left, top, right, bottom)
    text_width = bbox[2] - bbox[0]  # Ancho del texto
    text_height = bbox[3] - bbox[1]  # Altura del texto

    # Calcular la posición X para centrar el texto
    x_position = (image_width - text_width) // 2  # Centrar horizontalmente
    y_position = position[1]  # Mantener la coordenada Y original

    # Dibujar el texto centrado
    draw.text((x_position, y_position), text, font=font, fill=color)

# Insertar texto alineado a la izquierda
def draw_left_aligned_text(draw, text, position, font, color):
    # Usar la posición X original para alinear a la izquierda
    draw.text(position, text, font=font, fill=color)

# Función para generar un certificado
def generate_certificate(name, date, template_path, output_path):
    # Cargar la imagen de la plantilla (exportada de Figma)
    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)

    # Definir la fuente y tamaños
    name_font = ImageFont.truetype("C:/Windows/Fonts/seguisb.ttf", 28)
    date_font = ImageFont.truetype("C:/Windows/Fonts/segoeuil.ttf", 17)

    # Posiciones iniciales para nombre y fecha
    name_position = (0, 225)  # Solo la posición Y es relevante para el centrado
    date_position = (462, 434)

    # Obtener el ancho de la imagen para centrar texto
    image_width, image_height = image.size

    # Definir los colores
    # name_color = "#000000"
    # date_color = "#696969"

    # Escribir el nombre centrado
    draw_centered_text(draw, name, name_position, name_font, image_width, color="#000000")

    # Escribir la fecha alineada a la izquierda
    draw_left_aligned_text(draw, date, date_position, date_font, color="#696969")

    # Asegurarse de que la carpeta de salida exista
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # Guardar la imagen con el texto
    output_image_path = f"{output_path}/{name.replace(' ', '_')}.png"
    image.save(output_image_path, dpi=(300, 300))

    # Convertir las dimensiones de píxeles a milímetros
    width_mm = 842 * 0.352778
    height_mm = 595 * 0.352778

    # Crear un PDF con las mismas dimensiones que la imagen
    pdf = FPDF(unit="mm", format=[width_mm, height_mm])
    pdf.add_page()
    pdf.image(output_image_path, x=0, y=0, w=width_mm, h=height_mm)

    # Guardar el PDF final
    pdf_output = f"{output_path}/{name.replace(' ', '_')}.pdf"
    pdf.output(pdf_output)

    # Eliminar la imagen si no es necesaria
    # os.remove(output_image_path)

# Función principal
def main():
    # Ruta al archivo de datos (Excel o CSV) y a la plantilla del certificado
    directorio_datos = os.path.join(os.path.dirname(__file__), 'datos')
    data_file = os.path.join(directorio_datos, 'alumnos.xlsx')
    directorio_plantilla = os.path.join(os.path.dirname(__file__), 'plantilla')
    template_path = os.path.join(directorio_plantilla, 'plantilla_diploma.png')
    output_path = 'certificados'

    # Cargar los datos
    data = load_data(data_file)

    # Crear los certificados para cada nombre en la tabla
    for index, row in data.iterrows():
        name = row['Nombre']
        date = row['Fecha'].strftime('%d/%m/%Y')  # Formato de la fecha
        generate_certificate(name, date, template_path, output_path)

    print("Certificados generados con éxito.")

if __name__ == "__main__":
    main()
