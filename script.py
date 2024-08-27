import pandas as pd
import qrcode
from fpdf import FPDF
from io import BytesIO

# Cargar el archivo Excel
df = pd.read_excel('Base.xlsx')

# Función para generar el QR directamente en el PDF
def generar_qr_pdf(pdf, data, x, y, w, h):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    
    # Crear la imagen QR en memoria
    img = qr.make_image(fill='black', back_color='white')
    buffer = BytesIO()
    img.save(buffer, format='PNG')
    buffer.seek(0)

    # Guardar la imagen temporalmente en el disco
    temp_filename = "temp_qr.png"
    with open(temp_filename, 'wb') as f:
        f.write(buffer.getvalue())

    # Agregar la imagen QR al PDF
    pdf.image(temp_filename, x=x, y=y, w=w, h=h)

# Función para generar PDF con una imagen de fondo y el QR
def generar_pdf(identificacion, nombre, tipo, fondo):
    pdf = FPDF()
    pdf.add_page()

    # Agregar imagen de fondo
    pdf.image(fondo, x=0, y=0, w=210, h=297)

    # Agregar texto sobre la imagen de fondo
    pdf.set_font("Arial", size=16, style='B')
    pdf.set_text_color(255, 255, 255)  # Texto blanco

    # Ajustar márgenes
    margen_superior = 40
    pdf.set_xy(0, margen_superior)
    pdf.cell(210, 10, txt="Boleta de Ingreso", ln=True, align='C')

    # Ajustar tamaño del texto y posición
    pdf.set_font("Arial", size=14)
    pdf.set_xy(10, margen_superior + 20)  # Ajusta la posición del texto
    pdf.cell(190, 10, txt=f"Identificación: {identificacion}", ln=True, align='C')
    pdf.cell(190, 10, txt=f"Nombre: {nombre}", ln=True, align='C')
    pdf.cell(190, 10, txt=f"Tipo: {tipo}", ln=True, align='C')

    # Agregar el QR al PDF
    generar_qr_pdf(pdf, identificacion, x=80, y=margen_superior + 60, w=50, h=50)

    pdf.output(f"boleta_{identificacion}.pdf")

# Generar PDFs para cada registro
for index, row in df.iterrows():
    identificacion = row['IDENTIFICACIÓN']
    nombre = row['NOMBRE']
    tipo = row['TIPO']
    
    # Ruta de la imagen de fondo
    fondo = "Fondo.png"  # Reemplaza con el nombre de tu imagen de fondo
    
    generar_pdf(identificacion, nombre, tipo, fondo)
