import os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO

# Rutas
carpeta_pdfs = "C:/Users/gottec/Documents/ctrl"  # <-- CAMBIA esto
archivo_pie = "C:/Users/gottec/Documents/ctrl/pie_martens.pdf"      # <-- CAMBIA esto

# Cargar el PDF del pie de página
pie_pdf = PdfReader(archivo_pie).pages[0]

# Función para crear una página PDF con texto personalizado
def crear_overlay_texto(texto):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont("Helvetica", 10)
    c.drawString(46, 15, texto)  # Posición del texto
    c.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]

# Iterar sobre los archivos PDF
for archivo in os.listdir(carpeta_pdfs):
    if archivo.endswith(".pdf"):
        ruta_pdf = os.path.join(carpeta_pdfs, archivo)
        reader = PdfReader(ruta_pdf)
        writer = PdfWriter()
        total_paginas = len(reader.pages)

        for i, pagina in enumerate(reader.pages):
            #Comentar if si se necesita poner pie en todas las paginas 
            #if i == 0:
                # Página 1: solo agregarla sin cambios
                #writer.add_page(pagina)
                #continue

            # Clonar la página para evitar modificar el original
            nueva_pagina = pagina
            nueva_pagina.merge_page(pie_pdf)  # Pie de página

            # Texto personalizado
            texto = f"CALLE SACROMONTE NÚMERO 280, INTERIOR 1-A, COLONIA CHAPALITA, GUADALAJARA, JALISCO, C.P. 44500"
            overlay_texto = crear_overlay_texto(texto)
            nueva_pagina.merge_page(overlay_texto)

            writer.add_page(nueva_pagina)

        # Guardar archivo
        salida = os.path.join(carpeta_pdfs, f"con_pie_{archivo}")
        with open(salida, "wb") as f:
            writer.write(f)

print("✅ Proceso terminado sin modificar la portada.")
