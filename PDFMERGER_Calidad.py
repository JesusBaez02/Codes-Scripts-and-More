#needed libraries
import os
import re
import PyPDF2
from PyPDF2 import PdfReader , PdfWriter, PdfMerger
from natsort import natsorted

print("Start")

mes = input ("Mes a combinar ")
#sucursal = input ("sucursar a combinar ")
UB = input ("Tipo De Cuarto ")
#sucursal = input ("sucursal a combinar ")
contar = 0

#Set Origin Folder and make list
pdf_files = [_ for _ in os.listdir(r"C:/Users/gottec/Desktop/00 - PROPUESTA DM MITA/Entregable/MantenimientoYLimpieza/ProtocoloLimpInfraReportes/"+mes+"/"+UB) if _.endswith(".pdf")]

print("Running " + mes + " de " + "sucursal")

#define que eliminar para ordenar correctamente
#Mayo_Tel. Anonimo_

#print(natsorted(pdf_files))

truepdf_files = natsorted(pdf_files)

print("ListSorted")

print("TRABAJANDO")

#print("OG")
#print(pdf_files)
#print("NewList")
#print(truepdf_files)

#Activate and set Writer command
pdf_writer = PdfWriter()

print("Buscando Carpeta a juntar archivos")
#Loop for reading files
for pdf_file in truepdf_files:
    with open ("C:/Users/gottec/Desktop/00 - PROPUESTA DM MITA/Entregable/MantenimientoYLimpieza/ProtocoloLimpInfraReportes/"+mes+"/"+UB+"/{0}".format(pdf_file), "rb") as file:
        pdf_reader = PdfReader(file)
        contar = contar + 1
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

#print("Leido " + contar + " archivos")
#Set Output folder and create file
with open("C:/Users/gottec/Desktop/00 - PROPUESTA DM MITA/Entregable/MantenimientoYLimpieza/ProtocoloLimpInfraReportes/"+mes+"_"+UB+'.pdf', "wb") as output:
    pdf_writer.write(output)

contar = str(contar)
print("Finished " + contar + " archivos")
