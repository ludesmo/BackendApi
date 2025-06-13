from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd

app = FastAPI()  # Primero se crea la instancia

# Luego se agrega el middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://r-d-a-front-1.onrender.com"],  # O especifica el dominio de Angular si lo sabes
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/download-excel")
def download_excel():
    # Datos de ejemplo
    datos = [
        ["Juan Pérez", "12.345.678-9", "Informática", 20, 10, "Prof. Soto"],
        ["Ana Torres", "11.111.111-1", "Electrónica", 25, 15, "Prof. Reyes"],
        ["Luis Díaz", "22.222.222-2", "Mecánica", 30, 20, "Prof. Muñoz"],
        ["María Ruiz", "33.333.333-3", "Química", 18, 8, "Prof. Araya"],
    ]

    columnas = ["Alumno", "Rut Alumno", "Área", "Total Hras Realizadas", "Total BH OC", "Responsable"]

    # Crear libro y hoja
    wb = Workbook()
    ws = wb.active
    ws.title = "Informe de Alumnos Ayudantes"

    # Escribir encabezado
    ws.append(columnas)

    # Definir colores
    amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")  # LemonChiffon
    verde = PatternFill(start_color="C1E1C1", end_color="C1E1C1", fill_type="solid")     # LightGreen

    # Aplicar estilo al encabezado
    for cell in ws[1]:
        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Amarillo fuerte

    # Escribir datos y aplicar color alternado
    for i, fila in enumerate(datos):
        ws.append(fila)
        fill_color = amarillo if i % 2 == 0 else verde
        for cell in ws[i + 2]:  # +2 porque la primera fila es el encabezado
            cell.fill = fill_color

    # Guardar archivo
    ruta = "backend-fastapi/Informe de Alumnos.xlsx"
    wb.save(ruta)

    return FileResponse(ruta, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Informe de Alumnos.xlsx")

def read_root():
    return {"message": "API en funcionamiento"}