from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd

app = FastAPI()  # Primero se crea la instancia

# Luego se agrega el middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # O especifica el dominio de Angular si lo sabes
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/download-excel")
def download_excel():
    # Crear DataFrame vacío
    df = pd.DataFrame(columns=["Columna 1", "Columna 2", "Columna 3"])

    # Guardar en archivo Excel
    filename = "archivo_generado.xlsx"
    df.to_excel(filename, index=False)

    # Retornar archivo para descargar
    return FileResponse(
        path=filename,
        filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )