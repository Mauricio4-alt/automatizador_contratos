from docx import Document
from pathlib import Path
from datetime import date
import pandas as pd
OUTPUT_DIR = Path("contratos")
OUTPUT_DIR.mkdir(exist_ok=True)

CIUDAD = "Bogotá"
EMPRESA = "limpiaFacil"
NIT = "155779881"

CONTRATO_TEMPLATE = """
Entre los suscritos a saber, {empresa}, identificada con NIT {nit}, con domicilio en la ciudad de {ciudad},
quien para efectos de este contrato se denominará el EMPLEADOR, y por la otra parte {nombre},
identificado(a) con la cédula de ciudadanía No. {documento}, quien para efectos de este contrato se denominará
el TRABAJADOR, se ha convenido celebrar el presente contrato de trabajo bajo las siguientes cláusulas:

PRIMERA. OBJETO:
El EMPLEADOR contrata los servicios personales del TRABAJADOR para desempeñar el cargo de {cargo}.

TERCERA. DURACIÓN:
Contrato a TÉRMINO INDEFINIDO. Inicia el día {fecha}.
"""

def crear_contratos(df):
    columnas = {"Nombre", "Documento", "Cargo"}
    if not columnas.issubset(df.columns):
        raise ValueError("El Excel no tiene las columnas requeridas")

    fecha = date.today().strftime("%d/%m/%Y")

    for _, row in df.iterrows():
        doc = Document()

        texto = CONTRATO_TEMPLATE.format(
            empresa=EMPRESA,
            nit=NIT,
            ciudad=CIUDAD,
            nombre=row["Nombre"],
            documento=row["Documento"],
            cargo=row["Cargo"],
            fecha=fecha
        )

        doc.add_paragraph(texto)

        nombre_archivo = f"contrato_{row['Nombre']}_{row['Documento']}.docx"
        doc.save(OUTPUT_DIR / nombre_archivo)
