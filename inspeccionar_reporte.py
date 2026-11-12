# -*- coding: utf-8 -*-
"""
Paso previo: inspecciona la estructura de "Reporte 1.xlsx" SIN modificarlo.
Solo imprime en consola: hojas, fila de encabezados, columnas, valores distintos
de EMPRESA y de las columnas que contienen "ACTIVADA", y una muestra de fechas.
No escribe archivos ni envía nada a internet.

Uso (en Windows):
    python inspeccionar_reporte.py
"""
import unicodedata
from pathlib import Path

import pandas as pd

RUTA_ENTRADA = Path(
    r"C:\Users\LCUESTAC\Documents\prueba_automatizacion"
    r"\Validación clientes activados por fuera del horario laboral\Reporte 1.xlsx"
)
FILAS_A_REVISAR = 30  # filas iniciales donde se busca el encabezado


def normalizar(texto) -> str:
    """Quita tildes, espacios extra y pasa a mayúsculas."""
    if pd.isna(texto):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).upper()


def main():
    if not RUTA_ENTRADA.exists():
        print(f"No se encontró el archivo: {RUTA_ENTRADA}")
        return

    libro = pd.ExcelFile(RUTA_ENTRADA, engine="openpyxl")
    print("Hojas del libro:", libro.sheet_names)
    hoja = libro.sheet_names[0]
    print(f"Se inspecciona la primera hoja: '{hoja}'\n")

    crudo = pd.read_excel(libro, sheet_name=hoja, header=None, dtype=object)
    print(f"Dimensiones en bruto (filas x columnas): {crudo.shape}")

    # Buscar la fila de encabezados: la que contenga EMPRESA
    fila_enc = None
    for i in range(min(FILAS_A_REVISAR, len(crudo))):
        valores = [normalizar(v) for v in crudo.iloc[i].tolist()]
        if "EMPRESA" in valores:
            fila_enc = i
            break
    if fila_enc is None:
        print("No se encontró una fila con 'EMPRESA'. Primeras 10 filas en bruto:")
        print(crudo.head(10).to_string())
        return
    print(f"Fila de encabezados: fila {fila_enc + 1} de Excel\n")

    df = pd.read_excel(libro, sheet_name=hoja, header=fila_enc, dtype=object)
    print(f"Filas de datos: {len(df)}")
    print("Columnas encontradas:")
    for n, col in enumerate(df.columns, 1):
        tipos = df[col].dropna().map(lambda v: type(v).__name__).value_counts().to_dict()
        print(f"  {n:>2}. {col!r}  vacíos={df[col].isna().sum()}  tipos={tipos}")

    # Valores distintos de EMPRESA
    col_emp = next((c for c in df.columns if normalizar(c) == "EMPRESA"), None)
    if col_emp is not None:
        print("\nValores distintos de EMPRESA (conteo):")
        print(df[col_emp].map(repr).value_counts(dropna=False).to_string())

    # Columnas que contienen ACTIVADA
    print("\nColumnas que contienen el valor ACTIVADA:")
    for col in df.columns:
        norm = df[col].map(normalizar)
        if norm.eq("ACTIVADA").any():
            print(f"\n  -> {col!r}. Valores distintos (conteo):")
            print(df[col].map(repr).value_counts(dropna=False).head(40).to_string())

    # Muestra de fechas
    for col in df.columns:
        if "FECHA" in normalizar(col):
            print(f"\nMuestra de {col!r} (valor y tipo):")
            for v in df[col].dropna().head(8):
                print(f"   {v!r}  ({type(v).__name__})")
            textos = df[col].dropna()
            textos = textos[textos.map(lambda v: isinstance(v, str))]
            if len(textos):
                print(f"   Hay {len(textos)} valores como TEXTO. Ejemplos: "
                      f"{textos.head(8).tolist()}")


if __name__ == "__main__":
    main()
