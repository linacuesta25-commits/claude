# -*- coding: utf-8 -*-
"""
Validación de créditos ACTIVADOS fuera del horario laboral, organizada POR TRIMESTRE.

Lee todos los reportes de la carpeta de entrada (por ejemplo Reporte_ene-mar.xlsx,
Reporte_abr-jun.xlsx, ...), los junta, quita duplicados exactos entre archivos y
reparte cada fila en su trimestre según la FECHA DE CREACIÓN DE TICKET. Así no
importa cómo se hayan descargado los archivos: cada fila cae en el trimestre
correcto y un mismo ticket no se cuenta dos veces.

Genera, en la subcarpeta "Reportes por trimestre":
  - Un archivo por trimestre: "Reporte creditos fuera de horario 2026-T1 (ene-mar).xlsx"
    con las hojas "Fuera de horario", "Resumen" y "Fechas con problemas" (si hay).
  - "Resumen por trimestre.xlsx": comparación de todos los trimestres y detalle por mes.

Usa las mismas reglas y constantes de validar_fuera_horario.py (que debe estar en
la misma carpeta). No modifica los archivos originales y no se conecta a internet.

Uso:
    python validar_por_trimestre.py                     (usa CARPETA_ENTRADA)
    python validar_por_trimestre.py "C:\\otra\\carpeta"
"""
import sys
from pathlib import Path

import pandas as pd

import validar_fuera_horario as base
from validar_fuera_horario import ErrorValidacion

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
CARPETA_ENTRADA = Path(
    r"C:\Users\LCUESTAC\Documents\prueba_automatizacion"
    r"\Validación clientes activados por fuera del horario laboral"
)
PATRON_ARCHIVOS = "Reporte_*.xlsx"          # qué archivos de la carpeta se leen
NOMBRE_CARPETA_SALIDA = "Reportes por trimestre"
NOMBRE_RESUMEN = "Resumen por trimestre.xlsx"

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
COL_TRIMESTRE = "TRIMESTRE"


# ============================================================================
# 1. LEER Y UNIR LOS ARCHIVOS
# ============================================================================
def buscar_archivos(carpeta: Path) -> list:
    """Lista los reportes de entrada (ignora temporales de Excel '~$')."""
    if not carpeta.is_dir():
        raise ErrorValidacion(f"No existe la carpeta de entrada:\n    {carpeta}")
    archivos = sorted(a for a in carpeta.glob(PATRON_ARCHIVOS) if not a.name.startswith("~$"))
    if not archivos:
        raise ErrorValidacion(f"No hay archivos '{PATRON_ARCHIVOS}' en:\n    {carpeta}")
    return archivos


def leer_archivo(ruta: Path) -> pd.DataFrame:
    """Lee un reporte con la misma lógica del script base y valida sus columnas clave."""
    hojas, hoja, filas = base.leer_excel(ruta)
    aviso = f" (el libro tiene {len(hojas)} hojas; se usa la primera)" if len(hojas) > 1 else ""
    df = base.construir_tabla(filas)
    for col in (base.COLUMNA_EMPRESA, base.COLUMNA_CREACION, base.COLUMNA_MODIFICACION, base.COLUMNA_ESTADO):
        try:
            base.buscar_columna(df, col)
        except ErrorValidacion as error:
            raise ErrorValidacion(f"En el archivo '{ruta.name}': {error}")
    df.insert(0, base.COL_ARCHIVO, ruta.name)
    print(f"  - {ruta.name}: hoja '{hoja}'{aviso}, {len(df)} filas")
    return df


def unir_archivos(tablas: list) -> pd.DataFrame:
    """
    Junta los archivos alineando las columnas por nombre normalizado (tildes,
    mayúsculas y espacios no importan). Se usan los nombres del primer archivo.
    """
    referencia = {}
    for tabla in tablas:
        for col in tabla.columns:
            referencia.setdefault(base.normalizar_texto(col), col)
    alineadas = [t.rename(columns=lambda c: referencia[base.normalizar_texto(c)]) for t in tablas]
    return pd.concat(alineadas, ignore_index=True)


def quitar_duplicados_entre_archivos(df: pd.DataFrame, col_ticket):
    """
    Si un mismo ticket viene en dos archivos con TODOS los datos iguales, se deja
    una sola copia. Si viene con datos distintos, se conservan ambas y se avisa.
    Los duplicados dentro de un mismo archivo no se tocan (se informan en el Resumen).
    """
    if col_ticket is None:
        return df, 0, pd.DataFrame()
    datos = [c for c in df.columns if c not in (base.COL_ARCHIVO, base.COL_FILA_EXCEL)]
    firma = pd.util.hash_pandas_object(df[datos].astype(str), index=False)
    repetida = firma.duplicated(keep="first") & (
        df.groupby(firma)[base.COL_ARCHIVO].transform("nunique") > 1
    )
    df = df[~repetida].reset_index(drop=True)

    varios = df.groupby(col_ticket)[base.COL_ARCHIVO].transform("nunique") > 1
    conflictos = df.loc[varios, [base.COL_ARCHIVO, base.COL_FILA_EXCEL, col_ticket]]
    return df, int(repetida.sum()), conflictos


# ============================================================================
# 2. ASIGNAR TRIMESTRE
# ============================================================================
def nombre_trimestre(fecha) -> str:
    t = (fecha.month - 1) // 3 + 1
    return f"{fecha.year}-T{t} ({MESES[(t - 1) * 3]}-{MESES[(t - 1) * 3 + 2]})"


def asignar_trimestre(df: pd.DataFrame, col_creacion) -> pd.DataFrame:
    """El trimestre sale de la FECHA DE CREACIÓN DE TICKET; sin fecha válida no hay trimestre."""
    fechas = pd.to_datetime(df[col_creacion].map(base.convertir_fecha), errors="coerce").dt.round("s")
    df["_FECHA_CREACION"] = fechas
    df[COL_TRIMESTRE] = fechas.map(lambda f: "" if pd.isna(f) else nombre_trimestre(f))
    return df


# ============================================================================
# 3. ANALIZAR UN TRIMESTRE (mismas reglas que el script base)
# ============================================================================
def analizar(df: pd.DataFrame, cols: dict) -> dict:
    ssi, activadas, otros = base.filtrar_ssi_activadas(df, cols["empresa"], cols["estado"])
    activadas, problemas = base.normalizar_fechas(activadas, [cols["creacion"], cols["modificacion"]])
    activadas = base.evaluar_horario(activadas, cols["creacion"], cols["modificacion"])
    fuera = activadas[activadas[base.COL_MOTIVO] != ""]

    duplicados = int(activadas[cols["ticket"]].duplicated(keep=False).sum()) if cols["ticket"] else 0
    tabla_problemas = pd.DataFrame(problemas, columns=[base.COL_ARCHIVO, "FILA EXCEL ORIGINAL", "COLUMNA",
                                                       "VALOR ENCONTRADO", "PROBLEMA"])
    archivos = ", ".join(sorted(df[base.COL_ARCHIVO].unique()))
    info = {
        "archivo": archivos, "hoja": "Primera hoja de cada archivo", "fila_encabezado": "Detectada por archivo",
        "leidas": len(df), "ssi": len(ssi), "activadas": len(activadas),
        "filas_problema": len(tabla_problemas.drop_duplicates([base.COL_ARCHIVO, "FILA EXCEL ORIGINAL"])),
        "duplicados": duplicados,
    }
    resumen = base.armar_resumen(info, otros, fuera, cols["creacion"], cols["modificacion"], cols["estado"])
    resumen.loc[resumen["CONCEPTO"] == "Archivo de entrada", "CONCEPTO"] = "Archivos de entrada"
    resumen.loc[resumen["CONCEPTO"] == "Fila de encabezados en Excel", "CONCEPTO"] = "Fila de encabezados"
    return {"info": info, "fuera": fuera, "activadas": activadas, "otros": otros,
            "problemas": tabla_problemas, "resumen": resumen}


def cobertura(fechas: pd.Series) -> tuple:
    """Rango de fechas con datos y si cubre el trimestre completo (de día 1 al último día)."""
    inicio_t = fechas.min().to_period("Q").start_time
    fin_t = fechas.min().to_period("Q").end_time
    desde, hasta = fechas.min(), fechas.max()
    completo = desde.date() == inicio_t.date() and hasta.date() == fin_t.date()
    return desde, hasta, completo


def filas_extra_resumen(res: dict, cols: dict, desde, hasta, completo) -> pd.DataFrame:
    """Filas adicionales para el Resumen de cada trimestre: cobertura y detalle por mes."""
    filas = [
        ("Cobertura", "Datos desde (creación de ticket)", desde.strftime("%d/%m/%Y %H:%M:%S")),
        ("Cobertura", "Datos hasta (creación de ticket)", hasta.strftime("%d/%m/%Y %H:%M:%S")),
        ("Cobertura", "Trimestre completo", "Sí" if completo else "No: los datos no cubren todo el trimestre"),
    ]
    act = res["activadas"]
    mes_act = act[cols["creacion"]].dt.month.value_counts()
    mes_fuera = res["fuera"][cols["creacion"]].dt.month.value_counts()
    for mes in sorted(mes_act.index):
        filas.append(("Por mes de creación", f"{MESES[mes - 1]} - SSI {base.ESTADO_VALIDO}", int(mes_act.get(mes, 0))))
        filas.append(("Por mes de creación", f"{MESES[mes - 1]} - fuera de horario", int(mes_fuera.get(mes, 0))))
    return pd.DataFrame(filas, columns=["SECCIÓN", "CONCEPTO", "VALOR"])


# ============================================================================
# 4. EXPORTAR
# ============================================================================
def exportar_resumen_general(ruta: Path, comparativo: pd.DataFrame, por_mes: pd.DataFrame,
                             sin_trimestre: pd.DataFrame, conflictos: pd.DataFrame):
    try:
        with pd.ExcelWriter(ruta, engine="openpyxl") as escritor:
            comparativo.to_excel(escritor, sheet_name="Por trimestre", index=False)
            por_mes.to_excel(escritor, sheet_name="Por mes", index=False)
            if not sin_trimestre.empty:
                sin_trimestre.to_excel(escritor, sheet_name="Sin trimestre", index=False)
            if not conflictos.empty:
                conflictos.to_excel(escritor, sheet_name="Tickets en varios archivos", index=False)
            for hoja in escritor.book.worksheets:
                base.dar_formato(hoja)
    except PermissionError:
        raise ErrorValidacion(f"No se pudo guardar porque el archivo está abierto:\n    {ruta}\n"
                              "Ciérralo en Excel y vuelve a ejecutar el script.")


# ============================================================================
# PROGRAMA PRINCIPAL
# ============================================================================
def procesar(carpeta: Path) -> dict:
    archivos = buscar_archivos(carpeta)
    print(f"Archivos encontrados ({len(archivos)}):")
    df = unir_archivos([leer_archivo(a) for a in archivos])

    cols = {
        "empresa": base.buscar_columna(df, base.COLUMNA_EMPRESA),
        "creacion": base.buscar_columna(df, base.COLUMNA_CREACION),
        "modificacion": base.buscar_columna(df, base.COLUMNA_MODIFICACION),
        "estado": base.buscar_columna(df, base.COLUMNA_ESTADO),
        "ticket": base.buscar_columna(df, base.COLUMNA_TICKET, obligatoria=False),
    }
    leidas_total = len(df)
    df, quitados, conflictos = quitar_duplicados_entre_archivos(df, cols["ticket"])
    df = asignar_trimestre(df, cols["creacion"])

    # Filas que no se pueden ubicar en ningún trimestre (no entran a ningún reporte)
    sin_trimestre = df.loc[df[COL_TRIMESTRE] == "",
                           [base.COL_ARCHIVO, base.COL_FILA_EXCEL, cols["empresa"], cols["estado"], cols["creacion"]]]
    sin_trimestre = sin_trimestre.rename(columns={cols["creacion"]: "VALOR ENCONTRADO EN FECHA DE CREACIÓN"})
    sin_trimestre = sin_trimestre.fillna("(vacío)")

    carpeta_salida = carpeta / NOMBRE_CARPETA_SALIDA
    carpeta_salida.mkdir(exist_ok=True)
    comparativo, por_mes, resultados = [], [], {}

    for trimestre in sorted(t for t in df[COL_TRIMESTRE].unique() if t):
        parte = df[df[COL_TRIMESTRE] == trimestre].drop(columns=["_FECHA_CREACION", COL_TRIMESTRE])
        desde, hasta, completo = cobertura(df.loc[df[COL_TRIMESTRE] == trimestre, "_FECHA_CREACION"])
        res = analizar(parte, cols)
        resumen = pd.concat([res["resumen"], filas_extra_resumen(res, cols, desde, hasta, completo)],
                            ignore_index=True)
        ruta = carpeta_salida / f"Reporte creditos fuera de horario {trimestre}.xlsx"
        base.exportar(ruta, res["fuera"], resumen, res["problemas"])

        fuera, info = res["fuera"], res["info"]
        motivos = fuera[base.COL_MOTIVO].value_counts()
        comparativo.append({
            "TRIMESTRE": trimestre,
            "DATOS DESDE": desde.to_pydatetime(), "DATOS HASTA": hasta.to_pydatetime(),
            "TRIMESTRE COMPLETO": "Sí" if completo else "No",
            "FILAS LEÍDAS": info["leidas"], f"FILAS {base.EMPRESA_VALIDA}": info["ssi"],
            f"{base.EMPRESA_VALIDA} {base.ESTADO_VALIDO}": info["activadas"],
            "FUERA DE HORARIO": len(fuera),
            "% FUERA DE HORARIO": round(len(fuera) / info["activadas"] * 100, 2) if info["activadas"] else 0,
            base.MOTIVO_CREACION: int(motivos.get(base.MOTIVO_CREACION, 0)),
            base.MOTIVO_MODIFICACION: int(motivos.get(base.MOTIVO_MODIFICACION, 0)),
            base.MOTIVO_AMBAS: int(motivos.get(base.MOTIVO_AMBAS, 0)),
            "FECHAS CON PROBLEMAS": info["filas_problema"],
            "ARCHIVO GENERADO": ruta.name,
        })
        act = res["activadas"]
        for periodo, grupo in act.groupby(act[cols["creacion"]].dt.to_period("M")):
            m = grupo[base.COL_MOTIVO].value_counts()
            por_mes.append({
                "MES": f"{MESES[periodo.month - 1]}-{periodo.year}", "TRIMESTRE": trimestre,
                f"{base.EMPRESA_VALIDA} {base.ESTADO_VALIDO}": len(grupo),
                "FUERA DE HORARIO": int((grupo[base.COL_MOTIVO] != "").sum()),
                base.MOTIVO_CREACION: int(m.get(base.MOTIVO_CREACION, 0)),
                base.MOTIVO_MODIFICACION: int(m.get(base.MOTIVO_MODIFICACION, 0)),
                base.MOTIVO_AMBAS: int(m.get(base.MOTIVO_AMBAS, 0)),
            })
        resultados[trimestre] = res
        print(f"  {trimestre}: {info['activadas']} {base.EMPRESA_VALIDA} {base.ESTADO_VALIDO}, "
              f"{len(fuera)} fuera de horario{'' if completo else '  (trimestre INCOMPLETO)'}")

    comparativo = pd.DataFrame(comparativo)
    total = {"TRIMESTRE": "TOTAL"}
    for col in comparativo.columns:
        if col not in total and pd.api.types.is_integer_dtype(comparativo[col]):
            total[col] = int(comparativo[col].sum())
    if total.get(f"{base.EMPRESA_VALIDA} {base.ESTADO_VALIDO}"):
        total["% FUERA DE HORARIO"] = round(total["FUERA DE HORARIO"] / total[f"{base.EMPRESA_VALIDA} {base.ESTADO_VALIDO}"] * 100, 2)
    comparativo = pd.concat([comparativo, pd.DataFrame([total])], ignore_index=True)
    por_mes = pd.DataFrame(por_mes)

    ruta_resumen = carpeta_salida / NOMBRE_RESUMEN
    exportar_resumen_general(ruta_resumen, comparativo, por_mes, sin_trimestre, conflictos)
    return {"archivos": archivos, "leidas": leidas_total, "quitados": quitados, "conflictos": conflictos,
            "sin_trimestre": sin_trimestre, "comparativo": comparativo, "por_mes": por_mes,
            "resultados": resultados, "salida": carpeta_salida}


def main():
    carpeta = Path(sys.argv[1]) if len(sys.argv) > 1 else CARPETA_ENTRADA
    try:
        r = procesar(carpeta)
    except ErrorValidacion as error:
        print(f"\nERROR: {error}")
        sys.exit(1)

    print("\n================ RESUMEN ================")
    print(f"Archivos leídos:                 {len(r['archivos'])}")
    print(f"Filas leídas:                    {r['leidas']}")
    if r["quitados"]:
        print(f"Filas repetidas entre archivos:  {r['quitados']} (se dejó una sola copia)")
    if not r["conflictos"].empty:
        print(f"Aviso: {r['conflictos'].iloc[:, 2].nunique()} tickets aparecen en varios archivos con datos "
              "distintos (ver hoja 'Tickets en varios archivos').")
    if not r["sin_trimestre"].empty:
        print(f"Aviso: {len(r['sin_trimestre'])} filas sin fecha de creación válida (ver hoja 'Sin trimestre').")
    print(r["comparativo"][["TRIMESTRE", "TRIMESTRE COMPLETO", f"{base.EMPRESA_VALIDA} {base.ESTADO_VALIDO}",
                            "FUERA DE HORARIO", "% FUERA DE HORARIO"]].to_string(index=False))
    print(f"Reportes guardados en:\n    {r['salida']}")


if __name__ == "__main__":
    main()
