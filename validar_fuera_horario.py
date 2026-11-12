# -*- coding: utf-8 -*-
"""
Validación de créditos ACTIVADOS por fuera del horario laboral.

Lee "Reporte 1.xlsx" (sin modificarlo nunca), filtra las solicitudes de la
empresa SSI que están ACTIVADAS y revisa si la FECHA DE CREACIÓN DE TICKET y/o
la FECHA ULTIMA MODIFICACION caen fuera del horario laboral. El resultado se
guarda en un archivo NUEVO: "Reporte creditos fuera de horario.xlsx".

Todo se procesa localmente; el script no se conecta a internet.

Uso:
    python validar_fuera_horario.py                 (usa RUTA_ENTRADA)
    python validar_fuera_horario.py "otra ruta.xlsx"

Librerías:
    pip install pandas openpyxl
    pip install pywin32   (solo si el Excel tiene etiqueta de confidencialidad /
                           protección de Microsoft; el script lo abre con Excel)
"""
import sys
import unicodedata
import zipfile
from datetime import datetime, time
from pathlib import Path

try:
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import InvalidFileException
except ImportError as error:
    print(f"Falta una librería ({error.name}). Instálala con:")
    print("    pip install pandas openpyxl")
    sys.exit(1)

# ============================================================================
# CONFIGURACIÓN (cambia aquí lo que necesites)
# ============================================================================
RUTA_ENTRADA = Path(
    r"C:\Users\LCUESTAC\Documents\prueba_automatizacion"
    r"\Validación clientes activados por fuera del horario laboral\Reporte 1.xlsx"
)
NOMBRE_SALIDA = "Reporte creditos fuera de horario.xlsx"

COLUMNA_EMPRESA = "EMPRESA"
COLUMNA_CREACION = "FECHA DE CREACIÓN DE TICKET"
COLUMNA_MODIFICACION = "FECHA ULTIMA MODIFICACION"
COLUMNA_ESTADO = "IDENTIFICADOR"   # confirmada: aquí aparece el valor ACTIVADA
COLUMNA_TICKET = "TICKET"          # solo para avisar de tickets duplicados

EMPRESA_VALIDA = "SSI"
ESTADO_VALIDO = "ACTIVADA"

HORA_INICIO = 8    # 08:00:00 ya es horario laboral
HORA_FIN = 20      # 20:00:00 ya es FUERA de horario
EVALUAR_FINES_DE_SEMANA = False  # True: sábado y domingo siempre fuera de horario

FILAS_BUSQUEDA_ENCABEZADO = 30  # filas iniciales donde se busca el encabezado
FORMATO_FECHA_EXCEL = "dd/mm/yyyy hh:mm:ss"

# Nombres de las columnas y valores que agrega el script
COL_FILA_EXCEL = "FILA EXCEL ORIGINAL"
COL_DIA = "DIA DE LA SEMANA"
COL_HORA_CREACION = "HORA CREACION"
COL_HORA_MODIFICACION = "HORA ULTIMA MODIFICACION"
COL_CREACION_FUERA = "CREACION FUERA DE HORARIO"
COL_MODIFICACION_FUERA = "MODIFICACION FUERA DE HORARIO"
COL_MOTIVO = "MOTIVO"
MOTIVO_CREACION = "Creación fuera de horario"
MOTIVO_MODIFICACION = "Última modificación fuera de horario"
MOTIVO_AMBAS = "Ambas fuera de horario"
SIN_FECHA = "Sin fecha válida"

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


class ErrorValidacion(Exception):
    """Error con un mensaje claro para mostrar en consola."""


# ============================================================================
# UTILIDADES
# ============================================================================
def normalizar_texto(valor) -> str:
    """Quita tildes, espacios extra (incluido el espacio duro) y pasa a mayúsculas."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).upper()


def nombres_unicos(nombres: list) -> list:
    """Evita columnas repetidas o vacías en el encabezado."""
    vistos, resultado = {}, []
    for i, nombre in enumerate(nombres, start=1):
        nombre = str(nombre).strip() if nombre is not None and str(nombre).strip() else f"COLUMNA {i}"
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre} ({vistos[nombre]})"
        else:
            vistos[nombre] = 1
        resultado.append(nombre)
    return resultado


# ============================================================================
# 1. LEER
# ============================================================================
def leer_con_openpyxl(ruta: Path):
    """Lee la primera hoja en modo solo lectura. Devuelve (hojas, hoja, filas)."""
    libro = load_workbook(ruta, read_only=True, data_only=True)
    try:
        hojas = libro.sheetnames
        hoja = libro[hojas[0]]
        filas = [list(f) for f in hoja.iter_rows(min_row=1, values_only=True)]
    finally:
        libro.close()
    return hojas, hojas[0], filas


def leer_con_excel(ruta: Path):
    """
    Para archivos con etiqueta de confidencialidad (protegidos por Microsoft),
    que pandas/openpyxl no pueden abrir: los abre Excel en segundo plano, en
    SOLO LECTURA, copia los valores a memoria y cierra SIN guardar.
    """
    try:
        import win32com.client  # pywin32
    except ImportError:
        raise ErrorValidacion(
            "El archivo está protegido (etiqueta de confidencialidad de Microsoft) y "
            "solo se puede abrir a través de Excel. Instala pywin32 con:\n"
            "    pip install pywin32"
        )
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    libro = None
    try:
        libro = excel.Workbooks.Open(str(ruta), UpdateLinks=0, ReadOnly=True)
        hojas = [h.Name for h in libro.Worksheets]
        hoja = libro.Worksheets(1)
        rango = hoja.UsedRange
        valores = rango.Value
        if not isinstance(valores, tuple):
            valores = ((valores,),)
        # Filas vacías antes del rango usado, para conservar el número de fila de Excel
        filas = [[None]] * (rango.Row - 1)
        for fila in valores:
            filas.append([
                v.replace(tzinfo=None) if isinstance(v, datetime) else v for v in fila
            ])
        return hojas, hojas[0], filas
    finally:
        if libro is not None:
            libro.Close(SaveChanges=False)
        excel.Quit()


def leer_excel(ruta: Path):
    """Lee la primera hoja del libro sin modificarlo. Devuelve (hojas, hoja, filas crudas)."""
    if not ruta.exists():
        raise ErrorValidacion(f"No se encontró el archivo de entrada:\n    {ruta}")
    try:
        return leer_con_openpyxl(ruta)
    except (zipfile.BadZipFile, InvalidFileException, KeyError):
        if sys.platform != "win32":
            raise ErrorValidacion(
                "El archivo no es un .xlsx legible (posiblemente está protegido con una "
                "etiqueta de confidencialidad). Solo se puede abrir con Excel en Windows."
            )
        print("Aviso: el archivo está protegido; se abre con Excel en modo solo lectura...")
        return leer_con_excel(ruta)


# ============================================================================
# 2. NORMALIZAR (encabezados, columnas clave y fechas)
# ============================================================================
def construir_tabla(filas: list) -> pd.DataFrame:
    """Detecta la fila de encabezados y arma el DataFrame con la fila de Excel de cada dato."""
    claves = {normalizar_texto(c) for c in (COLUMNA_EMPRESA, COLUMNA_CREACION, COLUMNA_MODIFICACION)}
    fila_encabezado, mejor = None, -1
    for i, fila in enumerate(filas[:FILAS_BUSQUEDA_ENCABEZADO]):
        coincidencias = len(claves & {normalizar_texto(v) for v in fila})
        if coincidencias > mejor:
            fila_encabezado, mejor = i, coincidencias
        if coincidencias == len(claves):
            break
    if fila_encabezado is None:
        raise ErrorValidacion("La hoja está vacía.")

    ancho = max(len(f) for f in filas)
    encabezado = nombres_unicos(list(filas[fila_encabezado]) + [None] * (ancho - len(filas[fila_encabezado])))
    datos, numeros = [], []
    for i in range(fila_encabezado + 1, len(filas)):
        fila = list(filas[i]) + [None] * (ancho - len(filas[i]))
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in fila):
            continue  # fila totalmente vacía
        datos.append(fila)
        numeros.append(i + 1)  # número de fila real en Excel
    df = pd.DataFrame(datos, columns=encabezado, dtype=object)
    df.insert(0, COL_FILA_EXCEL, numeros)
    df.attrs["fila_encabezado"] = fila_encabezado + 1
    return df


def buscar_columna(df: pd.DataFrame, nombre: str, obligatoria: bool = True):
    """Encuentra una columna ignorando tildes, mayúsculas y espacios extra."""
    objetivo = normalizar_texto(nombre)
    for col in df.columns:
        if normalizar_texto(col) == objetivo:
            return col
    if obligatoria:
        encontradas = "\n".join(f"    - {c}" for c in df.columns if c != COL_FILA_EXCEL)
        raise ErrorValidacion(
            f"Falta la columna obligatoria '{nombre}'.\nColumnas encontradas:\n{encontradas}"
        )
    return None


def convertir_fecha(valor):
    """Convierte fecha real de Excel, número de serie o texto (día/mes/año). Devuelve NaT si no se puede."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return pd.NaT
    if isinstance(valor, (datetime, pd.Timestamp)):
        fecha = pd.Timestamp(valor)
        return fecha.tz_localize(None) if fecha.tzinfo else fecha
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if 1 <= valor < 2958466:  # número de serie de fecha de Excel
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(valor, unit="D")
        return pd.NaT
    texto = " ".join(str(valor).split())
    if not texto:
        return pd.NaT
    # Formato colombiano "p. m." / "a. m." -> PM / AM
    texto_ampm = (texto.lower().replace("p. m.", "PM").replace("a. m.", "AM")
                  .replace("p.m.", "PM").replace("a.m.", "AM"))
    return pd.to_datetime(texto_ampm, dayfirst=True, errors="coerce")


def normalizar_fechas(df: pd.DataFrame, columnas: list):
    """
    Convierte las columnas de fecha clave. Redondea al segundo (Excel guarda las
    horas como decimales y 08:00:00 podría leerse como 07:59:59.999).
    Devuelve el DataFrame y la lista de fechas con problemas.
    """
    problemas = []
    for col in columnas:
        originales = df[col]
        convertidas = originales.map(convertir_fecha)
        convertidas = pd.to_datetime(convertidas, errors="coerce").dt.round("s")
        for idx in convertidas[convertidas.isna()].index:
            valor = originales.loc[idx]
            vacio = valor is None or (isinstance(valor, float) and pd.isna(valor)) or not str(valor).strip()
            problemas.append({
                "FILA EXCEL ORIGINAL": df.at[idx, COL_FILA_EXCEL],
                "COLUMNA": col,
                "VALOR ENCONTRADO": "(vacío)" if vacio else str(valor),
                "PROBLEMA": "Fecha vacía" if vacio else "No se pudo convertir a fecha",
            })
        df[col] = convertidas
    return df, problemas


# ============================================================================
# 3. FILTRAR
# ============================================================================
def filtrar_ssi_activadas(df, col_empresa, col_estado):
    """Separa las filas SSI; dentro de ellas, las ACTIVADAS y las de otros estados."""
    es_ssi = df[col_empresa].map(normalizar_texto) == normalizar_texto(EMPRESA_VALIDA)
    ssi = df[es_ssi]
    activada = ssi[col_estado].map(normalizar_texto) == normalizar_texto(ESTADO_VALIDO)
    return ssi, ssi[activada].copy(), ssi[~activada]


# ============================================================================
# 4. EVALUAR HORARIO
# ============================================================================
def esta_fuera_de_horario(fecha):
    """True/False según la regla de horario; None si la fecha no es válida."""
    if pd.isna(fecha):
        return None
    if EVALUAR_FINES_DE_SEMANA and fecha.weekday() >= 5:
        return True
    hora = fecha.time()
    return hora < time(HORA_INICIO) or hora >= time(HORA_FIN)


def franja_horaria(fecha) -> str:
    """Franja de una fecha marcada como fuera de horario (para el resumen)."""
    hora = fecha.time()
    if hora < time(HORA_INICIO):
        return f"Antes de las {HORA_INICIO:02d}:00"
    if hora >= time(HORA_FIN):
        return f"Desde las {HORA_FIN:02d}:00"
    return "Fin de semana (dentro del horario)"


def a_si_no(valor) -> str:
    return SIN_FECHA if valor is None else ("Sí" if valor else "No")


def evaluar_horario(df, col_creacion, col_modificacion) -> pd.DataFrame:
    """Agrega las columnas de evaluación y el MOTIVO a cada fila analizada."""
    creacion_fuera = df[col_creacion].map(esta_fuera_de_horario)
    modif_fuera = df[col_modificacion].map(esta_fuera_de_horario)

    df[COL_DIA] = df[col_creacion].map(lambda f: "" if pd.isna(f) else DIAS_SEMANA[f.weekday()])
    df[COL_HORA_CREACION] = df[col_creacion].map(lambda f: "" if pd.isna(f) else f.strftime("%H:%M:%S"))
    df[COL_HORA_MODIFICACION] = df[col_modificacion].map(lambda f: "" if pd.isna(f) else f.strftime("%H:%M:%S"))
    df[COL_CREACION_FUERA] = creacion_fuera.map(a_si_no)
    df[COL_MODIFICACION_FUERA] = modif_fuera.map(a_si_no)

    def motivo(c, m):
        if c is True and m is True:
            return MOTIVO_AMBAS
        if c is True:
            return MOTIVO_CREACION
        if m is True:
            return MOTIVO_MODIFICACION
        return ""

    df[COL_MOTIVO] = [motivo(c, m) for c, m in zip(creacion_fuera, modif_fuera)]
    return df


# ============================================================================
# 5. RESUMEN Y EXPORTAR
# ============================================================================
def armar_resumen(info: dict, otros_estados, fuera, col_creacion, col_modificacion, col_estado):
    """Arma la hoja Resumen como una tabla de 3 columnas (Sección, Concepto, Valor)."""
    filas = [
        ("General", "Archivo de entrada", info["archivo"]),
        ("General", "Hoja usada", info["hoja"]),
        ("General", "Fila de encabezados en Excel", info["fila_encabezado"]),
        ("General", "Columna de estado usada", col_estado),
        ("General", "Horario laboral", f"{HORA_INICIO:02d}:00:00 a {HORA_FIN:02d}:00:00 (sin incluir {HORA_FIN:02d}:00:00)"),
        ("General", "Fines de semana fuera de horario", "Sí" if EVALUAR_FINES_DE_SEMANA else "No (solo se evalúa la hora)"),
        ("Totales", "Total filas leídas", info["leidas"]),
        ("Totales", f"Total filas {EMPRESA_VALIDA}", info["ssi"]),
        ("Totales", f"Total filas {EMPRESA_VALIDA} con estado {ESTADO_VALIDO}", info["activadas"]),
        ("Totales", f"Total filas {EMPRESA_VALIDA} en otros estados (descartadas)", len(otros_estados)),
        ("Totales", "Total fuera de horario", len(fuera)),
        ("Totales", "Filas analizadas con fechas con problemas", info["filas_problema"]),
        ("Totales", f"Tickets duplicados entre las filas {EMPRESA_VALIDA} {ESTADO_VALIDO} (se conservan)", info["duplicados"]),
    ]
    desglose = otros_estados[col_estado].map(lambda v: str(v).strip() if pd.notna(v) and str(v).strip() else "(vacío)")
    for estado, n in desglose.value_counts().items():
        filas.append(("Descartadas por estado", estado, int(n)))
    for motivo in (MOTIVO_CREACION, MOTIVO_MODIFICACION, MOTIVO_AMBAS):
        filas.append(("Fuera de horario por MOTIVO", motivo, int((fuera[COL_MOTIVO] == motivo).sum())))

    franjas = [f"Antes de las {HORA_INICIO:02d}:00", f"Desde las {HORA_FIN:02d}:00"]
    if EVALUAR_FINES_DE_SEMANA:
        franjas.append("Fin de semana (dentro del horario)")
    for col, marca, etiqueta in ((col_creacion, COL_CREACION_FUERA, "Creación"),
                                 (col_modificacion, COL_MODIFICACION_FUERA, "Última modificación")):
        conteo = fuera.loc[fuera[marca] == "Sí", col].map(franja_horaria).value_counts()
        for franja in franjas:
            filas.append((f"Franja horaria - {etiqueta}", franja, int(conteo.get(franja, 0))))
    return pd.DataFrame(filas, columns=["SECCIÓN", "CONCEPTO", "VALOR"])


def dar_formato(hoja):
    """Encabezado en negrita con color, filtros, encabezado congelado, anchos y formato de fecha."""
    relleno = PatternFill("solid", start_color="1F4E78")
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(vertical="center", wrap_text=True)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.iter_cols(min_row=1, max_row=hoja.max_row):
        ancho = 8
        for celda in columna:
            if isinstance(celda.value, datetime):
                celda.number_format = FORMATO_FECHA_EXCEL
                largo = 19
            else:
                largo = len(str(celda.value)) if celda.value is not None else 0
            ancho = max(ancho, min(largo, 60))
        hoja.column_dimensions[get_column_letter(columna[0].column)].width = ancho + 2


def exportar(ruta_salida: Path, fuera: pd.DataFrame, resumen: pd.DataFrame, problemas: pd.DataFrame):
    """Escribe el archivo nuevo. Si está abierto en Excel, avisa en vez de romperse."""
    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as escritor:
            fuera.to_excel(escritor, sheet_name="Fuera de horario", index=False)
            resumen.to_excel(escritor, sheet_name="Resumen", index=False)
            if not problemas.empty:
                problemas.to_excel(escritor, sheet_name="Fechas con problemas", index=False)
            for hoja in escritor.book.worksheets:
                dar_formato(hoja)
    except PermissionError:
        raise ErrorValidacion(
            f"No se pudo guardar el reporte porque el archivo está abierto o bloqueado:\n"
            f"    {ruta_salida}\nCiérralo en Excel y vuelve a ejecutar el script."
        )


# ============================================================================
# PROGRAMA PRINCIPAL
# ============================================================================
def procesar(ruta_entrada: Path, ruta_salida: Path) -> dict:
    if ruta_salida.resolve() == ruta_entrada.resolve():
        raise ErrorValidacion("El archivo de salida no puede ser el mismo archivo de entrada.")

    hojas, hoja, filas = leer_excel(ruta_entrada)
    if len(hojas) > 1:
        print(f"Aviso: el libro tiene {len(hojas)} hojas {hojas}; se usa la primera.")
    print(f"Hoja usada: '{hoja}'")

    df = construir_tabla(filas)
    col_empresa = buscar_columna(df, COLUMNA_EMPRESA)
    col_creacion = buscar_columna(df, COLUMNA_CREACION)
    col_modificacion = buscar_columna(df, COLUMNA_MODIFICACION)
    col_estado = buscar_columna(df, COLUMNA_ESTADO)
    col_ticket = buscar_columna(df, COLUMNA_TICKET, obligatoria=False)

    ssi, activadas, otros_estados = filtrar_ssi_activadas(df, col_empresa, col_estado)
    activadas, problemas = normalizar_fechas(activadas, [col_creacion, col_modificacion])
    activadas = evaluar_horario(activadas, col_creacion, col_modificacion)
    fuera = activadas[activadas[COL_MOTIVO] != ""]

    tabla_problemas = pd.DataFrame(problemas, columns=["FILA EXCEL ORIGINAL", "COLUMNA", "VALOR ENCONTRADO", "PROBLEMA"])
    duplicados = int(activadas[col_ticket].duplicated(keep=False).sum()) if col_ticket else 0
    info = {
        "archivo": str(ruta_entrada), "hoja": hoja, "fila_encabezado": df.attrs["fila_encabezado"],
        "leidas": len(df), "ssi": len(ssi), "activadas": len(activadas),
        "filas_problema": tabla_problemas["FILA EXCEL ORIGINAL"].nunique(), "duplicados": duplicados,
    }
    resumen = armar_resumen(info, otros_estados, fuera, col_creacion, col_modificacion, col_estado)
    exportar(ruta_salida, fuera, resumen, tabla_problemas)
    info.update(fuera=len(fuera), salida=str(ruta_salida), tabla_fuera=fuera,
                tabla_problemas=tabla_problemas, resumen=resumen)
    return info


def main():
    ruta_entrada = Path(sys.argv[1]) if len(sys.argv) > 1 else RUTA_ENTRADA
    ruta_salida = ruta_entrada.parent / NOMBRE_SALIDA
    try:
        info = procesar(ruta_entrada, ruta_salida)
    except ErrorValidacion as error:
        print(f"\nERROR: {error}")
        sys.exit(1)

    print("\n================ RESUMEN ================")
    print(f"Filas leídas:                  {info['leidas']}")
    print(f"Filas {EMPRESA_VALIDA}:                     {info['ssi']}")
    print(f"Filas {EMPRESA_VALIDA} {ESTADO_VALIDO}:            {info['activadas']}")
    print(f"Filas fuera de horario:        {info['fuera']}")
    if info["filas_problema"]:
        print(f"Filas con fechas con problemas: {info['filas_problema']} (ver hoja 'Fechas con problemas')")
    if info["duplicados"]:
        print(f"Aviso: {info['duplicados']} filas comparten TICKET (se conservan todas).")
    print(f"Reporte guardado en:\n    {info['salida']}")


if __name__ == "__main__":
    main()
