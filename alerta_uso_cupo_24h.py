# -*- coding: utf-8 -*-
"""
ALERTA: USO DEL 90 % O MÁS DEL CUPO EN LAS PRIMERAS 24 HORAS DESDE LA ACTIVACIÓN

Cruza las activaciones de crédito (reportes de solicitudes) con las utilizaciones
(consumos) y marca cada activación cuyo consumo ACUMULADO, dentro de las 24 horas
corridas siguientes a la activación, llega al 90 % o más del cupo aprobado.

Pensado para Spyder: cambia solo el bloque CONFIGURACIÓN y ejecuta con F5.
Todo se procesa localmente; no modifica los archivos de entrada.

Librerías: pandas, numpy, openpyxl  ->  pip install pandas numpy openpyxl
"""
import re
import unicodedata
from datetime import timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================================
# CONFIGURACIÓN (cambia solo este bloque)
# ============================================================================
CARPETA = Path(
    r"C:\Users\LCUESTAC\Documents\prueba_automatizacion"
    r"\Validación clientes activados por fuera del horario laboral"
)

# --- Activaciones: uno o varios reportes de solicitudes -----------------------
ARCHIVOS_ACTIVACIONES = sorted(CARPETA.glob("Reporte_*.xlsx"))   # o [CARPETA / "archivo.xlsx"]
ACT_CEDULA = "DOCUMENTO DE IDENTIDAD"
ACT_FECHA = "FECHA FIRMA PAGARÉ"          # inicio de la ventana (fecha y hora de activación)
ACT_CUPO = "VALOR CUPO APROBADO"
ACT_ESTADO = "IDENTIFICADOR"
ACT_ESTADOS_VALIDOS = ["ACTIVADA"]
ACT_EMPRESA = "EMPRESA"                   # None para no filtrar por empresa
ACT_EMPRESAS_VALIDAS = ["SSI"]
ACT_ID = "TICKET"                         # identifica cada activación (None si no existe)

# --- Transacciones / utilizaciones -------------------------------------------
ARCHIVO_TRANSACCIONES = CARPETA / "reporte_1.xlsx"
TX_CEDULA = "Identificación"
TX_FECHA = "Fecha Utilización"
TX_VALOR = "Monto"
TX_ID = "Número de crédito"               # une una reversa con su transacción original
TX_ESTADO = None                          # None = el archivo no trae estado: todas cuentan como aprobadas
TX_ESTADOS_APROBADOS = ["APROBADA"]
TX_ESTADOS_REVERSA = ["REVERSADA", "ANULADA", "REVERSO"]
# Cualquier otro estado (p. ej. RECHAZADA) se excluye del consumo.

# Tratamiento de reversas/anulaciones (DECISIÓN EXPLÍCITA):
#   "excluir": la transacción original que tiene reversa (mismo TX_ID) no se suma,
#              y la reversa tampoco. Es la opción por defecto.
#   "restar" : la reversa se suma como valor NEGATIVO en el momento en que ocurre.
TRATAMIENTO_REVERSAS = "excluir"

# --- Reversiones (consolidado de reversiones; None si no se tiene) -----------
# Regla aplicada (DECISIÓN EXPLÍCITA): valor que cuenta = valor de la compra - valor reversado.
#   * Reversión total  -> la compra queda en 0 (equivale a excluirla).
#   * Reversión parcial -> solo se descuenta lo reversado.
#   * Solo se descuentan las reversiones en estado REV_ESTADOS_APLICADOS; las pendientes se reportan.
# Cruce: cédula + fecha de la transacción (día) + valor total de la transacción. Si ese día el
# cliente tiene dos compras iguales, la reversión se asigna a la más reciente (caso típico de cobro doble).
ARCHIVO_REVERSIONES = CARPETA / "0.CONSOLIDADO REVERSIONES.xlsx"
HOJAS_REVERSIONES = ["AGIL"]              # AGIL = CrediSomos Ágil, la línea de las utilizaciones
REV_CEDULA = "Cédula Cliente"
REV_FECHA_TX = "Fecha Transacción"        # texto se lee como día/mes/año
REV_FECHA_TX_ALTERNA = "Fecha organizada" # aaaammdd; segundo intento si la anterior no cruza (None para no usar)
REV_VALOR_TX = "Valor Total TX"
REV_VALOR = "Valor Reversión"
REV_ESTADO = "Estado Reversión (Aplicada/Pendiente)"
REV_ESTADOS_APLICADOS = ["Aplicada"]
REV_TIPIFICACION = "Tipificación"
REV_FECHA_SOLICITUD = "Fecha Ingreso Solicitud"
# Tipificaciones que NO indican un retracto inocente (coincidencia parcial, sin importar tildes).
# Son solo informativas: se muestran en las hojas Alertas y Reversiones.
TIPIFICACIONES_DE_RIESGO = [
    "POSIBLE FRAUDE", "SUPLANTACION", "VALIDACION DE IDENTIDAD", "BIOMETRIA", "VINCULADO",
    "CEDULA EN MAL ESTADO", "CEDULA MAL ESTADO", "INCONSISTENCIA EN INFORMACION",
    "NO REALIZO LA ENTREGA", "NO RECIBE PRODUCTO", "NO ADQUIERE PRODUCTO", "ZONA NO AUTORIZADA",
]

# --- Reglas -------------------------------------------------------------------
UMBRAL = 0.90          # 90 % del cupo
QUITAR_CEROS_EN_CRUCE = True   # True: "00012345" y "12345" se consideran la misma cédula al cruzar
                               # (la cédula se muestra siempre tal como viene en el archivo)
HORAS_VENTANA = 24     # 24 horas corridas; el límite 24:00:00 exacto SÍ se incluye

# --- Salida -------------------------------------------------------------------
RUTA_SALIDA = CARPETA / "Alertas uso cupo 24h.xlsx"
FILAS_BUSQUEDA_ENCABEZADO = 30
# ============================================================================


class ErrorDatos(Exception):
    """Error con mensaje claro para la consola."""


# ============================================================================
# UTILIDADES DE LIMPIEZA
# ============================================================================
def normalizar_texto(valor) -> str:
    """Mayúsculas, sin tildes y sin espacios extra (incluye el espacio duro)."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return ""
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).upper()


def limpiar_llave(serie: pd.Series) -> pd.Series:
    """
    Cédulas y llaves como texto: sin espacios, puntos, comas ni guiones, sin '.0'
    de Excel y conservando ceros a la izquierda.
    """
    texto = serie.astype("string").str.strip()
    texto = texto.str.replace(r"\.0+$", "", regex=True)          # 39209430.0 -> 39209430
    texto = texto.str.replace(r"[\s.,\-]", "", regex=True)       # 39.209.430 -> 39209430
    return texto.replace("", pd.NA)


def _texto_a_numero(texto: str) -> float:
    """Convierte un monto en formato colombiano ('$ 1.234.567,89') a número."""
    s = re.sub(r"[^\d,.\-]", "", texto)
    if not s or s in "-.,":
        return np.nan
    if "," in s and "." in s:                 # 1.234.567,89 -> punto = miles, coma = decimal
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:                            # 1,5 decimal  |  1,234,567 miles
        s = s.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):   # 1.234.567 -> miles
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return np.nan


def limpiar_monto(serie: pd.Series) -> pd.Series:
    """Montos a número. Los numéricos se respetan; el texto se convierte una vez por valor distinto."""
    numeros = pd.to_numeric(serie, errors="coerce")
    es_texto = serie.map(lambda v: isinstance(v, str))
    if es_texto.any():
        unicos = serie[es_texto].unique()
        mapa = {u: _texto_a_numero(u) for u in unicos}
        numeros[es_texto] = serie[es_texto].map(mapa)
    return numeros.astype(float)


def limpiar_fecha(serie: pd.Series) -> pd.Series:
    """
    Fechas a datetime. Fecha real de Excel se respeta; el texto ISO (aaaa-mm-dd) se
    lee como tal y cualquier otro texto como día/mes/año (formato colombiano).
    Se quita la zona horaria si viniera y se redondea al segundo.
    """
    resultado = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    es_texto = serie.map(lambda v: isinstance(v, str))
    no_texto = ~es_texto & serie.notna()
    if no_texto.any():
        resultado[no_texto] = pd.to_datetime(serie[no_texto], errors="coerce")
    if es_texto.any():
        texto = serie[es_texto].str.strip()
        iso = texto.str.match(r"^\d{4}-\d{2}-\d{2}")
        resultado[es_texto & iso.reindex(serie.index, fill_value=False)] = pd.to_datetime(
            texto[iso], errors="coerce", format="ISO8601")
        resultado[es_texto & ~iso.reindex(serie.index, fill_value=True)] = pd.to_datetime(
            texto[~iso], errors="coerce", dayfirst=True)
    return resultado.dt.round("s")


# ============================================================================
# 1. CARGAR
# ============================================================================
def leer_tabla(ruta: Path, columnas: dict, hoja=0, opcionales=()) -> pd.DataFrame:
    """
    Lee la primera hoja, detecta la fila de encabezados y trae solo las columnas
    pedidas (buscadas sin importar tildes/mayúsculas/espacios). Si falta alguna,
    se detiene con un mensaje claro. 'columnas' = {nombre interno: nombre en Excel}.
    Las columnas cuyo nombre interno esté en 'opcionales' pueden faltar.
    """
    if not Path(ruta).exists():
        raise ErrorDatos(f"No se encontró el archivo:\n    {ruta}")
    buscadas = {k: v for k, v in columnas.items() if v}
    objetivos = {normalizar_texto(v) for v in buscadas.values()}

    previa = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=FILAS_BUSQUEDA_ENCABEZADO, dtype=object)
    puntajes = previa.apply(lambda f: len(objetivos & {normalizar_texto(v) for v in f}), axis=1)
    fila_enc = int(puntajes.idxmax()) if len(puntajes) else 0

    df = pd.read_excel(ruta, sheet_name=hoja, header=fila_enc, dtype=object)
    por_nombre = {normalizar_texto(c): c for c in df.columns}
    buscadas = {k: v for k, v in buscadas.items() if k not in opcionales or normalizar_texto(v) in por_nombre}
    faltan = [v for v in buscadas.values() if normalizar_texto(v) not in por_nombre]
    if faltan:
        encontradas = "\n".join(f"    - {c}" for c in df.columns)
        donde = f"'{Path(ruta).name}'" + (f", hoja '{hoja}'" if isinstance(hoja, str) else "")
        raise ErrorDatos(f"En {donde} no se encontró la(s) columna(s): {faltan}\n"
                         f"Columnas encontradas:\n{encontradas}")
    salida = pd.DataFrame({k: df[por_nombre[normalizar_texto(v)]] for k, v in buscadas.items()})
    salida.insert(0, "ARCHIVO", Path(ruta).name + (f" [{hoja}]" if isinstance(hoja, str) else ""))
    salida.insert(1, "FILA EXCEL", np.arange(len(df)) + fila_enc + 2)
    return salida


def cargar() -> tuple:
    """Carga activaciones (uno o varios archivos) y transacciones."""
    if not ARCHIVOS_ACTIVACIONES:
        raise ErrorDatos("No hay archivos de activaciones (revisa ARCHIVOS_ACTIVACIONES).")
    col_act = {"cedula": ACT_CEDULA, "fecha": ACT_FECHA, "cupo": ACT_CUPO, "estado": ACT_ESTADO,
               "empresa": ACT_EMPRESA, "id_activacion": ACT_ID}
    col_tx = {"cedula": TX_CEDULA, "fecha": TX_FECHA, "valor": TX_VALOR, "id_tx": TX_ID, "estado": TX_ESTADO}
    print("Leyendo activaciones...")
    partes = []
    for ruta in ARCHIVOS_ACTIVACIONES:
        parte = leer_tabla(ruta, col_act)
        print(f"  - {Path(ruta).name}: {len(parte)} filas")
        partes.append(parte)
    act = pd.concat(partes, ignore_index=True)
    print("Leyendo transacciones...")
    tx = leer_tabla(ARCHIVO_TRANSACCIONES, col_tx)
    print(f"  - {Path(ARCHIVO_TRANSACCIONES).name}: {len(tx)} filas")

    rev = None
    if ARCHIVO_REVERSIONES:
        print("Leyendo reversiones...")
        col_rev = {"cedula": REV_CEDULA, "fecha": REV_FECHA_TX, "fecha_alt": REV_FECHA_TX_ALTERNA,
                   "valor": REV_VALOR_TX, "valor_rev": REV_VALOR, "estado": REV_ESTADO,
                   "tipificacion": REV_TIPIFICACION, "fecha_solicitud": REV_FECHA_SOLICITUD}
        partes = []
        for hoja in HOJAS_REVERSIONES:
            parte = leer_tabla(ARCHIVO_REVERSIONES, col_rev, hoja=hoja,
                               opcionales=("fecha_alt", "estado", "tipificacion", "fecha_solicitud", "valor_rev"))
            print(f"  - {Path(ARCHIVO_REVERSIONES).name} [{hoja}]: {len(parte)} filas")
            partes.append(parte)
        rev = pd.concat(partes, ignore_index=True)
    return act, tx, rev


# ============================================================================
# 2. LIMPIAR
# ============================================================================
def llave_cruce(cedula: pd.Series) -> pd.Series:
    """Llave para cruzar por cédula (sin ceros a la izquierda si así se configuró)."""
    return cedula.str.lstrip("0").replace("", "0") if QUITAR_CEROS_EN_CRUCE else cedula


def marcar(df: pd.DataFrame, mascara, tipo: str, detalle_col: str = None) -> pd.DataFrame:
    """Devuelve las filas de 'df' que cumplen 'mascara' con el TIPO de excepción (y DETALLE)."""
    filas = df[mascara].copy()
    filas["TIPO"] = tipo
    if detalle_col:
        filas["DETALLE"] = filas[detalle_col].astype(str)
    return filas


def quitar_duplicados(df: pd.DataFrame, nombre: str, excepciones: list) -> pd.DataFrame:
    """Elimina duplicados exactos (mismos datos, sin mirar archivo ni fila) y los reporta."""
    datos = [c for c in df.columns if c not in ("ARCHIVO", "FILA EXCEL")]
    dup = df.duplicated(subset=datos, keep="first")
    if dup.any():
        excepciones.append(marcar(df, dup, f"Duplicado exacto eliminado ({nombre})"))
    return df[~dup].copy()


def limpiar(act: pd.DataFrame, tx: pd.DataFrame, rev) -> tuple:
    excepciones, conteos = [], {"act_leidas": len(act), "tx_leidas": len(tx)}

    # --- Activaciones ---
    act = quitar_duplicados(act, "activaciones", excepciones)
    conteos["act_duplicadas"] = conteos["act_leidas"] - len(act)
    valido = act["estado"].map(normalizar_texto).isin({normalizar_texto(e) for e in ACT_ESTADOS_VALIDOS})
    if ACT_EMPRESA:
        valido &= act["empresa"].map(normalizar_texto).isin({normalizar_texto(e) for e in ACT_EMPRESAS_VALIDAS})
    act = act[valido].copy()
    conteos["act_validas"] = len(act)
    act["cedula"] = limpiar_llave(act["cedula"])
    act["fecha_original"] = act["fecha"]
    act["fecha"] = limpiar_fecha(act["fecha"])
    act["cupo"] = limpiar_monto(act["cupo"])
    if "id_activacion" not in act:
        act["id_activacion"] = pd.NA
    act["id_activacion"] = act["id_activacion"].astype("string").fillna(
        "FILA " + act["FILA EXCEL"].astype(str) + " " + act["ARCHIVO"])

    malas = act["cedula"].isna() | act["fecha"].isna()
    excepciones.append(marcar(act, act["cedula"].isna(), "Activación sin cédula"))
    excepciones.append(marcar(act, act["fecha"].isna() & act["cedula"].notna(),
                              "Activación con fecha inválida o vacía", "fecha_original"))
    sin_cupo = ~malas & (act["cupo"].isna() | (act["cupo"] <= 0))
    excepciones.append(marcar(act, sin_cupo, "Cupo aprobado cero, nulo o no numérico"))
    act = act[~malas & ~sin_cupo].copy()

    # --- Transacciones ---
    tx = quitar_duplicados(tx, "transacciones", excepciones)
    conteos["tx_duplicadas"] = conteos["tx_leidas"] - len(tx)
    tx["cedula"] = limpiar_llave(tx["cedula"])
    tx["fecha_original"] = tx["fecha"]
    tx["fecha"] = limpiar_fecha(tx["fecha"])
    tx["valor"] = limpiar_monto(tx["valor"])
    tx["id_tx"] = limpiar_llave(tx["id_tx"]) if "id_tx" in tx else pd.NA

    malas = tx["cedula"].isna() | tx["fecha"].isna() | tx["valor"].isna()
    excepciones.append(marcar(tx, tx["fecha"].isna(), "Transacción con fecha inválida o vacía", "fecha_original"))
    excepciones.append(marcar(tx, tx["valor"].isna() & tx["fecha"].notna(), "Transacción con valor no numérico"))
    excepciones.append(marcar(tx, tx["cedula"].isna() & tx["fecha"].notna() & tx["valor"].notna(),
                              "Transacción sin cédula"))
    tx = tx[~malas].copy()
    tx = aplicar_estados(tx, excepciones, conteos)

    # --- Reversiones del consolidado ---
    act["llave"], tx["llave"] = llave_cruce(act["cedula"]), llave_cruce(tx["cedula"])
    tx["valor_reversado"], tx["tipificacion_rev"], tx["riesgo_rev"], tx["fecha_solicitud_rev"] = 0.0, "", False, pd.NaT
    conteos.update(rev_leidas=0, rev_fuera_periodo=0, rev_pendientes=0, rev_aplicadas=0,
                   rev_cruzadas=0, rev_sin_cruce=0, rev_riesgo=0)
    if rev is not None:
        tx = aplicar_reversiones(tx, rev, excepciones, conteos)
    return act, tx, excepciones, conteos


def es_tipificacion_de_riesgo(serie: pd.Series) -> pd.Series:
    terminos = [normalizar_texto(t) for t in TIPIFICACIONES_DE_RIESGO]
    return serie.map(normalizar_texto).map(lambda t: any(x in t for x in terminos) if t else False)


def emparejar(tx: pd.DataFrame, rev: pd.DataFrame, col_fecha: str) -> pd.DataFrame:
    """
    Empareja reversiones con utilizaciones 1 a 1 por cédula + día + valor total.
    Si hay varias utilizaciones iguales ese día, la primera reversión va a la más reciente.
    Devuelve pares (_rev, _tx).
    """
    t = tx[["llave", "fecha", "valor"]].assign(dia=tx["fecha"].dt.normalize(), _tx=tx.index)
    t["orden"] = t.sort_values("fecha", ascending=False).groupby(["llave", "dia", "valor"]).cumcount()
    r = rev.dropna(subset=[col_fecha])[["llave", col_fecha, "valor", "fecha_solicitud", "FILA EXCEL"]].copy()
    r["_rev"] = r.index
    r["orden"] = r.sort_values(["fecha_solicitud", "FILA EXCEL"]).groupby(["llave", col_fecha, "valor"]).cumcount()
    pares = r.merge(t, left_on=["llave", col_fecha, "valor", "orden"],
                    right_on=["llave", "dia", "valor", "orden"], how="inner")
    return pares[["_rev", "_tx"]]


def aplicar_reversiones(tx: pd.DataFrame, rev: pd.DataFrame, excepciones: list, conteos: dict) -> pd.DataFrame:
    """Descuenta del valor de cada utilización lo reversado en el consolidado (ver regla arriba)."""
    for col in ("fecha_alt", "estado", "tipificacion", "fecha_solicitud", "valor_rev"):
        if col not in rev:
            rev[col] = pd.NA
    # Filas vacías de la plantilla (sin cédula ni valor; la fecha trae fórmulas como "//")
    rev = rev[rev[["cedula", "valor"]].notna().any(axis=1)]
    conteos["rev_leidas"] = len(rev)
    rev = quitar_duplicados(rev, "reversiones", excepciones)
    rev["cedula"] = limpiar_llave(rev["cedula"])
    rev["llave"] = llave_cruce(rev["cedula"])
    rev["fecha_original"] = rev["fecha"]
    rev["fecha"] = limpiar_fecha(rev["fecha"]).dt.normalize()
    rev["fecha_alt"] = pd.to_datetime(rev["fecha_alt"].astype("string").str[:8], format="%Y%m%d", errors="coerce")
    rev["valor"] = limpiar_monto(rev["valor"])
    rev["valor_rev"] = limpiar_monto(rev["valor_rev"]).fillna(rev["valor"])
    rev["fecha_solicitud"] = limpiar_fecha(rev["fecha_solicitud"])
    rev["riesgo"] = es_tipificacion_de_riesgo(rev["tipificacion"])

    # Solo el periodo de las utilizaciones (las reversiones de años anteriores no aplican)
    desde, hasta = tx["fecha"].min().normalize(), tx["fecha"].max().normalize()
    en_periodo = rev["fecha"].between(desde, hasta) | rev["fecha_alt"].between(desde, hasta)
    invalida = rev["fecha"].isna() & rev["fecha_alt"].isna()
    excepciones.append(marcar(rev, invalida, "Reversión con fecha inválida o vacía", "fecha_original"))
    excepciones.append(marcar(rev, en_periodo & (rev["llave"].isna() | rev["valor"].isna()),
                              "Reversión sin cédula o sin valor"))
    conteos["rev_fuera_periodo"] = int((~en_periodo & ~invalida).sum())
    rev = rev[en_periodo & rev["llave"].notna() & rev["valor"].notna()]

    aplicada = rev["estado"].map(normalizar_texto).isin({normalizar_texto(e) for e in REV_ESTADOS_APLICADOS})
    if REV_ESTADO and rev["estado"].notna().any():
        excepciones.append(marcar(rev, ~aplicada, "Reversión no aplicada (no se descuenta)", "estado"))
        conteos["rev_pendientes"] = int((~aplicada).sum())
        rev = rev[aplicada]
    conteos["rev_aplicadas"] = len(rev)

    pares = emparejar(tx, rev, "fecha")
    if REV_FECHA_TX_ALTERNA:
        resto_rev = rev[~rev.index.isin(pares["_rev"])]
        resto_tx = tx[~tx.index.isin(pares["_tx"])]
        pares = pd.concat([pares, emparejar(resto_tx, resto_rev, "fecha_alt")], ignore_index=True)
    excepciones.append(marcar(rev, ~rev.index.isin(pares["_rev"]), "Reversión sin utilización (no cruza)"))

    datos = rev.loc[pares["_rev"], ["valor_rev", "tipificacion", "riesgo", "fecha_solicitud"]].set_index(pares["_tx"].values)
    tx.loc[datos.index, "valor_reversado"] = datos["valor_rev"].astype(float)
    tx.loc[datos.index, "tipificacion_rev"] = datos["tipificacion"].astype(str).str.strip()
    tx.loc[datos.index, "riesgo_rev"] = datos["riesgo"].astype(bool)
    tx.loc[datos.index, "fecha_solicitud_rev"] = datos["fecha_solicitud"]
    tx["valor_computable"] = (tx["valor_computable"] - tx["valor_reversado"]).clip(lower=0)
    conteos["rev_cruzadas"] = len(pares)
    conteos["rev_sin_cruce"] = conteos["rev_aplicadas"] - len(pares)
    conteos["rev_riesgo"] = int(datos["riesgo"].sum())
    return tx


def aplicar_estados(tx: pd.DataFrame, excepciones: list, conteos: dict) -> pd.DataFrame:
    """Aplica estado aprobado/rechazado y el tratamiento de reversas definido arriba."""
    conteos.update(tx_rechazadas=0, tx_reversas=0, tx_originales_reversadas=0)
    if not TX_ESTADO or "estado" not in tx:
        tx["valor_computable"] = tx["valor"]          # sin estado: todas son aprobadas
        return tx

    estado = tx["estado"].map(normalizar_texto)
    es_reversa = estado.isin({normalizar_texto(e) for e in TX_ESTADOS_REVERSA})
    es_aprobada = estado.isin({normalizar_texto(e) for e in TX_ESTADOS_APROBADOS})
    rechazada = ~es_reversa & ~es_aprobada
    conteos["tx_rechazadas"] = int(rechazada.sum())
    conteos["tx_reversas"] = int(es_reversa.sum())
    if rechazada.any():
        excepciones.append(marcar(tx, rechazada, "Transacción excluida por estado (no aprobada)", "estado"))

    if TRATAMIENTO_REVERSAS == "excluir":
        ids_reversados = set(tx.loc[es_reversa, "id_tx"].dropna())
        original = es_aprobada & tx["id_tx"].isin(ids_reversados)
        conteos["tx_originales_reversadas"] = int(original.sum())
        if (original | es_reversa).any():
            excepciones.append(marcar(tx, original | es_reversa,
                                      "Transacción excluida por reversa (original y reversa)"))
        tx = tx[es_aprobada & ~original].copy()
        tx["valor_computable"] = tx["valor"]
    elif TRATAMIENTO_REVERSAS == "restar":
        tx = tx[es_aprobada | es_reversa].copy()
        tx["valor_computable"] = np.where(es_reversa[tx.index], -tx["valor"].abs(), tx["valor"])
    else:
        raise ErrorDatos("TRATAMIENTO_REVERSAS debe ser 'excluir' o 'restar'.")
    return tx


# ============================================================================
# 3. CRUZAR
# ============================================================================
def cruzar(act: pd.DataFrame, tx: pd.DataFrame, excepciones: list, conteos: dict) -> pd.DataFrame:
    """
    Une por cédula y deja solo las transacciones dentro de la ventana de cada
    activación: activación <= transacción <= activación + 24 h (límite incluido).
    Cada activación se evalúa por separado aunque el cliente tenga varias.
    """
    ced_act, ced_tx = set(act["llave"]), set(tx["llave"])
    excepciones.append(marcar(act, ~act["llave"].isin(ced_tx), "Activación sin transacciones (no cruza)"))
    excepciones.append(marcar(tx, ~tx["llave"].isin(ced_act), "Transacción sin activación (no cruza)"))
    conteos["act_sin_tx"] = int((~act["llave"].isin(ced_tx)).sum())
    conteos["tx_sin_act"] = int((~tx["llave"].isin(ced_act)).sum())

    a = act[["id_activacion", "llave", "cedula", "fecha", "cupo", "ARCHIVO", "FILA EXCEL"]].rename(
        columns={"fecha": "fecha_activacion", "ARCHIVO": "archivo_act", "FILA EXCEL": "fila_act"})
    t = tx[["llave", "cedula", "fecha", "valor", "valor_computable", "id_tx", "valor_reversado",
            "tipificacion_rev", "riesgo_rev", "fecha_solicitud_rev", "ARCHIVO", "FILA EXCEL"]].rename(
        columns={"cedula": "cedula_tx", "fecha": "fecha_tx", "ARCHIVO": "archivo_tx", "FILA EXCEL": "fila_tx"})
    cruce = a.merge(t, on="llave", how="inner")
    fin = cruce["fecha_activacion"] + timedelta(hours=HORAS_VENTANA)
    en_ventana = (cruce["fecha_tx"] >= cruce["fecha_activacion"]) & (cruce["fecha_tx"] <= fin)
    conteos["tx_antes_activacion"] = int((cruce["fecha_tx"] < cruce["fecha_activacion"]).sum())
    return cruce[en_ventana].copy()


# ============================================================================
# 4. CALCULAR
# ============================================================================
def calcular(act: pd.DataFrame, ventana: pd.DataFrame) -> tuple:
    """
    Acumulado por activación en orden cronológico (vectorizado con groupby + cumsum).
    Se marca la activación si en algún momento el acumulado llega al UMBRAL; se
    guarda la primera transacción con la que se cruza.
    """
    v = ventana.sort_values(["id_activacion", "fecha_tx", "id_tx"]).copy()
    g = v.groupby("id_activacion", sort=False)
    v["N° TRANSACCIÓN EN VENTANA"] = g.cumcount() + 1
    v["ACUMULADO"] = g["valor_computable"].cumsum()
    v["% ACUMULADO"] = v["ACUMULADO"] / v["cupo"]
    v["CRUZA UMBRAL"] = v["% ACUMULADO"] >= UMBRAL
    v["% BRUTO"] = g["valor"].cumsum() / v["cupo"]                 # sin descontar reversiones
    v["tiene_rev"] = v["valor_reversado"] > 0
    v["rev_riesgo"] = v["tiene_rev"] & v["riesgo_rev"].astype(bool)

    por_act = g.agg(consumo=("valor_computable", "sum"), n_tx=("valor_computable", "size"),
                    max_pct=("% ACUMULADO", "max"), max_pct_bruto=("% BRUTO", "max"),
                    valor_rev=("valor_reversado", "sum"), n_rev=("tiene_rev", "sum"), riesgo=("rev_riesgo", "any"))
    tipif = (v[v["tiene_rev"]].groupby("id_activacion")["tipificacion_rev"]
             .agg(lambda s: " | ".join(sorted(set(s)))))
    por_act = por_act.join(tipif.rename("tipificaciones"))
    cruces = v[v["CRUZA UMBRAL"]].groupby("id_activacion").head(1).set_index("id_activacion")

    base = act.set_index("id_activacion")[["cedula", "fecha", "cupo", "ARCHIVO", "FILA EXCEL"]]
    evaluacion = base.join(por_act, how="left").fillna(
        {"consumo": 0, "n_tx": 0, "max_pct": 0, "max_pct_bruto": 0, "valor_rev": 0, "n_rev": 0,
         "riesgo": False, "tipificaciones": ""})
    evaluacion["pct"] = evaluacion["consumo"] / evaluacion["cupo"]
    evaluacion["marcado"] = evaluacion["max_pct"] >= UMBRAL
    evaluacion["marcado_sin_reversiones"] = evaluacion["max_pct_bruto"] >= UMBRAL

    marc = evaluacion[evaluacion["marcado"]].join(
        cruces[["fecha_tx", "N° TRANSACCIÓN EN VENTANA"]], how="left")
    horas = (marc["fecha_tx"] - marc["fecha"]).dt.total_seconds() / 3600
    alertas = pd.DataFrame({
        "CÉDULA": marc["cedula"],
        "ID ACTIVACIÓN": marc.index,
        "FECHA Y HORA ACTIVACIÓN": marc["fecha"],
        "CUPO APROBADO": marc["cupo"],
        f"CONSUMO EN {HORAS_VENTANA} HORAS": marc["consumo"],
        "% DE USO": marc["pct"],
        "% MÁXIMO ALCANZADO": marc["max_pct"],
        f"FECHA Y HORA EN QUE SE ALCANZÓ EL {UMBRAL:.0%}": marc["fecha_tx"],
        "HORAS DESDE ACTIVACIÓN HASTA EL CRUCE": horas.round(2),
        "TRANSACCIONES HASTA EL CRUCE": marc["N° TRANSACCIÓN EN VENTANA"].astype(int),
        "CANTIDAD DE TRANSACCIONES EN VENTANA": marc["n_tx"].astype(int),
        "RANGO DE USO": pd.cut(marc["max_pct"] * 100, [-np.inf, 95, 100, np.inf], right=False,
                               labels=["90 a 94", "95 a 99", "100 o más"]).astype(str),
        "REVERSIONES EN VENTANA": marc["n_rev"].astype(int),
        "VALOR REVERSADO EN VENTANA": marc["valor_rev"],
        "TIPIFICACIÓN DE REVERSIONES": marc["tipificaciones"],
        "REVERSIÓN CON TIPIFICACIÓN DE RIESGO": np.where(marc["riesgo"].astype(bool), "Sí", "No"),
        "ARCHIVO ACTIVACIÓN": marc["ARCHIVO"], "FILA ACTIVACIÓN": marc["FILA EXCEL"],
    }).sort_values("FECHA Y HORA ACTIVACIÓN").reset_index(drop=True)

    det = v[v["id_activacion"].isin(marc.index)]
    detalle = pd.DataFrame({
        "CÉDULA": det["cedula"], "ID ACTIVACIÓN": det["id_activacion"],
        "FECHA Y HORA ACTIVACIÓN": det["fecha_activacion"], "CUPO APROBADO": det["cupo"],
        "N° TRANSACCIÓN EN VENTANA": det["N° TRANSACCIÓN EN VENTANA"],
        "FECHA Y HORA TRANSACCIÓN": det["fecha_tx"], "ID TRANSACCIÓN": det["id_tx"],
        "VALOR ORIGINAL": det["valor"], "VALOR REVERSADO": det["valor_reversado"],
        "VALOR": det["valor_computable"], "ACUMULADO": det["ACUMULADO"], "% ACUMULADO": det["% ACUMULADO"],
        "TIPIFICACIÓN REVERSIÓN": det["tipificacion_rev"],
        "HORAS DESDE ACTIVACIÓN": ((det["fecha_tx"] - det["fecha_activacion"]).dt.total_seconds() / 3600).round(2),
        "CRUZA UMBRAL": np.where(det["CRUZA UMBRAL"], "Sí", "No"),
        "ARCHIVO TRANSACCIÓN": det["archivo_tx"], "FILA TRANSACCIÓN": det["fila_tx"],
    }).sort_values(["FECHA Y HORA ACTIVACIÓN", "ID ACTIVACIÓN", "FECHA Y HORA TRANSACCIÓN"])

    # Hoja Reversiones: toda utilización reversada dentro de la ventana de alguna activación
    r = v[v["tiene_rev"]].join(evaluacion[["max_pct", "max_pct_bruto", "marcado", "marcado_sin_reversiones"]],
                               on="id_activacion")
    resultado = np.select([r["marcado"], r["marcado_sin_reversiones"]],
                          ["Sigue en alerta", "La alerta se cae por la reversión"], "No estaba en alerta")
    reversiones = pd.DataFrame({
        "CÉDULA": r["cedula"], "ID ACTIVACIÓN": r["id_activacion"], "FECHA Y HORA ACTIVACIÓN": r["fecha_activacion"],
        "CUPO APROBADO": r["cupo"], "FECHA Y HORA TRANSACCIÓN": r["fecha_tx"], "ID TRANSACCIÓN": r["id_tx"],
        "VALOR ORIGINAL": r["valor"], "VALOR REVERSADO": r["valor_reversado"], "VALOR NETO": r["valor_computable"],
        "TIPIFICACIÓN": r["tipificacion_rev"],
        "TIPIFICACIÓN DE RIESGO": np.where(r["riesgo_rev"].astype(bool), "Sí", "No"),
        "FECHA SOLICITUD REVERSIÓN": r["fecha_solicitud_rev"],
        "% MÁX. SIN REVERSIONES": r["max_pct_bruto"], "% MÁX. CON REVERSIONES": r["max_pct"],
        "RESULTADO": resultado,
    }).sort_values(["FECHA Y HORA ACTIVACIÓN", "FECHA Y HORA TRANSACCIÓN"])
    return evaluacion, alertas, detalle, reversiones


def armar_resumen(conteos: dict, evaluacion: pd.DataFrame, alertas: pd.DataFrame,
                  reversiones: pd.DataFrame, excepciones: pd.DataFrame) -> pd.DataFrame:
    evaluadas = len(evaluacion)
    con_uso = int((evaluacion["n_tx"] > 0).sum())
    marcadas = len(alertas)
    filas = [
        ("Parámetros", "Umbral", f"{UMBRAL:.0%}"),
        ("Parámetros", "Ventana", f"{HORAS_VENTANA} horas corridas desde la activación (límite incluido)"),
        ("Parámetros", "Fecha de activación usada", ACT_FECHA),
        ("Parámetros", "Tratamiento de reversas", TRATAMIENTO_REVERSAS if TX_ESTADO else "No aplica (sin columna de estado)"),
        ("Datos", "Activaciones leídas (todas las filas)", conteos["act_leidas"]),
        ("Datos", "Duplicados exactos eliminados (activaciones)", conteos["act_duplicadas"]),
        ("Datos", "Activaciones válidas (estado y empresa)", conteos["act_validas"]),
        ("Datos", "Transacciones leídas", conteos["tx_leidas"]),
        ("Datos", "Duplicados exactos eliminados (transacciones)", conteos["tx_duplicadas"]),
        ("Datos", "Transacciones excluidas por estado", conteos["tx_rechazadas"]),
        ("Datos", "Reversas / originales reversadas", f"{conteos['tx_reversas']} / {conteos['tx_originales_reversadas']}"),
        ("Cruce", "Activaciones sin transacciones", conteos["act_sin_tx"]),
        ("Cruce", "Transacciones sin activación", conteos["tx_sin_act"]),
        ("Resultado", "Clientes (activaciones) evaluados", evaluadas),
        ("Resultado", f"Con al menos una transacción en las {HORAS_VENTANA} horas", con_uso),
        ("Resultado", "Total marcados", marcadas),
        ("Resultado", "% marcado sobre evaluados", marcadas / evaluadas if evaluadas else 0),
        ("Resultado", f"% marcado sobre los que usaron el cupo en {HORAS_VENTANA} h", marcadas / con_uso if con_uso else 0),
    ]
    for rango in ["90 a 94", "95 a 99", "100 o más"]:
        filas.append(("Distribución por % de uso", rango, int((alertas["RANGO DE USO"] == rango).sum())))
    if ARCHIVO_REVERSIONES:
        se_caen = evaluacion["marcado_sin_reversiones"] & ~evaluacion["marcado"]
        filas += [
            ("Reversiones", "Hojas usadas", ", ".join(HOJAS_REVERSIONES)),
            ("Reversiones", "Regla", "Valor que cuenta = valor de la compra - valor reversado (solo aplicadas)"),
            ("Reversiones", "Filas con datos leídas", conteos["rev_leidas"]),
            ("Reversiones", "Fuera del periodo de las utilizaciones", conteos["rev_fuera_periodo"]),
            ("Reversiones", "En el periodo y no aplicadas (no se descuentan)", conteos["rev_pendientes"]),
            ("Reversiones", "Aplicadas en el periodo", conteos["rev_aplicadas"]),
            ("Reversiones", "Cruzadas con una utilización", conteos["rev_cruzadas"]),
            ("Reversiones", "Sin utilización (no cruzan)", conteos["rev_sin_cruce"]),
            ("Reversiones", "Cruzadas con tipificación de riesgo", conteos["rev_riesgo"]),
            ("Reversiones", "Marcadas si NO se descontaran las reversiones", int(evaluacion["marcado_sin_reversiones"].sum())),
            ("Reversiones", "Alertas que se caen por la reversión", int(se_caen.sum())),
            ("Reversiones", "   de ellas con tipificación de riesgo (revisar)", int((se_caen & evaluacion["riesgo"].astype(bool)).sum())),
            ("Reversiones", "Alertas vigentes con alguna reversión en la ventana", int((alertas["REVERSIONES EN VENTANA"] > 0).sum())),
        ]
    if not excepciones.empty:
        for tipo, n in excepciones["TIPO"].value_counts().items():
            filas.append(("Excepciones", tipo, int(n)))
    return pd.DataFrame(filas, columns=["SECCIÓN", "CONCEPTO", "VALOR"])


# ============================================================================
# 5. EXPORTAR
# ============================================================================
def preparar_excepciones(lista: list) -> pd.DataFrame:
    columnas = ["TIPO", "ARCHIVO", "FILA EXCEL", "cedula", "fecha", "cupo", "valor", "valor_rev", "id_tx",
                "estado", "tipificacion", "DETALLE"]
    partes = [p for p in lista if p is not None and not p.empty]
    if not partes:
        return pd.DataFrame(columns=columnas)
    exc = pd.concat(partes, ignore_index=True)
    exc = exc[[c for c in columnas if c in exc.columns]]
    return exc.rename(columns={"cedula": "CÉDULA", "fecha": "FECHA", "cupo": "CUPO", "valor": "VALOR",
                               "valor_rev": "VALOR REVERSIÓN", "id_tx": "ID TRANSACCIÓN", "estado": "ESTADO",
                               "tipificacion": "TIPIFICACIÓN"})


def dar_formato(hoja):
    """Encabezado con color, filtros, panel congelado, anchos, % y separador de miles."""
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", start_color="1F4E78")
        celda.alignment = Alignment(vertical="center", wrap_text=True)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.iter_cols(min_row=1, max_row=hoja.max_row):
        titulo = str(columna[0].value or "").upper()
        es_pct = "%" in titulo
        es_monto = any(p in titulo for p in ("CUPO", "CONSUMO", "VALOR", "ACUMULADO")) and not es_pct
        ancho = len(titulo)
        for celda in columna[1:]:
            v = celda.value
            if hasattr(v, "year") and hasattr(v, "hour"):
                celda.number_format = "dd/mm/yyyy hh:mm:ss"
                ancho = max(ancho, 19)
                continue
            if isinstance(v, (int, float)):
                if es_pct or (titulo == "VALOR" and hoja.title == "Resumen" and isinstance(v, float) and v < 1):
                    celda.number_format = "0.00%"
                elif es_monto:
                    celda.number_format = "#,##0"
            ancho = max(ancho, min(len(str(v)) if v is not None else 0, 60))
        hoja.column_dimensions[get_column_letter(columna[0].column)].width = ancho + 3


def exportar(alertas, detalle, resumen, reversiones, excepciones):
    try:
        with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl") as escritor:
            alertas.to_excel(escritor, sheet_name="Alertas", index=False)
            detalle.to_excel(escritor, sheet_name="Detalle", index=False)
            resumen.to_excel(escritor, sheet_name="Resumen", index=False)
            if ARCHIVO_REVERSIONES:
                reversiones.to_excel(escritor, sheet_name="Reversiones", index=False)
            excepciones.to_excel(escritor, sheet_name="Excepciones", index=False)
            for hoja in escritor.book.worksheets:
                dar_formato(hoja)
    except PermissionError:
        raise ErrorDatos(f"No se pudo guardar porque el archivo está abierto:\n    {RUTA_SALIDA}\n"
                         "Ciérralo en Excel y vuelve a ejecutar.")


# ============================================================================
# PROGRAMA PRINCIPAL
# ============================================================================
def main():
    try:
        act, tx, rev = cargar()
        act, tx, excepciones, conteos = limpiar(act, tx, rev)
        ventana = cruzar(act, tx, excepciones, conteos)
        evaluacion, alertas, detalle, reversiones = calcular(act, ventana)
        excepciones = preparar_excepciones(excepciones)
        resumen = armar_resumen(conteos, evaluacion, alertas, reversiones, excepciones)
        exportar(alertas, detalle, resumen, reversiones, excepciones)
    except ErrorDatos as error:
        print(f"\nERROR: {error}")
        return None

    evaluadas, marcadas = len(evaluacion), len(alertas)
    print("\n============== RESUMEN ==============")
    print(f"Activaciones evaluadas:        {evaluadas}")
    print(f"Con uso en {HORAS_VENTANA} h:                {int((evaluacion['n_tx'] > 0).sum())}")
    print(f"Marcadas (>= {UMBRAL:.0%}):           {marcadas} ({marcadas / evaluadas:.2%} de las evaluadas)" if evaluadas else "")
    for rango in ["90 a 94", "95 a 99", "100 o más"]:
        print(f"   {rango:10}: {int((alertas['RANGO DE USO'] == rango).sum())}")
    if ARCHIVO_REVERSIONES:
        se_caen = int((evaluacion["marcado_sin_reversiones"] & ~evaluacion["marcado"]).sum())
        print(f"Reversiones cruzadas:          {conteos['rev_cruzadas']} de {conteos['rev_aplicadas']} aplicadas "
              f"en el periodo ({conteos['rev_riesgo']} con tipificación de riesgo)")
        print(f"Alertas que se caen por reversión: {se_caen}")
    print(f"Duplicados eliminados:         act {conteos['act_duplicadas']} / tx {conteos['tx_duplicadas']}")
    print(f"Sin cruce:                     act {conteos['act_sin_tx']} / tx {conteos['tx_sin_act']}")
    print(f"Excepciones:                   {len(excepciones)} filas (hoja 'Excepciones')")
    print(f"Archivo generado:\n    {RUTA_SALIDA}")
    return {"evaluacion": evaluacion, "alertas": alertas, "detalle": detalle, "reversiones": reversiones,
            "resumen": resumen, "excepciones": excepciones}


if __name__ == "__main__":
    resultado = main()
