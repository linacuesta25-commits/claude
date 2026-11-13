# -*- coding: utf-8 -*-
"""Prueba de casos borde para alerta_uso_cupo_24h.py con datos INVENTADOS."""
import tempfile
from pathlib import Path

import pandas as pd

import alerta_uso_cupo_24h as m

ACT = [  # TICKET, cédula, fecha firma, cupo, estado
    ("A1", "1.020.304", "2026-03-01 10:00:00", "1.000.000", "ACTIVADA"),  # 24:00:00 exacto entra
    ("A2", "00012345", "01/03/2026 10:00:00", "$ 1.000.000", "ACTIVADA"),  # 24:00:01 no entra
    ("A3", "555", "2026-03-01 10:00:00", 1000000, "ACTIVADA"),            # muchas compras pequeñas
    ("A4", "777", "2026-03-01 08:00:00", 1000000, "ACTIVADA"),            # 2 activaciones del mismo cliente
    ("A5", "777", "2026-03-10 08:00:00", 2000000, "ACTIVADA"),
    ("A6", "888", "2026-03-01 08:00:00", 0, "ACTIVADA"),                  # cupo cero -> excepción
    ("A7", "999", "fecha mala", 1000000, "ACTIVADA"),                     # fecha inválida -> excepción
    ("A8", "321", "2026-03-01 08:00:00", 1000000, "SOLICITUD EXPIRADA"),  # no es activada -> no se evalúa
    ("A9", "444", "2026-03-01 08:00:00", 1000000, "ACTIVADA"),            # reversa
    ("A10", "666", "2026-03-01 08:00:00", 1000000, "ACTIVADA"),           # sin transacciones
]
TX = [  # cédula, fecha, valor, id, estado
    ("1020304", "2026-03-02 10:00:00", 950000, "T1", "APROBADA"),
    ("12345", "2026-03-02 10:00:01", 950000, "T2", "APROBADA"),  # misma cédula que 00012345
    *[("555", f"2026-03-01 1{i}:00:00", 100000, f"T3{i}", "APROBADA") for i in range(1, 10)],
    ("555", "2026-03-01 19:30:00", 500000, "T3X", "RECHAZADA"),
    ("777", "2026-03-01 09:00:00", 950000, "T4", "APROBADA"),
    ("777", "2026-03-10 09:00:00", 500000, "T5", "APROBADA"),
    ("444", "2026-03-01 09:00:00", 950000, "T9", "APROBADA"),
    ("444", "2026-03-01 09:30:00", 950000, "T9", "REVERSADA"),
    ("444", "2026-03-01 10:00:00", 200000, "T9B", "APROBADA"),
    ("12312", "2026-03-01 10:00:00", 50000, "T0", "APROBADA"),        # sin activación
]
ESPERADO = {  # (excluir, restar)
    "A1": (True, True), "A2": (False, False), "A3": (True, True), "A4": (True, True),
    "A5": (False, False), "A9": (False, True), "A10": (False, False),
}


def correr(carpeta: Path, modo: str):
    act = pd.DataFrame(ACT, columns=["TICKET", "Documento de identidad", "Fecha firma pagare", "VALOR CUPO APROBADO", "IDENTIFICADOR"])
    act.insert(0, "EMPRESA", "SSI")
    with pd.ExcelWriter(carpeta / "act.xlsx") as w:   # encabezado en la fila 3
        pd.DataFrame([["Reporte"], [None]]).to_excel(w, index=False, header=False)
        act.to_excel(w, index=False, startrow=2)
    pd.DataFrame(TX, columns=["Identificación", "Fecha Utilización", "Monto", "Número de crédito", "Estado"]).to_excel(
        carpeta / "tx.xlsx", index=False)
    m.ARCHIVOS_ACTIVACIONES = [carpeta / "act.xlsx"]
    m.ARCHIVO_TRANSACCIONES = carpeta / "tx.xlsx"
    m.TX_ESTADO = "Estado"
    m.TRATAMIENTO_REVERSAS = modo
    m.RUTA_SALIDA = carpeta / f"salida_{modo}.xlsx"
    m.ARCHIVO_REVERSIONES = None
    return m.main()


def main():
    carpeta = Path(tempfile.mkdtemp())
    errores = 0
    for i, modo in enumerate(["excluir", "restar"]):
        print(f"\n######## TRATAMIENTO_REVERSAS = '{modo}' ########")
        r = correr(carpeta, modo)
        marcadas = set(r["alertas"]["ID ACTIVACIÓN"])
        for ticket, esperado in ESPERADO.items():
            ok = (ticket in marcadas) == esperado[i]
            errores += not ok
            fila = r["evaluacion"].loc[ticket]
            print(f"  {ticket}: consumo {fila['consumo']:>10,.0f}  uso {fila['pct']:6.1%}  marcado={ticket in marcadas}  {'OK' if ok else 'ERROR'}")
        tipos = set(r["excepciones"]["TIPO"])
        for t in ["Cupo aprobado cero, nulo o no numérico", "Activación con fecha inválida o vacía",
                  "Transacción sin activación (no cruza)", "Activación sin transacciones (no cruza)",
                  "Transacción excluida por estado (no aprobada)"]:
            errores += t not in tipos
        a3 = r["alertas"].set_index("ID ACTIVACIÓN").loc["A3"]
        cruce_ok = a3["TRANSACCIONES HASTA EL CRUCE"] == 9 and str(a3["FECHA Y HORA EN QUE SE ALCANZÓ EL 90%"]) == "2026-03-01 19:00:00"
        errores += not cruce_ok
        print(f"  A3 cruza el 90% en la transacción {a3['TRANSACCIONES HASTA EL CRUCE']} a las {a3["FECHA Y HORA EN QUE SE ALCANZÓ EL 90%"]}: {'OK' if cruce_ok else 'ERROR'}")
    print("\nRESULTADO:", "TODAS LAS PRUEBAS PASARON" if errores == 0 else f"{errores} PRUEBAS FALLARON")


if __name__ == "__main__":
    main()


# ============================================================================
# Prueba del consolidado de reversiones (datos INVENTADOS)
# ============================================================================
def prueba_reversiones():
    carpeta = Path(tempfile.mkdtemp())
    act = pd.DataFrame([  # TICKET, cédula, firma, cupo
        ("R1", "100", "2026-05-01 10:00:00", 1000000),   # reversión total -> se cae
        ("R2", "200", "2026-05-01 10:00:00", 1000000),   # reversión parcial 200.000 -> 80 %, se cae
        ("R3", "300", "2026-05-01 10:00:00", 1000000),   # reversión PENDIENTE -> no se descuenta, sigue
        ("R4", "400", "2026-05-01 10:00:00", 2000000),   # cobro doble: 2 x 1.000.000, una reversada -> 50 %
        ("R5", "500", "2026-05-01 10:00:00", 1000000),   # fecha texto 01/05/2026 (día/mes) + riesgo
        ("R6", "600", "2026-05-01 10:00:00", 1000000),   # fecha mala, cruza por 'Fecha organizada'
    ], columns=["TICKET", "DOCUMENTO DE IDENTIDAD", "FECHA FIRMA PAGARÉ", "VALOR CUPO APROBADO"])
    act["EMPRESA"], act["IDENTIFICADOR"] = "SSI", "ACTIVADA"
    act.to_excel(carpeta / "act.xlsx", index=False)
    pd.DataFrame([
        ("100", "2026-05-01 11:00:00", 1000000, "U1"), ("200", "2026-05-01 11:00:00", 1000000, "U2"),
        ("300", "2026-05-01 11:00:00", 1000000, "U3"), ("400", "2026-05-01 11:00:00", 1000000, "U4"),
        ("400", "2026-05-01 11:00:05", 1000000, "U5"), ("500", "2026-05-01 11:00:00", 1000000, "U6"),
        ("600", "2026-05-01 11:00:00", 1000000, "U7"),
    ], columns=["Identificación", "Fecha Utilización", "Monto", "Número de crédito"]).to_excel(carpeta / "tx.xlsx", index=False)
    cols = ["Cédula Cliente", "Fecha Transacción", "Fecha organizada", "Valor Total TX", "Valor Reversión",
            "Estado Reversión (Aplicada/Pendiente)", "Tipificación", "Fecha Ingreso Solicitud"]
    rev = pd.DataFrame([
        (100, pd.Timestamp("2026-05-01"), 20260501, 1000000, 1000000, "Aplicada", "DESISTE DE LA COMPRA", pd.Timestamp("2026-05-05")),
        (100, pd.Timestamp("2026-05-01"), 20260501, 1000000, 1000000, "Aplicada", "DESISTE DE LA COMPRA", pd.Timestamp("2026-05-05")),  # duplicada
        (200, pd.Timestamp("2026-05-01"), 20260501, 1000000, 200000, "Aplicada", "ERROR EN VALOR", pd.Timestamp("2026-05-05")),
        (300, pd.Timestamp("2026-05-01"), 20260501, 1000000, 1000000, "PENDIENTE", "DESISTE DE LA COMPRA", pd.Timestamp("2026-05-05")),
        (400, pd.Timestamp("2026-05-01"), 20260501, 1000000, 1000000, "Aplicada", "COBRO DOBLE", pd.Timestamp("2026-05-05")),
        (500, "01/05/2026", 20260501, 1000000, 1000000, "Aplicada", "NO PASA VALIDACIÓN DE IDENTIDAD", pd.Timestamp("2026-05-05")),
        (600, "//", 20260501, 1000000, 1000000, "Aplicada", "DESISTE DE LA COMPRA", pd.Timestamp("2026-05-05")),
        (None, "//", None, None, None, None, None, None),                                   # fila vacía de plantilla
        (700, pd.Timestamp("2024-01-01"), 20240101, 500000, 500000, "Aplicada", "DESISTE", pd.Timestamp("2024-01-05")),  # fuera de periodo
    ], columns=cols)
    with pd.ExcelWriter(carpeta / "rev.xlsx") as w:
        rev.to_excel(w, sheet_name="AGIL", index=False)

    m.ARCHIVOS_ACTIVACIONES, m.ARCHIVO_TRANSACCIONES = [carpeta / "act.xlsx"], carpeta / "tx.xlsx"
    m.ARCHIVO_REVERSIONES, m.HOJAS_REVERSIONES = carpeta / "rev.xlsx", ["AGIL"]
    m.TX_ESTADO, m.RUTA_SALIDA = None, carpeta / "salida_rev.xlsx"
    r = m.main()
    ev = r["evaluacion"]
    esperado = {"R1": (0.0, False), "R2": (0.8, False), "R3": (1.0, True), "R4": (0.5, False),
                "R5": (0.0, False), "R6": (0.0, False)}
    errores = 0
    for t, (pct, marcado) in esperado.items():
        ok = abs(ev.loc[t, "pct"] - pct) < 1e-9 and bool(ev.loc[t, "marcado"]) == marcado
        errores += not ok
        print(f"  {t}: uso {ev.loc[t, 'pct']:.0%} marcado={bool(ev.loc[t, 'marcado'])}  {'OK' if ok else 'ERROR'}")
    rv = r["reversiones"].set_index("ID ACTIVACIÓN")
    chequeos = {
        "cobro doble: se reversa la compra más reciente (U5)": rv.loc["R4", "ID TRANSACCIÓN"] == "U5",
        "R5 marcada con tipificación de riesgo": rv.loc["R5", "TIPIFICACIÓN DE RIESGO"] == "Sí",
        "R1 'La alerta se cae por la reversión'": rv.loc["R1", "RESULTADO"] == "La alerta se cae por la reversión",
        "1 duplicado y 1 pendiente reportados": set(r["excepciones"]["TIPO"]) >= {
            "Duplicado exacto eliminado (reversiones)", "Reversión no aplicada (no se descuenta)"},
    }
    for nombre, ok in chequeos.items():
        errores += not ok
        print(f"  {nombre}: {'OK' if ok else 'ERROR'}")
    print("\nRESULTADO REVERSIONES:", "TODAS LAS PRUEBAS PASARON" if errores == 0 else f"{errores} PRUEBAS FALLARON")


if __name__ == "__main__":
    prueba_reversiones()
