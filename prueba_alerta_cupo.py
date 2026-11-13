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
