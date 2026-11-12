# -*- coding: utf-8 -*-
"""
Prueba de casos límite para validar_fuera_horario.py con datos INVENTADOS.
Crea un Excel de prueba en una carpeta temporal, ejecuta el análisis y compara
cada fila contra el resultado esperado.

Uso:
    python prueba_casos_limite.py
"""
import tempfile
from datetime import datetime
from pathlib import Path
from unittest import mock

import pandas as pd

import validar_fuera_horario as v

D = datetime
# (empresa, estado, creación, última modificación, motivo esperado o None si no debe salir)
CASOS = [
    ("SSI", "ACTIVADA", D(2026, 11, 10, 7, 59, 59), D(2026, 11, 10, 10, 0), "Creación fuera de horario"),
    ("SSI", "ACTIVADA", D(2026, 11, 10, 8, 0, 0), D(2026, 11, 10, 10, 0), None),
    ("SSI", "ACTIVADA", D(2026, 11, 10, 19, 59, 59), D(2026, 11, 10, 19, 59, 59), None),
    ("SSI", "ACTIVADA", D(2026, 11, 10, 20, 0, 0), D(2026, 11, 10, 20, 0, 0), "Ambas fuera de horario"),
    ("SSI", "ACTIVADA", D(2026, 11, 10, 0, 0, 0), D(2026, 11, 10, 9, 0), "Creación fuera de horario"),
    ("SSI", "ACTIVADA", D(2026, 11, 10, 9, 0), D(2026, 11, 10, 20, 0, 0), "Última modificación fuera de horario"),
    ("SSI", "ACTIVADA", None, D(2026, 11, 10, 9, 0), None),                       # creación vacía
    ("SSI", " activada ", "10/11/2026 7:59:59 p. m.", "10/11/2026 20:30", "Última modificación fuera de horario"),
    ("SSI", "Activada", "03/02/2026 07:00", "03/02/2026 09:00", "Creación fuera de horario"),  # 3 de febrero
    ("SSI", "ACTIVADA", "texto raro", D(2026, 11, 10, 9, 0), None),               # no convertible
    ("SSI", "ACTIVADA", D(2026, 11, 14, 10, 0), D(2026, 11, 14, 11, 0), None),    # sábado, dentro de la hora
    ("SSI", "RECHAZADO PERFIL DE RIESGO", D(2026, 11, 10, 22, 0), D(2026, 11, 10, 22, 0), None),
    ("ESSA", "ACTIVADA", D(2026, 11, 10, 22, 0), D(2026, 11, 10, 22, 0), None),
]


def crear_excel(ruta: Path):
    """Encabezado en la fila 3, con tildes/mayúsculas/espacios distintos a propósito."""
    filas = [["Reporte de prueba", None, None, None, None], [None] * 5,
             ["TICKET", " empresa ", "Fecha de Creacion  de Ticket", "FECHA ÚLTIMA MODIFICACIÓN", "IDENTIFICADOR"]]
    for i, (emp, est, cre, mod, _) in enumerate(CASOS, start=1):
        filas.append([f"T{i:02d}", emp, cre, mod, est])
    pd.DataFrame(filas).to_excel(ruta, header=False, index=False)


def main():
    carpeta = Path(tempfile.mkdtemp())
    entrada = carpeta / "Validación prueba con tildes" / "Reporte 1.xlsx"
    entrada.parent.mkdir()
    crear_excel(entrada)
    info = v.procesar(entrada, entrada.parent / v.NOMBRE_SALIDA)

    fuera = info["tabla_fuera"].set_index("TICKET")
    errores = 0
    print(f"{'TICKET':6} {'CREACIÓN':22} {'MODIFICACIÓN':22} {'ESPERADO':38} OBTENIDO")
    for i, (emp, est, cre, mod, esperado) in enumerate(CASOS, start=1):
        t = f"T{i:02d}"
        obtenido = fuera.at[t, v.COL_MOTIVO] if t in fuera.index else None
        ok = obtenido == esperado
        errores += not ok
        print(f"{t:6} {str(cre):22} {str(mod):22} {str(esperado):38} {obtenido}  {'OK' if ok else 'ERROR'}")

    print("\nFechas con problemas:")
    print(info["tabla_problemas"].to_string(index=False))
    esperado_problemas = {(10, "Fecha de Creacion  de Ticket"), (13, "Fecha de Creacion  de Ticket")}
    obtenido_problemas = set(zip(info["tabla_problemas"]["FILA EXCEL ORIGINAL"], info["tabla_problemas"]["COLUMNA"]))
    errores += obtenido_problemas != esperado_problemas
    resumen_ok = (info["leidas"], info["ssi"], info["activadas"]) == (13, 12, 11)
    errores += not resumen_ok
    print(f"\nLeídas/SSI/Activadas = {info['leidas']}/{info['ssi']}/{info['activadas']} (esperado 13/12/11)")

    # Fines de semana activados: el sábado (T11) debe salir marcado
    with mock.patch.object(v, "EVALUAR_FINES_DE_SEMANA", True):
        info2 = v.procesar(entrada, entrada.parent / "fds.xlsx")
    sabado = "T11" in set(info2["tabla_fuera"]["TICKET"])
    errores += not sabado
    print(f"EVALUAR_FINES_DE_SEMANA=True marca el sábado T11: {sabado}")

    # Columna faltante -> mensaje claro
    with mock.patch.object(v, "COLUMNA_EMPRESA", "COMPAÑIA"):
        try:
            v.procesar(entrada, entrada.parent / "x.xlsx")
            print("Columna faltante: NO se detuvo  ERROR"); errores += 1
        except v.ErrorValidacion as e:
            print(f"Columna faltante detectada: {str(e).splitlines()[0]}")

    # Archivo de salida abierto -> mensaje claro
    with mock.patch.object(pd, "ExcelWriter", side_effect=PermissionError):
        try:
            v.procesar(entrada, entrada.parent / v.NOMBRE_SALIDA)
            errores += 1
        except v.ErrorValidacion as e:
            print(f"Archivo bloqueado detectado: {str(e).splitlines()[0]}")

    print("\nRESULTADO:", "TODAS LAS PRUEBAS PASARON" if errores == 0 else f"{errores} PRUEBAS FALLARON")


if __name__ == "__main__":
    main()
