"""
Exportación del reporte de Mercado Libre a Excel.
"""

from datetime import datetime

import pandas as pd

from config import (
    CARPETA_SALIDA,
    NOMBRE_ARCHIVO_SALIDA,
    NOMBRE_HOJA_SALIDA,
)


def construir_ruta_salida():
    """
    Crea la carpeta de salida y genera el nombre final.
    """

    CARPETA_SALIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    return (
        CARPETA_SALIDA
        / f"{NOMBRE_ARCHIVO_SALIDA}_{fecha}.xlsx"
    )


def ajustar_hoja(hoja):
    """
    Ajusta filtros, encabezados y ancho de columnas.
    """

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions

    hoja.column_dimensions["A"].width = 24
    hoja.column_dimensions["B"].width = 60
    hoja.column_dimensions["C"].width = 28

    for celda in hoja["A"]:
        celda.number_format = "@"


def guardar_reporte(reporte):
    """
    Guarda el reporte y devuelve la ruta generada.
    """

    ruta_salida = construir_ruta_salida()

    with pd.ExcelWriter(
        ruta_salida,
        engine="openpyxl",
    ) as writer:
        reporte.to_excel(
            writer,
            index=False,
            sheet_name=NOMBRE_HOJA_SALIDA,
        )

        ajustar_hoja(
            writer.sheets[NOMBRE_HOJA_SALIDA]
        )

    return ruta_salida
