#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Punto de entrada del reporte de ventas pendientes de Mercado Libre.
"""

import sys

from config import ESTADO_FILTRO
from exporter import guardar_reporte
from loader import (
    cargar_archivo,
    solicitar_ruta_entrada,
)
from processor import generar_reporte


def main():
    """
    Ejecuta el flujo completo del reporte.
    """

    try:
        print("=" * 70)
        print("REPORTE DE VENTAS PENDIENTES - MERCADO LIBRE")
        print("=" * 70)

        ruta_entrada = solicitar_ruta_entrada()
        dataframe = cargar_archivo(ruta_entrada)
        reporte = generar_reporte(dataframe)

        if reporte.empty:
            print(
                "\nNo se encontraron ventas con el estado:\n"
                f"{ESTADO_FILTRO}"
            )
            return

        ruta_salida = guardar_reporte(reporte)

        print()
        print("=" * 70)
        print("PROCESO COMPLETADO")
        print("=" * 70)
        print(f"Filas procesadas: {len(dataframe)}")
        print(f"SKU únicos: {len(reporte)}")
        print(
            "Total pendiente: "
            f"{reporte['Suma de PENDIENTE MELI'].sum():g}"
        )
        print(f"Reporte creado en:\n{ruta_salida}")

    except Exception as error:
        print()
        print("=" * 70)
        print("ERROR EN EL REPORTE")
        print("=" * 70)
        print(f"\n{error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
