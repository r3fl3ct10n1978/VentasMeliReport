"""
Configuración central del reporte de ventas pendientes de Mercado Libre.
"""

from pathlib import Path


# =============================================================================
# SALIDA
# =============================================================================

CARPETA_SALIDA = Path(
    r"C:\Users\CGormaz\Desktop\Scripts\Reportes Ventas Meli\output"
)

NOMBRE_ARCHIVO_SALIDA = "Reporte_Pendientes_Meli"
NOMBRE_HOJA_SALIDA = "Pendientes Meli"


# =============================================================================
# REGLAS DEL REPORTE
# =============================================================================

ESTADO_FILTRO = "Etiqueta lista para imprimir"

EXTENSIONES_VALIDAS = {
    ".xlsx",
    ".xls",
    ".csv",
    ".tsv",
    ".txt",
}


# =============================================================================
# COLUMNAS
# =============================================================================

COLUMNAS_REQUERIDAS = {
    "# de venta": [
        "# de venta",
        "n de venta",
        "numero de venta",
        "número de venta",
        "id de venta",
        "order id",
        "order",
        "venta",
    ],
    "Estado": [
        "estado",
        "status",
        "estatus",
    ],
    "Unidades": [
        "unidades",
        "unidad",
        "cantidad",
        "quantity",
        "qty",
    ],
    "SKU": [
        "sku",
        "sku id",
        "seller sku",
    ],
    "Título de la publicación": [
        "titulo de la publicacion",
        "título de la publicación",
        "titulo",
        "título",
        "title",
        "publication title",
    ],
}

COLUMNAS_SALIDA = [
    "CODIGO",
    "PRODUCTO",
    "Suma de PENDIENTE MELI",
]
