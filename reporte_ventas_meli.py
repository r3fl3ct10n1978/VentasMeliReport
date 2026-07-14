#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Genera un reporte de ventas pendientes de Mercado Libre.

Funcionamiento:
1. Toma automáticamente el archivo más reciente desde la carpeta "input".
2. Filtra el estado "Etiqueta lista para imprimir".
3. Agrupa por SKU y suma las unidades.
4. Guarda el resultado en la carpeta "Procesado".
"""

import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Falta pandas. Instala las dependencias con:\n"
        "py -m pip install -r requirements.txt"
    ) from exc


# =============================================================================
# CONFIGURACIÓN
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent
CARPETA_INPUT = BASE_DIR / "input"
CARPETA_SALIDA = BASE_DIR / "Procesado"

ESTADO_FILTRO = "Etiqueta lista para imprimir"
EXTENSIONES_VALIDAS = {".xlsx", ".xls", ".csv", ".tsv", ".txt"}

COLUMNAS_REQUERIDAS = {
    "# de venta": [
        "# de venta", "n de venta", "numero de venta",
        "id de venta", "order id", "order", "venta"
    ],
    "Estado": ["estado", "status", "estatus"],
    "Unidades": ["unidades", "unidad", "cantidad", "quantity", "qty"],
    "SKU": ["sku", "sku id"],
    "Título de la publicación": [
        "titulo de la publicacion", "titulo",
        "title", "publication title"
    ],
}

COLUMNAS_SALIDA = ["CODIGO", "PRODUCTO", "Suma de PENDIENTE MELI"]


# =============================================================================
# ARCHIVOS
# =============================================================================

def obtener_archivo_entrada():
    """Devuelve el archivo válido más reciente de la carpeta input."""
    CARPETA_INPUT.mkdir(exist_ok=True)

    archivos = [
        archivo
        for archivo in CARPETA_INPUT.iterdir()
        if archivo.is_file() and archivo.suffix.lower() in EXTENSIONES_VALIDAS
    ]

    if not archivos:
        raise FileNotFoundError(
            f"No hay archivos válidos en:\n{CARPETA_INPUT}\n"
            "Copia allí el archivo exportado de Mercado Libre."
        )

    archivo = max(archivos, key=lambda ruta: ruta.stat().st_mtime)
    print(f"[INFO] Archivo seleccionado: {archivo.name}")
    return archivo


def obtener_ruta_salida():
    """Crea la carpeta de salida y genera el nombre del reporte."""
    CARPETA_SALIDA.mkdir(exist_ok=True)
    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return CARPETA_SALIDA / f"Reporte_Pendientes_Meli_{fecha}.xlsx"


# =============================================================================
# NORMALIZACIÓN
# =============================================================================

def normalizar_texto(valor):
    """Convierte un valor a texto limpio."""
    if pd.isna(valor):
        return ""
    return " ".join(str(valor).strip().split())


def normalizar_columna(valor):
    """Normaliza encabezados para compararlos sin tildes ni símbolos."""
    texto = normalizar_texto(valor).lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )
    return re.sub(r"[^a-z0-9]+", "", texto)


def encontrar_columna(columnas, aliases):
    """Busca una columna usando sus posibles nombres."""
    columnas_normalizadas = {
        normalizar_columna(columna): columna
        for columna in columnas
    }

    for alias in aliases:
        alias_normalizado = normalizar_columna(alias)

        if alias_normalizado in columnas_normalizadas:
            return columnas_normalizadas[alias_normalizado]

        for nombre_normalizado, columna_original in columnas_normalizadas.items():
            if alias_normalizado and alias_normalizado in nombre_normalizado:
                return columna_original

    return None


def limpiar_sku(valor):
    """Limpia y estandariza un SKU."""
    sku = normalizar_texto(valor).upper()
    return sku or None


# =============================================================================
# CARGA DEL ARCHIVO
# =============================================================================

def detectar_fila_encabezado(tabla):
    """Busca en las primeras 15 filas la fila que contiene más encabezados."""
    aliases = {
        normalizar_columna(alias)
        for lista_aliases in COLUMNAS_REQUERIDAS.values()
        for alias in lista_aliases
    }

    mejor_fila = None
    mayor_coincidencia = 0

    for indice in range(min(15, len(tabla))):
        fila = {
            normalizar_columna(valor)
            for valor in tabla.iloc[indice].tolist()
        }
        coincidencias = len(fila & aliases)

        if coincidencias > mayor_coincidencia:
            mejor_fila = indice
            mayor_coincidencia = coincidencias

    return mejor_fila if mayor_coincidencia >= 3 else None


def construir_dataframe(tabla):
    """Convierte una tabla sin encabezados en un DataFrame utilizable."""
    fila_encabezado = detectar_fila_encabezado(tabla)

    if fila_encabezado is None:
        raise ValueError("No se pudo identificar la fila de encabezados.")

    encabezados = [
        normalizar_texto(valor)
        for valor in tabla.iloc[fila_encabezado].tolist()
    ]

    dataframe = tabla.iloc[fila_encabezado + 1:].copy()
    dataframe.columns = encabezados
    columnas_sin_duplicados = ~pd.Index(dataframe.columns).duplicated()
    dataframe = dataframe.loc[:, columnas_sin_duplicados]
    dataframe = dataframe.dropna(how="all")

    return dataframe


def cargar_archivo(ruta):
    """Carga Excel, CSV, TSV o TXT y detecta sus encabezados."""
    extension = ruta.suffix.lower()

    if extension in {".xlsx", ".xls"}:
        engine = "openpyxl" if extension == ".xlsx" else "xlrd"

        try:
            libro = pd.ExcelFile(ruta, engine=engine)
        except ImportError as exc:
            raise RuntimeError(
                "Para leer archivos .xls instala xlrd:\n"
                "py -m pip install xlrd"
            ) from exc

        for hoja in libro.sheet_names:
            tabla = pd.read_excel(ruta, sheet_name=hoja, header=None, engine=engine)

            try:
                return construir_dataframe(tabla)
            except ValueError:
                continue

        raise ValueError("No se encontraron encabezados válidos en el Excel.")

    for encoding in ("utf-8-sig", "latin-1"):
        try:
            tabla = pd.read_csv(
                ruta,
                sep=None,
                engine="python",
                header=None,
                encoding=encoding
            )
            return construir_dataframe(tabla)
        except UnicodeDecodeError:
            continue

    raise RuntimeError("No se pudo leer el archivo de entrada.")


# =============================================================================
# PROCESAMIENTO
# =============================================================================

def preparar_columnas(dataframe):
    """Encuentra y renombra las columnas necesarias."""
    renombrar = {}
    faltantes = []

    for nombre_final, aliases in COLUMNAS_REQUERIDAS.items():
        encontrada = encontrar_columna(dataframe.columns, aliases)

        if encontrada is None:
            faltantes.append(nombre_final)
        else:
            renombrar[encontrada] = nombre_final

    if faltantes:
        raise ValueError(
            f"Faltan columnas requeridas: {', '.join(faltantes)}\n"
            f"Columnas disponibles: {list(dataframe.columns)}"
        )

    return dataframe.rename(columns=renombrar)[
        list(COLUMNAS_REQUERIDAS)
    ].copy()


def generar_reporte(dataframe):
    """Filtra, limpia y agrupa las ventas por SKU."""
    dataframe = preparar_columnas(dataframe)

    dataframe["Estado"] = dataframe["Estado"].apply(normalizar_texto)
    dataframe = dataframe[
        dataframe["Estado"].str.casefold() == ESTADO_FILTRO.casefold()
    ].copy()

    if dataframe.empty:
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    dataframe["CODIGO"] = dataframe["SKU"].apply(limpiar_sku)
    dataframe["PRODUCTO"] = dataframe[
        "Título de la publicación"
    ].apply(normalizar_texto)

    dataframe["Unidades"] = pd.to_numeric(
        dataframe["Unidades"],
        errors="coerce"
    ).fillna(0)

    dataframe = dataframe.dropna(subset=["CODIGO"])

    reporte = (
        dataframe
        .groupby("CODIGO", as_index=False)
        .agg(
            PRODUCTO=("PRODUCTO", "first"),
            **{"Suma de PENDIENTE MELI": ("Unidades", "sum")}
        )
        .sort_values("CODIGO")
        .reset_index(drop=True)
    )

    return reporte[COLUMNAS_SALIDA]


def guardar_reporte(reporte, ruta_salida):
    """Guarda el resultado en Excel."""
    try:
        reporte.to_excel(
            ruta_salida,
            index=False,
            sheet_name="Pendientes Meli"
        )
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "Falta openpyxl. Instálalo con:\n"
            "py -m pip install openpyxl"
        ) from exc


# =============================================================================
# EJECUCIÓN
# =============================================================================

def main():
    try:
        print("=" * 65)
        print("REPORTE DE VENTAS PENDIENTES - MERCADO LIBRE")
        print("=" * 65)

        archivo_entrada = obtener_archivo_entrada()
        ruta_salida = obtener_ruta_salida()

        dataframe = cargar_archivo(archivo_entrada)
        reporte = generar_reporte(dataframe)
        guardar_reporte(reporte, ruta_salida)

        print(f"[OK] Filas procesadas: {len(dataframe)}")
        print(f"[OK] SKU únicos: {len(reporte)}")
        print(
            "[OK] Total pendiente: "
            f"{reporte['Suma de PENDIENTE MELI'].sum():g}"
        )
        print(f"[OK] Reporte creado en:\n{ruta_salida}")

    except Exception as error:
        print(f"\n[ERROR] {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()