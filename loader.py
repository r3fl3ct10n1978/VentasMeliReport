"""
Solicitud, validación y lectura del archivo exportado de Mercado Libre.
"""

import unicodedata
import re
from pathlib import Path

import pandas as pd

from config import (
    COLUMNAS_REQUERIDAS,
    EXTENSIONES_VALIDAS,
)


def solicitar_ruta_entrada():
    """
    Solicita obligatoriamente la ruta del archivo.
    """

    print("Ingresa la ruta completa del archivo exportado de Mercado Libre.")
    print("También puedes arrastrar el archivo hacia esta terminal.")

    entrada = input("\nRuta del archivo:\n> ").strip()

    if not entrada:
        raise ValueError(
            "Debes entregar una ruta de archivo para generar el reporte."
        )

    entrada = entrada.strip('"').strip("'").strip()

    if not entrada:
        raise ValueError("La ruta ingresada está vacía.")

    return Path(entrada)


def validar_ruta(ruta):
    """
    Valida existencia, tipo y extensión del archivo.
    """

    ruta = Path(ruta)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo:\n{ruta}"
        )

    if not ruta.is_file():
        raise ValueError(
            f"La ruta no corresponde a un archivo:\n{ruta}"
        )

    if ruta.suffix.lower() not in EXTENSIONES_VALIDAS:
        raise ValueError(
            f"Formato no compatible: {ruta.suffix}\n"
            f"Permitidos: {sorted(EXTENSIONES_VALIDAS)}"
        )

    return ruta


def normalizar_texto(valor):
    """
    Convierte un valor a texto limpio.
    """

    if pd.isna(valor):
        return ""

    return " ".join(str(valor).strip().split())


def normalizar_columna(valor):
    """
    Normaliza encabezados para compararlos sin tildes ni símbolos.
    """

    texto = normalizar_texto(valor).casefold()
    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    return re.sub(r"[^a-z0-9]+", "", texto)


def detectar_fila_encabezado(tabla):
    """
    Busca en las primeras 15 filas la fila con más encabezados reconocibles.
    """

    aliases = {
        normalizar_columna(alias)
        for lista in COLUMNAS_REQUERIDAS.values()
        for alias in lista
    }

    mejor_fila = None
    mayor_coincidencia = 0

    for indice in range(min(15, len(tabla))):
        valores = {
            normalizar_columna(valor)
            for valor in tabla.iloc[indice].tolist()
        }

        coincidencias = len(valores & aliases)

        if coincidencias > mayor_coincidencia:
            mejor_fila = indice
            mayor_coincidencia = coincidencias

    return mejor_fila if mayor_coincidencia >= 3 else None


def construir_dataframe(tabla):
    """
    Convierte una tabla sin encabezados en un DataFrame utilizable.
    """

    fila_encabezado = detectar_fila_encabezado(tabla)

    if fila_encabezado is None:
        raise ValueError(
            "No se pudo identificar la fila de encabezados."
        )

    encabezados = [
        normalizar_texto(valor)
        for valor in tabla.iloc[fila_encabezado].tolist()
    ]

    dataframe = tabla.iloc[fila_encabezado + 1:].copy()
    dataframe.columns = encabezados

    dataframe = dataframe.loc[
        :,
        ~pd.Index(dataframe.columns).duplicated(),
    ]

    return dataframe.dropna(how="all").reset_index(drop=True)


def cargar_excel(ruta):
    """
    Lee todas las hojas hasta encontrar encabezados válidos.
    """

    extension = ruta.suffix.lower()
    motor = "openpyxl" if extension == ".xlsx" else "xlrd"

    try:
        libro = pd.ExcelFile(ruta, engine=motor)
    except ImportError as error:
        raise RuntimeError(
            "Falta una dependencia para leer el archivo.\n"
            "Para .xlsx instala openpyxl.\n"
            "Para .xls instala xlrd."
        ) from error

    for hoja in libro.sheet_names:
        tabla = pd.read_excel(
            ruta,
            sheet_name=hoja,
            header=None,
            dtype=str,
            keep_default_na=False,
            engine=motor,
        )

        try:
            dataframe = construir_dataframe(tabla)
            print(f"Encabezados detectados en la hoja: {hoja}")
            return dataframe
        except ValueError:
            continue

    raise ValueError(
        "No se encontraron encabezados válidos en el Excel."
    )


def cargar_texto(ruta):
    """
    Lee CSV, TSV o TXT detectando separador y codificación.
    """

    ultimo_error = None

    for encoding in (
        "utf-8-sig",
        "utf-8",
        "latin-1",
        "cp1252",
    ):
        try:
            tabla = pd.read_csv(
                ruta,
                sep=None,
                engine="python",
                header=None,
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
            )

            return construir_dataframe(tabla)

        except (
            UnicodeDecodeError,
            pd.errors.ParserError,
            ValueError,
        ) as error:
            ultimo_error = error

    raise RuntimeError(
        "No se pudo leer el archivo de entrada.\n"
        f"Detalle: {ultimo_error}"
    )


def cargar_archivo(ruta):
    """
    Valida y carga el archivo indicado.
    """

    ruta = validar_ruta(ruta)

    print(f"Archivo seleccionado: {ruta.name}")
    print(f"Ruta utilizada: {ruta}")

    if ruta.suffix.lower() in {".xlsx", ".xls"}:
        return cargar_excel(ruta)

    return cargar_texto(ruta)
