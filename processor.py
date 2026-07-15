"""
Limpieza, filtrado y agrupación del reporte de Mercado Libre.
"""

import pandas as pd

from config import (
    COLUMNAS_REQUERIDAS,
    COLUMNAS_SALIDA,
    ESTADO_FILTRO,
)
from loader import (
    normalizar_columna,
    normalizar_texto,
)


def encontrar_columna(columnas, aliases):
    """
    Busca una columna usando sus posibles nombres.
    """

    normalizadas = {
        normalizar_columna(columna): columna
        for columna in columnas
    }

    for alias in aliases:
        alias_normalizado = normalizar_columna(alias)

        if alias_normalizado in normalizadas:
            return normalizadas[alias_normalizado]

        for nombre_normalizado, original in normalizadas.items():
            if (
                alias_normalizado
                and alias_normalizado in nombre_normalizado
            ):
                return original

    return None


def preparar_columnas(dataframe):
    """
    Encuentra y renombra las columnas necesarias.
    """

    renombrar = {}
    faltantes = []

    for nombre_final, aliases in COLUMNAS_REQUERIDAS.items():
        encontrada = encontrar_columna(
            dataframe.columns,
            aliases,
        )

        if encontrada is None:
            faltantes.append(nombre_final)
        else:
            renombrar[encontrada] = nombre_final

    if faltantes:
        raise ValueError(
            "Faltan columnas requeridas:\n"
            + "\n".join(f"- {columna}" for columna in faltantes)
            + "\n\nColumnas disponibles:\n"
            + "\n".join(f"- {columna}" for columna in dataframe.columns)
        )

    return dataframe.rename(columns=renombrar)[
        list(COLUMNAS_REQUERIDAS)
    ].copy()


def convertir_unidades(serie):
    """
    Convierte unidades a valores numéricos.
    """

    texto = (
        serie
        .astype("string")
        .fillna("")
        .str.strip()
        .str.replace(" ", "", regex=False)
    )

    tiene_coma = texto.str.contains(",", regex=False, na=False)
    tiene_punto = texto.str.contains(".", regex=False, na=False)

    formato_latino = (
        tiene_coma
        & tiene_punto
        & (texto.str.rfind(",") > texto.str.rfind("."))
    )

    texto.loc[formato_latino] = (
        texto.loc[formato_latino]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    formato_internacional = (
        tiene_coma
        & tiene_punto
        & ~formato_latino
    )

    texto.loc[formato_internacional] = (
        texto.loc[formato_internacional]
        .str.replace(",", "", regex=False)
    )

    solo_coma = tiene_coma & ~tiene_punto

    texto.loc[solo_coma] = (
        texto.loc[solo_coma]
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(
        texto,
        errors="coerce",
    ).fillna(0)


def generar_reporte(dataframe):
    """
    Filtra el estado requerido y agrupa por SKU.
    """

    dataframe = preparar_columnas(dataframe)

    dataframe["Estado"] = (
        dataframe["Estado"]
        .apply(normalizar_texto)
    )

    dataframe = dataframe[
        dataframe["Estado"].str.casefold()
        == ESTADO_FILTRO.casefold()
    ].copy()

    if dataframe.empty:
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    dataframe["CODIGO"] = (
        dataframe["SKU"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    dataframe["PRODUCTO"] = (
        dataframe["Título de la publicación"]
        .apply(normalizar_texto)
    )

    dataframe["Unidades"] = convertir_unidades(
        dataframe["Unidades"]
    )

    dataframe = dataframe[
        (dataframe["CODIGO"] != "")
        & (dataframe["CODIGO"].str.casefold() != "nan")
        & (dataframe["CODIGO"].str.casefold() != "none")
    ].copy()

    reporte = (
        dataframe
        .groupby(
            "CODIGO",
            as_index=False,
            sort=False,
            dropna=False,
        )
        .agg(
            PRODUCTO=("PRODUCTO", "first"),
            **{
                "Suma de PENDIENTE MELI": (
                    "Unidades",
                    "sum",
                )
            },
        )
        .sort_values("CODIGO", kind="stable")
        .reset_index(drop=True)
    )

    return reporte[COLUMNAS_SALIDA]
