#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SCRIPT DE AUTOMATIZACIÓN - REPORTE DE VENTAS SIN PROCESAR (MERCADO LIBRE)
================================================================================
Autor: Cristobal Gormaz
Fecha: 2026-07-08
Descripción: 
    Este script automatiza la limpieza y generación de reporte de ventas 
    "Etiqueta lista para imprimir" desde Mercado Libre.

    Proceso:
    1. Filtra transacciones con estado "Etiqueta lista para imprimir"
    2. Extrae SKU único por tipo
    3. Extrae nombre del producto por SKU
    4. Suma todas las unidades pendientes por cada SKU
    5. Genera tabla resultado: CODIGO | PRODUCTO | Suma de PENDIENTE MELI

Uso:
    python reporte_ventas_meli.py --input archivo.tsv --output reporte.xlsx

    O simplemente:
    python reporte_ventas_meli.py
    (buscará automáticamente archivos .tsv en el directorio)
================================================================================
"""

try:
    import pandas as pd
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Error: pandas no está instalado. Instala las dependencias con:\n"
        "  pip install pandas openpyxl"
    ) from exc

import os
import sys
import re
import argparse
import unicodedata
from datetime import datetime
import glob

# =============================================================================
# CONFIGURACIÓN GENERAL DEL SCRIPT
# =============================================================================
# Estado que se usará para filtrar las ventas que interesan al reporte.
ESTADO_FILTRO = "Etiqueta lista para imprimir"

# Columnas que deben existir en el archivo de entrada para poder procesarlo.
# Se aceptan variantes de nombres que suelen venir en exportaciones de Mercado Libre.
COLUMNAS_REQUERIDAS = {
    '# de venta': ['# de venta', 'n de venta', 'numero de venta', 'id de venta', 'order id', 'order', 'venta'],
    'Estado': ['estado', 'status', 'estatus'],
    'Unidades': ['unidades', 'unidad', 'cantidad', 'quantity', 'qty'],
    'SKU': ['sku', 'sku id'],
    'Título de la publicación': ['titulo de la publicacion', 'titulo', 'title', 'publication title']
}

# Columnas que tendrá la tabla final de salida.
COLUMNAS_SALIDA = ['CODIGO', 'PRODUCTO', 'Suma de PENDIENTE MELI']


def encontrar_archivo_entrada(ruta=None):
    """
    Busca y devuelve la ruta del archivo de entrada.
    Si no se entrega una ruta, intenta encontrar automáticamente un archivo
    en Descargas/Downloads que contenga las palabras clave "Ventas" y "Mercado Libre".
    """
    # Si el usuario ya pasó una ruta válida, se usa directamente.
    if ruta:
        if os.path.exists(ruta):
            return ruta
        raise FileNotFoundError(f"No se encontró el archivo especificado: {ruta}")

    carpetas_busqueda = []
    home = os.path.expanduser("~")
    if home:
        carpetas_busqueda.extend([
            os.path.join(home, "Downloads"),
            os.path.join(home, "Descargas"),
        ])
    carpetas_busqueda.append(os.getcwd())

    extensiones = {".xlsx", ".xls"}
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')

    candidatos = []
    for carpeta in carpetas_busqueda:
        if not carpeta or not os.path.isdir(carpeta):
            continue

        for root, _, files in os.walk(carpeta):
            for nombre_archivo in files:
                ruta_archivo = os.path.join(root, nombre_archivo)
                nombre_bajo = os.path.basename(ruta_archivo).lower()
                ext = os.path.splitext(nombre_bajo)[1].lower()
                if ext not in extensiones:
                    continue

                try:
                    fecha_modificacion = datetime.fromtimestamp(os.path.getmtime(ruta_archivo)).strftime('%Y-%m-%d')
                except OSError:
                    continue

                if fecha_modificacion == fecha_hoy:
                    candidatos.append(ruta_archivo)

    if candidatos:
        candidatos.sort(key=os.path.getmtime, reverse=True)
        print(f"[INFO] Archivo detectado automáticamente: {candidatos[0]}")
        return candidatos[0]

    raise FileNotFoundError(
        "No se encontró un archivo Excel descargado hoy en Descargas/Downloads. "
        "Por favor especifica la ruta con --input"
    )


def normalizar_nombre_columna(nombre):
    """Convierte un nombre de columna a una forma simple para compararlo."""
    if nombre is None:
        return ""
    texto = str(nombre).strip().lower()
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r'[^a-z0-9]+', '', texto)
    return texto


def encontrar_columna(df, aliases):
    """Busca una columna equivalente dentro del DataFrame usando alias posibles."""
    for alias in aliases:
        alias_norm = normalizar_nombre_columna(alias)
        for columna in df.columns:
            columna_norm = normalizar_nombre_columna(columna)
            if columna_norm == alias_norm:
                return columna
            if alias_norm and alias_norm in columna_norm:
                return columna
    return None


def cargar_dataframe_desde_archivo(archivo_entrada):
    """Carga el archivo de entrada y detecta encabezados incluso si vienen en otra fila."""
    ext = os.path.splitext(archivo_entrada)[1].lower()

    if ext in {'.xlsx', '.xls'}:
        try:
            excel_file = pd.ExcelFile(archivo_entrada, engine='openpyxl' if ext == '.xlsx' else 'xlrd')
        except Exception as exc:
            raise RuntimeError(f"No se pudo abrir el archivo Excel: {exc}") from exc

        alias_sets = {
            columna: {normalizar_nombre_columna(alias) for alias in aliases}
            for columna, aliases in COLUMNAS_REQUERIDAS.items()
        }

        for sheet_name in excel_file.sheet_names:
            try:
                raw_df = pd.read_excel(archivo_entrada, sheet_name=sheet_name, header=None)
            except Exception:
                continue

            if raw_df.empty:
                continue

            mejor_fila_idx = None
            mejor_coincidencias = []

            for fila_idx in range(min(15, len(raw_df))):
                fila = raw_df.iloc[fila_idx].fillna("")
                nombres = [normalizar_valor_texto(celda) for celda in fila.tolist()]
                coincidencias = []
                for columna_requerida, alias_set in alias_sets.items():
                    if any(normalizar_nombre_columna(nombre) in alias_set for nombre in nombres):
                        coincidencias.append(columna_requerida)

                if len(coincidencias) > len(mejor_coincidencias):
                    mejor_fila_idx = fila_idx
                    mejor_coincidencias = coincidencias

                if len(coincidencias) >= 4:
                    break

            if mejor_fila_idx is not None and mejor_coincidencias:
                fila = raw_df.iloc[mejor_fila_idx].fillna("")
                nombres = [normalizar_valor_texto(celda) for celda in fila.tolist()]
                datos = raw_df.iloc[mejor_fila_idx + 1:].copy()
                datos.columns = nombres
                datos = datos.loc[:, ~datos.columns.duplicated()].copy()
                return datos

        # Fallback simple.
        try:
            return pd.read_excel(archivo_entrada, sheet_name=0)
        except Exception as exc:
            raise RuntimeError(f"No se pudo leer el archivo Excel: {exc}") from exc

    try:
        return pd.read_csv(archivo_entrada, sep='\t', encoding='utf-8', low_memory=False)
    except Exception as exc:
        raise RuntimeError(f"No se pudo leer el archivo de entrada: {exc}") from exc


def normalizar_valor_texto(valor):
    """Convierte un valor a texto limpio para comparar encabezados."""
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def limpiar_sku(sku):
    """Normaliza un SKU para que quede en un formato consistente."""
    # Si el valor está vacío o nulo, se devuelve None.
    if pd.isna(sku):
        return None

    # Se elimina espacio en blanco y se convierte a mayúsculas.
    return str(sku).strip().upper()


def limpiar_nombre_producto(nombre):
    """Limpia el nombre del producto para dejarlo más legible."""
    # Si falta el nombre, se asigna un valor por defecto.
    if pd.isna(nombre):
        return "SIN NOMBRE"

    # Se quitan espacios extras y se deja el texto limpio.
    nombre = str(nombre).strip()
    nombre = " ".join(nombre.split())
    return nombre


def preparar_ruta_salida(archivo_salida=None):
    """Construye la ruta de salida dentro de la carpeta Procesado con fecha y hora."""
    if archivo_salida:
        return archivo_salida

    carpeta_procesado = os.path.join(os.getcwd(), 'Procesado')
    os.makedirs(carpeta_procesado, exist_ok=True)

    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    nombre_base = f'Reporte_Pendientes_Meli_{timestamp}.xlsx'
    return os.path.join(carpeta_procesado, nombre_base)


def procesar_ventas(archivo_entrada, archivo_salida=None):
    """
    Procesa el archivo de ventas y genera el reporte final.

    Flujo principal del script:
    1. Carga el archivo fuente.
    2. Verifica que tenga las columnas necesarias.
    3. Filtra las ventas que correspondan al estado esperado.
    4. Limpia los datos para trabajar con ellos con mayor seguridad.
    5. Agrupa y resume la información por SKU.
    6. Muestra el resultado y opcionalmente lo guarda en Excel o CSV.
    """

    # Encabezado general del proceso para identificar la ejecución.
    print("=" * 70)
    print("  REPORTE DE VENTAS SIN PROCESAR - MERCADO LIBRE")
    print("=" * 70)
    print(f"[INFO] Fecha de procesamiento: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[INFO] Archivo fuente: {archivo_entrada}")

    # -------------------------------------------------------------------------
    # PASO 1: CARGAR DATOS DESDE EL ARCHIVO DE ENTRADA
    # -------------------------------------------------------------------------
    # Aquí se lee el archivo TSV/TXT con pandas para convertirlo en una tabla.
    print("\n[1/5] Cargando datos desde archivo...")
    try:
        df = cargar_dataframe_desde_archivo(archivo_entrada)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        df.columns = [normalizar_valor_texto(col) for col in df.columns]
        print(f"       [OK] {len(df)} transacciones cargadas")
        print(f"       [OK] {len(df.columns)} columnas detectadas")
    except Exception as e:
        print(f"       [ERROR] Error al cargar archivo: {e}")
        raise

    # -------------------------------------------------------------------------
    # PASO 2: VALIDAR QUE LAS COLUMNAS REQUERIDAS EXISTAN
    # -------------------------------------------------------------------------
    # Se comprueba que el archivo tenga los campos mínimos para trabajar.
    print("\n[2/5] Validando estructura de columnas...")
    columnas_renombradas = {}
    columnas_faltantes = []

    for columna_requerida, aliases in COLUMNAS_REQUERIDAS.items():
        columna_encontrada = encontrar_columna(df, aliases)
        if columna_encontrada is None:
            columnas_faltantes.append(columna_requerida)
        else:
            columnas_renombradas[columna_encontrada] = columna_requerida

    if columnas_faltantes:
        print(f"       [ERROR] Columnas faltantes: {columnas_faltantes}")
        print(f"       [DEBUG] Columnas disponibles: {list(df.columns)}")
        raise ValueError(f"Columnas requeridas no encontradas: {columnas_faltantes}")

    df = df.rename(columns=columnas_renombradas)
    columnas_esperadas = list(COLUMNAS_REQUERIDAS.keys())
    df = df[[col for col in columnas_esperadas if col in df.columns]].copy()
    df.columns = [col.strip() for col in df.columns]
    print(f"       [OK] Todas las columnas requeridas presentes")

    # -------------------------------------------------------------------------
    # PASO 3: FILTRAR SOLO LAS VENTAS DEL ESTADO INTERESADO
    # -------------------------------------------------------------------------
    # Se dejan solo las filas que correspondan al estado de interés.
    print(f"\n[3/5] Filtrando por estado '{ESTADO_FILTRO}'...")

    # Se revisan los estados presentes para confirmar qué trae el archivo.
    df['Estado'] = df['Estado'].fillna('').astype(str).str.strip()
    df['Estado'] = df['Estado'].replace(r'\s+', ' ', regex=True)
    estados_unicos = df['Estado'].dropna().unique()
    print(f"       Estados encontrados: {list(estados_unicos)}")

    # Se crea una copia del DataFrame solo con las filas válidas.
    df_filtrado = df[df['Estado'].str.lower() == ESTADO_FILTRO.lower()].copy()
    print(f"       [OK] {len(df_filtrado)} transacciones con estado '{ESTADO_FILTRO}'")

    # Si no hay filas después del filtro, se termina de forma limpia.
    if len(df_filtrado) == 0:
        print("       [AVISO] No hay transacciones para procesar")
        return pd.DataFrame(columns=COLUMNAS_SALIDA)

    # -------------------------------------------------------------------------
    # PASO 4: LIMPIEZA Y NORMALIZACIÓN DE LOS DATOS
    # -------------------------------------------------------------------------
    # Aquí se corrigen valores como SKU vacíos, nombres mal formateados y unidades no numéricas.
    print("\n[4/5] Limpiando y normalizando datos...")

    # Se eliminan filas sin SKU porque no aportan información útil para el resumen.
    df_filtrado = df_filtrado[df_filtrado['SKU'].notna()].copy()
    print(f"       [OK] {len(df_filtrado)} transacciones con SKU válido")

    # Se limpian los SKU y los nombres de producto para estandarizarlos.
    df_filtrado['SKU_LIMPIO'] = df_filtrado['SKU'].apply(limpiar_sku)
    df_filtrado['PRODUCTO_LIMPIO'] = df_filtrado['Título de la publicación'].apply(limpiar_nombre_producto)

    # Se convierten las unidades a valor numérico para poder sumarlas correctamente.
    df_filtrado['Unidades'] = pd.to_numeric(df_filtrado['Unidades'], errors='coerce').fillna(0)
    print(f"       [OK] Datos numericos normalizados")

    # -------------------------------------------------------------------------
    # PASO 5: AGRUPAR POR SKU Y RESUMIR LAS UNIDADES PENDIENTES
    # -------------------------------------------------------------------------
    # Se agrupan las filas por SKU para obtener un reporte resumido por producto.
    print("\n[5/5] Agrupando por SKU y sumando unidades...")

    # Para cada SKU se toma el primer nombre encontrado y se suman las unidades.
    reporte = df_filtrado.groupby('SKU_LIMPIO').agg({
        'PRODUCTO_LIMPIO': 'first',
        'Unidades': 'sum'
    }).reset_index()

    # Se renombra la tabla final con los nombres pedidos por el usuario.
    reporte.columns = ['CODIGO', 'PRODUCTO', 'Suma de PENDIENTE MELI']

    # Se ordena la tabla por código para que sea más fácil de revisar.
    reporte = reporte.sort_values('CODIGO').reset_index(drop=True)

    print(f"       [OK] {len(reporte)} SKUs unicos encontrados")
    print(f"       [OK] Total de unidades pendientes: {int(reporte['Suma de PENDIENTE MELI'].sum())}")

    # -------------------------------------------------------------------------
    # MOSTRAR EL RESULTADO EN PANTALLA
    # -------------------------------------------------------------------------
    # Se imprime la tabla final para que el usuario pueda verla en consola.
    print("\n" + "=" * 70)
    print("  TABLA RESULTADO")
    print("=" * 70)
    print(reporte.to_string(index=False))
    print("=" * 70)

    # -------------------------------------------------------------------------
    # GUARDAR EL RESULTADO EN EL FORMATO SOLICITADO
    # -------------------------------------------------------------------------
    # Si se indicó un nombre de salida, se guarda el reporte en Excel o CSV.
    if archivo_salida:
        # Se determina la extensión para elegir el formato adecuado.
        ext = os.path.splitext(archivo_salida)[1].lower()

        if ext == '.xlsx':
            try:
                import openpyxl  # noqa: F401
            except ModuleNotFoundError as exc:
                raise SystemExit(
                    "Error: openpyxl no está instalado. Instala las dependencias con:\n"
                    "  pip install pandas openpyxl"
                ) from exc

            reporte.to_excel(archivo_salida, index=False, sheet_name='Pendientes Meli')
            print(f"\n[INFO] Reporte guardado en Excel: {archivo_salida}")
        elif ext == '.csv':
            reporte.to_csv(archivo_salida, index=False, encoding='utf-8-sig')
            print(f"\n[INFO] Reporte guardado en CSV: {archivo_salida}")
        else:
            # Si la extensión no es válida, se asume Excel por defecto.
            archivo_salida = archivo_salida.replace(ext, '.xlsx')
            reporte.to_excel(archivo_salida, index=False, sheet_name='Pendientes Meli')
            print(f"\n[INFO] Reporte guardado en Excel: {archivo_salida}")

    return reporte


def main():
    """Función principal que organiza la ejecución del script."""
    # Se prepara el parser de argumentos para aceptar entrada y salida desde consola.
    parser = argparse.ArgumentParser(
        description='Genera reporte de ventas sin procesar desde Mercado Libre',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python reporte_ventas_meli.py
  python reporte_ventas_meli.py --input ventas.tsv --output reporte.xlsx
  python reporte_ventas_meli.py --input datos.txt --output resultado.csv
        """
    )
    parser.add_argument(
        '--input', '-i',
        help='Ruta al archivo de entrada (TSV/TXT de Mercado Libre)'
    )
    parser.add_argument(
        '--output', '-o',
        help='Ruta del archivo de salida (.xlsx o .csv)'
    )

    # Se leen los argumentos ingresados por el usuario.
    args = parser.parse_args()

    try:
        # Se busca el archivo de entrada, ya sea por ruta explícita o automáticamente.
        archivo_entrada = encontrar_archivo_entrada(args.input)

        # Si no se indicó salida, se genera un nombre por defecto dentro de la carpeta Procesado.
        if not args.output:
            args.output = preparar_ruta_salida()

        # Se ejecuta el procesamiento principal del reporte.
        reporte = procesar_ventas(archivo_entrada, args.output)

        print("\n✅ PROCESO COMPLETADO EXITOSAMENTE")

    except FileNotFoundError as e:
        print(f"\n❌ ERROR: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"\n[ERROR] ERROR DE VALIDACION: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] ERROR INESPERADO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


# =============================================================================
# EJECUCIÓN DIRECTA (SIN ARGUMENTOS)
# =============================================================================
# Este bloque permite ejecutar el script de forma simple, sin necesidad de pasar parámetros.
if __name__ == "__main__":
    # Si se ejecuta sin argumentos, se usa el modo interactivo con detección automática del archivo.
    if len(sys.argv) == 1:
        print("=" * 70)
        print("  MODO INTERACTIVO - Procesando archivo detectado automáticamente")
        print("=" * 70)
        try:
            # Se busca el archivo automáticamente y se genera un nombre de salida por defecto.
            archivo_entrada = encontrar_archivo_entrada()
            archivo_salida = preparar_ruta_salida()
            procesar_ventas(archivo_entrada, archivo_salida)
            print("\n[OK] PROCESO COMPLETADO EXITOSAMENTE")
        except Exception as e:
            print(f"\n[ERROR] {e}")
            sys.exit(1)
    else:
        # Si se pasaron argumentos, se usa la función principal con parser.
        main()
