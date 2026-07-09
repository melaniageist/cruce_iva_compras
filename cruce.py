"""
cruce.py
--------
Módulo propio del proyecto Cruce IVA Compras.
Contiene la lógica de cruce entre el archivo de Contabilidad y el de ARCA.
"""

import pandas as pd
from datetime import datetime


# Columnas que se usan para identificar un comprobante de forma única
# Nota: tipo_comprobante se excluye del cruce porque puede diferir
# entre sistemas (ej. "FAC A" vs "FACTURA A")
COLUMNAS_CRUCE = [
    "fecha",
    "punto_de_venta",
    "nro_factura",
    "cuit",
]

# Columnas de importes donde se detectan diferencias
COLUMNAS_IMPORTES = [
    "importe_gravado",
    "importe_no_gravado",
    "iva_21",
    "iva_27",
    "iva_105",
    "importe_total",
]

# Todas las columnas esperadas en los archivos de entrada
COLUMNAS_ESPERADAS = COLUMNAS_CRUCE + ["tipo_comprobante", "nombre_empresa"] + COLUMNAS_IMPORTES


def validar_columnas(df: pd.DataFrame, nombre_archivo: str) -> list[str]:
    """
    Verifica que el DataFrame tenga todas las columnas esperadas.

    Args:
        df: DataFrame leído desde el archivo Excel.
        nombre_archivo: Nombre del archivo para mostrar en los mensajes de error.

    Returns:
        Lista de columnas faltantes. Lista vacía si el archivo es válido.
    """
    faltantes = [col for col in COLUMNAS_ESPERADAS if col not in df.columns]
    return faltantes

def _normalizar_fecha(valor) -> str:

    """
    Normaliza un valor de fecha eliminando la parte de hora si existe.

    Acepta formatos como:
    - '01/06/2025'
    - '01/06/2025 00:00:00'
    - '2025-06-01 00:00:00'
    - '2025-06-01'

    Args:
        valor: Cadena de texto con la fecha a normalizar.

    Returns:
        Fecha en formato DD/MM/AAAA como string.
    """
    
    if pd.isna(valor):
        return ""

    # Si ya es una fecha de Excel
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%d/%m/%Y")

    texto = str(valor).strip()

    # Si viene con hora
    texto = texto.split(" ")[0]

    # Si ya está DD/MM/AAAA
    if "/" in texto:
        try:
            return datetime.strptime(texto, "%d/%m/%Y").strftime("%d/%m/%Y")
        except ValueError:
            pass

    # Si viene AAAA-MM-DD
    if "-" in texto:
        try:
            return datetime.strptime(texto, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass

    return texto


def normalizar(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza el DataFrame para asegurar consistencia antes del cruce.

    Convierte columnas de texto a mayúsculas sin espacios extra,
    normaliza fechas eliminando la parte de hora, y convierte
    columnas numéricas a float redondeado a 2 decimales.

    Args:
        df: DataFrame a normalizar.

    Returns:
        DataFrame normalizado.
    """
    df = df.copy()

    # Normalizar fecha: eliminar hora si viene con timestamp (ej. ARCA)
    if "fecha" in df.columns:
        df["fecha"] = df["fecha"].apply(_normalizar_fecha)

    # Normalizar columnas de texto del cruce (excepto fecha, ya procesada)
    columnas_texto = [c for c in COLUMNAS_CRUCE if c != "fecha"] + ["tipo_comprobante", "nombre_empresa"]
    for col in columnas_texto:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    # Normalizar columnas numéricas
    for col in COLUMNAS_IMPORTES:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round(2)

    return df


def cruzar(df_cont: pd.DataFrame, df_arca: pd.DataFrame) -> list[dict]:
    """
    Realiza el cruce entre los datos de Contabilidad y ARCA.

    Compara fila por fila usando las columnas clave (fecha, PDV,
    número de factura y CUIT). Para cada comprobante determina si hay
    match, diferencia de importes, o si está solo en uno de los dos archivos.

    Args:
        df_cont: DataFrame con los datos de Contabilidad, ya normalizado.
        df_arca: DataFrame con los datos de ARCA, ya normalizado.

    Returns:
        Lista de diccionarios, uno por comprobante, con las claves:
        - estado: "match" | "diferencia" | "solo_contabilidad" | "solo_arca"
        - contabilidad: dict con los datos del lado contabilidad (o None)
        - arca: dict con los datos del lado ARCA (o None)
        - diferencias: lista de columnas con diferencia de importe
    """
    print("=== CONTABILIDAD ===")
    print(df_cont[["fecha", "punto_de_venta", "nro_factura", "cuit"]].head())

    print("\n=== ARCA ===")
    print(df_arca[["fecha", "punto_de_venta", "nro_factura", "cuit"]].head())

    df_cont = df_cont.set_index(COLUMNAS_CRUCE)
    df_arca = df_arca.set_index(COLUMNAS_CRUCE)

    print("\nPrimer índice Contabilidad:")
    print(repr(df_cont.index[0]))

    print("\nPrimer índice ARCA:")
    print(repr(df_arca.index[0]))

    todos_los_indices = df_cont.index.union(df_arca.index)
    resultados = []

    for clave in todos_los_indices:
        en_cont = clave in df_cont.index
        en_arca = clave in df_arca.index

        fila_cont = df_cont.loc[clave].to_dict() if en_cont else None
        fila_arca = df_arca.loc[clave].to_dict() if en_arca else None

        if en_cont and en_arca:
            diferencias = _detectar_diferencias(fila_cont, fila_arca)
            estado = "diferencia" if diferencias else "match"
        elif en_cont:
            estado = "solo_contabilidad"
            diferencias = []
        else:
            estado = "solo_arca"
            diferencias = []

        resultados.append({
            "estado": estado,
            "clave": clave,
            "contabilidad": fila_cont,
            "arca": fila_arca,
            "diferencias": diferencias,
        })

    return resultados




def _detectar_diferencias(fila_cont: dict, fila_arca: dict) -> list[str]:
    """
    Compara los importes entre dos filas y devuelve las columnas que difieren.

    Args:
        fila_cont: Diccionario con los datos de Contabilidad.
        fila_arca: Diccionario con los datos de ARCA.

    Returns:
        Lista de nombres de columnas donde los valores difieren.
    """
    return [
        col for col in COLUMNAS_IMPORTES
        if round(float(fila_cont.get(col, 0)), 2) != round(float(fila_arca.get(col, 0)), 2)
    ]


def resumen(resultados: list[dict]) -> dict:
    """
    Genera un resumen con la cantidad de comprobantes por estado.

    Args:
        resultados: Lista de resultados devuelta por cruzar().

    Returns:
        Diccionario con conteos por estado y total general.
    """
    conteos = {"match": 0, "diferencia": 0, "solo_contabilidad": 0, "solo_arca": 0}
    for r in resultados:
        conteos[r["estado"]] += 1
    conteos["total"] = len(resultados)
    return conteos