"""
validar.py
----------
Módulo de validación de archivos de entrada.
Verifica que los archivos Excel cargados tengan la estructura correcta
antes de intentar el cruce.
"""

import pandas as pd
from cruce import COLUMNAS_ESPERADAS, validar_columnas, normalizar


def cargar_y_validar(ruta: str, nombre: str) -> tuple[pd.DataFrame | None, list[str]]:
    """
    Carga un archivo Excel y valida que tenga las columnas esperadas.

    Args:
        ruta: Ruta al archivo Excel o objeto de archivo (Streamlit UploadedFile).
        nombre: Nombre descriptivo del archivo para los mensajes de error.

    Returns:
        Tupla (DataFrame normalizado, lista de errores).
        Si hay errores, el DataFrame es None.
    """
    errores = []

    try:
        df = pd.read_excel(ruta, skiprows=2, dtype=str, engine="openpyxl")
        df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    except Exception as e:
        return None, [f"No se pudo leer el archivo {nombre}: {e}"]

    faltantes = validar_columnas(df, nombre)
    if faltantes:
        errores.append(
            f"El archivo '{nombre}' no tiene las siguientes columnas: "
            + ", ".join(faltantes)
        )
        errores.append("Revisá la plantilla en ejemplos/plantilla_contabilidad.xlsx o ejemplos/plantilla_arca.xlsx")
        return None, errores

    df = normalizar(df)
    return df, []


def describir_columnas() -> str:
    """
    Devuelve una descripción de las columnas esperadas para mostrar al usuario.

    Returns:
        Texto con la descripción de cada columna.
    """
    descripciones = {
        "fecha": "Fecha del comprobante en formato DD/MM/AAAA",
        "tipo_comprobante": "Tipo de comprobante (FAC A, FAC B, NC A, etc.)",
        "punto_de_venta": "Punto de venta con ceros a la izquierda (ej: 0001)",
        "nro_factura": "Número de factura con ceros a la izquierda (ej: 00000123)",
        "cuit": "CUIT del proveedor en formato XX-XXXXXXXX-X",
        "nombre_empresa": "Nombre o razón social del proveedor (no se usa para el cruce)",
        "importe_gravado": "Base imponible gravada",
        "importe_no_gravado": "Importes no gravados o exentos",
        "iva_21": "IVA al 21%",
        "iva_27": "IVA al 27%",
        "iva_105": "IVA al 10,5%",
        "importe_total": "Total del comprobante",
    }
    lineas = [f"  • {col}: {desc}" for col, desc in descripciones.items()]
    return "\n".join(lineas)