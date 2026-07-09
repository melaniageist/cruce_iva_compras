"""
Punto de entrada de la aplicación Cruce IVA Compras.
Interfaz web construida con Streamlit.

Uso:
    streamlit run main.py
"""

import streamlit as st
import pandas as pd
from cruce import COLUMNAS_CRUCE, cruzar, resumen
from validar import cargar_y_validar
from exportar import exportar_excel


# ── Configuración de la página ─────────────────────────────────────────────
st.set_page_config(
    page_title="Cruce IVA Compras",
    page_icon="🧾",
    layout="wide",
)

# ── Estilos personalizados ─────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .estado-badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 700;
    }
    div[data-testid="metric-container"] {
        background: #f8f9fa;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 12px;
    }
</style>
""", unsafe_allow_html=True)

COLORES_ESTADO = {
    "match":             {"bg": "#e6f4ea", "color": "#1e8c45", "label": "✅ Match"},
    "diferencia":        {"bg": "#fff8e1", "color": "#b06000", "label": "⚠️ Diferencia"},
    "solo_contabilidad": {"bg": "#fce8e6", "color": "#c5221f", "label": "❌ Solo Contabilidad"},
    "solo_arca":         {"bg": "#e8f0fe", "color": "#1a73e8", "label": "🔵 Solo ARCA"},
}

COLUMNAS_CRUCE_DISPLAY = [
    "fecha", "tipo_comprobante", "punto_de_venta",
    "nro_factura", "cuit", "nombre_empresa",
    "importe_gravado", "importe_no_gravado",
    "iva_21", "iva_27", "iva_105", "importe_total",
]


def badge(estado: str) -> str:

    """Devuelve HTML de un badge de color según el estado del comprobante."""

    cfg = COLORES_ESTADO[estado]
    return (
        f'<span class="estado-badge" '
        f'style="background:{cfg["bg"]};color:{cfg["color"]}">'
        f'{cfg["label"]}</span>'
    )


def tabla_resultados(resultados: list[dict], filtro: str) -> None:
    """
    Muestra la tabla de resultados filtrada por estado.

    Args:
        resultados: Lista de resultados del cruce.
        filtro: Estado por el cual filtrar ("todos" o un estado específico).
    """
    
    filas = resultados if filtro == "todos" else [r for r in resultados if r["estado"] == filtro]

    if not filas:
        st.info("No hay comprobantes para este filtro.")
        return

    registros = []
    for r in filas:
        datos = r["contabilidad"] or r["arca"]
        clave = r["clave"]
        fila = dict(zip(COLUMNAS_CRUCE, clave))
        fila.update(datos or {})
        fila["estado"] = COLORES_ESTADO[r["estado"]]["label"]
        fila["diferencias"] = ", ".join(r["diferencias"]) if r["diferencias"] else "—"
        registros.append(fila)

    df_display = pd.DataFrame(registros)

    cols_mostrar = [
        "estado", "fecha", "tipo_comprobante", "punto_de_venta",
        "nro_factura", "cuit", "nombre_empresa", "importe_total", "diferencias"
    ]
    cols_presentes = [c for c in cols_mostrar if c in df_display.columns]
    st.dataframe(
        df_display[cols_presentes],
        use_container_width=True,
        hide_index=True,
        column_config={
            "estado":           st.column_config.TextColumn("Estado", width=140),
            "fecha":            st.column_config.TextColumn("Fecha", width=100),
            "tipo_comprobante": st.column_config.TextColumn("Tipo", width=80),
            "punto_de_venta":   st.column_config.TextColumn("PDV", width=70),
            "nro_factura":      st.column_config.TextColumn("Nro. Factura", width=120),
            "cuit":             st.column_config.TextColumn("CUIT", width=150),
            "nombre_empresa":   st.column_config.TextColumn("Nombre"),
            "importe_total":    st.column_config.NumberColumn("Total", format="$%.2f", width=120),
            "diferencias":      st.column_config.TextColumn("Campos con diferencia"),
        }
    )
    st.caption(f"{len(filas)} comprobante(s) mostrados.")


# ── Layout principal ───────────────────────────────────────────────────────

st.title("🧾 Cruce IVA Compras")
st.markdown("**Contabilidad vs. ARCA** — Cargá los dos archivos para iniciar el cruce.")

# ── Carga de archivos ──────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📒 Contabilidad")
    archivo_cont = st.file_uploader(
        "Subí el archivo de Contabilidad",
        type=["xlsx"],
        key="cont",
        help="Debe tener el formato de la plantilla. Descargala desde el sidebar."
    )

with col2:
    st.subheader("🏛 ARCA")
    archivo_arca = st.file_uploader(
        "Subí el archivo de ARCA",
        type=["xlsx"],
        key="arca",
        help="Debe tener el formato de la plantilla. Descargala desde el sidebar."
    )

# ── Sidebar: plantillas y ayuda ────────────────────────────────────────────
with st.sidebar:
    st.header("📥 Plantillas")
    st.markdown("Usá estas plantillas para preparar tus archivos:")

    for nombre_archivo in ["plantilla_contabilidad.xlsx", "plantilla_arca.xlsx"]:
        ruta = f"ejemplos/{nombre_archivo}"
        try:
            with open(ruta, "rb") as f:
                datos = f.read()
            st.download_button(
                label=f"⬇️ {nombre_archivo}",
                data=datos,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{nombre_archivo}",
            )
        except FileNotFoundError:
            st.warning(f"No se encontró {nombre_archivo}")

    st.divider()
    st.header("🧪 Archivos de ejemplo")
    st.markdown("Probá el programa con estos datos ficticios:")

    for nombre_archivo in ["contabilidad_ejemplo.xlsx", "arca_ejemplo.xlsx"]:
        ruta = f"ejemplos/{nombre_archivo}"
        try:
            with open(ruta, "rb") as f:
                datos = f.read()
            st.download_button(
                label=f"⬇️ {nombre_archivo}",
                data=datos,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{nombre_archivo}",
            )
        except FileNotFoundError:
            st.warning(f"No se encontró {nombre_archivo}")

    st.divider()
    st.header("ℹ️ Columnas esperadas")
    st.markdown("""
    Ambos archivos deben tener exactamente estas columnas:
    - `fecha`
    - `tipo_comprobante`
    - `punto_de_venta`
    - `nro_factura`
    - `cuit`
    - `nombre_empresa`
    - `importe_gravado`
    - `importe_no_gravado`
    - `iva_21`
    - `iva_27`
    - `iva_105`
    - `importe_total`

    El cruce se realiza por todas las columnas **excepto** `tipo_comprobante` y `nombre_empresa`.
    """)

# ── Procesamiento ──────────────────────────────────────────────────────────
if archivo_cont and archivo_arca:
    st.divider()

    with st.spinner("Validando y procesando archivos..."):
        df_cont, errores_cont = cargar_y_validar(archivo_cont, "Contabilidad")
        df_arca, errores_arca = cargar_y_validar(archivo_arca, "ARCA")

    # Mostrar errores de validación si los hay
    if errores_cont:
        st.error("**Errores en el archivo de Contabilidad:**")
        for e in errores_cont:
            st.markdown(f"- {e}")

    if errores_arca:
        st.error("**Errores en el archivo de ARCA:**")
        for e in errores_arca:
            st.markdown(f"- {e}")

    if errores_cont or errores_arca:
        st.stop()

    # Cruce
    resultados = cruzar(df_cont, df_arca)
    datos_resumen = resumen(resultados)

    # ── Tarjetas de resumen ────────────────────────────────────────────────
    st.subheader("📊 Resumen del cruce")
    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Total", datos_resumen["total"])
    c2.metric("✅ Match", datos_resumen["match"])
    c3.metric("⚠️ Diferencias", datos_resumen["diferencia"])
    c4.metric("❌ Solo Contab.", datos_resumen["solo_contabilidad"])
    c5.metric("🔵 Solo ARCA", datos_resumen["solo_arca"])

    # ── Filtros y tabla ────────────────────────────────────────────────────
    st.subheader("📋 Detalle por comprobante")

    opciones = {
        "Todos":                "todos",
        "✅ Match":             "match",
        "⚠️ Diferencias":      "diferencia",
        "❌ Solo Contabilidad": "solo_contabilidad",
        "🔵 Solo ARCA":        "solo_arca",
    }

    col_filtro, col_busqueda = st.columns([3, 2])
    with col_filtro:
        filtro_label = st.radio(
            "Filtrar por estado:",
            options=list(opciones.keys()),
            horizontal=True,
        )
    with col_busqueda:
        busqueda = st.text_input("🔍 Buscar por CUIT, nro. factura o nombre", "")

    filtro = opciones[filtro_label]

    # Aplicar búsqueda de texto
    resultados_filtrados = resultados
    if busqueda.strip():
        q = busqueda.strip().upper()
        resultados_filtrados = [
            r for r in resultados
            if any(
                q in str(v).upper()
                for v in ((r["contabilidad"] or r["arca"]) or {}).values()
            )
            or any(q in str(v).upper() for v in r["clave"])
        ]

    tabla_resultados(resultados_filtrados, filtro)

    # ── Exportar ───────────────────────────────────────────────────────────
    st.divider()
    st.subheader("📥 Exportar resultados")

    col_exp1, col_exp2 = st.columns([2, 3])
    with col_exp1:
        if st.button("Generar Excel de resultados", type="primary", use_container_width=True):
            with st.spinner("Generando archivo..."):
                excel_bytes = exportar_excel(resultados, datos_resumen)
            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_bytes,
                file_name="cruce_iva_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with col_exp2:
        st.info("El Excel incluye una hoja por estado (Match, Diferencias, Solo Contabilidad, Solo ARCA) y una hoja de Resumen.")

else:
    st.info("👆 Cargá los dos archivos para iniciar el cruce.")