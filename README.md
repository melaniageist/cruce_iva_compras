# 🧾 Cruce IVA Compras

Herramienta para cruzar registros de IVA Compras entre el sistema de contabilidad (Contabilium) y los registros de ARCA, detectando matches, diferencias y comprobantes faltantes en cada fuente.

---

## ¿Qué hace?

Dado un archivo Excel de **Contabilidad** y uno de **ARCA**, el programa:

1. Valida que ambos archivos tengan la estructura correcta.
2. Cruza los comprobantes por: `fecha`, `tipo_comprobante`, `punto_de_venta`, `nro_factura` y `cuit`.
3. Clasifica cada comprobante en uno de cuatro estados:
   - ✅ **Match** — figura en ambos archivos con los mismos importes.
   - ⚠️ **Diferencia** — figura en ambos pero con importe o fecha distintos.
   - ❌ **Solo Contabilidad** — está en Contabilidad pero no en ARCA.
   - 🔵 **Solo ARCA** — está en ARCA pero no en Contabilidad.
4. Muestra los resultados en una tabla filtrable e interactiva.
5. Exporta un Excel con una hoja por estado y hoja de resumen.

---

## Estructura del proyecto

```
cruce_iva/
│
├── main.py              # Punto de entrada — interfaz Streamlit
├── cruce.py             # Lógica del cruce
├── validar.py           # Validación de archivos de entrada
├── exportar.py          # Generación del Excel de resultados
│
├── ejemplos/
│   ├── plantilla_contabilidad.xlsx   # Plantilla vacía para Contabilidad
│   ├── plantilla_arca.xlsx           # Plantilla vacía para ARCA
│   ├── contabilidad_ejemplo.xlsx     # Datos ficticios de ejemplo
│   └── arca_ejemplo.xlsx             # Datos ficticios de ejemplo
│
├── requirements.txt     # Dependencias del proyecto
└── README.md            # Este archivo
```

---

## Formato de los archivos

Ambos archivos Excel deben tener exactamente estas columnas (en cualquier orden):

| Columna | Tipo | Descripción |
|---|---|---|
| `fecha` | texto | Fecha del comprobante en formato DD/MM/AAAA |
| `tipo_comprobante` | texto | FAC A, FAC B, NC A, etc. |
| `punto_de_venta` | texto | Con ceros a la izquierda (ej: 0001) |
| `nro_factura` | texto | Con ceros a la izquierda (ej: 00000123) |
| `cuit` | texto | En formato XX-XXXXXXXX-X |
| `nombre_empresa` | texto | Razón social (no se usa para el cruce) |
| `importe_gravado` | número | Base imponible gravada |
| `importe_no_gravado` | número | Importes no gravados o exentos |
| `iva_21` | número | IVA al 21% |
| `iva_27` | número | IVA al 27% |
| `iva_105` | número | IVA al 10,5% |
| `importe_total` | número | Total del comprobante |

> 💡 Descargá las plantillas vacías desde el sidebar de la app para preparar tus archivos.

---

## Cómo usar la app (versión web)

Accedé directamente desde el navegador, sin instalar nada:

🔗 **[Abrir Cruce IVA Compras](https://cruceiva.streamlit.app)**

---

## Instalación local

### Requisitos
- Python 3.10 o superior
- Git

### Pasos

```bash
# 1 — Clonar el repositorio
git clone https://github.com/tu-usuario/cruce_iva_compras.git
cd cruce_iva_compras

# 2 — Crear el entorno virtual
python -m venv venv

# 3 — Activar el entorno virtual
# Windows:
venv\Scripts\activate
# Mac / Linux:
source venv/bin/activate

# 4 — Instalar dependencias
pip install -r requirements.txt

# 5 — Correr la app
streamlit run main.py
```

La app se abre automáticamente en `http://localhost:8501`

---

## Librerías utilizadas

| Librería | Versión | Uso |
|---|---|---|
| `streamlit` | 1.58.0 | Interfaz web interactiva |
| `pandas` | 3.0.2 | Lectura y procesamiento de archivos Excel |
| `openpyxl` | 3.1.5 | Exportación de Excel con formato y colores |

---

## Módulos propios

| Módulo | Descripción |
|---|---|
| `cruce.py` | Lógica de cruce, normalización y detección de diferencias |
| `validar.py` | Validación de estructura de archivos de entrada |
| `exportar.py` | Generación del Excel de resultados con formato |

---

## Autora

Melania Geist
Proyecto desarrollado como trabajo final de **Programación 1**  
Carrera: Ciencia de Datos — ISTEA  
Ciclo lectivo: 2026