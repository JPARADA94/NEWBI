import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ===================== CONFIGURACIÓN =====================
st.set_page_config(page_title="Validación estricta de encabezados", layout="wide")
st.title("📄 Construcción de Excel – Todas las columnas requeridas (estricto)")

# ===================== UTILIDADES =====================
def df_to_xlsx_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
    buf.seek(0)
    return buf

def normalizar(col):
    return (
        str(col)
        .strip()
        .replace("–", "-")
        .replace("μ", "Μ")
        .replace("  ", " ")
        .upper()
    )

# ===================== COLUMNAS BASE (REQUERIDAS) =====================
REQUERIDOS = [
    "NOMBRE_CLIENTE","NOMBRE_OPERACION","N_MUESTRA","CORRELATIVO","FECHA_MUESTREO","FECHA_INGRESO",
    "FECHA_RECEPCION","FECHA_INFORME","EDAD_COMPONENTE","UNIDAD_EDAD_COMPONENTE","EDAD_PRODUCTO",
    "UNIDAD_EDAD_PRODUCTO","CANTIDAD_ADICIONADA","UNIDAD_CANTIDAD_ADICIONADA","PRODUCTO","TIPO_PRODUCTO",
    "EQUIPO","TIPO_EQUIPO","MARCA_EQUIPO","MODELO_EQUIPO","COMPONENTE","MARCA_COMPONENTE","MODELO_COMPONENTE",
    "DESCRIPTOR_COMPONENTE","ESTADO_REPORTE","NIVEL_DE_SERVICIO",
    "ÍNDICE PQ (PQI) - 3","PLATA (AG) - 19","ALUMINIO (AL) - 20","CROMO (CR) - 24",
    "COBRE (CU) - 25","HIERRO (FE) - 26","TITANIO (TI) - 38","PLOMO (PB) - 35",
    "NÍQUEL (NI) - 32","MOLIBDENO (MO) - 30","SILICIO (SI) - 36","SODIO (NA) - 31",
    "POTASIO (K) - 27","VANADIO (V) - 39","BORO (B) - 18","BARIO (BA) - 21",
    "CALCIO (CA) - 22","CADMIO (CD) - 23","MAGNESIO (MG) - 28","MANGANESO (MN) - 29",
    "FÓSFORO (P) - 34","ZINC (ZN) - 40","CÓDIGO ISO (4/6/14) - 47",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49","CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48","OXIDACIÓN - 80","NITRACIÓN - 82",
    "NÚMERO ÁCIDO (AN) - 43","NÚMERO BÁSICO (BN) - 12","NÚMERO BÁSICO (BN) - 17",
    "HOLLÍN - 79","DILUCIÓN POR COMBUSTIBLE - 46","AGUA (IR) - 81",
    "CONTENIDO AGUA (KARL FISCHER) - 41","CONTENIDO GLICOL - 105",
    "VISCOSIDAD A 100 °C - 13","VISCOSIDAD A 40 °C - 14",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51","AGUA CUALITATIVA (PLANCHA) - 360",
    "AGUA LIBRE - 416","ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45","COBRE (CU) - 119",
    "ESPUMA SEC 1 - ESTABILIDAD - 60","ESPUMA SEC 1 - TENDENCIA - 59",
    "ESTAÑO (SN) - 37","ÍNDICE VISCOSIDAD - 359","RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "**ULTRACENTRÍFUGA (UC) - 1",
    "ESTADO_PRODUCTO","ESTADO_DESGASTE","ESTADO_CONTAMINACION",
    "N_SOLICITUD","CAMBIO_DE_PRODUCTO","CAMBIO_DE_FILTRO",
    "TEMPERATURA_RESERVORIO","UNIDAD_TEMPERATURA_RESERVORIO",
    "COMENTARIO_CLIENTE","TIPO_DE_COMBUSTIBLE","TIPO_DE_REFRIGERANTE",
    "USUARIO","COMENTARIO_REPORTE","id_muestra"
]

# ===================== COLUMNAS ESTADO (REQUERIDAS) =====================
NUEVAS_ESTADO = [
    "ESTADO_MUESTRA",
    "AGUA CUALITATIVA (PLANCHA) - 360 - Estado",
    "AGUA (IR) - 81 - Estado",
    "ALUMINIO (AL) - 20 - Estado",
    "HIERRO (FE) - 26 - Estado",
    "OXIDACIÓN - 80 - Estado",
    "NITRACIÓN - 82 - Estado",
    "VISCOSIDAD A 40 °C - 14 - Estado",
    "VISCOSIDAD A 100 °C - 13 - Estado"
]

TODAS_REQUERIDAS = REQUERIDOS + NUEVAS_ESTADO

# ===================== CARGA DE ARCHIVOS =====================
files = st.file_uploader(
    "📤 Sube uno o varios Excel (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if files:
    dfs = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")

        # Mapa normalizado → nombre real
        cols_norm = {normalizar(c): c for c in df.columns}

        # ================= VALIDACIÓN ESTRICTA =================
        faltantes = [
            col for col in TODAS_REQUERIDAS
            if normalizar(col) not in cols_norm
        ]

        if faltantes:
            st.error(f"❌ Archivo {f.name} NO cumple con los encabezados requeridos")
            st.dataframe(pd.DataFrame({"Columna faltante": faltantes}))
            st.stop()

        # ================= CONSTRUCCIÓN ORDENADA =================
        df_out = pd.DataFrame()

        # BASE
        for col in REQUERIDOS:
            real = cols_norm[normalizar(col)]
            df_out[col] = df[real]

        # Renombre puntual
        if "ESTADO_REPORTE" in df_out.columns:
            df_out.rename(columns={"ESTADO_REPORTE": "ESTADO"}, inplace=True)

        # Archivo origen
        df_out["Archivo_Origen"] = f.name

        # ESTADOS
        for col in NUEVAS_ESTADO:
            real = cols_norm[normalizar(col)]
            df_out[col] = df[real]

        dfs.append(df_out)

    df_final = pd.concat(dfs, ignore_index=True)

    st.success("✅ Archivo generado correctamente (validación estricta)")
    st.dataframe(df_final.head(20), use_container_width=True)

    nombre = f"resultado_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "📥 Descargar archivo final",
        df_to_xlsx_bytes(df_final),
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


