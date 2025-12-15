import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ===================== CONFIGURACIÓN =====================
st.set_page_config(page_title="Control total de columnas", layout="wide")
st.title("📄 Validación estricta de encabezados y control de datos no usados")

# ===================== UTILIDADES =====================
def col_index_to_letter(idx: int) -> str:
    s = ""
    i = int(idx)
    while i >= 0:
        s = chr(i % 26 + 65) + s
        i = i // 26 - 1
    return s

def df_to_xlsx_bytes(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
    buf.seek(0)
    return buf

def normalizar(col: str) -> str:
    return (
        str(col)
        .strip()
        .replace("–", "-")
        .replace("μ", "Μ")
        .replace("  ", " ")
        .upper()
    )

# ===================== ENCABEZADOS BASE (REQUERIDOS) =====================
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

# ===================== ENCABEZADOS ESTADO (REQUERIDOS) =====================
NUEVAS_ESTADO = [
    "ESTADO_MUESTRA",

    # ----- AGUA -----
    "AGUA (IR) - 74",
    "AGUA (IR) - 74 - Estado",
    "AGUA (IR) - 81 - Estado",
    "AGUA LIBRE - 416 - Estado",
    "AGUA CUALITATIVA (PLANCHA) - 360 - Estado",

    # ----- METALES -----
    "ALUMINIO (AL) - 20 - Estado",
    "BARIO (BA) - 21 - Estado",
    "BORO (B) - 18 - Estado",
    "CALCIO (CA) - 22 - Estado",
    "CADMIO (CD) - 23 - Estado",
    "COBRE (CU) - 25 - Estado",
    "COBRE (CU) - 119 - Estado",
    "CROMO (CR) - 24 - Estado",
    "HIERRO (FE) - 26 - Estado",
    "MAGNESIO (MG) - 28 - Estado",
    "MANGANESO (MN) - 29 - Estado",
    "MOLIBDENO (MO) - 30 - Estado",
    "NÍQUEL (NI) - 32 - Estado",
    "PLATA (AG) - 19 - Estado",
    "PLOMO (PB) - 35 - Estado",
    "POTASIO (K) - 27 - Estado",
    "SILICIO (SI) - 36 - Estado",
    "SODIO (NA) - 31 - Estado",
    "TITANIO (TI) - 38 - Estado",
    "VANADIO (V) - 39 - Estado",
    "ZINC (ZN) - 40 - Estado",
    "ESTAÑO (SN) - 37 - Estado", 
    "FÓSFORO (P) - 34 - Estado", 

    # ----- PARTÍCULAS / LIMPIEZA -----
    "CÓDIGO ISO (4/6/14) - 47 - Estado",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49 - Estado",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50 - Estado",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48 - Estado",

    # ----- OXIDACIÓN / NITRACIÓN / PQ -----
    "OXIDACIÓN - 80 - Estado",
    "NITRACIÓN - 82 - Estado",
    "ÍNDICE PQ (PQI) - 3 - Estado",

    # ----- QUÍMICA DEL ACEITE -----
    "NÚMERO ÁCIDO (AN) - 43 - Estado",
    "NÚMERO BÁSICO (BN) - 12 - Estado",
    "NÚMERO BÁSICO (BN) - 17 - Estado",
    "CONTENIDO AGUA (KARL FISCHER) - 41 - Estado",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44 - Estado",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45 - Estado",

    # ----- HOLLÍN / COMBUSTIBLE -----
    "HOLLÍN - 73",
    "HOLLÍN - 73 - Estado",
    "HOLLÍN - 79 - Estado",
    "DILUCIÓN POR COMBUSTIBLE - 46 - Estado",

    # ----- VISCOSIDAD -----
    "VISCOSIDAD A 40 °C - 14 - Estado",
    "VISCOSIDAD A 100 °C - 13 - Estado",
    "ÍNDICE VISCOSIDAD - 359 - Estado",

    # ----- ESPUMA -----
    "ESPUMA SEC 1 - ESTABILIDAD - 60 - Estado",
    "ESPUMA SEC 1 - TENDENCIA - 59 - Estado",

    # ----- MPC / DEPÓSITOS -----
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado",
    "RESIDUO CARBÓN (MCR) - 361",
    "RESIDUO CARBÓN (MCR) - 361 - Estado",

    # ----- SEGURIDAD / ENVEJECIMIENTO -----
    "PUNTO DE INFLAMACIÓN (PMA) - 61",
    "PUNTO DE INFLAMACIÓN (PMA) - 61 - Estado",
    "RPVOT - 10 - Estado",

    # ----- DEMULSIBILIDAD -----
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83 - Estado",

    # ----- ESPECIALES -----
    "**ULTRACENTRÍFUGA (UC) - 1 - Estado"
]


COLUMNAS_USADAS = REQUERIDOS + NUEVAS_ESTADO

# ===================== CARGA DE ARCHIVOS =====================
files = st.file_uploader("📤 Sube uno o varios Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if files:
    dfs_out = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()
        cols_norm = {normalizar(c): c for c in cols}

        # -------- VALIDACIÓN DE ENCABEZADOS --------
        faltantes = [c for c in COLUMNAS_USADAS if normalizar(c) not in cols_norm]
        if faltantes:
            st.error(f"❌ {f.name} – Faltan encabezados requeridos")
            st.dataframe(pd.DataFrame({"Encabezado faltante": faltantes}), use_container_width=True)
            st.stop()

        # -------- DETECCIÓN DE COLUMNAS CON DATOS NO USADAS --------
        usadas_norm = {normalizar(c) for c in COLUMNAS_USADAS}
        extras = []

        for idx, c in enumerate(cols):
            if normalizar(c) in usadas_norm:
                continue
            serie = df[c].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA})
            n = int(serie.notna().sum())
            if n > 0:
                extras.append({
                    "Archivo": f.name,
                    "Encabezado NO usado": c,
                    "Registros con datos": n,
                    "Posición": col_index_to_letter(idx)
                })

        if extras:
            st.warning(f"⚠️ {f.name}: columnas con datos NO usadas en la salida")
            st.dataframe(pd.DataFrame(extras), use_container_width=True)

        # -------- CONSTRUCCIÓN DEL EXCEL FINAL --------
        df_out = pd.DataFrame()
        for c in REQUERIDOS:
            df_out[c] = df[cols_norm[normalizar(c)]]

        df_out.rename(columns={"ESTADO_REPORTE": "ESTADO"}, inplace=True)
        df_out["Archivo_Origen"] = f.name

        for c in NUEVAS_ESTADO:
            df_out[c] = df[cols_norm[normalizar(c)]]

        dfs_out.append(df_out)

    df_final = pd.concat(dfs_out, ignore_index=True)

    st.success("✅ Proceso completado correctamente")
    st.dataframe(df_final.head(20), use_container_width=True)

    nombre = f"resultado_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "📥 Descargar archivo final",
        df_to_xlsx_bytes(df_final),
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

