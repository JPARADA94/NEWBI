import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ===================== CONFIG =====================
st.set_page_config(page_title="Control total de columnas", layout="wide")
st.title("📄 Validación y control: columnas requeridas + columnas con datos no usadas")

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
        .replace("–", "-")   # guion largo
        .replace("μ", "Μ")   # mu
        .replace("  ", " ")
        .upper()
    )

# ===================== ENCABEZADOS REQUERIDOS (BASE) =====================
REQUERIDOS = [
    "NOMBRE_CLIENTE","NOMBRE_OPERACION","N_MUESTRA","CORRELATIVO","FECHA_MUESTREO","FECHA_INGRESO",
    "FECHA_RECEPCION","FECHA_INFORME","EDAD_COMPONENTE","UNIDAD_EDAD_COMPONENTE","EDAD_PRODUCTO",
    "UNIDAD_EDAD_PRODUCTO","CANTIDAD_ADICIONADA","UNIDAD_CANTIDAD_ADICIONADA","PRODUCTO","TIPO_PRODUCTO",
    "EQUIPO","TIPO_EQUIPO","MARCA_EQUIPO","MODELO_EQUIPO","COMPONENTE","MARCA_COMPONENTE","MODELO_COMPONENTE",
    "DESCRIPTOR_COMPONENTE","ESTADO_REPORTE","NIVEL_DE_SERVICIO","ÍNDICE PQ (PQI) - 3","PLATA (AG) - 19",
    "ALUMINIO (AL) - 20","CROMO (CR) - 24","COBRE (CU) - 25","HIERRO (FE) - 26","TITANIO (TI) - 38",
    "PLOMO (PB) - 35","NÍQUEL (NI) - 32","MOLIBDENO (MO) - 30","SILICIO (SI) - 36","SODIO (NA) - 31",
    "POTASIO (K) - 27","VANADIO (V) - 39","BORO (B) - 18","BARIO (BA) - 21","CALCIO (CA) - 22",
    "CADMIO (CD) - 23","MAGNESIO (MG) - 28","MANGANESO (MN) - 29","FÓSFORO (P) - 34","ZINC (ZN) - 40",
    "CÓDIGO ISO (4/6/14) - 47","CONTEO PARTÍCULAS >= 4 ΜM - 49","CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48","OXIDACIÓN - 80","NITRACIÓN - 82","NÚMERO ÁCIDO (AN) - 43",
    "NÚMERO BÁSICO (BN) - 12","NÚMERO BÁSICO (BN) - 17","HOLLÍN - 79","DILUCIÓN POR COMBUSTIBLE - 46",
    "AGUA (IR) - 81","CONTENIDO AGUA (KARL FISCHER) - 41","CONTENIDO GLICOL - 105",
    "VISCOSIDAD A 100 °C - 13","VISCOSIDAD A 40 °C - 14","COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AGUA CUALITATIVA (PLANCHA) - 360","AGUA LIBRE - 416","ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45","COBRE (CU) - 119","ESPUMA SEC 1 - ESTABILIDAD - 60",
    "ESPUMA SEC 1 - TENDENCIA - 59","ESTAÑO (SN) - 37","ÍNDICE VISCOSIDAD - 359","RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6","SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8","SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "**ULTRACENTRÍFUGA (UC) - 1","ESTADO_PRODUCTO","ESTADO_DESGASTE","ESTADO_CONTAMINACION",
    "N_SOLICITUD","CAMBIO_DE_PRODUCTO","CAMBIO_DE_FILTRO","TEMPERATURA_RESERVORIO",
    "UNIDAD_TEMPERATURA_RESERVORIO","COMENTARIO_CLIENTE","TIPO_DE_COMBUSTIBLE","TIPO_DE_REFRIGERANTE",
    "USUARIO","COMENTARIO_REPORTE","id_muestra"
]

# ===================== ENCABEZADOS REQUERIDOS (ESTADOS) =====================
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

COLUMNAS_USADAS = REQUERIDOS + NUEVAS_ESTADO

# ===================== CARGA =====================
files = st.file_uploader("📤 Sube uno o varios Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if files:
    dfs_out = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()

        # Mapa normalizado → nombre real
        cols_norm = {normalizar(c): c for c in cols}

        # ========= VALIDACIÓN (TODAS requeridas) =========
        faltantes = [c for c in COLUMNAS_USADAS if normalizar(c) not in cols_norm]
        if faltantes:
            st.error(f"❌ Archivo {f.name}: faltan encabezados requeridos")
            st.dataframe(pd.DataFrame({"Encabezado faltante": faltantes}), use_container_width=True)
            st.stop()

        # ========= REPORTAR COLUMNAS CON DATOS NO USADAS =========
        usadas_norm = {normalizar(c) for c in COLUMNAS_USADAS}
        extras_con_datos = []

        for idx, c in enumerate(cols):
            if normalizar(c) in usadas_norm:
                continue

            serie = df[c].astype(str).str.strip()
            serie = serie.replace({"": pd.NA, "nan": pd.NA})

            n_datos = int(serie.notna().sum())
            if n_datos > 0:
                extras_con_datos.append({
                    "Archivo": f.name,
                    "Encabezado NO usado": c,
                    "Registros con datos": n_datos,
                    "Posición original": col_index_to_letter(idx)
                })

        if extras_con_datos:
            st.warning(f"⚠️ {f.name}: columnas con DATOS que NO se usan en el archivo de salida")
            st.dataframe(pd.DataFrame(extras_con_datos), use_container_width=True)
        else:
            st.info(f"✅ {f.name}: no hay columnas con datos fuera de las listas usadas.")

        # ========= CONSTRUIR EXCEL DE SALIDA (orden exacto) =========
        df_out = pd.DataFrame()

        for c in REQUERIDOS:
            df_out[c] = df[cols_norm[normalizar(c)]]

        # Renombre
        if "ESTADO_REPORTE" in df_out.columns:
            df_out.rename(columns={"ESTADO_REPORTE": "ESTADO"}, inplace=True)

        # Archivo origen
        df_out["Archivo_Origen"] = f.name

        for c in NUEVAS_ESTADO:
            df_out[c] = df[cols_norm[normalizar(c)]]

        dfs_out.append(df_out)

    df_final = pd.concat(dfs_out, ignore_index=True)

    st.success("✅ Proceso completado correctamente")
    st.subheader("📋 Vista previa")
    st.dataframe(df_final.head(20), use_container_width=True)

    nombre = f"resultado_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "📥 Descargar archivo final",
        df_to_xlsx_bytes(df_final),
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

