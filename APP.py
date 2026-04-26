import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

# ===================== CONFIGURACIÓN =====================
st.set_page_config(page_title="SmartAssistence a Power BI", layout="wide")
st.title("📄 Convertidor de formato SmartAssistence a Power BI")

st.info(
    "Instrucciones importantes:\n\n"
    "1. Sube el archivo Excel original exportado desde SmartAssistence.\n"
    "2. Descarga el archivo generado por esta herramienta.\n"
    "3. El archivo descargado NO debe ser modificado.\n"
    "4. No cambies el nombre del archivo generado.\n"
    "5. No edites columnas, encabezados, hojas, datos ni formato interno.\n"
    "6. Sube el archivo descargado tal cual a la base de datos para evitar errores en Power BI."
)

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

def df_to_csv_bytes(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
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

def limpiar_nombre_archivo(texto: str) -> str:
    texto = str(texto).strip()
    texto = re.sub(r"[\\/*?:\"<>|]", "", texto)
    texto = re.sub(r"\s+", "_", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto[:80] if texto else "CUENTA_SIN_NOMBRE"

def obtener_nombre_cuenta(df: pd.DataFrame) -> str:
    for col in ["NOMBRE_CLIENTE", "NOMBRE_OPERACION"]:
        if col in df.columns:
            serie = df[col].dropna().astype(str).str.strip()
            serie = serie[serie != ""]
            if not serie.empty:
                return limpiar_nombre_archivo(serie.iloc[0])
    return "CUENTA_SIN_NOMBRE"

# ===================== ENCABEZADOS BASE =====================
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
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AGUA CUALITATIVA (PLANCHA) - 360",
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

# ===================== ENCABEZADOS ESTADO =====================
NUEVAS_ESTADO = [
    "ESTADO_MUESTRA",
    "AGUA (IR) - 74",
    "AGUA (IR) - 74 - Estado",
    "AGUA (IR) - 81 - Estado",
    "AGUA LIBRE - 416 - Estado",
    "AGUA CUALITATIVA (PLANCHA) - 360 - Estado",
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
    "CÓDIGO ISO (4/6/14) - 47 - Estado",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49 - Estado",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50 - Estado",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48 - Estado",
    "OXIDACIÓN - 80 - Estado",
    "NITRACIÓN - 82 - Estado",
    "ÍNDICE PQ (PQI) - 3 - Estado",
    "NÚMERO ÁCIDO (AN) - 43 - Estado",
    "NÚMERO BÁSICO (BN) - 12 - Estado",
    "NÚMERO BÁSICO (BN) - 17 - Estado",
    "CONTENIDO AGUA (KARL FISCHER) - 41 - Estado",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44 - Estado",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45 - Estado",
    "HOLLÍN - 73",
    "HOLLÍN - 73 - Estado",
    "HOLLÍN - 79 - Estado",
    "DILUCIÓN POR COMBUSTIBLE - 46 - Estado",
    "VISCOSIDAD A 40 °C - 14 - Estado",
    "VISCOSIDAD A 100 °C - 13 - Estado",
    "ÍNDICE VISCOSIDAD - 359 - Estado",
    "ESPUMA SEC 1 - ESTABILIDAD - 60 - Estado",
    "ESPUMA SEC 1 - TENDENCIA - 59 - Estado",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado",
    "RESIDUO CARBÓN (MCR) - 361",
    "RESIDUO CARBÓN (MCR) - 361 - Estado",
    "PUNTO DE INFLAMACIÓN (PMA) - 61",
    "PUNTO DE INFLAMACIÓN (PMA) - 61 - Estado",
    "RPVOT - 10 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83 - Estado",
    "**ULTRACENTRÍFUGA (UC) - 1 - Estado"
]

# ===================== ALIAS DE ENTRADA =====================
ALIASES_ENTRADA = {
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51": [
        "** COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51"
    ],
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado": [
        "** COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado"
    ],
}

# ===================== FUNCIONES PARA MAPEAR ENCABEZADOS =====================
def posibles_entradas(nombre_salida: str) -> list[str]:
    return [nombre_salida] + ALIASES_ENTRADA.get(nombre_salida, [])

def encontrar_columna_origen(cols_norm_map: dict, nombre_salida: str) -> str | None:
    for candidato in posibles_entradas(nombre_salida):
        key = normalizar(candidato)
        if key in cols_norm_map:
            return cols_norm_map[key]
    return None

# ===================== COLUMNAS USADAS =====================
COLUMNAS_USADAS = REQUERIDOS + NUEVAS_ESTADO

# ===================== CARGA DE ARCHIVOS =====================
files = st.file_uploader(
    "📤 Sube uno o varios Excel exportados desde SmartAssistence (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if files:
    dfs_out = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()
        cols_norm = {normalizar(c): c for c in cols}

        # -------- VALIDACIÓN DE ENCABEZADOS --------
        faltantes = []
        for col_salida in COLUMNAS_USADAS:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            if col_origen is None:
                faltantes.append(col_salida)

        if faltantes:
            st.error(f"❌ {f.name} – Faltan encabezados requeridos")
            st.dataframe(
                pd.DataFrame({"Encabezado faltante esperado en salida": faltantes}),
                use_container_width=True
            )
            st.stop()

        # -------- DETECCIÓN DE COLUMNAS CON DATOS NO USADAS --------
        usadas_norm = set()
        for c in COLUMNAS_USADAS:
            for cand in posibles_entradas(c):
                usadas_norm.add(normalizar(cand))

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

        # -------- CONSTRUCCIÓN DEL ARCHIVO FINAL --------
        df_out = pd.DataFrame()

        for col_salida in REQUERIDOS:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            df_out[col_salida] = df[col_origen]

        df_out.rename(columns={"ESTADO_REPORTE": "ESTADO"}, inplace=True)

        df_out["Archivo_Origen"] = f.name

        for col_salida in NUEVAS_ESTADO:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            df_out[col_salida] = df[col_origen]

        dfs_out.append(df_out)

    df_final = pd.concat(dfs_out, ignore_index=True)

    st.success("✅ Conversión de SmartAssistence a Power BI completada correctamente")

    st.warning(
        "Antes de subir a la base de datos: descarga este archivo y súbelo tal cual. "
        "No modifiques el nombre, no edites el contenido y no cambies la estructura del archivo."
    )

    st.dataframe(df_final.head(20), use_container_width=True)

    nombre_cuenta = obtener_nombre_cuenta(df_final)
    fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")

    nombre_excel = f"{nombre_cuenta}_{fecha_hora}.xlsx"
    nombre_csv = f"{nombre_cuenta}_{fecha_hora}.csv"

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Descargar archivo final en Excel",
            df_to_xlsx_bytes(df_final),
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        st.download_button(
            "📥 Descargar archivo final en CSV",
            df_to_csv_bytes(df_final),
            file_name=nombre_csv,
            mime="text/csv"
        )


