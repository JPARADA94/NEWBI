import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ===================== Configuración =====================
st.set_page_config(page_title="Filtrar por Encabezados EXACTOS", layout="wide")
st.title("📄 Construir Excel solo con encabezados requeridos (EXACTOS) + Archivo_Origen")
st.caption(
    "Se detiene si falta alguna columna requerida. Si todo está OK, se genera el archivo final en la hoja 'Sheet1' "
    "sin tablas de Excel, con 'Archivo_Origen' como última columna, seguido de columnas Estado."
)

# ===================== Utilitarios =====================
def col_index_to_letter(idx: int) -> str:
    s = ""
    i = int(idx)
    while i >= 0:
        s = chr(i % 26 + 65) + s
        i = i // 26 - 1
    return s

def df_to_xlsx_bytes(df: pd.DataFrame, sheet: str = "Sheet1") -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    return buf

# ===================== Función para detectar columnas faltantes =====================
def verificar_columnas_faltantes(cols_archivo, cols_requeridos):
    faltantes = [c for c in cols_requeridos if c not in cols_archivo]
    if faltantes:
        st.error("❌ Este archivo NO cumple con los encabezados requeridos.")
        st.dataframe(pd.DataFrame({"Columnas faltantes": faltantes}), use_container_width=True)
    return faltantes

# ===================== Encabezados requeridos (EXACTOS y ORDENADOS) =====================
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
    "CONTEO PARTÍCULAS >= 14 ΜM - 48","OXIDACIÓN - 80","NITRACIÓN - 82",
    "NÚMERO ÁCIDO (AN) - 43","NÚMERO BÁSICO (BN) - 12","NÚMERO BÁSICO (BN) - 17","HOLLÍN - 79",
    "DILUCIÓN POR COMBUSTIBLE - 46","AGUA (IR) - 81","CONTENIDO AGUA (KARL FISCHER) - 41",
    "CONTENIDO GLICOL - 105","VISCOSIDAD A 100 °C - 13","VISCOSIDAD A 40 °C - 14",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51","AGUA CUALITATIVA (PLANCHA) - 360","AGUA LIBRE - 416",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44","ANÁLISIS ANTIOXIDANTES (FENOL) - 45","COBRE (CU) - 119",
    "ESPUMA SEC 1 - ESTABILIDAD - 60","ESPUMA SEC 1 - TENDENCIA - 59","ESTAÑO (SN) - 37",
    "ÍNDICE VISCOSIDAD - 359","RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6","SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8","SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "**ULTRACENTRÍFUGA (UC) - 1","ESTADO_PRODUCTO","ESTADO_DESGASTE","ESTADO_CONTAMINACION",
    "N_SOLICITUD","CAMBIO_DE_PRODUCTO","CAMBIO_DE_FILTRO","TEMPERATURA_RESERVORIO",
    "UNIDAD_TEMPERATURA_RESERVORIO","COMENTARIO_CLIENTE","TIPO_DE_COMBUSTIBLE",
    "TIPO_DE_REFRIGERANTE","USUARIO","COMENTARIO_REPORTE","id_muestra"
]

# ===================== Nuevas columnas ESTADO al final =====================
NUEVAS_ESTADO = [
    "ESTADO_MUESTRA","AGUA CUALITATIVA (PLANCHA) - 360 - Estado","AGUA (IR) - 81 - Estado",
    "ALUMINIO (AL) - 20 - Estado","BARIO (BA) - 21 - Estado","BORO (B) - 18 - Estado",
    "CADMIO (CD) - 23 - Estado","CALCIO (CA) - 22 - Estado","COBRE (CU) - 25 - Estado",
    "CÓDIGO ISO (4/6/14) - 47 - Estado","CONTENIDO AGUA (KARL FISCHER) - 41 - Estado",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48 - Estado","CONTEO PARTÍCULAS >= 4 ΜM - 49 - Estado",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50 - Estado","CROMO (CR) - 24 - Estado",
    "DILUCIÓN POR COMBUSTIBLE - 46 - Estado","ESTAÑO (SN) - 37 - Estado",
    "FÓSFORO (P) - 34 - Estado","HIERRO (FE) - 26 - Estado","HOLLÍN - 79 - Estado",
    "ÍNDICE PQ (PQI) - 3 - Estado","MAGNESIO (MG) - 28 - Estado","MANGANESO (MN) - 29 - Estado",
    "MOLIBDENO (MO) - 30 - Estado","NÍQUEL (NI) - 32 - Estado","NITRACIÓN - 82 - Estado",
    "NÚMERO ÁCIDO (AN) - 43 - Estado","NÚMERO BÁSICO (BN) - 17 - Estado",
    "NÚMERO BÁSICO (BN) - 12 - Estado","OXIDACIÓN - 80 - Estado","PLATA (AG) - 19 - Estado",
    "PLOMO (PB) - 35 - Estado","POTASIO (K) - 27 - Estado","SILICIO (SI) - 36 - Estado",
    "SODIO (NA) - 31 - Estado","TITANIO (TI) - 38 - Estado","VANADIO (V) - 39 - Estado",
    "VISCOSIDAD A 100 °C - 13 - Estado","VISCOSIDAD A 40 °C - 14 - Estado",
    "ZINC (ZN) - 40 - Estado"
]

# ===================== Carga de archivos =====================
files = st.file_uploader("📤 Sube uno o varios Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if files:
    faltantes_global = []
    extras_tabla = []
    dfs_filtrados = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()

        faltantes = verificar_columnas_faltantes(cols, REQUERIDOS)
        if faltantes:
            for col in faltantes:
                faltantes_global.append({"Archivo": f.name,"Columna requerida NO encontrada": col})
            continue  

        df_out = df[REQUERIDOS].copy()

        # RENOMBRES (sin afectar orden existente)
        rename_map = {}
        if "ESTADO_REPORTE" in df_out.columns:
            rename_map["ESTADO_REPORTE"] = "ESTADO"
        if rename_map:
            df_out = df_out.rename(columns=rename_map)

        # Añadir archivo origen
        df_out["Archivo_Origen"] = f.name

        # AGREGAR TODAS LAS NUEVAS COLUMNAS AL FINAL (sin alterar el orden previo)
        for col in NUEVAS_ESTADO:
            if col not in df_out.columns:
                df_out[col] = ""

        dfs_filtrados.append(df_out)

        # Columnas no requeridas
        requeridos_set = set(REQUERIDOS)
        for idx, col in enumerate(cols):
            if col in requeridos_set:
                continue
            serie = df[col].astype(str).str.strip()
            serie = serie.replace({"": pd.NA, "nan": pd.NA})
            if serie.notna().sum() > 1:
                extras_tabla.append({
                    "Archivo": f.name,
                    "Encabezado (no requerido)": col,
                    "Registros con datos (>1)": serie.notna().sum(),
                    "Posición original": col_index_to_letter(idx)
                })

    if faltantes_global:
        st.error("❌ Existen archivos con columnas faltantes.")
        st.dataframe(pd.DataFrame(faltantes_global), use_container_width=True)
        st.stop()

    st.success("✅ Todos los archivos cumplen con los encabezados requeridos.")

    df_final = pd.concat(dfs_filtrados, ignore_index=True)
    st.subheader("📋 Vista previa final")
    st.dataframe(df_final.head(20), use_container_width=True)

    cliente = str(df_final["NOMBRE_CLIENTE"].dropna().iloc[0]).strip().replace(" ", "_")
    fecha_actual = datetime.now().strftime("%Y%m%d")
    nombre_archivo = f"{cliente}_{fecha_actual}.xlsx"

    xlsx_bytes = df_to_xlsx_bytes(df_final)
    st.download_button(
        "📥 Descargar archivo final",
        xlsx_bytes,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


