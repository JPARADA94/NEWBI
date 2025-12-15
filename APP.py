import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ===================== CONFIG =====================
st.set_page_config(page_title="Control total de columnas", layout="wide")
st.title("📄 Validación y control de columnas usadas vs ignoradas")

# ===================== UTILIDADES =====================
def col_index_to_letter(idx: int) -> str:
    s = ""
    i = int(idx)
    while i >= 0:
        s = chr(i % 26 + 65) + s
        i = i // 26 - 1
    return s

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

# ===================== ENCABEZADOS =====================
REQUERIDOS = [ ... ]          # <-- TU LISTA COMPLETA (sin cambios)
NUEVAS_ESTADO = [ ... ]       # <-- TU LISTA COMPLETA (sin cambios)

COLUMNAS_USADAS = REQUERIDOS + NUEVAS_ESTADO

# ===================== CARGA =====================
files = st.file_uploader(
    "📤 Sube uno o varios Excel (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if files:
    dfs = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()

        cols_norm = {normalizar(c): c for c in cols}

        # ========= VALIDACIÓN =========
        faltantes = [
            c for c in COLUMNAS_USADAS
            if normalizar(c) not in cols_norm
        ]

        if faltantes:
            st.error(f"❌ Archivo {f.name} – FALTAN ENCABEZADOS")
            st.dataframe(pd.DataFrame({"Encabezado faltante": faltantes}))
            st.stop()

        # ========= DETECCIÓN DE DATOS NO USADOS =========
        columnas_usadas_norm = {normalizar(c) for c in COLUMNAS_USADAS}
        extras_con_datos = []

        for idx, col in enumerate(cols):
            if normalizar(col) in columnas_usadas_norm:
                continue

            serie = df[col].astype(str).str.strip()
            serie = serie.replace({"": pd.NA, "nan": pd.NA})

            if serie.notna().sum() > 0:
                extras_con_datos.append({
                    "Encabezado NO usado": col,
                    "Registros con datos": serie.notna().sum(),
                    "Posición original": col_index_to_letter(idx)
                })

        if extras_con_datos:
            st.warning(f"⚠️ {f.name} contiene columnas con datos que NO se usan")
            st.dataframe(pd.DataFrame(extras_con_datos), use_container_width=True)

        # ========= CONSTRUCCIÓN FINAL =========
        df_out = pd.DataFrame()

        for col in REQUERIDOS:
            df_out[col] = df[cols_norm[normalizar(col)]]

        df_out.rename(columns={"ESTADO_REPORTE": "ESTADO"}, inplace=True)
        df_out["Archivo_Origen"] = f.name

        for col in NUEVAS_ESTADO:
            df_out[col] = df[cols_norm[normalizar(col)]]

        dfs.append(df_out)

    df_final = pd.concat(dfs, ignore_index=True)

    st.success("✅ Proceso finalizado correctamente")
    st.dataframe(df_final.head(20), use_container_width=True)

    nombre = f"resultado_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "📥 Descargar archivo final",
        df_to_xlsx_bytes(df_final),
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

