import streamlit as st
import pandas as pd
from io import BytesIO

# ===================== Configuración =====================
st.set_page_config(page_title="Consolidar (encabezados exactos)", layout="wide")
st.title("📄 Consolidar Excel con ENCABEZADOS EXACTOS")
st.caption(
    "Se detiene si falta 1 columna requerida. "
    "Si todo está OK, se crea el archivo final sólo con las columnas pedidas y en el mismo orden."
)

# ===================== Utilitarios =====================
def col_index_to_letter(idx: int) -> str:
    s = ""
    i = int(idx)
    while i >= 0:
        s = chr(i % 26 + 65) + s
        i = i // 26 - 1
    return s

def df_to_xlsx_bytes(df: pd.DataFrame, sheet: str = "Consolidado") -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    return buf

# ===================== Encabezados requeridos (EXACTOS y ORDEN) =====================
REQUERIDOS = [
    "NOMBRE_CLIENTE","NOMBRE_OPERACION","N_MUESTRA","CORRELATIVO","FECHA_MUESTREO","FECHA_INGRESO",
    "FECHA_RECEPCION","FECHA_INFORME","EDAD_COMPONENTE","UNIDAD_EDAD_COMPONENTE","EDAD_PRODUCTO",
    "UNIDAD_EDAD_PRODUCTO","CANTIDAD_ADICIONADA","UNIDAD_CANTIDAD_ADICIONADA","PRODUCTO","TIPO_PRODUCTO",
    "EQUIPO","TIPO_EQUIPO","MARCA_EQUIPO","MODELO_EQUIPO","COMPONENTE","MARCA_COMPONENTE","MODELO_COMPONENTE",
    "DESCRIPTOR_COMPONENTE","ESTADO","NIVEL_DE_SERVICIO","ÍNDICE PQ (PQI) - 3","PLATA (AG) - 19","ALUMINIO (AL) - 20",
    "CROMO (CR) - 24","COBRE (CU) - 25","HIERRO (FE) - 26","TITANIO (TI) - 38","PLOMO (PB) - 35","NÍQUEL (NI) - 32",
    "MOLIBDENO (MO) - 30","SILICIO (SI) - 36","SODIO (NA) - 31","POTASIO (K) - 27","VANADIO (V) - 39","BORO (B) - 18",
    "BARIO (BA) - 21","CALCIO (CA) - 22","CADMIO (CD) - 23","MAGNESIO (MG) - 28","MANGANESO (MN) - 29",
    "FÓSFORO (P) - 34","ZINC (ZN) - 40","CÓDIGO ISO (4/6/14) - 47","CONTEO PARTÍCULAS >= 4 ΜM - 49",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50","CONTEO PARTÍCULAS >= 14 ΜM - 48","**OXIDACIÓN - 80","**NITRACIÓN - 82",
    "NÚMERO ÁCIDO (AN) - 43","NÚMERO BÁSICO (BN) - 12","NÚMERO BÁSICO (BN) - 17","**HOLLÍN - 79",
    "DILUCIÓN POR COMBUSTIBLE - 46","**AGUA (IR) - 81","CONTENIDO AGUA (KARL FISCHER) - 41","CONTENIDO GLICOL  - 105",
    "VISCOSIDAD A 100 °C - 13","VISCOSIDAD A 40 °C - 14","COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AGUA CUALITATIVA (PLANCHA) - 360","AGUA LIBRE - 416","ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45","COBRE (CU) - 119","ESPUMA SEC 1 - ESTABILIDAD - 60",
    "ESPUMA SEC 1 - TENDENCIA - 59","ESTAÑO (SN) - 37","**ÍNDICE VISCOSIDAD - 359","RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6","SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8","SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83","**ULTRACENTRÍFUGA (UC) - 1"
]

# ===================== Subida de archivos =====================
files = st.file_uploader(
    "📤 Sube uno o varios Excel (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if files:
    errores_lectura = []
    faltantes_global = []
    extras_tabla = []
    dfs_filtrados = []

    for f in files:
        # 1) Leer archivo
        try:
            df = pd.read_excel(f, dtype=str, engine="openpyxl")
        except Exception as e:
            errores_lectura.append({"Archivo": f.name, "Error de lectura": str(e)})
            continue

        # 2) Normalizar espacios en encabezados
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        cols = df.columns.tolist()

        # 3) Verificar faltantes EXACTOS
        faltantes = [c for c in REQUERIDOS if c not in cols]
        if faltantes:
            for col in faltantes:
                faltantes_global.append({
                    "Archivo": f.name,
                    "Columna requerida NO encontrada": col
                })
            continue

        # 4) Armar salida SOLO con requeridos (orden exacto)
        df_out = df[REQUERIDOS].copy()
        dfs_filtrados.append(df_out)

        # 5) Analizar columnas NO requeridas con >1 dato (ignorando celdas = nombre encabezado)
        req_set = set(REQUERIDOS)
        for idx, col in enumerate(cols):
            if col in req_set:
                continue
            serie = df[col].astype(str).str.strip()
            serie = serie.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA})
            mask_valido = serie.notna() & (serie.str.casefold() != str(col).strip().casefold())
            cnt = int(mask_valido.sum())
            if cnt > 1:
                extras_tabla.append({
                    "Archivo": f.name,
                    "Encabezado (no requerido)": col,
                    "Registros con datos (>1, sin repetir encabezado)": cnt,
                    "Posición original (n)": idx + 1,
                    "Posición original (Excel)": col_index_to_letter(idx)
                })

    # 6) Mostrar errores de lectura
    if errores_lectura:
        st.subheader("❗ Errores de lectura")
        st.dataframe(pd.DataFrame(errores_lectura), use_container_width=True)

    # 7) Si hay faltantes en cualquier archivo → detener
    if faltantes_global:
        st.error("❌ Faltan columnas REQUERIDAS (coincidencia EXACTA). Proceso detenido.")
        st.dataframe(
            pd.DataFrame(faltantes_global, columns=["Archivo", "Columna requerida NO encontrada"]),
            use_container_width=True
        )
        st.stop()

    # 8) Unir y descargar
    if not dfs_filtrados:
        st.warning("No hubo archivos válidos para consolidar (todos fallaron al leer o tenían faltantes).")
    else:
        st.success("✅ Todos los archivos válidos contienen TODAS las columnas requeridas.")

        df_final = pd.concat(dfs_filtrados, ignore_index=True)
        st.subheader("📋 Vista previa del archivo final (solo columnas requeridas y en orden)")
        st.dataframe(df_final.head(15), use_container_width=True)

        # Descarga del consolidado
        xlsx_bytes = df_to_xlsx_bytes(df_final, sheet="Consolidado")
        st.download_button(
            label="📥 Descargar archivo final (XLSX)",
            data=xlsx_bytes,
            file_name="consolidado_requeridos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 9) Tabla de NO requeridas con >1 dato
        st.subheader("🟠 Columnas NO requeridas con >1 dato (ignorando celdas iguales al encabezado)")
        if extras_tabla:
            df_extras = pd.DataFrame(
                extras_tabla,
                columns=[
                    "Archivo",
                    "Encabezado (no requerido)",
                    "Registros con datos (>1, sin repetir encabezado)",
                    "Posición original (n)",
                    "Posición original (Excel)"
                ]
            )
            st.dataframe(df_extras, use_container_width=True)

            extras_xlsx = df_to_xlsx_bytes(df_extras, sheet="Extras_con_datos")
            st.download_button(
                label="📥 Descargar tabla de extras (XLSX)",
                data=extras_xlsx,
                file_name="extras_con_datos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No se encontraron columnas NO requeridas con más de 1 dato.")


