import streamlit as st
import pandas as pd
from io import BytesIO

# ===================== Configuración =====================
st.set_page_config(page_title="Filtrar por Encabezados EXACTOS", layout="wide")
st.title("📄 Construir Excel solo con encabezados requeridos (coincidencia EXACTA)")
st.caption("Si falta AL MENOS una columna requerida en cualquier archivo, se mostrará un aviso y se detendrá el proceso.")

# ===================== Utilitarios =====================
def col_index_to_letter(idx: int) -> str:
    """0->A, 25->Z, 26->AA, etc."""
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

# ===================== Encabezados requeridos (EXACTOS y en ORDEN) =====================
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

# ===================== Carga de archivos =====================
files = st.file_uploader("📤 Sube uno o varios Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if files:
    # 1) Validación de encabezados por archivo (EXACTOS). Si falta alguno: avisar y detener.
    faltantes_reg = []      # lista de dicts con faltantes por archivo
    extras_tabla = []       # para tabla de extras (solo si todo OK)
    dfs_filtrados = []      # para concatenar si todo OK

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()

        # --- Validación exacta ---
        faltantes = [c for c in REQUERIDOS if c not in cols]
        if faltantes:
            for col in faltantes:
                faltantes_reg.append({"Archivo": f.name, "Columna requerida NO encontrada": col})
            # seguimos revisando los demás archivos para mostrar todos los errores,
            # pero NO generaremos salida si hay al menos un faltante.

        # Si no hay faltantes en este archivo, preparamos su versión filtrada y
        # recolectamos info de extras (no requeridas) con datos.
        if not faltantes:
            # Filtrar y ordenar exactamente como REQUERIDOS
            df_out = df[REQUERIDOS].copy()
            df_out["Archivo_Origen"] = f.name  # trazabilidad (no está en REQUERIDOS; puedes quitarlo si no lo quieres)
            dfs_filtrados.append(df_out)

            # Evaluar otras columnas no relacionadas con datos (tabla extra)
            requeridos_set = set(REQUERIDOS)
            for idx, col in enumerate(cols):
                if col not in requeridos_set:
                    serie = df[col]
                    # contar datos no vacíos/ni NaN (considera strings vacíos como vacíos)
                    datos = serie.astype(str).str.strip().replace({"": pd.NA}).notna().sum()
                    if datos > 0:
                        extras_tabla.append({
                            "Archivo": f.name,
                            "Encabezado (no requerido)": col,
                            "Registros con datos": int(datos),
                            "Posición original (n)": idx + 1,
                            "Posición original (Excel)": col_index_to_letter(idx)
                        })

    # Si hubo faltantes en CUALQUIER archivo → avisar y detener.
    if faltantes_reg:
        st.error("❌ Se detectaron columnas REQUERIDAS que NO aparecen con nombre EXACTO.")
        df_falt = pd.DataFrame(faltantes_reg, columns=["Archivo","Columna requerida NO encontrada"])
        st.dataframe(df_falt, use_container_width=True)
        st.stop()

    # 2) Si TODO está OK → unimos, mostramos tabla de extras y generamos archivo
    st.success("✅ Todos los archivos contienen TODAS las columnas requeridas con nombre EXACTO.")

    # Tabla de extras (si hubo)
    st.subheader("🟠 Columnas NO requeridas con datos (por archivo)")
    if extras_tabla:
        df_extras = pd.DataFrame(extras_tabla, columns=[
            "Archivo","Encabezado (no requerido)","Registros con datos",
            "Posición original (n)","Posición original (Excel)"
        ])
        st.dataframe(df_extras, use_container_width=True)
        extras_xlsx = df_to_xlsx_bytes(df_extras, sheet="Extras_con_datos")
        st.download_button("📥 Descargar tabla de extras (XLSX)", extras_xlsx,
                           file_name="extras_con_datos.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No se encontraron columnas adicionales con datos.")

    # Unir y descargar resultado final
    df_final = pd.concat(dfs_filtrados, ignore_index=True)
    st.subheader("📋 Vista previa del archivo final")
    st.dataframe(df_final.head(15), use_container_width=True)

    xlsx_bytes = df_to_xlsx_bytes(df_final, sheet="Consolidado")
    st.download_button("📥 Descargar archivo final (XLSX)", xlsx_bytes,
                       file_name="consolidado_requeridos.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

