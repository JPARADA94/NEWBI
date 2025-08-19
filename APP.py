import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata

# —————— Configuración general ——————
st.set_page_config(page_title="Validación Global – Excel Consolidado v4.2", layout="wide")
st.title("📊 Validación Global de Encabezados – Excel Consolidado – Mobil v4.2")
st.markdown("**Responsables:** Grupo de Soporte en Campo – Mobil")

st.markdown(
    """
### 🧾 Instrucciones de uso:
1. Sube **uno o varios archivos Excel (.xlsx)**.
2. El sistema unirá todos los archivos en un solo conjunto.
3. Validará los encabezados sobre el conjunto completo.
4. Generará dos reportes:
   - **Tabla de desalineaciones**: posición esperada vs. posición encontrada o ausencia.
   - **Tabla de columnas con datos no mapeadas** (se agregarán al final).
5. Genera **un único archivo Excel consolidado** y los reportes descargables.
"""
)

# —————— Utilitarios ——————
def col_index_to_letter(idx: int) -> str:
    """Convierte índice base 0 a letra(s) de columna de Excel (A, Z, AA...)."""
    letter = ""
    while idx >= 0:
        letter = chr(idx % 26 + ord('A')) + letter
        idx = idx // 26 - 1
    return letter

def normalize_header(s: str) -> str:
    """Normaliza encabezados para coincidencias tolerantes."""
    if s is None:
        return ""
    s = s.strip()
    s = s.replace("≥", ">=").replace("Μ", "µ").replace("\u00A0", " ")  # NBSP → espacio
    s = s.replace("**", "")
    s = unicodedata.normalize('NFKD', s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def df_to_xlsx_bytes(df: pd.DataFrame, sheet: str = "Hoja") -> BytesIO:
    """Convierte un DataFrame a bytes XLSX usando openpyxl (sin XlsxWriter)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    return buf

def make_downloads(df: pd.DataFrame, base_name: str, sheet: str):
    """Botones de descarga CSV/XLSX para un DataFrame."""
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    xlsx_bytes = df_to_xlsx_bytes(df, sheet=sheet)
    c1, c2 = st.columns(2)
    c1.download_button(
        f"📥 {base_name} (CSV)",
        data=csv_bytes,
        file_name=f"{base_name}.csv",
        mime="text/csv",
    )
    c2.download_button(
        f"📥 {base_name} (XLSX)",
        data=xlsx_bytes,
        file_name=f"{base_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ——— Orden base EXACTO solicitado (sin id_muestra) ———
expected_names = [
    "NOMBRE_CLIENTE",
    "NOMBRE_OPERACION",
    "N_MUESTRA",
    "CORRELATIVO",
    "FECHA_MUESTREO",
    "FECHA_INGRESO",
    "FECHA_RECEPCION",
    "FECHA_INFORME",
    "EDAD_COMPONENTE",
    "UNIDAD_EDAD_COMPONENTE",
    "EDAD_PRODUCTO",
    "UNIDAD_EDAD_PRODUCTO",
    "CANTIDAD_ADICIONADA",
    "UNIDAD_CANTIDAD_ADICIONADA",
    "PRODUCTO",
    "TIPO_PRODUCTO",
    "EQUIPO",
    "TIPO_EQUIPO",
    "MARCA_EQUIPO",
    "MODELO_EQUIPO",          # ← posición 19
    "COMPONENTE",
    "MARCA_COMPONENTE",
    "MODELO_COMPONENTE",
    "DESCRIPTOR_COMPONENTE",
    "ESTADO",
    "NIVEL_DE_SERVICIO",
    "ÍNDICE PQ (PQI) - 3",
    "PLATA (AG) - 19",
    "ALUMINIO (AL) - 20",
    "CROMO (CR) - 24",
    "COBRE (CU) - 25",
    "HIERRO (FE) - 26",
    "TITANIO (TI) - 38",
    "PLOMO (PB) - 35",
    "NÍQUEL (NI) - 32",
    "MOLIBDENO (MO) - 30",
    "SILICIO (SI) - 36",
    "SODIO (NA) - 31",
    "POTASIO (K) - 27",
    "VANADIO (V) - 39",
    "BORO (B) - 18",
    "BARIO (BA) - 21",
    "CALCIO (CA) - 22",
    "CADMIO (CD) - 23",
    "MAGNESIO (MG) - 28",
    "MANGANESO (MN) - 29",
    "FÓSFORO (P) - 34",
    "ZINC (ZN) - 40",
    "CÓDIGO ISO (4/6/14) - 47",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48",
    "**OXIDACIÓN - 80",
    "**NITRACIÓN - 82",
    "NÚMERO ÁCIDO (AN) - 43",
    "NÚMERO BÁSICO (BN) - 12",
    "NÚMERO BÁSICO (BN) - 17",
    "**HOLLÍN - 79",
    "DILUCIÓN POR COMBUSTIBLE - 46",
    "**AGUA (IR) - 81",
    "CONTENIDO AGUA (KARL FISCHER) - 41",
    "CONTENIDO GLICOL  - 105",
    "VISCOSIDAD A 100 °C - 13",
    "VISCOSIDAD A 40 °C - 14",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AGUA CUALITATIVA (PLANCHA) - 360",
    "AGUA LIBRE - 416",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45",
    "COBRE (CU) - 119",
    "ESPUMA SEC 1 - ESTABILIDAD - 60",
    "ESPUMA SEC 1 - TENDENCIA - 59",
    "ESTAÑO (SN) - 37",
    "**ÍNDICE VISCOSIDAD - 359",
    "RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "**ULTRACENTRÍFUGA (UC) - 1",
    "Archivo_Origen"          # última fija
]

# —————— Subida de múltiples archivos ——————
uploaded_files = st.file_uploader(
    "📤 Sube uno o varios archivos Excel (.xlsx):",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:
    # Unir todo como texto
    dfs = []
    for uploaded in uploaded_files:
        df = pd.read_excel(uploaded, header=0, dtype=str, engine="openpyxl")
        df["Archivo_Origen"] = uploaded.name
        dfs.append(df)
    df_global = pd.concat(dfs, ignore_index=True)

    # Columnas reales y mapas auxiliares
    columnas_reales = [c.strip() for c in df_global.columns.tolist()]
    mapa_nombre_a_indice = {col: i for i, col in enumerate(columnas_reales)}
    mapa_norm_a_nombre = {normalize_header(col): col for col in columnas_reales}
    expected_set_norm = {normalize_header(v) for v in expected_names}

    # —— Reporte de desalineaciones ——
    des_rows = []
    for pos_esp, esperado in enumerate(expected_names):
        letra_esp = col_index_to_letter(pos_esp)
        if esperado in mapa_nombre_a_indice:
            pos_real = mapa_nombre_a_indice[esperado]
            if pos_real != pos_esp:
                des_rows.append({
                    "Posición esperada": f"{pos_esp+1} ({letra_esp})",
                    "Encabezado esperado": esperado,
                    "Posición encontrada": f"{pos_real+1} ({col_index_to_letter(pos_real)})",
                })
        else:
            norm = normalize_header(esperado)
            if norm in mapa_norm_a_nombre:
                casi = mapa_norm_a_nombre[norm]
                pos_real = mapa_nombre_a_indice[casi]
                des_rows.append({
                    "Posición esperada": f"{pos_esp+1} ({letra_esp})",
                    "Encabezado esperado": esperado,
                    "Posición encontrada": f"{pos_real+1} ({col_index_to_letter(pos_real)}) – (variante '{casi}')",
                })
            else:
                des_rows.append({
                    "Posición esperada": f"{pos_esp+1} ({letra_esp})",
                    "Encabezado esperado": esperado,
                    "Posición encontrada": "(no existe)",
                })

    st.subheader("📋 Tabla de Desalineaciones")
    if des_rows:
        st.dataframe(pd.DataFrame(des_rows), use_container_width=True)
    else:
        st.success("✅ Todas las columnas están en la posición esperada.")

    st.divider()

    # —— Columnas no mapeadas con datos ——
    st.subheader("🟠 Columnas con datos no mapeadas (se agregarán al final)")
    extra_rows = []
    extra_cols_ordered = []
    for idx, nombre in enumerate(columnas_reales):
        if normalize_header(nombre) not in expected_set_norm:
            datos = df_global.iloc[:, idx].notna().sum()
            if datos > 0:
                extra_rows.append({
                    "Letra actual": col_index_to_letter(idx),
                    "Encabezado no considerado": nombre,
                    "Registros con datos": int(datos),
                })
                extra_cols_ordered.append(nombre)

    if extra_rows:
        st.dataframe(pd.DataFrame(extra_rows), use_container_width=True)
    else:
        st.info("No se encontraron columnas adicionales con datos.")

    st.divider()

    # —— Construcción del archivo final ——
    st.subheader("🧩 Construcción del archivo final (orden fijo + extras al final)")

    columnas_finales = []
    for esperado in expected_names:
        if esperado in mapa_nombre_a_indice:
            columnas_finales.append(df_global.iloc[:, mapa_nombre_a_indice[esperado]].rename(esperado))
        else:
            columnas_finales.append(pd.Series([None]*len(df_global), name=esperado))

    for nombre in extra_cols_ordered:
        if nombre not in [s.name for s in columnas_finales]:
            columnas_finales.append(df_global[nombre])

    df_resultado = pd.concat(columnas_finales, axis=1)

    st.subheader("📋 Vista previa – Archivo Final")
    st.dataframe(df_resultado.head(10), use_container_width=True)
    make_downloads(df_resultado, "archivo_consolidado", sheet="Consolidado")



