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
   - **Tabla de desalineaciones**: ubicación original, encabezado esperado, lo encontrado y nueva ubicación del esperado.
   - **Tabla de columnas con datos no mapeadas**.
5. Permite **incluir columnas extra con datos**.
6. Genera **un único archivo Excel consolidado** y los reportes descargables.
"""
)

# —————— Utilitarios ——————
def col_letter_to_index(letter: str) -> int:
    idx = 0
    for c in letter.upper():
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1


def col_index_to_letter(idx: int) -> str:
    letter = ""
    while idx >= 0:
        letter = chr(idx % 26 + ord('A')) + letter
        idx = idx // 26 - 1
    return letter


def normalize_header(s: str) -> str:
    if s is None:
        return ""
    s = s.strip()
    s = s.replace("≥", ">=").replace("Μ", "µ").replace(" ", " ")  # NBSP a espacio
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
    """Muestra botones de descarga CSV/XLSX para un DataFrame."""
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    xlsx_bytes = df_to_xlsx_bytes(df, sheet=sheet)
    c1, c2 = st.columns(2)
    c1.download_button(
        f"📥 {base_name} (CSV)", data=csv_bytes, file_name=f"{base_name}.csv", mime="text/csv"
    )
    c2.download_button(
        f"📥 {base_name} (XLSX)", data=xlsx_bytes, file_name=f"{base_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# —————— Diccionario actualizado de columnas esperadas (ajustado) ——————
columnas_esperadas = {
    "A": "NOMBRE_CLIENTE",
    "B": "NOMBRE_OPERACION",
    "C": "N_MUESTRA",
    "D": "CORRELATIVO",
    "E": "FECHA_MUESTREO",
    "F": "FECHA_INGRESO",
    "G": "FECHA_RECEPCION",
    "H": "FECHA_INFORME",
    "I": "EDAD_COMPONENTE",
    "J": "UNIDAD_EDAD_COMPONENTE",
    "K": "EDAD_PRODUCTO",
    "L": "UNIDAD_EDAD_PRODUCTO",
    "M": "CANTIDAD_ADICIONADA",
    "N": "UNIDAD_CANTIDAD_ADICIONADA",
    "O": "PRODUCTO",
    "U": "COMPONENTE",
    "V": "MARCA_COMPONENTE",
    "W": "MODELO_COMPONENTE",
    "X": "DESCRIPTOR_COMPONENTE",
    "Y": "ESTADO",
    "Z": "NIVEL_DE_SERVICIO",
    "IQ": "ÍNDICE PQ (PQI) - 3",
    "MK": "PLATA (AG) - 19",
    "AK": "ALUMINIO (AL) - 20",
    "FM": "CROMO (CR) - 24",
    "BX": "COBRE (CU) - 25",
    "IF": "HIERRO (FE) - 26",
    "PB": "TITANIO (TI) - 38",
    "MN": "PLOMO (PB) - 35",
    "JS": "NÍQUEL (NI) - 32",
    "JM": "MOLIBDENO (MO) - 30",
    "OE": "SILICIO (SI) - 36",
    "OH": "SODIO (NA) - 31",
    "MP": "POTASIO (K) - 27",
    "PF": "VANADIO (V) - 39",
    "BK": "BORO (B) - 18",
    "BE": "BARIO (BA) - 21",
    "BO": "CALCIO (CA) - 22",
    "BM": "CADMIO (CD) - 23",
    "JG": "MAGNESIO (MG) - 28",
    "JH": "MANGANESO (MN) - 29",
    "HR": "FÓSFORO (P) - 34",
    "PQ": "ZINC (ZN) - 40",
    "CA": "CÓDIGO ISO (4/6/14) - 47",
    "FC": "CONTEO PARTÍCULAS >= 4 ΜM - 49",
    "FD": "CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "FB": "CONTEO PARTÍCULAS >= 14 ΜM - 48",
    "KD": "**OXIDACIÓN - 80",
    "JT": "**NITRACIÓN - 82",
    "JW": "NÚMERO ÁCIDO (AN) - 43",
    "JY": "NÚMERO BÁSICO (BN) - 12",
    "JX": "NÚMERO BÁSICO (BN) - 17",
    "IH": "**HOLLÍN - 79",
    "GP": "DILUCIÓN POR COMBUSTIBLE - 46",
    "AF": "**AGUA (IR) - 81",
    "CT": "CONTENIDO AGUA (KARL FISCHER) - 41",
    "ES": "CONTENIDO GLICOL  - 105",
    "PI": "VISCOSIDAD A 100 °C - 13",
    "PJ": "VISCOSIDAD A 40 °C - 14",
    "CF": "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AE": "AGUA CUALITATIVA (PLANCHA) - 360",
    "AH": "AGUA LIBRE - 416",
    "AL": "ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "AM": "ANÁLISIS ANTIOXIDANTES (FENOL) - 45",
    "BW": "COBRE (CU) - 119",
    "GU": "ESPUMA SEC 1 - ESTABILIDAD - 60",
    "GV": "ESPUMA SEC 1 - TENDENCIA - 59",
    "HL": "ESTAÑO (SN) - 37",
    "IT": "**ÍNDICE VISCOSIDAD - 359",
    "NX": "RPVOT - 10",
    "NZ": "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6",
    "OA": "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "OB": "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8",
    "OC": "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "PE": "**ULTRACENTRÍFUGA (UC) - 1",
}

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

    columnas_reales = [c.strip() for c in df_global.columns.tolist()]
    expected_names = list(columnas_esperadas.values())

    # —— Reporte de desalineaciones ——
    des_rows = []
    for letra, esperado in columnas_esperadas.items():
        idx = col_letter_to_index(letra)
        if idx < len(columnas_reales):
            encontrado = columnas_reales[idx]
            if encontrado != esperado:
                nueva_letra = col_index_to_letter(columnas_reales.index(esperado)) if esperado in columnas_reales else "—"
                des_rows.append({
                    "Ubicación original": letra,
                    "Encabezado esperado": esperado,
                    "Encontrado en origen": encontrado,
                    "Nueva ubicación del esperado": nueva_letra,
                })
        else:
            des_rows.append({
                "Ubicación original": letra,
                "Encabezado esperado": esperado,
                "Encontrado en origen": "(no existe)",
                "Nueva ubicación del esperado": "—",
            })

    st.subheader("📋 Tabla de Desalineaciones")
    if des_rows:
        df_des = pd.DataFrame(des_rows, columns=[
            "Ubicación original","Encabezado esperado","Encontrado en origen","Nueva ubicación del esperado"
        ])
        st.dataframe(df_des, use_container_width=True)
        make_downloads(df_des, "reporte_desalineaciones", sheet="Desalineaciones")
    else:
        st.success("✅ Todas las columnas coinciden con lo esperado.")

    st.divider()

    # —— Columnas no mapeadas con datos ——
    st.subheader("🟠 Columnas con datos que no estaban en el mapa")
    expected_set_norm = {normalize_header(v) for v in columnas_esperadas.values()}
    extra_rows = []
    for idx, nombre in enumerate(columnas_reales):
        if normalize_header(nombre) not in expected_set_norm:
            datos = df_global.iloc[:, idx].notna().sum()
            if datos > 0:
                extra_rows.append({
                    "Letra": col_index_to_letter(idx),
                    "Encabezado no considerado": nombre,
                    "Registros con datos": int(datos),
                })
    if extra_rows:
        df_extra = pd.DataFrame(extra_rows, columns=["Letra","Encabezado no considerado","Registros con datos"])
        st.dataframe(df_extra, use_container_width=True)
        make_downloads(df_extra, "no_mapeadas_con_datos", sheet="No_mapeadas")
    else:
        st.info("No se encontraron columnas adicionales con datos.")

    st.divider()

    # —— Construcción del archivo final por NOMBRE ——
    st.subheader("🧩 Construcción del archivo final")
    usar_normalizado = st.checkbox("Sugerir coincidencias usando comparación normalizada (aproximada)", value=False)

    mapa_nombre_a_indice = {col: i for i, col in enumerate(columnas_reales)}
    mapa_norm_a_nombre = {normalize_header(col): col for col in columnas_reales}

    columnas_finales = []
    faltantes = []
    sugerencias = []
    for esperado in expected_names:
        if esperado in mapa_nombre_a_indice:
            columnas_finales.append(df_global.iloc[:, mapa_nombre_a_indice[esperado]].rename(esperado))
        else:
            if usar_normalizado:
                norm = normalize_header(esperado)
                if norm in mapa_norm_a_nombre:
                    casi = mapa_norm_a_nombre[norm]
                    sugerencias.append({"Esperado": esperado, "Coincidencia aproximada": casi})
                    columnas_finales.append(df_global.iloc[:, mapa_nombre_a_indice[casi]].rename(esperado))
                else:
                    faltantes.append(esperado)
                    columnas_finales.append(pd.Series([None]*len(df_global), name=esperado))
            else:
                faltantes.append(esperado)
                columnas_finales.append(pd.Series([None]*len(df_global), name=esperado))

    df_resultado = pd.concat(columnas_finales, axis=1)

    # Incluir columnas extra seleccionadas
    st.subheader("📌 Columnas extra con datos para incluir en el final (opcional)")
    if extra_rows:
        opciones_extra = {f"{r['Letra']} – {r['Encabezado no considerado']}": r['Letra'] for r in extra_rows}
        seleccionadas = st.multiselect("Selecciona las columnas extra a incluir:", options=list(opciones_extra.keys()))
        if seleccionadas:
            letras_sel = [opciones_extra[s] for s in seleccionadas]
            idx_sel = [col_letter_to_index(L) for L in letras_sel]
            df_resultado = pd.concat([df_resultado, df_global.iloc[:, idx_sel]], axis=1)
    else:
        st.caption("No hay columnas extra con datos disponibles para añadir.")

    # Añadir origen
    if "Archivo_Origen" in df_global.columns:
        df_resultado["Archivo_Origen"] = df_global["Archivo_Origen"]

    st.subheader("📋 Vista previa – Archivo Final")
    st.dataframe(df_resultado.head(10), use_container_width=True)
    make_downloads(df_resultado, "archivo_consolidado", sheet="Consolidado")

    # Mostrar sugerencias/faltantes si aplica
    if sugerencias:
        with st.expander("Coincidencias aproximadas aplicadas"):
            st.write(pd.DataFrame(sugerencias))
    if faltantes:
        with st.expander("Encabezados faltantes en los archivos cargados"):
            st.write(pd.DataFrame({"Esperado": faltantes}))
