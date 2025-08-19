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
4. Si hay desalineaciones, verás **dos tablas**:
   - **Migraciones de encabezados:** ubicación donde estaba, el encabezado que se suponía y la nueva ubicación del esperado.
   - **Columnas con datos no contempladas:** encabezados con datos que no están en el mapa esperado.
5. Opcional: reordenar por **NOMBRE** con coincidencia aproximada.
6. Descargar **Excel consolidado** y **reportes** (CSV/XLSX).
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
    s = s.replace("≥", ">=").replace("Μ", "µ").replace(" ", " ")
    s = s.replace("**", "")
    s = unicodedata.normalize('NFKD', s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

# —————— Diccionario esperado (actualizado con tus ubicaciones) ——————
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
    "P": "TIPO_PRODUCTO",
    "Q": "EQUIPO",
    "R": "TIPO_EQUIPO",
    "S": "MARCA_EQUIPO",
    "T": "MODELO_EQUIPO",
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

st.caption(f"Mapa esperado con {len(columnas_esperadas)} ubicaciones.")

if uploaded_files:
    dfs = []
    for up in uploaded_files:
        df = pd.read_excel(up, header=0, dtype=str, engine="openpyxl")
        df["Archivo_Origen"] = up.name
        dfs.append(df)
    df_global = pd.concat(dfs, ignore_index=True)

    columnas_reales = [c.strip() for c in df_global.columns.tolist()]
    expected_names = list(columnas_esperadas.values())
    expected_set = set(expected_names)

    # ====== TABLA 1: Migraciones de encabezados ======
    migraciones_rows = []
    columnas_ok = []
    for letra, esperado in columnas_esperadas.items():
        idx = col_letter_to_index(letra)
        if idx < len(columnas_reales):
            encabezado_en_origen = columnas_reales[idx]
            if encabezado_en_origen == esperado:
                columnas_ok.append(idx)
            else:
                nueva_letra = col_index_to_letter(columnas_reales.index(esperado)) if esperado in columnas_reales else "—"
                migraciones_rows.append({
                    "UBICACIÓN ORIGEN": letra,
                    "ENCABEZADO ESPERADO EN ORIGEN": esperado,
                    "ENCABEZADO QUE ESTÁ EN ORIGEN": encabezado_en_origen,
                    "NUEVA UBICACIÓN DEL ESPERADO": nueva_letra,
                })
        else:
            migraciones_rows.append({
                "UBICACIÓN ORIGEN": letra,
                "ENCABEZADO ESPERADO EN ORIGEN": esperado,
                "ENCABEZADO QUE ESTÁ EN ORIGEN": "(no existe)",
                "NUEVA UBICACIÓN DEL ESPERADO": "—",
            })

    total_esperadas = len(columnas_esperadas)
    total_ok = len(columnas_ok)
    total_err = total_esperadas - total_ok

    c1, c2, c3 = st.columns(3)
    c1.metric("Encabezados esperados", total_esperadas)
    c2.metric("Correctos en posición", total_ok)
    c3.metric("Con desalineación", total_err)

    if migraciones_rows:
        df_migraciones = pd.DataFrame(migraciones_rows, columns=[
            "UBICACIÓN ORIGEN",
            "ENCABEZADO ESPERADO EN ORIGEN",
            "ENCABEZADO QUE ESTÁ EN ORIGEN",
            "NUEVA UBICACIÓN DEL ESPERADO",
        ])
        st.subheader("📌 Migraciones de encabezados (posición real vs. esperada)")
        st.dataframe(df_migraciones, use_container_width=True)
        # Descargas
        mig_csv = df_migraciones.to_csv(index=False).encode("utf-8-sig")
        mig_xlsx = BytesIO()
        with pd.ExcelWriter(mig_xlsx, engine="xlsxwriter") as w:
            df_migraciones.to_excel(w, index=False, sheet_name="Migraciones")
            w.sheets["Migraciones"].set_column(0, 3, 48)
        mig_xlsx.seek(0)
        m1, m2 = st.columns(2)
        m1.download_button("📥 Migraciones (CSV)", data=mig_csv, file_name="migraciones_columnas.csv", mime="text/csv")
        m2.download_button("📥 Migraciones (XLSX)", data=mig_xlsx, file_name="migraciones_columnas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.success("✅ No hay migraciones: todo coincide en su ubicación esperada.")

    st.divider()

    # ====== TABLA 2: Columnas con datos no contempladas ======
    st.subheader("🟠 Columnas con datos que no estaban en el mapa esperado")
    expected_set_norm = {normalize_header(n) for n in expected_set}
    extra_rows = []
    for idx, nombre in enumerate(columnas_reales):
        if normalize_header(nombre) not in expected_set_norm:
            datos = df_global.iloc[:, idx].notna().sum()
            if datos > 0:
                extra_rows.append({
                    "LETRA": col_index_to_letter(idx),
                    "ENCABEZADO NO CONSIDERADO": nombre,
                    "REGISTROS CON DATOS": int(datos),
                })
    if extra_rows:
        df_no_consideradas = pd.DataFrame(extra_rows, columns=["LETRA","ENCABEZADO NO CONSIDERADO","REGISTROS CON DATOS"]) 
        st.dataframe(df_no_consideradas, use_container_width=True)
        ex_csv = df_no_consideradas.to_csv(index=False).encode("utf-8-sig")
        ex_xlsx = BytesIO()
        with pd.ExcelWriter(ex_xlsx, engine="xlsxwriter") as w2:
            df_no_consideradas.to_excel(w2, index=False, sheet_name="No_mapeadas")
            w2.sheets["No_mapeadas"].set_column(0, 2, 42)
        ex_xlsx.seek(0)
        e1, e2 = st.columns(2)
        e1.download_button("📥 No mapeadas (CSV)", data=ex_csv, file_name="no_mapeadas_con_datos.csv", mime="text/csv")
        e2.download_button("📥 No mapeadas (XLSX)", data=ex_xlsx, file_name="no_mapeadas_con_datos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No se detectaron columnas con datos por fuera del mapa.")

    st.divider()

    # ====== REORDENAMIENTO POR NOMBRE (opcional) ======
    st.subheader("🧩 Construcción del archivo final")
    usar_normalizado = st.checkbox("Sugerir coincidencias usando comparación normalizada (aproximada)", value=False)

    mapa_nombre_a_indice = {col: i for i, col in enumerate(columnas_reales)}
    mapa_norm_a_nombre = {normalize_header(col): col for col in columnas_reales}

    columnas_finales = []
    faltantes = []
    sugerencias = []
    for esperado in expected_names:
        if esperado in mapa_nombre_a_indice:
            columnas_finales.append(df_global.iloc[:, mapa_nombre_a_indice[esperado]])
        else:
            if usar_normalizado:
                norm = normalize_header(esperado)
                if norm in mapa_norm_a_nombre:
                    casi = mapa_norm_a_nombre[norm]
                    sugerencias.append({"Esperado": esperado, "Coincidencia aproximada": casi})
                    columnas_finales.append(df_global.iloc[:, mapa_nombre_a_indice[casi]])
                else:
                    faltantes.append(esperado)
                    columnas_finales.append(pd.Series([None] * len(df_global), name=esperado))
            else:
                faltantes.append(esperado)
                columnas_finales.append(pd.Series([None] * len(df_global), name=esperado))

    df_resultado = pd.concat(columnas_finales, axis=1)

    if "Archivo_Origen" in df_global.columns:
        df_resultado["Archivo_Origen"] = df_global["Archivo_Origen"]

    st.subheader("📋 Vista previa – Archivo Final")
    st.dataframe(df_resultado.head(10), use_container_width=True)

    buf_xlsx = BytesIO()
    with pd.ExcelWriter(buf_xlsx, engine="xlsxwriter") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Consolidado")
        writer.sheets["Consolidado"].set_column(0, df_resultado.shape[1]-1, 22)
    buf_xlsx.seek(0)

    st.download_button(
        label="📥 Descargar Excel Final Consolidado",
        data=buf_xlsx,
        file_name="archivo_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
