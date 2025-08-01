import streamlit as st
import pandas as pd
from io import BytesIO

# —————— Configuración general ——————
st.set_page_config(page_title="Validación de Encabezados – Consolidado Excel v3.2", layout="wide")
st.title("📊 Validación de Encabezados – Consolidado Excel – Mobil v3.2")
st.markdown("**Responsables:** Grupo de Soporte en Campo – Mobil")

st.markdown("""
### 🧾 Instrucciones de uso:
1. Sube uno o varios archivos Excel (.xlsx) con los datos originales.
2. El sistema validará que los encabezados no hayan sido modificados.
3. Para cada archivo:
   - Visualiza las columnas validadas.
   - Detecta columnas extra con datos y selecciona si deseas incluirlas.
4. Todos los archivos correctos se **consolidarán en un único Excel final**.
""")

# —————— Funciones utilitarias ——————
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

# —————— Diccionario actualizado de columnas esperadas ——————
columnas_esperadas = {
    "A": "NOMBRE_CLIENTE",
    "B": "NOMBRE_OPERACION",
    "C": "N_MUESTRA",
    "D": "CORRELATIVO",  # Nueva
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
    "P": "TIPO_PRODUCTO",  # Nueva
    "Q": "EQUIPO",         # Nueva
    "R": "TIPO_EQUIPO",    # Nueva
    "S": "MARCA_EQUIPO",   # Nueva
    "T": "MODELO_EQUIPO",  # Nueva
    "U": "COMPONENTE",
    "V": "MARCA_COMPONENTE",
    "W": "MODELO_COMPONENTE",
    "X": "DESCRIPTOR_COMPONENTE",
    "Y": "ESTADO",
    "Z": "NIVEL_DE_SERVICIO",
    "IP": "ÍNDICE PQ (PQI) - 3",
    "MJ": "PLATA (AG) - 19",
    "AJ": "ALUMINIO (AL) - 20",
    "FL": "CROMO (CR) - 24",
    "BW": "COBRE (CU) - 25",
    "IE": "HIERRO (FE) - 26",
    "PA": "TITANIO (TI) - 38",
    "MM": "PLOMO (PB) - 35",
    "JR": "NÍQUEL (NI) - 32",
    "JL": "MOLIBDENO (MO) - 30",
    "OD": "SILICIO (SI) - 36",
    "OG": "SODIO (NA) - 31",
    "MO": "POTASIO (K) - 27",
    "PE": "VANADIO (V) - 39",
    "BJ": "BORO (B) - 18",
    "BD": "BARIO (BA) - 21",
    "BN": "CALCIO (CA) - 22",
    "BL": "CADMIO (CD) - 23",
    "JF": "MAGNESIO (MG) - 28",
    "JG": "MANGANESO (MN) - 29",
    "HQ": "FÓSFORO (P) - 34",
    "PP": "ZINC (ZN) - 40",
    "BZ": "CÓDIGO ISO (4/6/14) - 47",
    "FB": "CONTEO PARTÍCULAS >= 4 ΜM - 49",
    "FC": "CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "FA": "CONTEO PARTÍCULAS >= 14 ΜM - 48",
    "KC": "**OXIDACIÓN - 80",
    "JS": "**NITRACIÓN - 82",
    "JV": "NÚMERO ÁCIDO (AN) - 43",
    "JX": "NÚMERO BÁSICO (BN) - 12",
    "JW": "NÚMERO BÁSICO (BN) - 17",
    "IG": "**HOLLÍN - 79",
    "GO": "DILUCIÓN POR COMBUSTIBLE - 46",
    "AE": "**AGUA (IR) - 81",
    "CS": "CONTENIDO AGUA (KARL FISCHER) - 41",
    "ER": "CONTENIDO GLICOL  - 105",
    "PH": "VISCOSIDAD A 100 °C - 13",
    "PI": "VISCOSIDAD A 40 °C - 14",
    "CE": "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51"  # Nueva
}

# —————— Subida de múltiples archivos ——————
uploaded_files = st.file_uploader("📤 Sube uno o varios archivos Excel (.xlsx):", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    df_consolidado = pd.DataFrame()  # DataFrame único final

    for uploaded in uploaded_files:
        st.markdown(f"---\n## 📂 Procesando archivo: **{uploaded.name}**")
        df_original = pd.read_excel(uploaded, header=0, dtype=str)
        columnas_reales = df_original.columns.tolist()

        errores = []
        columnas_validas = []
        nombres_validos = []
        resumen_validacion = []

        # ——— Validación de columnas esperadas ———
        for letra, nombre_esperado in columnas_esperadas.items():
            idx = col_letter_to_index(letra)
            if idx < len(columnas_reales):
                nombre_real = columnas_reales[idx].strip()
                if nombre_real == nombre_esperado.strip():
                    columnas_validas.append(idx)
                    nombres_validos.append(nombre_real)
                    resumen_validacion.append(f"✅ Columna {letra} = \"{nombre_real}\"")
                else:
                    if nombre_esperado in columnas_reales:
                        nueva_pos = columnas_reales.index(nombre_esperado)
                        nueva_letra = col_index_to_letter(nueva_pos)
                        errores.append(
                            f"- Columna {letra}: se esperaba \"{nombre_esperado}\" pero se encontró \"{nombre_real}\". "
                            f"⚠️ Encontrado en columna {nueva_letra}."
                        )
                    else:
                        errores.append(
                            f"- Columna {letra}: se esperaba \"{nombre_esperado}\" pero se encontró \"{nombre_real}\". "
                            f"⚠️ No se encontró en ninguna otra columna."
                        )
            else:
                errores.append(f"- Columna {letra}: se esperaba \"{nombre_esperado}\" pero no existe en el archivo.")

        if errores:
            st.error(f"❌ El archivo **{uploaded.name}** tiene errores y será omitido.")
            st.markdown("\n".join(errores))
            continue
        else:
            st.success(f"✅ {uploaded.name} validado correctamente")

            with st.expander("🔍 Ver columnas validadas"):
                for linea in resumen_validacion:
                    st.markdown(linea)

            df_resultado = df_original.iloc[:, columnas_validas]
            df_resultado.columns = nombres_validos
            df_resultado["Archivo_Origen"] = uploaded.name

            # —————— Detectar columnas no movidas con datos ——————
            st.subheader("📌 Columnas NO movidas que contienen datos")
            columnas_restantes = [i for i in range(len(columnas_reales)) if i not in columnas_validas]

            reporte_columnas_extra = []
            for idx in columnas_restantes:
                if df_original.iloc[1:, idx].notna().sum() > 0:
                    reporte_columnas_extra.append({
                        "Letra Excel": col_index_to_letter(idx),
                        "Encabezado": columnas_reales[idx],
                        "Ubicación Excel": f"Columna {col_index_to_letter(idx)}",
                        "Index": idx
                    })

            if reporte_columnas_extra:
                df_reporte = pd.DataFrame(reporte_columnas_extra)
                st.dataframe(df_reporte[["Letra Excel", "Encabezado", "Ubicación Excel"]])

                opciones_extra = {f"{row['Letra Excel']} – {row['Encabezado']}": row['Index'] for _, row in df_reporte.iterrows()}
                seleccionadas = st.multiselect(
                    f"Selecciona las columnas extra que deseas incluir para {uploaded.name}:",
                    options=list(opciones_extra.keys())
                )

                if seleccionadas:
                    idx_seleccionados = [opciones_extra[sel] for sel in seleccionadas]
                    df_extra = df_original.iloc[:, idx_seleccionados]
                    df_resultado = pd.concat([df_resultado, df_extra], axis=1)
            else:
                st.info("No hay columnas extra con datos.")

            # Agregar al consolidado
            df_consolidado = pd.concat([df_consolidado, df_resultado], ignore_index=True)

    # —————— Descargar archivo consolidado ——————
    if not df_consolidado.empty:
        st.subheader("📋 Vista previa – Archivo consolidado final")
        st.dataframe(df_consolidado.head(10))

        buffer = BytesIO()
        df_consolidado.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Excel Consolidado",
            data=buffer,
            file_name="archivo_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

