import streamlit as st
import pandas as pd

# —————— Configuración de la página ——————
st.set_page_config(page_title="Reordenador Excel MobilServ v2.0", layout="wide")
st.title("📊 Reordenador Excel a formato MobilServ – Versión 2.0")
st.markdown("**Creado por:** Javier Parada  \n**Ingeniero de Soporte en Campo**")

# —————— Instrucciones ——————
st.markdown("""
### 🧾 Instrucciones de uso:
1. Sube el archivo Excel (.xlsx) con los datos originales.
2. El sistema verificará que los encabezados no hayan sido modificados.
3. Si todo está correcto, se generará una nueva tabla solo con las columnas útiles.
""")

# —————— Utilitario: letra columna Excel → índice 0-based ——————
def col_letter_to_index(letter: str) -> int:
    idx = 0
    for c in letter.upper():
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1

# —————— Diccionario de columnas esperadas (Letra → Nombre) ——————
columnas_esperadas = {
    "A": "NOMBRE_CLIENTE", "Y": "ESTADO", "H": "FECHA_INFORME", "U": "COMPONENTE",
    "X": "DESCRIPTOR_COMPONENTE", "Z": "NIVEL_DE_SERVICIO", "V": "MARCA_COMPONENTE",
    "W": "MODELO_COMPONENTE", "E": "FECHA_MUESTREO", "F": "FECHA_INGRESO",
    "G": "FECHA_RECEPCION", "I": "EDAD_COMPONENTE", "J": "UNIDAD_EDAD_COMPONENTE",
    "K": "EDAD_PRODUCTO", "L": "UNIDAD_EDAD_PRODUCTO", "M": "CANTIDAD_ADICIONADA",
    "N": "UNIDAD_CANTIDAD_ADICIONADA", "O": "PRODUCTO", "B": "NOMBRE_OPERACION",
    "IO": "ÍNDICE PQ (PQI) - 3", "MI": "PLATA (AG) - 19", "AJ": "ALUMINIO (AL) - 20",
    "FK": "CROMO (CR) - 24", "BV": "COBRE (CU) - 25", "IE": "HIERRO (FE) - 26",
    "OZ": "TITANIO (TI) - 38", "MK": "PLOMO (PB) - 35", "JQ": "NÍQUEL (NI) - 32",
    "JJ": "MOLIBDENO (MO) - 30", "OB": "SILICIO (SI) - 36", "OG": "SODIO (NA) - 31",
    "MM": "POTASIO (K) - 27", "PD": "VANADIO (V) - 39", "BI": "BORO (B) - 18",
    "BD": "BARIO (BA) - 21", "BM": "CALCIO (CA) - 22", "BL": "CADMIO (CD) - 23",
    "JE": "MAGNESIO (MG) - 28", "JF": "MANGANESO (MN) - 29", "HQ": "FÓSFORO (P) - 34",
    "PO": "ZINC (ZN) - 40", "BZ": "CÓDIGO ISO (4/6/14) - 47", "FB": "CONTEO PARTÍCULAS >= 4 ΜM - 49",
    "FC": "CONTEO PARTÍCULAS >= 6 ΜM - 50", "FA": "CONTEO PARTÍCULAS >= 14 ΜM - 48",
    "KB": "**OXIDACIÓN - 80", "JR": "**NITRACIÓN - 82", "JU": "NÚMERO ÁCIDO (AN) - 43",
    "JW": "NÚMERO BÁSICO (BN) - 12", "JV": "NÚMERO BÁSICO (BN) - 17", "IG": "**HOLLÍN - 79",
    "GO": "DILUCIÓN POR COMBUSTIBLE - 46", "AE": "**AGUA (IR) - 81", "CS": "CONTENIDO AGUA (KARL FISCHER) - 41",
    "ER": "CONTENIDO GLICOL  - 105", "PG": "VISCOSIDAD A 100 °C - 13", "PH": "VISCOSIDAD A 40 °C - 14",
    "C": "N_MUESTRA"
}

# —————— Subida del archivo ——————
uploaded = st.file_uploader("📤 Sube tu archivo Excel (.xlsx):", type="xlsx")

if uploaded:
    df_original = pd.read_excel(uploaded, header=0, dtype=str)
    columnas_reales = df_original.columns.tolist()

    errores = []
    columnas_validas = []
    nombres_validos = []

    # — Verificar que cada letra tenga el encabezado correcto —
    for letra, nombre_esperado in columnas_esperadas.items():
        idx = col_letter_to_index(letra)
        if idx < len(columnas_reales):
            nombre_real = columnas_reales[idx].strip()
            if nombre_real == nombre_esperado.strip():
                columnas_validas.append(idx)
                nombres_validos.append(nombre_real)
            else:
                errores.append(f"- Columna {letra}: se esperaba **\"{nombre_esperado}\"**, se encontró **\"{nombre_real}\"**")
        else:
            errores.append(f"- Columna {letra}: se esperaba **\"{nombre_esperado}\"**, pero no existe en el archivo")

    # — Mostrar errores o continuar —
    if errores:
        st.error("❌ Las siguientes columnas tienen encabezados modificados o están en una posición incorrecta:")
        st.markdown("\n".join(errores))
        st.stop()
    else:
        st.success("✅ Encabezados verificados correctamente. Procediendo con las columnas válidas...")

        # — Extraer solo las columnas válidas y reordenadas —
        df_resultado = df_original.iloc[:, columnas_validas]
        df_resultado.columns = nombres_validos

        # — Vista previa del nuevo archivo —
        st.subheader("✅ Vista previa – Archivo limpio con columnas válidas")
        st.dataframe(df_resultado.head(10))
