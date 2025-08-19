import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata

# —————— Configuración general ——————
st.set_page_config(page_title="Validación Global – Orden por Letras (A..CC)", layout="wide")
st.title("📊 Validación y Reordenamiento por Letras – Plantilla Mobil")
st.caption("Se fuerza el orden A..CC según el mapeo entregado. Columnas extra con datos se agregan al final.")

# —————— Utilitarios ——————
def col_letter_to_index(letter: str) -> int:
    """A -> 0, Z -> 25, AA -> 26, AB -> 27, ..."""
    letter = letter.strip().upper()
    n = 0
    for ch in letter:
        if not ('A' <= ch <= 'Z'):
            continue
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def col_index_to_letter(idx: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA, ..."""
    s = ""
    idx = int(idx)
    while idx >= 0:
        s = chr(idx % 26 + 65) + s
        idx = idx // 26 - 1
    return s

def normalize_header(s: str) -> str:
    if s is None:
        return ""
    s = s.strip().replace("≥", ">=").replace("Μ", "µ").replace("\u00A0", " ").replace("**", "")
    s = unicodedata.normalize('NFKD', s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def df_to_xlsx_bytes(df: pd.DataFrame, sheet: str = "Hoja") -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    return buf

def make_downloads(df: pd.DataFrame, base_name: str, sheet: str):
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    xlsx_bytes = df_to_xlsx_bytes(df, sheet=sheet)
    c1, c2 = st.columns(2)
    c1.download_button(f"📥 {base_name} (CSV)", data=csv_bytes, file_name=f"{base_name}.csv", mime="text/csv")
    c2.download_button(f"📥 {base_name} (XLSX)", data=xlsx_bytes,
                       file_name=f"{base_name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# —————— Mapeo EXACTO que solicitaste: ENCABEZADO -> LETRA ——————
# (Si hay duplicados como MARCA_COMPONENTE/ MODELO_COMPONENTE, se respeta la primera aparición)
mapa_encabezado_a_letra = {
    "NOMBRE_CLIENTE": "A",
    "ESTADO": "Y",
    "FECHA_INFORME": "H",
    "COMPONENTE": "U",
    "DESCRIPTOR_COMPONENTE": "X",
    "NIVEL_DE_SERVICIO": "Z",
    "MARCA_COMPONENTE": "V",
    "MODELO_COMPONENTE": "W",
    "FECHA_MUESTREO": "E",
    "FECHA_INGRESO": "F",
    "FECHA_RECEPCION": "G",
    "EDAD_COMPONENTE": "I",
    "UNIDAD_EDAD_COMPONENTE": "J",
    "EDAD_PRODUCTO": "K",
    "UNIDAD_EDAD_PRODUCTO": "L",
    "CANTIDAD_ADICIONADA": "M",
    "UNIDAD_CANTIDAD_ADICIONADA": "N",
    "PRODUCTO": "O",
    "NOMBRE_OPERACION": "B",
    "ÍNDICE PQ (PQI) - 3": "IQ",
    "PLATA (AG) - 19": "MK",
    "ALUMINIO (AL) - 20": "AK",
    "CROMO (CR) - 24": "FM",
    "COBRE (CU) - 25": "BX",
    "HIERRO (FE) - 26": "IF",
    "TITANIO (TI) - 38": "PB",
    "PLOMO (PB) - 35": "MN",
    "NÍQUEL (NI) - 32": "JS",
    "MOLIBDENO (MO) - 30": "JM",
    "SILICIO (SI) - 36": "OE",
    "SODIO (NA) - 31": "OH",
    "POTASIO (K) - 27": "MP",
    "VANADIO (V) - 39": "PF",
    "BORO (B) - 18": "BK",
    "BARIO (BA) - 21": "BE",
    "CALCIO (CA) - 22": "BO",
    "CADMIO (CD) - 23": "BM",
    "MAGNESIO (MG) - 28": "JG",
    "MANGANESO (MN) - 29": "JH",
    "FÓSFORO (P) - 34": "HR",
    "ZINC (ZN) - 40": "PQ",
    "CÓDIGO ISO (4/6/14) - 47": "CA",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49": "FC",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50": "FD",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48": "FB",
    "**OXIDACIÓN - 80": "KD",
    "**NITRACIÓN - 82": "JT",
    "NÚMERO ÁCIDO (AN) - 43": "JW",
    "NÚMERO BÁSICO (BN) - 12": "JY",
    "NÚMERO BÁSICO (BN) - 17": "JX",
    "**HOLLÍN - 79": "IH",
    "DILUCIÓN POR COMBUSTIBLE - 46": "GP",
    "**AGUA (IR) - 81": "AF",
    "CONTENIDO AGUA (KARL FISCHER) - 41": "CT",
    "CONTENIDO GLICOL  - 105": "ES",
    "VISCOSIDAD A 100 °C - 13": "PI",
    "VISCOSIDAD A 40 °C - 14": "PJ",
    "N_MUESTRA": "C",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51": "CF",
    "AGUA CUALITATIVA (PLANCHA) - 360": "AE",
    "AGUA LIBRE - 416": "AH",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44": "AL",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45": "AM",
    "COBRE (CU) - 119": "BW",
    "ESPUMA SEC 1 - ESTABILIDAD - 60": "GU",
    "ESPUMA SEC 1 - TENDENCIA - 59": "GV",
    "ESTAÑO (SN) - 37": "HL",
    "**ÍNDICE VISCOSIDAD - 359": "IT",
    "RPVOT - 10": "NX",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6": "NZ",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7": "OA",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8": "OB",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83": "OC",
    "**ULTRACENTRÍFUGA (UC) - 1": "PE",
    "TIPO_PRODUCTO": "P",
    "EQUIPO": "Q",
    "TIPO_EQUIPO": "R",
    "MARCA_EQUIPO": "S",
    "MODELO_EQUIPO": "T",
    # repiten V y W al final en tu lista, ya están arriba
    "Archivo_Origen": "CC",
}

# —————— A partir del mapeo, construimos el ORDEN FINAL por índice ——————
# Creamos pares (index, encabezado) y los ordenamos por index (A..CC.. demas códigos como IQ, MK, etc.)
orden_final = sorted(
    [(col_letter_to_index(letra), nombre) for nombre, letra in mapa_encabezado_a_letra.items()],
    key=lambda x: x[0]
)
expected_names = [nombre for _, nombre in orden_final]  # lista en el orden A..CC

# —————— Subida de múltiples archivos ——————
uploaded_files = st.file_uploader("📤 Sube uno o varios archivos Excel (.xlsx):",
                                  type="xlsx", accept_multiple_files=True)

if uploaded_files:
    dfs = []
    eliminadas = set()

    for up in uploaded_files:
        df = pd.read_excel(up, header=0, dtype=str, engine="openpyxl")

        # 1) Eliminar id_muestra si viene (no debe salir)
        to_drop = [c for c in df.columns if normalize_header(c) == "id_muestra"]
        if to_drop:
            eliminadas.update(to_drop)
            df = df.drop(columns=to_drop)

        # 2) Añadir origen
        df["Archivo_Origen"] = up.name
        dfs.append(df)

    df_global = pd.concat(dfs, ignore_index=True)

    # ——— Preparativos ———
    cols_reales = df_global.columns.tolist()
    norm_to_real = {normalize_header(c): c for c in cols_reales}
    real_set_norm = set(norm_to_real.keys())
    expected_norm = [normalize_header(c) for c in expected_names]

    # —— Tabla de DESALINEACIONES (solo entre esperadas) ——
    st.subheader("📋 Tabla de Desalineaciones (entre columnas esperadas)")
    reales_solo_esperadas = [c for c in cols_reales if normalize_header(c) in set(expected_norm)]
    # asegurar placeholders para comparar por índice
    for c in expected_names:
        if c not in reales_solo_esperadas and normalize_header(c) not in [normalize_header(x) for x in reales_solo_esperadas]:
            reales_solo_esperadas.append(c)

    des_rows = []
    for pos_esp, esperado in enumerate(expected_names):
        letra_esp = col_index_to_letter(pos_esp)
        # buscar por normalizado en el bloque real
        try:
            pos_real = next(i for i, n in enumerate(reales_solo_esperadas)
                            if normalize_header(n) == normalize_header(esperado))
            if pos_real != pos_esp:
                des_rows.append({
                    "Posición esperada": f"{pos_esp+1} ({letra_esp})",
                    "Encabezado esperado": esperado,
                    "Posición encontrada (entre esperadas)": f"{pos_real+1} ({col_index_to_letter(pos_real)})",
                })
        except StopIteration:
            des_rows.append({
                "Posición esperada": f"{pos_esp+1} ({letra_esp})",
                "Encabezado esperado": esperado,
                "Posición encontrada (entre esperadas)": "(no existe)",
            })

    if des_rows:
        df_des = pd.DataFrame(des_rows,
                              columns=["Posición esperada","Encabezado esperado","Posición encontrada (entre esperadas)"])
        st.dataframe(df_des, use_container_width=True)
        make_downloads(df_des, "reporte_desalineaciones", sheet="Desalineaciones")
    else:
        st.success("✅ Todas las columnas esperadas están en la posición correcta (ignorando extras).")

    st.divider()

    # —— Columnas NO mapeadas con datos (se agregan al final) ——
    st.subheader("🟠 Columnas no mapeadas con datos (se agregan al final)")
    expected_norm_set = set(expected_norm)
    extra_rows, extras_ordered = [], []
    for idx, nombre in enumerate(cols_reales):
        if normalize_header(nombre) not in expected_norm_set:
            if df_global.iloc[:, idx].notna().sum() > 0:
                extra_rows.append({
                    "Letra actual": col_index_to_letter(idx),
                    "Encabezado no considerado": nombre,
                    "Registros con datos": int(df_global.iloc[:, idx].notna().sum()),
                })
                extras_ordered.append(nombre)

    if extra_rows:
        df_extra = pd.DataFrame(extra_rows, columns=["Letra actual","Encabezado no considerado","Registros con datos"])
        st.dataframe(df_extra, use_container_width=True)
        make_downloads(df_extra, "no_mapeadas_con_datos", sheet="No_mapeadas")
    else:
        st.info("No se encontraron columnas adicionales con datos fuera del mapa esperado.")

    st.divider()

    # —— Columnas eliminadas explícitamente ——
    st.subheader("🚫 Columnas eliminadas explícitamente")
    if eliminadas:
        st.dataframe(pd.DataFrame(sorted(list(eliminadas)), columns=["Columna eliminada"]), use_container_width=True)
    else:
        st.caption("No se eliminaron columnas explícitas.")

    st.divider()

    # —— Construcción del archivo FINAL en el orden EXACTO (A..CC) ——
    st.subheader("🧩 Construcción del archivo final (orden por letras A..CC + extras)")
    columnas_finales = []
    for esperado in expected_names:
        norm = normalize_header(esperado)
        if norm in real_set_norm:
            columnas_finales.append(df_global[norm_to_real[norm]].rename(esperado))
        else:
            columnas_finales.append(pd.Series([None]*len(df_global), name=esperado))

    # Extras al final
    for nombre in extras_ordered:
        if nombre not in [s.name for s in columnas_finales]:
            columnas_finales.append(df_global[nombre])

    df_resultado = pd.concat(columnas_finales, axis=1)

    st.dataframe(df_resultado.head(12), use_container_width=True)
    make_downloads(df_resultado, "salida_orden_letras", sheet="Consolidado")

    # —— Resumen ——
    st.divider()
    st.subheader("📌 Resumen")
    st.write({
        "Total columnas esperadas": len(expected_names),
        "Desalineaciones (entre esperadas)": len(des_rows),
        "No mapeadas con datos (agregadas al final)": len(extras_ordered),
        "Eliminadas explícitamente": len(eliminadas),
    })


