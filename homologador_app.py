import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Homologador Pro", layout="wide")

st.title("🚀 Homologador de Datos Eficiente")

# -----------------------------
# 🔧 FUNCIONES DE UTILIDAD
# -----------------------------
def limpiar_columnas_sin_eliminar(df):
    nuevos_nombres = []
    contador = 1
    nombres_vistos = set()
    for col in df.columns:
        if pd.isna(col) or str(col).startswith("Unnamed"):
            nuevo = f"col_vacia_{contador}"
            contador += 1
        else:
            nuevo = str(col)
        original = nuevo
        i = 1
        while nuevo in nombres_vistos:
            nuevo = f"{original}_{i}"
            i += 1
        nombres_vistos.add(nuevo)
        nuevos_nombres.append(nuevo)
    df.columns = nuevos_nombres
    return df

# -----------------------------
# 🧠 ESTADO DE LA SESIÓN
# -----------------------------
# Usamos un diccionario anidado: {columna: {valor_viejo: valor_nuevo}}
if "reemplazos_globales" not in st.session_state:
    st.session_state.reemplazos_globales = {}

# -----------------------------
# 📂 CARGA Y CONFIGURACIÓN
# -----------------------------
archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if archivo:
    xls = pd.ExcelFile(archivo)
    hojas = xls.sheet_names
    
    with st.sidebar:
        st.header("Configuración Global")
        hojas_seleccionadas = st.multiselect("Hojas a homologar", hojas, default=hojas)
        header_row = st.number_input("Fila de encabezado", min_value=0, value=0)
        
        if st.button("Limpiar todos los cambios"):
            st.session_state.reemplazos_globales = {}
            st.rerun()

    # Seleccionar hoja para trabajar
    hoja_activa = st.selectbox("Selecciona hoja para extraer valores", hojas_seleccionadas)
    df_preview = pd.read_excel(archivo, sheet_name=hoja_activa, header=header_row)
    df_preview = limpiar_columnas_sin_eliminar(df_preview)

    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("Definir Homologación")
        columna_target = st.selectbox("Columna objetivo", df_preview.columns)
        
        # Obtener valores únicos de esa columna
        valores_unicos = sorted(df_preview[columna_target].dropna().astype(str).unique())
        valor_original = st.selectbox("Valor original", valores_unicos)
        
        nuevo_valor = st.text_input("Nuevo valor (o selecciona de la lista)")
        sugerencia = st.selectbox("Existentes en la columna", [""] + valores_unicos)
        
        final_val = nuevo_valor if nuevo_valor else sugerencia

        if st.button("Registrar Cambio"):
            if final_val:
                if columna_target not in st.session_state.reemplazos_globales:
                    st.session_state.reemplazos_globales[columna_target] = {}
                st.session_state.reemplazos_globales[columna_target][valor_original] = final_val
                st.success("Cambio registrado localmente")

    with col2:
        st.subheader("Resumen de cambios acumulados")
        if st.session_state.reemplazos_globales:
            for col, mapping in st.session_state.reemplazos_globales.items():
                with st.expander(f"Columna: {col}"):
                    st.write(pd.DataFrame(mapping.items(), columns=["Original", "Nuevo"]))
        else:
            st.info("Aún no hay cambios registrados.")

    # -----------------------------
    # 🚀 PROCESAMIENTO FINAL
    # -----------------------------
    st.divider()
    if st.button("Generar Archivo con los cambios", type="primary"):
        output = BytesIO()
        progress_bar = st.progress(0)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for i, nombre_hoja in enumerate(hojas):
                # Leer hoja
                df_temp = pd.read_excel(archivo, sheet_name=nombre_hoja, header=header_row)
                df_temp = limpiar_columnas_sin_eliminar(df_temp)
                
                # Aplicar todos los cambios acumulados si la hoja está seleccionada
                if nombre_hoja in hojas_seleccionadas:
                    for col_map, mapeo in st.session_state.reemplazos_globales.items():
                        if col_map in df_temp.columns:
                            df_temp[col_map] = df_temp[col_map].astype(str).str.strip().replace(mapeo)
                
                # Limpiar nombres de columnas vacías para el Excel final
                df_temp.columns = ["" if "col_vacia_" in str(c) else c for c in df_temp.columns]
                
                df_temp.to_excel(writer, sheet_name=nombre_hoja, index=False)
                progress_bar.progress((i + 1) / len(hojas))

        st.success("¡Homologación completa aplicada a todas las hojas!")
        st.download_button(
            label="📥 Descargar Excel Final",
            data=output.getvalue(),
            file_name="archivo_homologado_total.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )