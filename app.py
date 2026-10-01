import streamlit as st
import pandas as pd

st.set_page_config(page_title="Buscador Hikvision", layout="wide", page_icon="🔍")

# Carga de datos
@st.cache_data
def load_data():
    df = pd.read_csv('Lista Hikvision.txt', sep='\t')
    df['Stock'] = pd.to_numeric(df['Stock'], errors='coerce').fillna(0)
    return df

try:
    df = load_data()
except Exception as e:
    st.error("No se pudo cargar el archivo 'Lista Hikvision.txt'. Asegúrate de que esté en la misma carpeta.")
    st.stop()

st.title("🔍 Buscador y Selector de Productos Hikvision")
st.caption(f"Base de datos activa: **{len(df):,}** artículos")

# --- BARRA LATERAL: FILTROS ---
st.sidebar.header("🎯 Filtros Principales")

# 1. Filtro Marca
marcas = ['Todas'] + sorted(df['Marca'].dropna().unique().tolist())
marca_sel = st.sidebar.selectbox("Marca", marcas)

df_filtered = df if marca_sel == 'Todas' else df[df['Marca'] == marca_sel]

# 2. Filtro Categoría Madre
cats = ['Todas'] + sorted(df_filtered['Cat. madre'].dropna().unique().tolist())
cat_sel = st.sidebar.selectbox("Categoría Madre", cats)

if cat_sel != 'Todas':
    df_filtered = df_filtered[df_filtered['Cat. madre'] == cat_sel]

# 3. Filtro Line Focus
lines = ['Todas'] + sorted(df_filtered['Line Focus'].dropna().unique().tolist())
line_sel = st.sidebar.selectbox("Línea (Line Focus)", lines)

if line_sel != 'Todas':
    df_filtered = df_filtered[df_filtered['Line Focus'] == line_sel]

# 4. Filtro por Segmento
segmentos = ['Todos'] + sorted(df_filtered['Segmento'].dropna().unique().tolist())
seg_sel = st.sidebar.selectbox("Segmento", segmentos)

if seg_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Segmento'] == seg_sel]

# 5. Filtro de Stock
solo_stock = st.sidebar.checkbox("Mostrar solo con Stock disponible (>0)")
if solo_stock:
    df_filtered = df_filtered[df_filtered['Stock'] > 0]

# --- FILTROS POR CARACTERÍSTICAS TÉCNICAS (EN LA DESCRIPCIÓN) ---
st.sidebar.markdown("---")
st.sidebar.header("🛠️ Filtro por Características")

# Selección rápida de características frecuentes
caracteristicas_populares = st.sidebar.multiselect(
    "Selecciona palabras clave técnicas:",
    ["1080P", "4K", "Wi-Fi", "GPS", "PoE", "4G", "BSD", "ADAS", "Microphone", "SD card", "Audio", "Intercom"]
)

for kw in caracteristicas_populares:
    df_filtered = df_filtered[
        df_filtered['Descripción'].astype(str).str.contains(kw, case=False, na=False) |
        df_filtered['Modelo'].astype(str).str.contains(kw, case=False, na=False)
    ]

# Búsqueda libre
busqueda_libre = st.sidebar.text_input("O escribe otra característica o modelo:")
if busqueda_libre:
    mask = (
        df_filtered['Modelo'].astype(str).str.contains(busqueda_libre, case=False, na=False) |
        df_filtered['SAP'].astype(str).str.contains(busqueda_libre, case=False, na=False) |
        df_filtered['Descripción'].astype(str).str.contains(busqueda_libre, case=False, na=False)
    )
    df_filtered = df_filtered[mask]

# --- RESULTADOS ---
col1, col2 = st.columns([1, 3])
with col1:
    st.metric("Productos encontrados", len(df_filtered))

# Columnas para mostrar en la tabla interactiva
cols_a_mostrar = ['SAP', 'Modelo', 'Marca', 'Cat. madre', 'Line Focus', 'Segmento', 'Status', 'Stock', 'MAP ARS', 'MAP Web', 'Descripción']
cols_existentes = [c for c in cols_a_mostrar if c in df_filtered.columns]

st.dataframe(
    df_filtered[cols_existentes],
    use_container_width=True,
    height=500
)

# Exportación de resultados
@st.cache_data
def convert_df(df_input):
    return df_input.to_csv(index=False).encode('utf-8')

st.download_button(
    label="📥 Descargar resultados filtrados (Excel/CSV)",
    data=convert_df(df_filtered),
    file_name="seleccion_hikvision.csv",
    mime="text/csv"
)
