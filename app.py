import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Buscador Hikvision Pro", layout="wide", page_icon="🔍")

# Carga y preparación de datos
@st.cache_data
def load_data():
    df = pd.read_csv('Lista Hikvision.txt', sep='\t')
    df['Stock'] = pd.to_numeric(df['Stock'], errors='coerce').fillna(0)
    
    # Clasificación por Macro Tipo de Producto
    def clasificar_tipo(row):
        cat = str(row['Cat. madre']).upper()
        line = str(row['Line Focus']).upper()
        mod = str(row['Modelo']).upper()
        
        if 'SWITCH' in line or 'NETWORKING' in cat or 'ROUTER' in line:
            return 'Switches / Networking'
        elif 'NVR' in line or 'NVR' in mod or 'NVR' in cat:
            return 'NVR'
        elif 'DVR' in line or 'DVR' in mod or 'TURBO HD' in line or ('BACK-END' in cat and ('DVR' in mod or 'XVR' in mod)):
            return 'DVR / XVR'
        elif 'P-FRONT-END' in cat or 'D-FRONT-END' in cat or 'CAMERA' in line or 'CAM' in line or 'TURBO HD' in line:
            if 'PTZ' in line or 'PTZ' in mod:
                return 'Cámaras PTZ'
            return 'Cámaras (IP / Análogas)'
        elif 'ACCESS CONTROL' in cat:
            return 'Control de Acceso'
        elif 'INTERCOM' in cat or 'INTERCOM' in line:
            return 'Videoporteros / Intercom'
        elif 'ALARM' in cat:
            return 'Alarmas'
        elif 'HIKMICRO' in cat or 'THERMAL' in line:
            return 'Térmicas / Hikmicro'
        elif 'ACCESSORY' in line or 'HOUSING' in line or 'BRAKET' in line or 'BRACKET' in line:
            return 'Accesorios y Montajes'
        elif 'DISPLAY' in cat or 'LED' in cat or 'MONITOR' in line:
            return 'Monitores / Pantallas / LED'
        else:
            return 'Otros / Varios'

    df['Tipo de Producto'] = df.apply(clasificar_tipo, axis=1)
    return df

try:
    df = load_data()
except Exception as e:
    st.error("No se pudo cargar el archivo 'Lista Hikvision.txt'. Asegúrate de que esté en la misma carpeta.")
    st.stop()

st.title("🔍 Buscador y Selector Comercial Hikvision")
st.caption(f"Base de datos activa: **{len(df):,}** artículos")

# --- BARRA LATERAL: FILTROS ---
st.sidebar.header("🎯 Filtros Generales")

# 1. Filtro Tipo de Producto (Macro Filtro)
tipos_producto = ['Todos'] + sorted(df['Tipo de Producto'].unique().tolist())
tipo_sel = st.sidebar.selectbox("Tipo de Producto", tipos_producto)

df_filtered = df if tipo_sel == 'Todos' else df[df['Tipo de Producto'] == tipo_sel]

# 2. Filtro Marca
marcas = ['Todas'] + sorted(df_filtered['Marca'].dropna().unique().tolist())
marca_sel = st.sidebar.selectbox("Marca", marcas)

if marca_sel != 'Todas':
    df_filtered = df_filtered[df_filtered['Marca'] == marca_sel]

# 3. Filtro Categoría Madre
cats = ['Todas'] + sorted(df_filtered['Cat. madre'].dropna().unique().tolist())
cat_sel = st.sidebar.selectbox("Categoría Madre", cats)

if cat_sel != 'Todas':
    df_filtered = df_filtered[df_filtered['Cat. madre'] == cat_sel]

# 4. Filtro Line Focus
lines = ['Todas'] + sorted(df_filtered['Line Focus'].dropna().unique().tolist())
line_sel = st.sidebar.selectbox("Línea (Line Focus)", lines)

if line_sel != 'Todas':
    df_filtered = df_filtered[df_filtered['Line Focus'] == line_sel]

# 5. Filtro por Segmento
segmentos = ['Todos'] + sorted(df_filtered['Segmento'].dropna().unique().tolist())
seg_sel = st.sidebar.selectbox("Segmento", segmentos)

if seg_sel != 'Todos':
    df_filtered = df_filtered[df_filtered['Segmento'] == seg_sel]

# 6. Filtro de Stock
solo_stock = st.sidebar.checkbox("Mostrar solo con Stock disponible (>0)")
if solo_stock:
    df_filtered = df_filtered[df_filtered['Stock'] > 0]

# --- FILTROS POR CARACTERÍSTICAS TÉCNICAS Y DESCRIPCIÓN ---
st.sidebar.markdown("---")
st.sidebar.header("🛠️ Características Técnicas")

tecnologias_frecuentes = {
    "Resolución / Calidad": ["1080P", "2MP", "4MP", "5MP", "8MP", "4K", "HD"],
    "Funciones / IA": ["AcuSense", "ColorVu", "DarkFighter", "WDR", "Microphone", "Audio", "Built-in MIC", "Deep learning", "Human/Vehicle classification", "Face Recognition", "GPS", "Wi-Fi", "4G", "PoE", "BSD", "ADAS"],
    "Protección / Construcción": ["IP67", "IK10", "Water resistant", "Vandal", "Metal"],
    "Compresión / Visión": ["H.265+", "H.264", "EXIR", "Hybrid light", "Smart Light", "Dual Light"]
}

for cat_tech, tags in tecnologias_frecuentes.items():
    seleccionados = st.sidebar.multiselect(f"{cat_tech}:", tags)
    for kw in seleccionados:
        df_filtered = df_filtered[
            df_filtered['Descripción'].astype(str).str.contains(kw, case=False, na=False) |
            df_filtered['Modelo'].astype(str).str.contains(kw, case=False, na=False)
        ]

busqueda_libre = st.sidebar.text_input("O escribe una palabra clave específica:")
if busqueda_libre:
    mask = (
        df_filtered['Modelo'].astype(str).str.contains(busqueda_libre, case=False, na=False) |
        df_filtered['SAP'].astype(str).str.contains(busqueda_libre, case=False, na=False) |
        df_filtered['Descripción'].astype(str).str.contains(busqueda_libre, case=False, na=False)
    )
    df_filtered = df_filtered[mask]

# --- VISTA PRINCIPAL ---
st.metric("Productos encontrados", len(df_filtered))

precios_cols = ['Diamond', 'Black', 'Ruby', 'Platinum', 'Gold Gremio', 'Gold Plus', 'Gold Prime', 'MAP ARS', 'MAP Web']
cols_a_mostrar = ['SAP', 'Modelo', 'Marca', 'Tipo de Producto', 'Cat. madre', 'Line Focus', 'Segmento', 'Status', 'Stock'] + precios_cols + ['Descripción']
cols_existentes = [c for c in cols_a_mostrar if c in df_filtered.columns]

st.dataframe(
    df_filtered[cols_existentes],
    use_container_width=True,
    height=500
)

# --- FUNCIÓN PARA GENERAR EXCEL CON DISEÑO PROFESIONAL ---
def generar_excel_estetico(dataframe):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lista de Precios Hikvision"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for r in dataframe_to_rows(dataframe[cols_existentes], index=False, header=True):
        ws.append(r)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = data_font
            cell.border = border_thin
            cell.alignment = Alignment(vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

excel_data = generar_excel_estetico(df_filtered)

st.download_button(
    label="📊 Descargar Excel Presentación (.xlsx)",
    data=excel_data,
    file_name="Lista_Precios_Hikvision.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
