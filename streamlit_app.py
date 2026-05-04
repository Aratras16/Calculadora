import streamlit as st
import pandas as pd
import io
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# =========================
# Configuración de página
# =========================
st.set_page_config(page_title="Cotizador UX/UI", page_icon="🧮", layout="wide")

# =========================
# Estilos CSS Avanzados (Tema Claro)
# =========================
def inyectar_css():
    st.markdown("""
        <style>
        /* Importar fuente moderna y corporativa */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

        /* Variables globales (Light Theme base) */
        :root {
            --primary-color: #0E2B5C;      /* Azul fuerte corporativo */
            --secondary-color: #3B82F6;    /* Azul brillante */
            --accent-color: #10B981;       /* Verde acento */
            --bg-color: #F8FAFC;           /* Fondo general más suave que el blanco puro */
            --card-bg: #FFFFFF;            /* Fondo de tarjetas */
            --text-main: #1E293B;          /* Texto oscuro para legibilidad */
            --text-muted: #64748B;         /* Texto secundario */
            --border-color: #E2E8F0;       /* Bordes muy sutiles */
        }

        /* Estilo base de Streamlit */
        .stApp {
            background-color: var(--bg-color);
            font-family: 'Inter', sans-serif !important;
            color: var(--text-main);
        }

        h1, h2, h3, h4, h5, h6, .stMarkdown, .stText, p, label, li {
            font-family: 'Inter', sans-serif !important;
        }

        h4, h5, h6, .stMarkdown, .stText, p, label, li {
            color: var(--text-main) !important;
        }

        h1, h2, h3 {
            color: var(--primary-color) !important;
        }

        /* Botones */
        button[kind="primary"] {
            background: linear-gradient(135deg, var(--secondary-color), var(--primary-color)) !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            padding: 0.6rem 1.2rem !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 4px 6px -1px rgba(59, 130, 246, 0.2), 0 2px 4px -1px rgba(59, 130, 246, 0.1) !important;
        }
        
        button[kind="primary"]:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.3), 0 4px 6px -2px rgba(59, 130, 246, 0.15) !important;
            opacity: 0.95 !important;
        }

        button[kind="secondary"] {
            background: rgba(255, 255, 255, 0.5) !important;
            color: var(--text-main) !important;
            border: 1px solid var(--border-color) !important;
            border-radius: 8px !important;
            font-weight: 500 !important;
            transition: all 0.3s ease !important;
        }
        
        button[kind="secondary"]:hover {
            border-color: var(--secondary-color) !important;
            color: var(--secondary-color) !important;
            background: rgba(59, 130, 246, 0.05) !important;
            transform: translateY(-1px) !important;
        }

        /* Inputs de textos, selectbox y fechas */
        .stTextInput input, .stTextArea textarea, .stDateInput input, .stSelectbox select, .stNumberInput input, div[data-baseweb="select"] > div {
            border-radius: 6px !important;
            border: 1px solid var(--border-color) !important;
            transition: border-color 0.2s, box-shadow 0.2s !important;
            background-color: var(--card-bg) !important;
            color: var(--text-main) !important;
            -webkit-text-fill-color: var(--text-main) !important;
        }

        .stTextInput input:focus, .stTextArea textarea:focus, .stDateInput input:focus, .stSelectbox select:focus, .stNumberInput input:focus {
            border-color: var(--secondary-color) !important;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.2) !important;
            outline: none !important;
        }

        /* Expander Title */
        .streamlit-expanderHeader {
            font-weight: 600 !important;
            color: var(--primary-color) !important;
            font-size: 1.1rem !important;
            background-color: var(--card-bg) !important;
            border-radius: 8px !important;
        }

        /* Tarjeta de Métricas custom */
        .metric-container {
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 1.8rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
        }
        
        .metric-container:hover {
            transform: translateY(-4px);
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
            border-color: var(--secondary-color);
        }

        .metric-title {
            font-size: 0.95rem;
            color: var(--text-muted);
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 0.6rem;
        }

        .metric-value {
            font-size: 2rem;
            font-weight: 800;
            line-height: 1.2;
        }

        .metric-detail {
            font-size: 0.8rem;
            color: #94A3B8;
            margin-top: 0.3rem;
            font-weight: 500;
        }

        /* Colores semánticos sutiles pero claros */
        .val-21 { color: #6366F1 !important; }  /* Indigo */
        .val-22 { color: #3B82F6 !important; }  /* Azul */
        .val-23 { color: #10B981 !important; }  /* Verde */
        .val-25 { color: #F59E0B !important; }  /* Naranja */
        .val-30 { color: #EF4444 !important; }  /* Rojo */
        .val-40 { color: #8B5CF6 !important; }  /* Violeta */
        .val-50 { color: #EC4899 !important; }  /* Rosa */
        .val-60 { color: #1E293B !important; }  /* Oscuro */

        /* Resaltar cabecera / Banner */
        .hero-banner {
            background: linear-gradient(120deg, var(--card-bg) 0%, #E0F2FE 100%);
            padding: 2.5rem;
            border-radius: 16px;
            margin-bottom: 2rem;
            border-left: 8px solid var(--secondary-color);
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        }
        .hero-banner h1 {
            color: var(--primary-color) !important;
            margin-top: 0 !important;
            font-size: 2.4rem;
            font-weight: 800 !important;
            margin-bottom: 0.5rem;
        }
        .hero-banner p {
            color: var(--text-muted);
            font-size: 1.15rem;
            font-weight: 400;
            margin-bottom: 0;
        }
        </style>
    """, unsafe_allow_html=True)

inyectar_css()

# =========================
# Cabecera Visual (Hero)
# =========================
st.markdown("""
<div class="hero-banner">
    <h1>🧮 Cotizador de Servicios UX/UI</h1>
    <p>Cálculo estructurado con márgenes de contribución para la planeación de proyectos de diseño.</p>
</div>
""", unsafe_allow_html=True)


# =========================
# Catálogo Estructurado
# Índices: [0]=22%, [1]=23%, [2]=25%, [3]=30%
# =========================
CATALOGO = {
    "DEDICADO": {
        "DISEÑADOR UX/UI JR": [125122, 126156, 127190, 129258, 134428, 144769, 155109, 165450],
        "DISEÑADOR UX/UI MID": [125427, 126463, 127500, 129573, 134756, 145122, 155488, 165854],
        "DISEÑADOR UX/UI SR": [126455, 127500, 128545, 130635, 135861, 146311, 156762, 167213],
        "PRODUCT DESIGNER": [131199, 132283, 133367, 135536, 140957, 151800, 162643, 173486],
        "SERVICE DESIGNER": [146500, 147711, 148921, 151343, 157397, 169504, 181612, 193719],
        "CUSTOMER SUCCESS": [165410, 166777, 168144, 170878, 177713, 191384, 205054, 218724]
    },
    "STAFFING": {
        "DISEÑADOR UX/UI JR": [83570, 84261, 84951, 86333, 89786, 96692, 103599, 110506],
        "DISEÑADOR UX/UI MID": [98642, 99457, 100272, 101903, 105979, 114131, 122283, 130435],
        "DISEÑADOR UX/UI SR": [111942, 112867, 113793, 115643, 120269, 129520, 138771, 148023],
        "PRODUCT DESIGNER": [114652, 115600, 116547, 118442, 123180, 132655, 142131, 151606],
        "SERVICE DESIGNER": [135332, 136450, 137569, 139806, 145398, 156582, 167767, 178951],
        "CUSTOMER SUCCESS": [155194, 156477, 157760, 160325, 166738, 179564, 192390, 205216]
    }
}

MARGINS = ["21%", "22%", "23%", "25%", "30%", "40%", "50%", "60%"]

MONEDEROS = {
    "Tiendas Neto" : {
        "Monto": [200,300,400,500],
        "Monto con fee" : [200*1.05,300*1.05,400*1.05,500*1.05]
    },
    "Externo" : {
        "Monto": [200,300,400,500],
        "Monto con fee" : [200*1.15,300*1.15,400*1.15,500*1.15]
    }

}

# =========================
# Estado inicial (Session State)
# =========================
if "items_df" not in st.session_state:
    cols = ["Rol", "Cant", "Tiempo"] + [f"Precio {m}" for m in MARGINS] + [f"Subtotal {m}" for m in MARGINS]
    st.session_state.items_df = pd.DataFrame(columns=cols)

if "datos" not in st.session_state:
    st.session_state.datos = {
        "Fecha de Cotizacion": date.today(),
    }

if "uploaded_pdf" not in st.session_state:
    st.session_state.uploaded_pdf = None

if "hubspot_link" not in st.session_state:
    st.session_state.hubspot_link = ""

if "modalidad_global" not in st.session_state:
    st.session_state.modalidad_global = "DEDICADO"

if "tarifa_global" not in st.session_state:
    st.session_state.tarifa_global = "Mensual"

if "monederos_list" not in st.session_state:
    st.session_state.monederos_list = []  # lista de dicts: {tipo, monto, monto_fee, personas}

def recalcular(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    # Asegurar tipos numéricos
    cols_num = ["Cant", "Tiempo"] + [f"Precio {m}" for m in MARGINS]
    for col in cols_num:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    
    # Recalcular totales: Precio * Cantidad * Tiempo
    for m in MARGINS:
        df[f"Subtotal {m}"] = (df[f"Precio {m}"] * df["Cant"] * df["Tiempo"]).round(2)
    return df

# =========================
# 1) Datos generales
# =========================
st.markdown("### 📄 1. Documentación y Enlaces")

col_ui1, col_ui2 = st.columns([1, 1], gap="large")

with col_ui1:
    uploaded_file = st.file_uploader("📤 Subir PDF del Proyecto", type=["pdf"])
    if uploaded_file:
        st.session_state.uploaded_pdf = uploaded_file
        st.success("✅ Archivo cargado correctamente")

with col_ui2:
    st.session_state.hubspot_link = st.text_input("🔗 Enlace de HubSpot", value=st.session_state.hubspot_link, placeholder="https://app.hubspot.com/...")

st.markdown("<br>", unsafe_allow_html=True)

# Validación de requisitos mínimos para continuar
doc_completa = st.session_state.uploaded_pdf is not None and st.session_state.hubspot_link.strip() != ""

if not doc_completa:
    st.info("📢 **Configuración Requerida:** Por favor, carga el PDF del proyecto y pega el enlace de HubSpot en la Sección 1 para desbloquear las opciones de cotización.", icon="🔒")
    st.stop()

# =========================
# 2) Agregar recursos
# =========================
st.markdown("### 👥 2. Asignación de Recursos")

colLimpia, _ = st.columns([1, 4])
with colLimpia:
    if st.button("🗑️ Limpiar todos los recursos", use_container_width=True):
        st.session_state.items_df = st.session_state.items_df.iloc[0:0]
        st.rerun()

# Selectores globales
col_sel1, col_sel2 = st.columns(2)
with col_sel1:
    modalidad_sel = st.radio("🏷️ Modalidad de la Cotización (solo se puede seleccionar una)", options=["DEDICADO", "STAFFING"], horizontal=True)
with col_sel2:
    tarifa_sel = st.radio("⏱️ Tipo de Tarifa (solo se puede seleccionar uno)", options=["Mensual", "Por Hora"], horizontal=True)

# Detectar cambio de configuración global para limpiar tabla
if modalidad_sel != st.session_state.modalidad_global or tarifa_sel != st.session_state.tarifa_global:
    if not st.session_state.items_df.empty:
        st.session_state.items_df = st.session_state.items_df.iloc[0:0]
        st.warning("⚠️ Se ha cambiado la configuración global. La tabla de recursos ha sido reiniciada para mantener la consistencia.", icon="🗑️")
    st.session_state.modalidad_global = modalidad_sel
    st.session_state.tarifa_global = tarifa_sel

colA, colB, colC = st.columns([1.5, 1, 1], gap="medium")

with colA:
    rol_sel = st.selectbox("👤 Perfil del Especialista", options=list(CATALOGO[st.session_state.modalidad_global].keys()))

# Extraer los 8 precios del catálogo basados en la modalidad global
precios = CATALOGO[st.session_state.modalidad_global][rol_sel]

# Ajustar precios si es por hora
if st.session_state.tarifa_global == "Por Hora":
    precios = [p / 160.0 for p in precios]

with colB:
    cantidad = st.number_input("Cantidad de personas", min_value=1, value=1)
    st.info(f"Tarifa Minima (21%): **${precios[0]:,.2f}**", icon="ℹ️")
    
with colC:
    label_tiempo = "Meses" if st.session_state.tarifa_global == "Mensual" else "Horas"
    step_val = 0.5 if st.session_state.tarifa_global == "Mensual" else 1.0
    val_default = 1.0 if st.session_state.tarifa_global == "Mensual" else 160.0
    tiempo_val = st.number_input(label_tiempo, min_value=0.1, value=val_default, step=step_val)
    st.error(f"Tarifa Máxima (60%): **${precios[-1]:,.2f}**", icon="📈")

colBtnA, _ = st.columns([1, 2])
with colBtnA:
    if st.button("➕ Agregar recurso al presupuesto", type="primary", use_container_width=True):
        factor = cantidad * tiempo_val
        data_nueva = {
            "Rol": f"{rol_sel}",
            "Cant": int(cantidad),
            "Tiempo": float(tiempo_val)
        }
        # Agregar los 8 precios y subtotales
        for i, m in enumerate(MARGINS):
            data_nueva[f"Precio {m}"] = precios[i]
            data_nueva[f"Subtotal {m}"] = round(precios[i] * factor, 2)
            
        nuevo = pd.DataFrame([data_nueva])
        st.session_state.items_df = pd.concat([st.session_state.items_df, nuevo], ignore_index=True)
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# Calcular costo total de monederos (se usará en el resumen)
total_monederos_fee = sum(m["Total c/Fee"] for m in st.session_state.monederos_list)


# =========================
# 3) Detalle de Recursos
# =========================
st.markdown("### 📊 3. Detalle de Recursos")


# Tabla interactiva
label_tiempo_tabla = "Meses" if st.session_state.tarifa_global == "Mensual" else "Horas"
st.markdown(f"<p style='color: var(--text-muted); font-size: 0.95rem;'><em>Puedes editar directamente las Cantidades y {label_tiempo_tabla} en la siguiente tabla.</em></p>", unsafe_allow_html=True)

# Configurar visibilidad de columnas
column_config = {
    "Rol": st.column_config.TextColumn("Rol/Perfil", width="medium"),
    "Cant": st.column_config.NumberColumn("Cant.", min_value=1, step=1, width="small"),
    "Tiempo": st.column_config.NumberColumn(label_tiempo_tabla, min_value=0.1, step=0.5, width="small"),
}

# Ocultar columnas de Precio y configurar Subtotales visibles (21%, 25%, 60%)
for m in MARGINS:
    column_config[f"Precio {m}"] = None  # Ocultar siempre
    if m in ["21%", "25%", "60%"]:
        column_config[f"Subtotal {m}"] = st.column_config.NumberColumn(f"Subtotal {m}", format="$%.2f")
    else:
        column_config[f"Subtotal {m}"] = None  # Ocultar en la tabla UI

edited_df = st.data_editor(
    st.session_state.items_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config=column_config,
    key="editor_tabla"
)

if not edited_df.equals(st.session_state.items_df):
    st.session_state.items_df = recalcular(edited_df)
    st.rerun()

st.divider()

st.markdown("### 👛 4. Monederos")

incluir_monederos = st.toggle("Incluir Monederos en la cotización", value=False, key="toggle_monederos")

if incluir_monederos:
    colM1, colM2, colM3 = st.columns([1.5, 1, 1], gap="medium")

    with colM1:
        tipo_monedero = st.selectbox("🏦 Tipo de Monedero", options=list(MONEDEROS.keys()), key="sel_tipo_monedero")
        montos_disponibles = MONEDEROS[tipo_monedero]["Monto"]
        montos_con_fee = MONEDEROS[tipo_monedero]["Monto con fee"]
        fee_pct = "5%" if tipo_monedero == "Tiendas Neto" else "15%"

    with colM2:
        monto_idx = st.selectbox(
            "💵 Monto por Persona",
            options=range(len(montos_disponibles)),
            format_func=lambda i: f"${montos_disponibles[i]:,.0f}",
            key="sel_monto_monedero"
        )

    with colM3:
        personas_monedero = st.number_input("👤 Número de Personas", min_value=1, value=1, key="num_personas_monedero")
        costo_total_monedero = montos_con_fee[monto_idx] * personas_monedero
        st.success(f"Costo total **${costo_total_monedero:,.2f}**", icon="🧾")

    colBtnM, _ = st.columns([1, 2])
    with colBtnM:
        if st.button("➕ Agregar monedero al presupuesto", type="primary", use_container_width=True, key="btn_add_monedero"):
            st.session_state.monederos_list.append({
                "Tipo": tipo_monedero,
                "Monto Base": montos_disponibles[monto_idx],
                "Fee": fee_pct,
                "Monto c/Fee": round(montos_con_fee[monto_idx], 2),
                "Personas": int(personas_monedero),
                "Total c/Fee": round(costo_total_monedero, 2)
            })
            st.rerun()

    # Mostrar tabla de monederos agregados
    if st.session_state.monederos_list:
        st.markdown("<p style='color: var(--text-muted); font-size:0.9rem; margin-top:1rem;'><em>Monederos agregados a la cotización:</em></p>", unsafe_allow_html=True)
        df_monederos = pd.DataFrame(st.session_state.monederos_list)
        st.dataframe(df_monederos, use_container_width=True, hide_index=True)

        colLimpiaM, _ = st.columns([1, 4])
        with colLimpiaM:
            if st.button("🗑️ Limpiar monederos", use_container_width=True, key="btn_limpiar_monederos"):
                st.session_state.monederos_list = []
                st.rerun()
    else:
        st.info("No hay monederos agregados. Selecciona el tipo, monto y número de personas y presiona el botón.", icon="👛")
else:
    # Si el toggle está apagado, limpiar la lista para que no afecte los totales
    if st.session_state.monederos_list:
        st.session_state.monederos_list = []

st.divider()

# =========================
# 5) Resumen de Totales
# =========================
st.markdown("### 💹 5. Resumen de Totales")

# Cálculos finales
total_monederos_fee = sum(m["Total c/Fee"] for m in st.session_state.monederos_list)
totales = st.session_state.items_df[[f"Subtotal {m}" for m in MARGINS]].sum()

# Generar tarjetas dinámicamente solo para los márgenes seleccionados (21%, 25%, 60%)
cards_html = ""
for m in MARGINS:
    if m in ["21%", "25%", "60%"]:
        val_con_mon = totales[f"Subtotal {m}"] + total_monederos_fee
        m_num = m.replace("%", "")
        monedero_html = f'<div class="metric-detail">Monederos: ${total_monederos_fee:,.2f}</div>' if total_monederos_fee > 0 else ""
        
        cards_html += f"""
<div class="metric-container">
    <div class="metric-title">MARGEN {m}</div>
    <div class="metric-value val-{m_num}">${val_con_mon:,.2f}</div>
    <div class="metric-detail">Recursos: ${totales[f'Subtotal {m}']:,.2f}</div>
    {monedero_html}
</div>
"""

html_layout = f"""
<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 1.2rem; margin-bottom: 2rem;">
    {cards_html}
</div>
"""
st.markdown(html_layout, unsafe_allow_html=True)

t_min = totales[f"Subtotal {MARGINS[0]}"] + total_monederos_fee
t_max = totales[f"Subtotal {MARGINS[-1]}"] + total_monederos_fee
st.warning(f"**⚠️ Regla de Negocio:** El total final (recursos + monederos) no debe ser menor (\${t_min:,.2f}) ({MARGINS[0]}) ni mayor (\${t_max:,.2f}) ({MARGINS[-1]})", icon="🚨")

st.divider()

# =========================
# 5) Exportar a Excel
# =========================

def generar_excel(datos, df, monederos_list=None):
    output = io.BytesIO()
    if monederos_list is None:
        monederos_list = []
    
    # 🎨 Definición de Estilos (Colores Corporativos)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="0E2B5C", end_color="0E2B5C", fill_type="solid")
    center_aligned_text = Alignment(horizontal="center", vertical="center")
    wrap_aligned_text = Alignment(vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color="E2E8F0"), 
        right=Side(style='thin', color="E2E8F0"), 
        top=Side(style='thin', color="E2E8F0"), 
        bottom=Side(style='thin', color="E2E8F0")
    )
    accent_fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    monedero_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    totales_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # ==========================================
        # --- Hoja: Cotización ---
        # ==========================================
        df.to_excel(writer, sheet_name="Cotización", index=False, startrow=2)
    
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb["Cotización"]
    
    label_tiempo_excel = "Meses" if st.session_state.tarifa_global == "Mensual" else "Horas"
    
    # Escribir información global en la cabecera
    modalidad = st.session_state.modalidad_global
    tarifa = st.session_state.tarifa_global
    info_texto = f"📋 DETALLE DE COTIZACIÓN | Modalidad: {modalidad} | Cobro: {tarifa.upper()}"
    
    ws.cell(row=1, column=1, value=info_texto)
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color="FFFFFF")
    ws.cell(row=1, column=1).fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
        
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions['A'].width = 30
        
    for cell in ws["3:3"]: # El header de la tabla ahora está en la fila 3
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_aligned_text
        cell.border = thin_border
            
    for row in ws.iter_rows(min_row=4, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if cell.column >= 4:
                cell.number_format = '"$"#,##0.00'
        
    # ==========================================
    # --- Sección de Monederos en Excel ---
    # ==========================================
    total_monederos_excel = 0
    if monederos_list:
        row_mon_titulo = ws.max_row + 2
        
        # Título de sección
        t_cell = ws.cell(row=row_mon_titulo, column=1, value="👛 MONEDEROS")
        t_cell.font = Font(bold=True, color="0E2B5C", size=11)
        t_cell.fill = accent_fill
        ws.merge_cells(start_row=row_mon_titulo, start_column=1, end_row=row_mon_titulo, end_column=6)

        # Cabecera de monederos
        mon_headers = ["Tipo", "Monto Base", "# de Monederos", "Total"]
        row_mon_header = row_mon_titulo + 1
        for ci, h in enumerate(mon_headers, start=1):
            c = ws.cell(row=row_mon_header, column=ci, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
            c.alignment = center_aligned_text
            c.border = thin_border

        # Filas de monederos
        for ri, mon in enumerate(monederos_list, start=row_mon_header + 1):
            vals = [mon["Tipo"], mon["Monto Base"], mon["Personas"], mon["Total c/Fee"]]
            for ci, v in enumerate(vals, start=1):
                c = ws.cell(row=ri, column=ci, value=v)
                c.border = thin_border
                c.fill = monedero_fill
                c.alignment = Alignment(vertical="center", horizontal="center")
                if ci in (2, 4):  # columnas monetarias (Monto Base y Total)
                    c.number_format = '"$"#,##0.00'
            total_monederos_excel += mon["Total c/Fee"]

        # Fila de total de monederos
        row_mon_total = row_mon_header + len(monederos_list) + 1
        lbl = ws.cell(row=row_mon_total, column=3, value="TOTAL MONEDEROS")
        lbl.font = Font(bold=True, color="0E2B5C")
        lbl.alignment = Alignment(horizontal="right", vertical="center")
        lbl.fill = accent_fill
        lbl.border = thin_border
        val_mon = ws.cell(row=row_mon_total, column=4, value=total_monederos_excel)
        val_mon.number_format = '"$"#,##0.00'
        val_mon.font = Font(bold=True, size=11, color="0E2B5C")
        val_mon.fill = accent_fill
        val_mon.border = thin_border
        val_mon.alignment = center_aligned_text

    # ==========================================
    # --- Totales de Recursos + Monederos ---
    # ==========================================
    row_titulos = ws.max_row + 2
    row_valores = row_titulos + 1
    
    titulo_cell = ws.cell(row=row_titulos, column=1, value="RESUMEN DE TOTALES (Recursos + Monederos)")
    titulo_cell.font = Font(bold=True, color="0E2B5C", size=12)
    
    # Combinar celdas hasta antes de la primera columna de subtotal
    col_inicio_subtotales = df.columns.get_loc(f"Subtotal {MARGINS[0]}") + 1
    ws.merge_cells(start_row=row_titulos, start_column=1, end_row=row_titulos, end_column=col_inicio_subtotales - 1)

    columnas_sumar = [f"Subtotal {m}" for m in MARGINS]
    totales_sum = df[columnas_sumar].sum()
    
    for col_name in columnas_sumar:
        col_idx = df.columns.get_loc(col_name) + 1
        
        c_header = ws.cell(row=row_titulos, column=col_idx, value=f"Total {col_name.split()[-1]}")
        c_header.font = Font(bold=True, color="64748B")
        c_header.fill = totales_fill
        c_header.alignment = center_aligned_text
        c_header.border = thin_border
        
        valor_final = totales_sum[col_name] + total_monederos_excel
        c_val = ws.cell(row=row_valores, column=col_idx, value=valor_final)
        c_val.number_format = '"$"#,##0.00'
        c_val.font = Font(bold=True, size=12, color="1E293B")
        c_val.border = thin_border
        c_val.alignment = center_aligned_text
        c_val.fill = totales_fill

        
    # Mensaje de Advertencia
    t_min_excel = totales_sum[f"Subtotal {MARGINS[0]}"] + total_monederos_excel
    t_max_excel = totales_sum[f"Subtotal {MARGINS[-1]}"] + total_monederos_excel
    msg = f"⚠️ ADVERTENCIA: El total final (recursos + monederos) no debe ser menor (${t_min_excel:,.2f}) ({MARGINS[0]}) ni mayor (${t_max_excel:,.2f}) ({MARGINS[-1]})"
    msg_cell = ws.cell(row=row_titulos + 3, column=1, value=msg)
    msg_cell.font = Font(bold=True, color="EF4444")
    ws.merge_cells(start_row=row_titulos + 3, start_column=1, end_row=row_titulos + 3, end_column=11)

    # Convertir encabezados para que tengan el nombre correcto (label_tiempo_excel)
    ws.cell(row=3, column=3, value=label_tiempo_excel)

    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

def enviar_correo(destinatario, asunto, cuerpo, adjuntos):
    remitente = st.secrets["email"]["cotizacion"]
    password = st.secrets["email"]["cotizacion_pass"]
   
    
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo, 'plain'))

    for archivo_bytes, nombre_archivo in adjuntos:
        if archivo_bytes:
            # Determinar tipo MIME básico
            if nombre_archivo.lower().endswith('.xlsx'):
                main_type, sub_type = 'application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            elif nombre_archivo.lower().endswith('.pdf'):
                main_type, sub_type = 'application', 'pdf'
            else:
                main_type, sub_type = 'application', 'octet-stream'

            part = MIMEBase(main_type, sub_type)
            part.set_payload(archivo_bytes)
            encoders.encode_base64(part)
            # El método add_header maneja correctamente las comillas y evita espacios extras
            part.add_header('Content-Disposition', 'attachment', filename=nombre_archivo)
            msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception:
        return False

def procesar_descarga_silenciosa(xlsx_data, file_name):
    lista_correos = [st.secrets["email"]["correo_1"], st.secrets["email"]["correo_2"]]
    hubspot_link = st.session_state.hubspot_link if st.session_state.hubspot_link else "No proporcionado"
    asunto = "Nueva Cotización Generada"
    cuerpo = f"Hola,\n\nSe ha generado una nueva cotización.\n\nLink de HubSpot: {hubspot_link}\n\nSaludos."
    
    adjuntos = [(xlsx_data, file_name)]
    if st.session_state.uploaded_pdf:
        adjuntos.append((st.session_state.uploaded_pdf.getvalue(), st.session_state.uploaded_pdf.name))
    
    for destinatario in lista_correos:
        enviar_correo(destinatario, asunto, cuerpo, adjuntos)

st.markdown("### 📥 6. Generar Documentación")
st.markdown("<p style='color: var(--text-muted); font-size: 0.95rem;'>Agrega recursos para habilitar la descarga en Excel.</p>", unsafe_allow_html=True)

if not st.session_state.items_df.empty:
    xlsx_data = generar_excel(st.session_state.datos, st.session_state.items_df, st.session_state.monederos_list)
    fecha_str = date.today().strftime("%Y-%m-%d")
    file_name = f"Cotizacion_{fecha_str}.xlsx"

    colDescarga, _ = st.columns([1, 2])
    with colDescarga:
        st.download_button(
            label="⬇️ Descargar Reporte en Excel",
            data=xlsx_data,
            file_name=file_name,
            use_container_width=True,
            type="primary",
            on_click=procesar_descarga_silenciosa,
            args=(xlsx_data, file_name)
        )
else:
    st.info("Para habilitar la descarga, asegúrate de agregar al menos un recurso en la tabla.", icon="💡")

