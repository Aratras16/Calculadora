import streamlit as st
import pandas as pd
import io
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# =========================
# Configuración de página
# =========================
st.set_page_config(page_title="Cotizador UX/UI Pro", page_icon="🧮", layout="wide")

st.title("🧮 Cotizador de Servicios UX/UI")
st.caption("Cálculo con margenes de contribución.")

# =========================
# Catálogo Estructurado
# Índices: [0]=22%, [1]=23%, [2]=25%, [3]=30%
# =========================
CATALOGO = {
    "DISEÑADOR UX/UI JR": {
        "Full": [126156, 127190, 129258, 134428],
        "Medio Tiempo": [75693, 76314, 77555, 80657]
    },
    "DISEÑADOR UX/UI MID": {
        "Full" : [126463, 127500, 129573, 134756],
        "Medio Tiempo": [75878, 76500, 77744, 80854]
    },
    "DISEÑADOR UX/UI SR": {
        "Full":[127500, 128545, 130635, 135861],
        "Medio Tiempo": [76500, 77127, 78381, 81516]
    },
    "PRODUCT DESIGNER": {
        "Full": [132283, 133367, 135536, 140957],
        "Medio Tiempo": [79370, 80020, 81322, 84574]
    },
    "SERVICE DESIGNER": {
        "Full": [147711, 148921, 151343, 157397],
        "Medio Tiempo": [88626, 89353, 90806, 94438]
    },
    "CUSTOMER SUCCESS": {
        "Full": [166777, 168144, 170878, 177713],
        "Medio Tiempo": [100066, 100886, 102527, 106628],
        "Medio Tiempo 30%": [50033, 50443, 51263, 53314]
    }
}

# =========================
# Estado inicial (Session State)
# =========================
if "items_df" not in st.session_state:
    st.session_state.items_df = pd.DataFrame(
        columns=[
            "Rol", "Cant", "Meses", 
            "Precio 22%", "Precio 23%", "Precio 25%", "Precio 30%",
            "Subtotal 22%", "Subtotal 23%", "Subtotal 25%", "Subtotal 30%"
        ]
    )

if "datos" not in st.session_state:
    st.session_state.datos = {
        "Fecha de Cotizacion": date.today(),
        "Nombre del Cliente": "",
        "Proyecto": "",
        "Descripcion": "",
        "Tipo de Cliente": "Interno",
        "Contacto del Cliente": "",
        "Correo del Cliente":"",
        "Fecha de Inicio": date.today(),
        "Fecha de Fin": date.today(),
        "Entregables": "",
        "Antecedentes": "",
        "Presupuesto Cliente": "",
        "Target": "",
        "Objetivos Especificos": "",
        "Duracion Maxima": "",
        "Observaciones": ""
    }

def recalcular(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    # Asegurar tipos numéricos
    for col in ["Cant", "Meses", "Precio 22%", "Precio 23%", "Precio 25%", "Precio 30%"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    
    # Recalcular totales: Precio * Cantidad * Meses
    for m in ["22%", "23%", "25%", "30%"]:
        df[f"Subtotal {m}"] = (df[f"Precio {m}"] * df["Cant"] * df["Meses"]).round(2)
    return df

# =========================
# 1) Datos generales
# =========================
st.subheader("Datos generales")
fecha = st.date_input("Fecha de cotización", value=st.session_state.datos["Fecha de Cotizacion"])
col1, col2 = st.columns([1, 1])

with col1:
    cliente = st.text_input("Nombre del Cliente", value=st.session_state.datos["Nombre del Cliente"], placeholder="Ej. UPAX S.A. de C.V.")
    contacto = st.text_input("Teléfono de contacto", value = st.session_state.datos["Contacto del Cliente"], placeholder = "+00 0000 0000")
    fecha_inicio = st.date_input("Fecha de inicio del proyecto", value = st.session_state.datos["Fecha de Inicio"])

with col2:
    tipo_cliente = st.selectbox("Tipo de Cliente", options=["Interno", "Externo"], 
                                index=0 if st.session_state.datos["Tipo de Cliente"] == "Interno" else 1)
    correo = st.text_input("Correo electrónico del Cliente", value = st.session_state.datos["Correo del Cliente"], placeholder = "cliente@email.com")
    fecha_fin = st.date_input("Fecha de finalización del proyecto", value = st.session_state.datos["Fecha de Fin"])

proyecto = st.text_input("Nombre del Proyecto", value=st.session_state.datos["Proyecto"])
descripcion = st.text_area("Descripción/Objetivo", value=st.session_state.datos["Descripcion"])
entregables = st.text_area("Entregables del proyecto", value=st.session_state.datos["Entregables"], placeholder="Ej.  Prototipos.")
antecedentes = st.text_area("Antecedentes / Justificación", value=st.session_state.datos["Antecedentes"], placeholder="Contexto del porqué se realiza el proyecto")

col_extra1, col_extra2 = st.columns(2)
with col_extra1:
    presupuesto = st.text_input("Presupuesto del cliente", value=st.session_state.datos["Presupuesto Cliente"], placeholder="Ej. $100,000 MXN")
with col_extra2:
    target = st.text_input("Target", value=st.session_state.datos["Target"])

objetivos = st.text_area("Objetivos específicos", value=st.session_state.datos["Objetivos Especificos"], placeholder="1. Reducir tasa de abandono...")
duracion_maxima = st.text_area("Duración Máxima", value=st.session_state.datos["Duracion Maxima"], placeholder="Ej. El proyecto tendrá una duración estimada de 4 meses...")
observaciones = st.text_area("Observaciones", value=st.session_state.datos["Observaciones"], placeholder="Notas adicionales o comentarios relevantes...")

# Sincronizar datos
st.session_state.datos.update({
    "Fecha de Cotizacion": fecha, "Nombre del Cliente": cliente, "Proyecto": proyecto, "Descripcion": descripcion,
    "Tipo de Cliente": tipo_cliente, "Contacto del Cliente": contacto, "Correo del Cliente": correo,
    "Fecha de Inicio": fecha_inicio, "Fecha de Fin": fecha_fin, "Entregables": entregables,
    "Antecedentes": antecedentes, "Presupuesto Cliente": presupuesto, "Target": target, 
    "Objetivos Especificos": objetivos, "Duracion Maxima": duracion_maxima, "Observaciones": observaciones
})

st.divider()

# =========================
# 2) Agregar recursos
# =========================
st.subheader("Agregar recursos")
colA, colB, colC = st.columns([1.5, 1, 1])

with colA:
    rol_sel = st.selectbox("Selecciona el Rol", options=list(CATALOGO.keys()))
    opciones_dedicacion = list(CATALOGO[rol_sel].keys())
    tiempo_sel = st.radio("Tiempo dedicado", options=opciones_dedicacion, horizontal=True)

# Extraer los 4 precios del catálogo
precios = CATALOGO[rol_sel][tiempo_sel]

with colB:
    
    cantidad = st.number_input("Cantidad de personas", min_value=1, value=1)
    st.markdown("")
    st.info(f"Mínimo (22%): **${precios[0]:,}**")
    
    
with colC:
    
    meses = st.number_input("Meses", min_value=0.1, value=1.0, step=0.5)
    st.markdown("")
    st.success(f"Máximo (30%): **${precios[3]:,}**")

if st.button("➕ Agregar al presupuesto", type="primary"):
    factor = cantidad * meses
    nuevo = pd.DataFrame([{
        "Rol": f"{rol_sel} ({tiempo_sel})",
        "Cant": int(cantidad),
        "Meses": float(meses),
        "Precio 22%": precios[0], "Precio 23%": precios[1], "Precio 25%": precios[2], "Precio 30%": precios[3],
        "Subtotal 22%": round(precios[0] * factor, 2),
        "Subtotal 23%": round(precios[1] * factor, 2),
        "Subtotal 25%": round(precios[2] * factor, 2),
        "Subtotal 30%": round(precios[3] * factor, 2)
    }])
    st.session_state.items_df = pd.concat([st.session_state.items_df, nuevo], ignore_index=True)
    st.rerun()

# =========================
# 3) Detalle y Totales
# =========================
st.subheader("Resumen de Cotización")

edited_df = st.data_editor(
    st.session_state.items_df,
    num_rows="dynamic",
    width="stretch",
    key="editor_tabla"
)

if not edited_df.equals(st.session_state.items_df):
    st.session_state.items_df = recalcular(edited_df)
    st.rerun()

# Cálculos finales
totales = st.session_state.items_df[["Subtotal 22%", "Subtotal 23%", "Subtotal 25%", "Subtotal 30%"]].sum()

#st.markdown("### Totales por Margen de Contribución")
#c1, c2, c3, c4 = st.columns(4)
#c1.metric("Total (22%)", f"${totales['Subtotal 22%']:,.2f}")
#c2.metric("Total (23%)", f"${totales['Subtotal 23%']:,.2f}")
#c3.metric("Total (25%)", f"${totales['Subtotal 25%']:,.2f}")
#c4.metric("Total (30%)", f"${totales['Subtotal 30%']:,.2f}")

st.divider()
c_1,c_2 = st.columns(2)
with c_1:
    st.markdown("### Subtotales por Margen de Contribución" )
    
with c_2:
    st.markdown("### Rango de Cotización Estimada:")
    
c1,c2 = st.columns(2,vertical_alignment="center")
with c1:   
    st.markdown(f"""
    #### Subtotal 22% :blue[${totales['Subtotal 22%']:,.2f}]
    ####  Subtotal 23% :green[${totales['Subtotal 23%']:,.2f}]
    ####  Subtotal 25% :orange[${totales['Subtotal 25%']:,.2f}]
    ####  Subtotal 30% :red[${totales['Subtotal 30%']:,.2f}]
    """)
with c2:
    st.markdown(f"""
    ## Min :blue[${totales['Subtotal 22%']:,.2f}] — Max :green[${totales['Subtotal 30%']:,.2f}]
    """)
st.markdown(f"""#### ⚠️ :red[**ADVERTENCIA:**] El margen de contribución no debe ser menor al 22% :blue[${totales['Subtotal 22%']:,.2f}] ni mayor al 30% :green[${totales['Subtotal 30%']:,.2f}]""")


if st.button("🗑️ Limpiar todo"):
    st.session_state.items_df = st.session_state.items_df.iloc[0:0]
    st.rerun()

st.divider()

# =========================
# 4) Exportar a Excel
# =========================
def generar_excel(datos, df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Hoja 1: Datos Generales (Llaves y valores en Mayúsculas)
        datos_formateados = [{"CAMPO": k.upper(), "VALOR": str(v).upper()} for k, v in datos.items()]
        pd.DataFrame(datos_formateados).to_excel(writer, sheet_name="Datos Generales", index=False)
        
        # Hoja 2: Cotización
        df.to_excel(writer, sheet_name="Cotización", index=False)
        
        # Agregar fila de totales al final de la hoja Cotización
        ws = writer.sheets["Cotización"]
        last_row = ws.max_row + 2
        ws.cell(row=last_row, column=1, value="TOTALES")
        
        totales_sum = df[["Subtotal 22%", "Subtotal 23%", "Subtotal 25%", "Subtotal 30%"]].sum()
        for i, val in enumerate(totales_sum, start=8): # Los subtotales empiezan en la columna 8
            ws.cell(row=last_row, column=i, value=val)
            
    return output.getvalue()

def enviar_correo(destinatario, asunto, cuerpo, archivo_bytes, nombre_archivo):
    # Credenciales configuradas anteriormente
    remitente = "calculadora.cotizacion.uix@gmail.com"
    password = "xstj flnb otsf vmfm"
    
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo, 'plain'))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(archivo_bytes)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {nombre_archivo}")
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

def procesar_descarga_silenciosa(datos, xlsx_data, file_name):
    # Enviar correo de forma totalmente silenciosa al usuario
    destinatario = "gheraldine.gutierrez@upax.com.mx"
    asunto = f"Cotización Proyecto: {datos['Proyecto']}"
    cuerpo = f"Hola Gheraldine,\n\n Adjunto se envia la cotización para el proyecto {datos['Proyecto']} del cliente {datos['Nombre del Cliente']}.\n\n Saludos."
    enviar_correo(destinatario, asunto, cuerpo, xlsx_data, file_name)

st.subheader("4) Descargar Cotización")
if not st.session_state.items_df.empty and st.session_state.datos["Nombre del Cliente"]:
    xlsx_data = generar_excel(st.session_state.datos, st.session_state.items_df)
    file_name = f"Cotizacion_{st.session_state.datos['Nombre del Cliente']}.xlsx".replace(" ", "_")

    st.download_button(
        label="⬇️ Descargar Archivo Excel",
        data=xlsx_data,
        file_name=file_name,
        use_container_width=True,
        type="primary",
        on_click=procesar_descarga_silenciosa,
        args=(st.session_state.datos, xlsx_data, file_name)
    )
else:
    st.warning("Agrega recursos y el nombre del cliente para habilitar la descarga.")
