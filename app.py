import streamlit as st
import base64
from datetime import datetime, timedelta, timezone
import io
import pandas as pd
import openpyxl
import urllib.parse

# --- PASO 3: IMPORTACIÓN DE MÓDULO BASE DE DATOS CACHEADO ---
import database as db

# Mantenemos try/except para pytz para prevenir caídas de entorno
try:
    import pytz
except ImportError:
    pytz = None

# --- DEBE SER LA PRIMERA LÍNEA DE STREAMLIT EN TU CÓDIGO ---
st.set_page_config(
    page_title="HOME MEDIC",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="auto"
)

# --- IMPORTACIONES PARA GENERAR EL PDF ORDENADO ---
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import Image as RLImage

def generar_pdf_bonito(df_datos, titulo_reporte, subti_reporte, es_inventario=True):
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # 1. ESTILOS DE TEXTO
    estilo_titulo = ParagraphStyle(
        'TituloPDF', fontName='Helvetica-Bold', fontSize=16, leading=20, alignment=1, textColor=colors.HexColor("#0D47A1")
    )
    estilo_sub = ParagraphStyle(
        'SubPDF', fontName='Helvetica', fontSize=9, leading=14, alignment=1, textColor=colors.HexColor("#555555")
    )
    estilo_celda = ParagraphStyle(
        'CeldaPDF', fontName='Helvetica', fontSize=8, leading=11, alignment=1
    )
    estilo_cabecera = ParagraphStyle(
        'CabeceraPDF', fontName='Helvetica-Bold', fontSize=8, leading=11, alignment=1, textColor=colors.white
    )

    # 2. ENCABEZADO CON LOGO Y TEXTO CENTRADO
    texto_header = [
        Paragraph(titulo_reporte.upper(), estilo_titulo),
        Spacer(1, 4),
        Paragraph(subti_reporte, estilo_sub)
    ]
    
    if st.session_state.get("logo_bytes"):
        try:
            logo_data = base64.b64decode(st.session_state.logo_bytes)
            logo_img = RLImage(io.BytesIO(logo_data), width=90, height=90)
            logo_img.hAlign = 'LEFT'
            
            tabla_header = Table([[logo_img, texto_header]], colWidths=[90, 470])
            tabla_header.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (1,0), (1,0), 'CENTER'),
                ('RIGHTPADDING', (1,0), (1,0), 90),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
            ]))
            story.append(tabla_header)
        except Exception:
            story.extend(texto_header)
    else:
        story.extend(texto_header)
        
    story.append(Spacer(1, 5))
    
    # 3. MAPEADO DE COLUMNAS
    df_temp = df_datos.copy()
    
    col_vencimiento_real = None
    for col in df_temp.columns:
        if str(col).strip() == "Vence" or "vencimiento" in str(col).lower():
            col_vencimiento_real = col
            break
            
    if col_vencimiento_real:
        df_temp["Fecha Vencimiento"] = df_temp[col_vencimiento_real]

    if es_inventario:
        columnas_pdf = ["Nombre Comercial / Marca", "Principio Activo", "Laboratorio / Fabricante", "Presentación", "Stock Unit. Total", "Costo Unit. S/.", "Capital Invertido Total (S/.)", "Fecha Vencimiento"]
    else:
        columnas_pdf = [c for c in df_temp.columns if c != "N° Item"][:8]
        
    for col in columnas_pdf:
        if col not in df_temp.columns:
            df_temp[col] = ""
            
    tabla_data = []
    cabecera_fila = [Paragraph(str(col), estilo_cabecera) for col in columnas_pdf]
    tabla_data.append(cabecera_fila)
    
    total_capital_calculado = 0.0
    total_stock_calculado = 0
    
    # 4. PROCESAMIENTO DE FILAS
    for _, row in df_temp.iterrows():
        fila = []
        
        try:
            stock_val = float(row["Stock Unit. Total"]) if "Stock Unit. Total" in df_temp.columns and str(row["Stock Unit. Total"]).strip() != "" else 0.0
        except Exception:
            stock_val = 0.0
            
        try:
            costo_val = float(row["Costo Unit. S/."]) if "Costo Unit. S/." in df_temp.columns and str(row["Costo Unit. S/."]).strip() != "" else 0.0
        except Exception:
            costo_val = 0.0
            
        capital_fila = stock_val * costo_val
        total_capital_calculado += capital_fila
        total_stock_calculado += int(stock_val)
        
        for col in columnas_pdf:
            val = row[col]
            
            if es_inventario and col == "Stock Unit. Total":
                texto_celda = f"{int(stock_val)}"
            elif es_inventario and col == "Costo Unit. S/.":
                texto_celda = f"S/. {costo_val:,.2f}"
            elif es_inventario and col == "Capital Invertido Total (S/.)":
                texto_celda = f"S/. {capital_fila:,.2f}"
            elif es_inventario and col == "Fecha Vencimiento":
                if pd.notna(val) and str(val).strip() != "" and str(val) != "None" and str(val) != "-":
                    try:
                        texto_celda = pd.to_datetime(val).strftime('%d/%m/%Y')
                    except Exception:
                        texto_celda = str(val)
                else:
                    texto_celda = "-"
            elif any(x in str(col) for x in ["(S/.)", "Monto", "Ganancia", "Total", "Inversión"]):
                try:
                    texto_celda = f"S/. {float(val):,.2f}" if pd.notna(val) and str(val).strip() != "" else "S/. 0.00"
                except ValueError:
                    texto_celda = f"S/. {val}"
            else:
                texto_celda = str(val if pd.notna(val) else "")
                
            fila.append(Paragraph(texto_celda, estilo_celda))
        tabla_data.append(fila)
        
    # 5. FILA DE TOTALES
    fila_totales = []
    for col in columnas_pdf:
        if col == columnas_pdf[0]:
            texto_total = "<b>TOTAL GENERAL</b>"
        elif col == "Capital Invertido Total (S/.)":
            texto_total = f"<b>S/. {total_capital_calculado:,.2f}</b>"
        elif col in ["Stock Unit. Total", "Cant. Unidades Solicitadas", "Unidades Compradas", "Unidades Vendidas", "Unidades", "Entradas (Unid.)", "Salidas (Unid.)"]:
            if es_inventario and col == "Stock Unit. Total":
                texto_total = f"<b>{total_stock_calculado:,} u.</b>"
            else:
                suma_rep = int(pd.to_numeric(df_temp[col], errors='coerce').fillna(0).sum())
                texto_total = f"<b>{suma_rep:,} u.</b>"
        elif not es_inventario and any(x in str(col) for x in ["Monto", "Ganancia", "Total", "Inversión", "Recaudado"]):
            suma_rep = pd.to_numeric(df_temp[col], errors='coerce').fillna(0).sum()
            texto_total = f"<b>S/. {suma_rep:,.2f}</b>"
        else:
            texto_total = ""
            
        estilo_total_celda = ParagraphStyle('TotalCelda', fontName='Helvetica-Bold', fontSize=8, leading=11, alignment=1)
        fila_totales.append(Paragraph(texto_total, estilo_total_celda))
        
    tabla_data.append(fila_totales)
    
    if es_inventario:
        col_widths = [90, 80, 80, 65, 45, 55, 75, 70]
    else:
        num_cols = len(columnas_pdf)
        col_widths = [560 / num_cols] * num_cols

    t = Table(tabla_data, colWidths=col_widths, repeatRows=1)
    
    estilo_tabla = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0D47A1")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('INNERGRID', (0,0), (-1,-2), 0.5, colors.HexColor("#E0E0E0")),
        ('BOX', (0,0), (-1,-2), 1, colors.HexColor("#B0BEC5")),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#ECEFF1")),
        ('LINEABOVE', (0,-1), (-1,-1), 1.5, colors.HexColor("#0D47A1")),
        ('LINEBELOW', (0,-1), (-1,-1), 2.5, colors.HexColor("#0D47A1")),
        ('BOTTOMPADDING', (0,-1), (-1,-1), 6),
        ('TOPPADDING', (0,-1), (-1,-1), 6),
    ])
    
    for i in range(1, len(tabla_data) - 1):
        if i % 2 == 0:
            estilo_tabla.add('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F9F9F9"))
            
    t.setStyle(estilo_tabla)
    story.append(t)
    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

# =====================================================================
# INICIALIZACIÓN DE SESSION STATE Y CONFIGURACIÓN PERSISTENTE DIARIA
# =====================================================================
def obtener_fecha_hoy():
    if pytz:
        try:
            return datetime.now(pytz.timezone("America/Lima")).date()
        except Exception:
            pass
    tz_peru = timezone(timedelta(hours=-5))
    return datetime.now(tz_peru).date()

fecha_hoy_local = obtener_fecha_hoy()
fecha_hoy_str = fecha_hoy_local.strftime('%Y-%m-%d')

# PERSISTENCIA VIA COOKIES / PARÁMETROS NAVEGADOR AL REFRESCAR
cookie_params = st.query_params

if "pantalla" in cookie_params:
    st.session_state.pantalla_activa = cookie_params["pantalla"]

if "auth_user" in cookie_params and "auth_date" in cookie_params:
    user_cookie = cookie_params.get("auth_user")
    date_cookie = cookie_params.get("auth_date")
    
    if date_cookie == fecha_hoy_str and ("logged_in" not in st.session_state or not st.session_state.logged_in):
        res_u = db.ejecutar_query("SELECT usuario, rol FROM usuarios WHERE usuario = %s", (user_cookie,), fetch=True)
        if res_u:
            u_data = res_u[0]
            st.session_state.logged_in = True
            st.session_state.usuario_nombre = u_data[0]
            st.session_state.usuario_rol = u_data[1]
            st.session_state.fecha_login = fecha_hoy_str

if "fecha_login" not in st.session_state:
    st.session_state.fecha_login = None

if st.session_state.fecha_login and st.session_state.fecha_login != fecha_hoy_str:
    st.session_state.logged_in = False
    st.session_state.usuario_rol = None
    st.session_state.usuario_nombre = None

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "usuario_rol" not in st.session_state:
    st.session_state.usuario_rol = None
if "usuario_nombre" not in st.session_state:
    st.session_state.usuario_nombre = None
if "pantalla_activa" not in st.session_state:
    st.session_state.pantalla_activa = "buscar"

# --- PASO 3: REEMPLAZO DE CÁRGA DE CONFIGURACIÓN INICIAL POR FUNCIÓN CACHEADA ---
config = db.cargar_configuracion_db()

if config:
    NOMBRE_NEGOCIO = config["nombre_negocio"]
    TEMA_COLOR = config["tema_color"]
    LOGO_BYTES = config["logo_bytes"]
    FONDO_BYTES = config["fondo_bytes"]
else:
    NOMBRE_NEGOCIO = "Mi Negocio"
    TEMA_COLOR = "#000000"
    LOGO_BYTES = None
    FONDO_BYTES = None

if "nombre_negocio" not in st.session_state:
    st.session_state.nombre_negocio = NOMBRE_NEGOCIO

if "tema_color" not in st.session_state:
    st.session_state.tema_color = TEMA_COLOR

if "logo_bytes" not in st.session_state:
    st.session_state.logo_bytes = LOGO_BYTES

if "fondo_bytes" not in st.session_state:
    st.session_state.fondo_bytes = FONDO_BYTES

if "reset_form" not in st.session_state:
    st.session_state.reset_form = 0

if "boleta_ventas" not in st.session_state:
    st.session_state.boleta_ventas = []

if "carrito_compras" not in st.session_state:
    st.session_state.carrito_compras = []

if "reset_fecha_version" not in st.session_state:
    st.session_state.reset_fecha_version = 0

# =====================================================================
# MOTOR DE ESTILOS CSS DINÁMICOS Y TARJETAS DE PRECIO
# =====================================================================
config_temas = {
    "Celeste Pastel": {
        "bg_color": "#E3F2FD",
        "text_color": "#0D47A1",
        "container_bg": "rgba(255, 255, 255, 0.95)",
        "container_text": "#0D47A1",
        "sidebar_bg": "#E3F2FD",
        "sidebar_text": "#0D47A1"
    },
    "Gris Elegante": {
        "bg_color": "#F5F5F5",
        "text_color": "#212121",
        "container_bg": "rgba(255, 255, 255, 0.98)",
        "container_text": "#212121",
        "sidebar_bg": "#E0E0E0",
        "sidebar_text": "#212121"
    },
    "Azul Profesional": {
        "bg_color": "#0D1B2A",
        "text_color": "#E0E1DD",
        "container_bg": "rgba(27, 38, 59, 0.95)",
        "container_text": "#F5F5F5",
        "sidebar_bg": "#1B263B",
        "sidebar_text": "#E0E1DD"
    },
    "Blanco Puro": {
        "bg_color": "#FFFFFF",
        "text_color": "#1C1C1C",
        "container_bg": "rgba(245, 245, 247, 0.95)",
        "container_text": "#1C1C1C",
        "sidebar_bg": "#FAFAFA",
        "sidebar_text": "#1C1C1C"
    },
    "Oscuro Clásico": {
        "bg_color": "#121212",
        "text_color": "#E0E0E0",
        "container_bg": "rgba(30, 30, 30, 0.95)",
        "container_text": "#FFFFFF",
        "sidebar_bg": "#1E1E1E",
        "sidebar_text": "#E0E0E0"
    }
}

tema_actual = config_temas.get(st.session_state.tema_color, config_temas["Celeste Pastel"])

if st.session_state.fondo_bytes:
    style_bg = f"background-image: url('data:image/png;base64,{st.session_state.fondo_bytes}'); background-size: cover; background-position: center; background-repeat: no-repeat; background-attachment: fixed;"
else:
    style_bg = f"background-color: {tema_actual['bg_color']};"

style_css = f"""
<style>
html, body, .stApp, .main, [data-testid="stAppViewContainer"] {{
    overscroll-behavior-y: none !important;
    overscroll-behavior-x: none !important;
    {style_bg}
}}

div, section, table, .stDataFrame {{
    overscroll-behavior-y: contain !important;
}}

.main .block-container {{
    background-color: {tema_actual['container_bg']} !important;
    color: {tema_actual['container_text']} !important;
    border-radius: 15px;
    padding: 30px;
    box-shadow: 0px 4px 15px rgba(0,0,0,0.15);
    margin-top: 20px;
    margin-bottom: 20px;
}}
.main .block-container p, .main .block-container span, .main .block-container label, .main .block-container h1, .main .block-container h2, .main .block-container h3, .main .block-container h4 {{
    color: {tema_actual['container_text']} !important;
}}
section[data-testid="stSidebar"] {{
    background-color: {tema_actual['sidebar_bg']} !important;
    color: {tema_actual['sidebar_text']} !important;
    overflow-y: auto !important;
    -webkit-overflow-scrolling: touch !important;
}}

section[data-testid="stSidebar"] > div:first-child {{
    padding-top: 0.5rem !important;
}}
section[data-testid="stSidebar"] .block-container {{
    padding-top: 0.5rem !important;
    padding-bottom: 0.5rem !important;
}}

section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {{
    gap: 0.3rem !important;
}}

section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{
    color: {tema_actual['sidebar_text']} !important;
}}

.card-precio {{
    background: white;
    border: 2px solid #0D47A1;
    border-radius: 12px;
    padding: 15px;
    text-align: center;
    box-shadow: 0px 3px 8px rgba(0,0,0,0.12);
    margin-bottom: 15px;
}}
.card-precio-titulo {{
    font-size: 0.95rem;
    font-weight: bold;
    color: #555555;
    text-transform: uppercase;
    margin-bottom: 5px;
}}
.card-precio-monto {{
    font-size: 2.1rem;
    font-weight: 900;
    color: #0D47A1;
}}

.nav-link-btn {{
    display: block;
    width: 100%;
    padding: 8px 12px;
    margin-bottom: 6px;
    border-radius: 8px;
    text-decoration: none !important;
    font-weight: 700;
    font-size: 14px;
    text-align: left;
    transition: all 0.2s ease-in-out;
}}
.nav-link-active {{
    background-color: #0D47A1 !important;
    color: #FFFFFF !important;
    box-shadow: 0px 2px 6px rgba(0,0,0,0.2);
}}
.nav-link-inactive {{
    background-color: transparent !important;
    color: #FFFFFF !important;
    border: 1px solid transparent;
}}
.nav-link-inactive:hover {{
    background-color: rgba(255, 255, 255, 0.15) !important;
    color: #FFFFFF !important;
    border-color: rgba(255, 255, 255, 0.3);
}}
</style>
"""
st.markdown(style_css, unsafe_allow_html=True)

# =====================================================================
# SISTEMA DE AUTENTICACIÓN POP-UP
# =====================================================================
@st.dialog("🔒 Inicio de Sesión - Kardex Farmacia", width="small")
def modal_login():
    st.write("Ingresa tus credenciales para acceder al sistema:")
    with st.form("form_login_dialog", clear_on_submit=False):
        usr = st.text_input("Usuario", key="input_usr_login")
        pwd = st.text_input("Contraseña", type="password", key="input_pwd_login", autocomplete="current-password")
        submit = st.form_submit_button("Ingresar al Sistema", use_container_width=True, type="primary")

        if submit:
            res = db.ejecutar_query("SELECT usuario, password, rol FROM usuarios WHERE usuario = %s AND password = %s", (usr.strip(), pwd.strip()), fetch=True)
            if res:
                u_data = res[0]
                st.session_state.logged_in = True
                st.session_state.usuario_nombre = u_data[0]
                st.session_state.usuario_rol = u_data[2]
                st.session_state.fecha_login = fecha_hoy_str
                
                st.query_params["auth_user"] = u_data[0]
                st.query_params["auth_date"] = fecha_hoy_str
                
                st.toast(f"Bienvenido {u_data[0]} ({u_data[2]})", icon="🔑")
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos")

if not st.session_state.logged_in:
    col_top_ing1, col_top_ing2 = st.columns([3, 1])
    with col_top_ing1:
        st.info("👋 Por favor, inicia sesión para comenzar a trabajar en el sistema.")
    with col_top_ing2:
        if st.button("🔑 Ingrese", type="primary", use_container_width=True):
            modal_login()
            
    if "modal_visto" not in st.session_state:
        st.session_state.modal_visto = True
        modal_login()
    st.stop()

# =====================================================================
# PASO 4: MOVER CONSULTAS Y ALERTAS A SUS RESPECTIVAS FUNCIONES
# =====================================================================
@st.dialog("🚨 NOTIFICACIONES DE STOCK CRÍTICO")
def mostrar_modal_alertas():
    prods = db.cargar_productos_db()
    alertas_sin_stock = []
    alertas_tabletas_bajo = []
    alertas_otros_bajo = []

    if prods:
        for p in prods:
            stock = p['stock_actual_unidades']
            pres = p['presentacion']
            
            if stock == 0:
                alertas_sin_stock.append(f"{p['nombre']} ({p['marca']}) - [{pres}]")
            elif pres == "Tableta / Cápsula" and stock <= 10:
                alertas_tabletas_bajo.append(f"{p['nombre']} — [{pres}] — Quedan: {stock} u.")
            elif pres != "Tableta / Cápsula" and stock < 3:
                alertas_otros_bajo.append(f"{p['nombre']} — [{pres}] — Quedan: {stock} u.")

    with st.container(height=400):
        if alertas_sin_stock:
            st.error(f"🔴 **PRODUCTOS SIN STOCK (0 UNIDADES):** ({len(alertas_sin_stock)})")
            for prod in alertas_sin_stock:
                st.markdown(f"- **{prod}**")
            st.markdown("---")
            
        if alertas_tabletas_bajo:
            st.warning(f"🟡 **STOCK BAJO: TABLETAS / CÁPSULAS (≤ 10 u.):** ({len(alertas_tabletas_bajo)})")
            for prod in alertas_tabletas_bajo:
                st.markdown(f"- {prod}")
            st.markdown("---")
            
        if alertas_otros_bajo:
            st.warning(f"🟡 **STOCK BAJO: OTROS PRODUCTOS (< 3 u.):** ({len(alertas_otros_bajo)})")
            for prod in alertas_otros_bajo:
                st.markdown(f"- {prod}")

# =====================================================================
# MODAL PARA AÑADIR PRODUCTO A LA COMPRA
# =====================================================================
@st.dialog("➕ Añadir Producto a la Compra", width="large")
def modal_agregar_producto_compra(lista_productos):
    st.write("Busca el producto previamente registrado e ingresa el detalle de la compra:")
    
    opciones_prod = {f"{p['nombre']} ({p['marca']}) - [{p['presentacion']}] - Stock: {p['stock_actual_unidades']} u.": p for p in lista_productos}
    prod_sel_key = st.selectbox("🔎 Buscar Producto:", [""] + list(opciones_prod.keys()), index=0)
    
    if prod_sel_key != "":
        p_info = opciones_prod[prod_sel_key]
        
        st.info(f"📦 Configuración empaque: **{p_info['unidades_por_caja']} u./caja** | **{p_info['unidades_por_blister']} u./blíster**")
        
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            cant_cajas = st.number_input("Cajas / Packs Comprados", min_value=0.0, value=0.0, step=0.5, format="%.2f", key="compra_cant_cajas_input")
        with col_c2:
            if p_info['presentacion'] == "Tableta / Cápsula":
                cant_blisters = st.number_input("Blísters Sueltos", min_value=0, value=0, step=1)
            else:
                cant_blisters = 0
                st.caption("*(No aplica blíster)*")
        with col_c3:
            cant_unidades = st.number_input("Unidades Sueltas", min_value=0, value=0, step=1)
            
        col_venc, col_pagado = st.columns(2)
        with col_venc:
            fecha_venc_lote = st.date_input("Fecha Vencimiento del Lote", obtener_fecha_hoy() + timedelta(days=365), format="DD/MM/YYYY")
        with col_pagado:
            costo_total_item = st.number_input("Monto Total Pagado por este Producto (S/.)", min_value=0.0, value=0.0, step=0.5)
            
        tot_unidades_compra = int(round(cant_cajas * p_info['unidades_por_caja'])) + (cant_blisters * p_info['unidades_por_blister']) + cant_unidades
        
        if tot_unidades_compra > 0 and costo_total_item > 0:
            costo_u_calc = costo_total_item / tot_unidades_compra
            st.caption(f"💡 **Total a ingresar:** {tot_unidades_compra} unidades | **Costo unitario calculated:** S/. {costo_u_calc:.4f} por unidad.")
            
        st.markdown("---")
        col_m_btn1, col_m_btn2 = st.columns([1, 1])
        
        with col_m_btn2:
            if st.button("Aceptar", type="primary", use_container_width=True):
                if tot_unidades_compra <= 0:
                    st.error("⚠️ Debes ingresar una cantidad válida.")
                elif costo_total_item <= 0:
                    st.error("⚠️ El costo total pagado debe ser mayor a 0.")
                else:
                    st.session_state.carrito_compras.append({
                        "id_producto": p_info['id_producto'],
                        "nombre": p_info['nombre'],
                        "marca": p_info['marca'],
                        "presentacion": p_info['presentacion'],
                        "cajas": cant_cajas,
                        "blisters": cant_blisters,
                        "unidades": cant_unidades,
                        "total_unidades": tot_unidades_compra,
                        "costo_total": costo_total_item,
                        "fecha_vencimiento": fecha_venc_lote,
                        "unidades_por_caja": p_info['unidades_por_caja'],
                        "unidades_por_blister": p_info['unidades_por_blister']
                    })
                    st.toast(f"✅ Se agregó '{p_info['nombre']}' a la compra", icon="📦")
                    st.rerun()

# =====================================================================
# MODAL INTERACTIVO PARA EDITAR / ELIMINAR VENTAS POR FECHA
# =====================================================================
@st.dialog("Corregir venta", width="small")
def modal_editar_ventas():
    st.write("Selecciona la fecha exacta en el calendario para consultar y corregir los productos vendidos:")
    
    fecha_editar = st.date_input("📅 Fecha:", value=obtener_fecha_hoy(), format="DD/MM/YYYY", key="input_fecha_modal_editar")
    
    query_v_fecha = """
        SELECT m.id_movimiento, m.id_producto, p.nombre, p.marca, p.presentacion, 
               m.unidades, m.blisters, p.unidades_por_blister, m.monto_total
        FROM movimientos m
        JOIN productos p ON m.id_producto = p.id_producto
        WHERE m.fecha = %s AND m.tipo_movimiento = 'VENTA'
        ORDER BY p.nombre ASC, p.marca ASC, m.id_movimiento DESC
    """
    ventas_del_dia = db.ejecutar_query(query_v_fecha, (fecha_editar,), fetch=True)
    
    st.markdown("---")
    st.markdown(f"### 📋 Productos Vendidos el `{fecha_editar.strftime('%d/%m/%Y')}`")
    
    if not ventas_del_dia:
        st.info("ℹ️ No se registraron ventas en la fecha seleccionada.")
    else:
        for item_v in ventas_del_dia:
            id_mov = item_v[0]
            id_prod = item_v[1]
            nom_prod = item_v[2]
            marca_prod = item_v[3]
            pres_prod = item_v[4]
            u_sueltas = item_v[5]
            blis_cant = item_v[6]
            u_por_blis = item_v[7] or 1
            monto_v = float(item_v[8])
            
            tot_unid_mov = u_sueltas + (blis_cant * u_por_blis)
            
            cant_str = f"{u_sueltas} u."
            if blis_cant > 0:
                cant_str += f" + {blis_cant} blís."
                
            col_m_info, col_m_cant, col_m_monto, col_m_del = st.columns([4, 2, 2, 2])
            
            with col_m_info:
                st.markdown(f"**{nom_prod}**  \n<small style='color:gray;'>{marca_prod} - [{pres_prod}]</small>", unsafe_allow_html=True)
            with col_m_cant:
                st.markdown(f"📦 **{cant_str}**  \n<small style='color:gray;'>({tot_unid_mov} u. total)</small>", unsafe_allow_html=True)
            with col_m_monto:
                st.markdown(f"💰 **S/. {monto_v:.2f}**")
            with col_m_del:
                if st.button("🗑️ Eliminar", key=f"btn_del_mov_{id_mov}", type="secondary", use_container_width=True):
                    db.ejecutar_query("DELETE FROM movimientos WHERE id_movimiento = %s", (id_mov,), commit=True)
                    db.ejecutar_query("UPDATE productos SET stock_actual_unidades = stock_actual_unidades + %s WHERE id_producto = %s", (tot_unid_mov, id_prod), commit=True)
                    st.cache_data.clear()
                    st.toast(f"🗑️ Se eliminó la venta de {nom_prod} y se reintegraron {tot_unid_mov} u. al stock.", icon="✅")
                    st.rerun()
            st.markdown("<hr style='margin: 5px 0; border-color: #eee;'>", unsafe_allow_html=True)

# =====================================================================
# BARRA LATERAL MULTIPESTAÑA
# =====================================================================
col_info, col_alerta = st.sidebar.columns([3, 1])

with col_info:
    rol_actual = st.session_state.get("usuario_rol", "admin").capitalize()
    st.markdown(f"**Usuario:** {rol_actual}")

with col_alerta:
    if st.button("🚨", key="btn_alerta_icono", help="Notificaciones de stock bajo o agotado", use_container_width=False):
        mostrar_modal_alertas()

if st.session_state.get("logo_bytes"):
    st.sidebar.image(base64.b64decode(st.session_state.logo_bytes), use_container_width=True)

st.sidebar.markdown(
    f"<h3 style='text-align: center; margin-top: 5px; margin-bottom: 10px; font-size: 1.1rem;'>{st.session_state.nombre_negocio}</h3>", 
    unsafe_allow_html=True
)

opciones_menu = [
    ("buscar", "Buscar Producto"),
    ("registrar", "Registrar Productos"),
    ("venta", "Registrar Venta"),
    ("compra", "Ingresar Compra"),
    ("reportes", "Reportes y Finanzas"),
    ("config", "Configuración")
]

if st.session_state.get("usuario_rol") == "admin":
    opciones_menu.append(("usuarios", "Gestionar Usuarios"))

auth_usr_param = st.query_params.get("auth_user", "")
auth_date_param = st.query_params.get("auth_date", "")

for id_pantalla, nombre_boton in opciones_menu:
    es_activo = (st.session_state.pantalla_activa == id_pantalla)
    clase_css = "nav-link-btn nav-link-active" if es_activo else "nav-link-btn nav-link-inactive"
    label_btn = f"🔹 {nombre_boton}" if es_activo else f"▫️ {nombre_boton}"
    
    query_url = f"?pantalla={id_pantalla}"
    if auth_usr_param and auth_date_param:
        query_url += f"&auth_user={auth_usr_param}&auth_date={auth_date_param}"
        
    st.sidebar.markdown(
        f"<a href='{query_url}' target='_self' class='{clase_css}'>{label_btn}</a>",
        unsafe_allow_html=True
    )

st.sidebar.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

col_vacia, col_puerta = st.sidebar.columns([3, 1])
with col_puerta:
    if st.button("🚪", key="btn_cerrar_sesion", help="Cerrar Sesión", use_container_width=False):
        st.query_params.clear()
        st.session_state.clear()
        st.rerun()

menu_url = st.session_state.pantalla_activa

st.title(f"💊 Kardex - {st.session_state.nombre_negocio}")

# =====================================================================
# PANTALLA 1: BUSCAR Y CONSULTAR PRODUCTO
# =====================================================================
if menu_url == "buscar":
    st.header("🔍 Buscador de Medicamentos")
    
    todos_productos = db.cargar_productos_db()
    
    if not todos_productos:
        st.info("Aún no hay productos registrados.")
    else:
        total_invertido_capital = sum(float(p['stock_actual_unidades']) * float(p['precio_costo_unidad']) for p in todos_productos)
        total_valor_venta_esperado = sum(float(p['stock_actual_unidades']) * float(p['precio_venta_unidad']) for p in todos_productos)
        ganancia_inventario_potencial = total_valor_venta_esperado - total_invertido_capital
        
        st.markdown("### 💰 Valor Actual de tu Mercadería en Almacén")
        col_val1, col_val2, col_val3 = st.columns(3)
        with col_val1:
            st.metric(
                label="💵 Capital Total Invertido (Precio de Costo)", 
                value=f"S/. {total_invertido_capital:,.2f}",
                help="Suma total de lo que te costó comprar la mercadería que te queda en stock."
            )
        with col_val2:
            st.metric(
                label="📈 Valor de Retorno (Precio de Venta)", 
                value=f"S/. {total_valor_venta_esperado:,.2f}",
                help="Suma total de dinero que recibirás si vendes todo tu stock actual por unidades sueltas."
            )
        with col_val3:
            st.metric(
                label="✨ Ganancia Estimada en Almacén", 
                value=f"S/. {ganancia_inventario_potencial:,.2f}",
                help="Es la diferencia entre el valor de venta total y tu capital invertido."
            )
        
        st.markdown("---")
        
        opciones_busqueda = {f"{p['nombre']} ({p['marca']}) - [{p['presentacion']}] - Stock: {p['stock_actual_unidades']} u.": p for p in todos_productos}
        busqueda_sel = st.selectbox("Escribe o selecciona el medicamento para ver detalles:", [""] + list(opciones_busqueda.keys()), index=0)
        
        if busqueda_sel != "":
            p_sel = opciones_busqueda[busqueda_sel]
            id_p = p_sel['id_producto']
            
            st.markdown("---")
            st.markdown(f"### 🏷️ Precios Destacados — `{p_sel['nombre']}`")
            
            col_pre1, col_pre2, col_pre3 = st.columns(3)
            with col_pre1:
                st.markdown(f"""
                <div class="card-precio" style="border-color: #D32F2F;">
                    <div class="card-precio-titulo" style="color: #D32F2F;">🔴 Costo Unitario</div>
                    <div class="card-precio-monto" style="color: #D32F2F;">S/. {p_sel['precio_costo_unidad']:.2f}</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col_pre2:
                st.markdown(f"""
                <div class="card-precio" style="border-color: #2E7D32;">
                    <div class="card-precio-titulo" style="color: #2E7D32;">🟢 Venta por Unidad</div>
                    <div class="card-precio-monto" style="color: #2E7D32;">S/. {p_sel['precio_venta_unidad']:.2f}</div>
                </div>
                """, unsafe_allow_html=True)
                
            with col_pre3:
                val_blister_str = f"S/. {p_sel['precio_venta_blister']:.2f}" if p_sel['presentacion'] == "Tableta / Cápsula" else "N/A"
                st.markdown(f"""
                <div class="card-precio" style="border-color: #0D47A1;">
                    <div class="card-precio-titulo" style="color: #0D47A1;">🔵 Venta por Blíster</div>
                    <div class="card-precio-monto" style="color: #0D47A1;">{val_blister_str}</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            col_izq, col_der = st.columns([2, 3])
            
            with col_izq:
                st.subheader("📋 Ficha Técnica")
                u_caja = p_sel['unidades_por_caja']
                tot_u = p_sel['stock_actual_unidades']
                pres_sel = p_sel['presentacion']
                
                if pres_sel in ["Unidad", "Frasco / Pote", "Jarabe / Líquido", "Crema / Pomada"] or u_caja <= 1:
                    cajas_enteras = 0
                    sueltas = tot_u
                else:
                    cajas_enteras = tot_u // u_caja if u_caja > 0 else 0
                    sueltas = tot_u % u_caja if u_caja > 0 else tot_u
                    
                vence_fmt = p_sel['fecha_vencimiento'].strftime("%d/%m/%Y") if p_sel['fecha_vencimiento'] else "Sin fecha"
                
                capital_invertido_este_prod = tot_u * p_sel['precio_costo_unidad']
                retorno_esperado_este_prod = tot_u * p_sel['precio_venta_unidad']
                
                st.markdown(f"**Medicamento / Producto:** {p_sel['nombre']}")
                st.markdown(f"**Principio Activo:** {p_sel['marca']}")
                st.markdown(f"**Laboratorio / Fabricante:** {p_sel['laboratorio'] or 'No especificado'}")
                st.markdown(f"**Presentación:** {pres_sel}")
                st.markdown(f"**Unidades por Caja/Pack:** {u_caja} u.")
                if pres_sel == "Tableta / Cápsula":
                    st.markdown(f"**Unidades por Blíster:** {p_sel['unidades_por_blister']} u.")
                st.markdown(f"**Vencimiento:** {vence_fmt}")
                st.markdown(f"**Stock Físico:** {cajas_enteras} Cajas + {sueltas} Unidades *(Total: {tot_u} u.)*")
                
                st.markdown("---")
                st.markdown("**Valorización de este producto en almacén:**")
                st.markdown(f"- **Capital retenido en este producto:** S/. {capital_invertido_este_prod:.2f}")
                st.markdown(f"- **Venta estimada si vendes todo:** S/. {retorno_esperado_este_prod:.2f}")
                
                st.markdown("---")
                busqueda_query = f"site:vademecum.es {p_sel['marca']} {p_sel['nombre']}"
                url_vademecum = f"https://www.google.com/search?q={urllib.parse.quote(busqueda_query)}"
                
                st.link_button(
                    label=f"📖 Consultar Usos y Dosis de '{p_sel['nombre']}'",
                    url=url_vademecum,
                    use_container_width=True,
                    help="Abre una pestaña con la información médica, posología e indicaciones de Vademécum."
                )
                
                if st.session_state.usuario_rol == "admin":
                    if "editando_id" not in st.session_state:
                        st.session_state.editando_id = None
                    
                    st.markdown("---")
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.button("✏️ Editar Producto", key="btn_edit", use_container_width=True):
                            st.session_state.editando_id = id_p
                            st.rerun()
                    with col_btn2:
                        confirmar_borrado = st.button("🗑️ Eliminar Producto", key="btn_delete", use_container_width=True)
                    
                    if st.session_state.editando_id == id_p:
                        with st.form("edit_form_popup", clear_on_submit=False):
                            st.write("✏️ **Corregir Datos del Producto**")
                            
                            edit_nombre = st.text_input("Nombre Comercial / Marca", value=p_sel['nombre'], key=f"inp_edit_nom_{id_p}")
                            edit_marca = st.text_input("Principio Activo", value=p_sel['marca'], key=f"inp_edit_mar_{id_p}")
                            edit_laboratorio = st.text_input("Laboratorio / Fabricante", value=p_sel['laboratorio'] or '', key=f"inp_edit_lab_{id_p}")
                            
                            opciones_pres = ["Tableta / Cápsula", "Jarabe / Líquido", "Crema / Pomada", "Frasco / Pote", "Inyectable", "Unidad"]
                            idx_pres = opciones_pres.index(p_sel['presentacion']) if p_sel['presentacion'] in opciones_pres else 0
                            edit_presentacion = st.selectbox("Presentación del Producto", opciones_pres, index=idx_pres, key=f"inp_edit_pre_{id_p}")
                            
                            edit_unidades_caja = st.number_input("Unidades por Caja / Pack", min_value=1, value=int(p_sel['unidades_por_caja'] or 1), step=1, key=f"inp_edit_uc_{id_p}")
                            if edit_presentacion == "Tableta / Cápsula":
                                edit_unidades_blister = st.number_input("Unidades por Blíster", min_value=1, value=int(p_sel['unidades_por_blister'] or 1), step=1, key=f"inp_edit_ub_{id_p}")
                            else:
                                edit_unidades_blister = 1
                            
                            fecha_defecto = p_sel['fecha_vencimiento'] if p_sel['fecha_vencimiento'] else obtener_fecha_hoy()
                            edit_vence = st.date_input("Fecha de Vencimiento", value=fecha_defecto, format="DD/MM/YYYY", key=f"inp_edit_fec_{id_p}")
                            
                            edit_stock = st.number_input("Stock Total (unidades)", min_value=0, value=int(p_sel['stock_actual_unidades']), key=f"inp_edit_stk_{id_p}")
                            edit_costo = st.number_input("Costo Unitario (S/.)", min_value=0.0, value=float(p_sel['precio_costo_unidad']), key=f"inp_edit_cst_{id_p}")
                            edit_venta_u = st.number_input("Venta Unitario (S/.)", min_value=0.0, value=float(p_sel['precio_venta_unidad']), key=f"inp_edit_vtu_{id_p}")
                            edit_venta_b = st.number_input("Venta Blíster (S/.)", min_value=0.0, value=float(p_sel['precio_venta_blister']), key=f"inp_edit_vtb_{id_p}") if edit_presentacion == "Tableta / Cápsula" else 0.0
                            
                            guardar_edit = st.form_submit_button("💾 Guardar Cambios")
                            if guardar_edit:
                                edit_nombre_clean = str(edit_nombre).strip().encode('latin1', errors='ignore').decode('latin1')
                                edit_marca_clean = str(edit_marca).strip().encode('latin1', errors='ignore').decode('latin1')
                                edit_lab_clean = str(edit_laboratorio).strip().encode('latin1', errors='ignore').decode('latin1')
                                
                                db.ejecutar_query(
                                    "UPDATE productos SET nombre=%s, marca=%s, laboratorio=%s, presentacion=%s, unidades_por_caja=%s, unidades_por_blister=%s, fecha_vencimiento=%s, stock_actual_unidades=%s, precio_costo_unidad=%s, precio_venta_unidad=%s, precio_venta_blister=%s WHERE id_producto=%s",
                                    (edit_nombre_clean, edit_marca_clean, edit_lab_clean, edit_presentacion, edit_unidades_caja, edit_unidades_blister, edit_vence, edit_stock, edit_costo, edit_venta_u, edit_venta_b, id_p),
                                    commit=True
                                )
                                st.cache_data.clear()
                                st.toast("💾 ¡Cambios guardados en la base de datos!", icon="✅")
                                st.success("🎉 ¡Producto editado exitosamente!")
                                st.session_state.editando_id = None
                                st.rerun()
                                
                    if confirmar_borrado:
                        st.error(f"⚠️ ¿Eliminar permanentemente '{p_sel['nombre']}' de la base de datos?")
                        col_elim1, col_elim2 = st.columns(2)
                        with col_elim1:
                            if st.button("Sí, Eliminar", key="si_elim"):
                                db.ejecutar_query("DELETE FROM movimientos WHERE id_producto=%s", (id_p,), commit=True)
                                db.ejecutar_query("DELETE FROM productos WHERE id_producto=%s", (id_p,), commit=True)
                                st.cache_data.clear()
                                st.success("¡Eliminado correctamente!")
                                st.rerun()
                        with col_elim2:
                            if st.button("No, Cancelar", key="no_elim"):
                                st.rerun()
                else:
                    st.info("🔒 Como Vendedor solo puedes visualizar la información. No tienes permisos para editar o eliminar.")

            with col_der:
                st.subheader("📈 Rendimiento de Ventas")
                hoy = obtener_fecha_hoy()
                inicio_semana = hoy - timedelta(days=hoy.weekday())
                inicio_mes = hoy.replace(day=1)
                
                query_ventas = """
                    SELECT fecha, monto_total, costo_total_capital, ingreso_neto 
                    FROM movimientos 
                    WHERE id_producto = %s AND tipo_movimiento = 'VENTA'
                """
                ventas_prod = db.ejecutar_query(query_ventas, (id_p,), fetch=True)
                
                v_semana_tot, c_semana_tot, g_semana_tot = 0.0, 0.0, 0.0
                v_mes_tot, c_mes_tot, g_mes_tot = 0.0, 0.0, 0.0
                
                if ventas_prod:
                    for v in ventas_prod:
                        f_v = v[0]
                        if isinstance(f_v, str):
                            f_v = datetime.strptime(f_v, "%Y-%m-%d").date()
                        elif isinstance(f_v, datetime):
                            f_v = f_v.date()

                        if f_v >= inicio_semana:
                            v_semana_tot += float(v[1])
                            c_semana_tot += float(v[2])
                            g_semana_tot += float(v[3])
                        if f_v >= inicio_mes:
                            v_mes_tot += float(v[1])
                            c_mes_tot += float(v[2])
                            g_mes_tot += float(v[3])
                
                periodo_sel = st.radio("Filtrar rendimiento por:", ["Esta Semana", "Este Mes"], horizontal=True)
                col_m1, col_m2, col_m3 = st.columns(3)
                
                if periodo_sel == "Esta Semana":
                    with col_m1:
                        st.metric("Venta Total", f"S/. {v_semana_tot:.2f}")
                    with col_m2:
                        st.metric("Costo Capital", f"S/. {c_semana_tot:.2f}")
                    with col_m3:
                        st.metric("Ganancia Neta", f"S/. {g_semana_tot:.2f}")
                else:
                    with col_m1:
                        st.metric("Venta Total", f"S/. {v_mes_tot:.2f}")
                    with col_m2:
                        st.metric("Costo Capital", f"S/. {c_mes_tot:.2f}")
                    with col_m3:
                        st.metric("Ganancia Neta", f"S/. {g_mes_tot:.2f}")

        st.markdown("---")
        st.subheader("📋 Lista de Inventario (Orden Alfabético)")
        
        rows = []
        for p in todos_productos:
            u_caja = p['unidades_por_caja']
            tot_u = p['stock_actual_unidades']
            pres = p['presentacion']
            
            if pres in ["Unidad", "Inyectable", "Frasco / Pote", "Jarabe / Líquido", "Crema / Pomada"] or u_caja <= 1:
                cajas_enteras = 0
                sueltas = tot_u
            else:
                cajas_enteras = tot_u // u_caja if u_caja > 0 else 0
                sueltas = tot_u % u_caja if u_caja > 0 else tot_u
            
            rows.append({
                "Nombre Comercial / Marca": p['nombre'],
                "Principio Activo": p['marca'],
                "Laboratorio / Fabricante": p['laboratorio'] if p['laboratorio'] else "-",
                "Presentación": pres,
                "Unid./Caja": u_caja,
                "Stock (Cajas + Unid.)": f"{cajas_enteras} Cajas + {sueltas} Unid.",
                "Stock Unit. Total": tot_u,
                "Costo Unit. S/.": p['precio_costo_unidad'],
                "Venta Unit. S/.": p['precio_venta_unidad'],
                "Venta Blíster S/.": p['precio_venta_blister'] if pres == "Tableta / Cápsula" else 0.0,
                "Vence": p['fecha_vencimiento'].strftime("%d/%m/%Y") if p['fecha_vencimiento'] else ""
            })
            
        df_inventario = pd.DataFrame(rows)
        
        def colorear_filas(row):
            stock = row["Stock Unit. Total"]
            pres = row["Presentación"]
            if stock == 0:
                return ['background-color: #ffcccc; color: black'] * len(row)
            elif (pres == "Tableta / Cápsula" and stock <= 10) or (pres != "Tableta / Cápsula" and stock < 3):
                return ['background-color: #ffe5cc; color: black'] * len(row)
            return [''] * len(row)

        if not df_inventario.empty:
            st.dataframe(df_inventario.style.apply(colorear_filas, axis=1), use_container_width=True)
            
            df_inv_excel = df_inventario.copy()
            df_inv_excel["Capital Invertido Total (S/.)"] = df_inv_excel["Stock Unit. Total"] * df_inv_excel["Costo Unit. S/."]
            df_inv_excel["Retorno Venta Total (S/.)"] = df_inv_excel["Stock Unit. Total"] * df_inv_excel["Venta Unit. S/."]
            df_inv_excel["Ganancia Esperada (S/.)"] = df_inv_excel["Retorno Venta Total (S/.)"] - df_inv_excel["Capital Invertido Total (S/.)"]
            
            output_inv = io.BytesIO()
            with pd.ExcelWriter(output_inv, engine='openpyxl') as writer:
                df_inv_excel.to_excel(writer, sheet_name="Stock Actual", index=False, startrow=4)
                
                workbook = writer.book
                worksheet = writer.sheets["Stock Actual"]
                
                worksheet["A1"] = f"REPORTE DE INVENTARIO Y STOCK ACTUAL - {st.session_state.nombre_negocio.upper()}"
                worksheet["A2"] = f"Fecha de emision: {obtener_fecha_hoy().strftime('%d/%m/%Y')} {datetime.now().strftime('%H:%M')}"
                worksheet["A3"] = f"Total Capital Invertido en Almacén: S/. {total_invertido_capital:,.2f}"
                
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                    
            excel_inv_data = output_inv.getvalue()
            pdf_data_inv = generar_pdf_bonito(df_inventario, "INVENTARIO Y STOCK DE ALMACÉN", f"Empresa: {st.session_state.nombre_negocio} | Fecha y Hora de Reporte: {obtener_fecha_hoy().strftime('%d/%m/%Y')} {datetime.now().strftime('%H:%M')}", es_inventario=True)

            st.write("")
            col_cen1, col_cen2, col_cen3 = st.columns([2, 1, 2])
            with col_cen2:
                with st.popover("📥 Descargar", use_container_width=True):
                    st.download_button(
                        label="Excel (.xlsx)",
                        data=excel_inv_data,
                        file_name=f"Inventario_Actual_{st.session_state.nombre_negocio.replace(' ', '_')}_{obtener_fecha_hoy().strftime('%Y-%m-%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.download_button(
                        label="PDF (.pdf)",
                        data=pdf_data_inv,
                        file_name=f"Inventario_Actual_{st.session_state.nombre_negocio.replace(' ', '_')}_{obtener_fecha_hoy().strftime('%Y-%m-%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

# =====================================================================
# PANTALLA 2: REGISTRAR NUEVO PRODUCTO
# =====================================================================
elif menu_url == "registrar":
    st.header("📝 Registrar Nuevo Producto en el Inventario")
    
    todos_productos = db.cargar_productos_db()
    
    if st.session_state.usuario_rol != "admin":
        st.warning("🔒 Acceso Restringido: Tu cuenta de Vendedor te permite visualizar productos pero no registrar nuevos. Solicitante a un Administrador.")
    
    disabled_mode = (st.session_state.usuario_rol != "admin")

    if "reset_id" not in st.session_state:
        st.session_state.reset_id = 0
    k = st.session_state.reset_id
    
    labs_query = db.ejecutar_query("SELECT DISTINCT laboratorio FROM productos WHERE laboratorio IS NOT NULL AND laboratorio != '' ORDER BY laboratorio ASC", fetch=True)
    laboratorios_existentes = [l[0] for l in labs_query] if labs_query else []

    dict_plantillas = {f"{p['nombre']} ({p['marca']})": p for p in todos_productos} if todos_productos else {}
    prod_base_sel = st.selectbox(
        "💡 Basarse en un producto ya registrado (Opcional - Autorrellena los datos):",
        ["[ Crear nuevo producto desde cero ]"] + list(dict_plantillas.keys()),
        disabled=disabled_mode,
        key=f"plantilla_prod_{k}"
    )

    val_nombre_defecto = ""
    val_marca_defecto = ""
    val_lab_defecto = "[ Escribir un nuevo laboratorio ]"
    val_pres_defecto = "Tableta / Cápsula"

    if prod_base_sel != "[ Crear nuevo producto desde cero ]":
        p_base = dict_plantillas[prod_base_sel]
        val_nombre_defecto = p_base['nombre']
        val_marca_defecto = p_base['marca']
        val_pres_defecto = p_base['presentacion'] if p_base['presentacion'] in ["Tableta / Cápsula", "Jarabe / Líquido", "Crema / Pomada", "Frasco / Pote", "Inyectable", "Unidad"] else "Tableta / Cápsula"
        if p_base['laboratorio'] and p_base['laboratorio'] in laboratorios_existentes:
            val_lab_defecto = p_base['laboratorio']
    
    st.subheader("Información Básica")
    
    nombre_prod = st.text_input(
        "Nombre Comercial del Producto / Marca", 
        value=val_nombre_defecto, 
        placeholder="Escriba o registre un nuevo producto aquí...", 
        key=f"nombre_{k}", 
        disabled=disabled_mode
    )
    
    col_basica1, col_basica2 = st.columns(2)
    with col_basica1:
        marca_principio = st.text_input(
            "Principio Activo", 
            value=val_marca_defecto, 
            placeholder="Ej: Paracetamol / Piritionato de zinc", 
            key=f"marca_{k}", 
            disabled=disabled_mode
        )
    with col_basica2:
        idx_lab = (["[ Escribir un nuevo laboratorio ]"] + laboratorios_existentes).index(val_lab_defecto) if val_lab_defecto in (["[ Escribir un nuevo laboratorio ]"] + laboratorios_existentes) else 0
        opcion_lab = st.selectbox(
            "Laboratorio / Fabricante (Sugerencia o Ingreso)",
            ["[ Escribir un nuevo laboratorio ]"] + laboratorios_existentes,
            index=idx_lab,
            help="Selecciona de los laboratorios anteriormente registrados o escribe uno nuevo abajo.",
            key=f"opcion_lab_{k}",
            disabled=disabled_mode
        )
        if opcion_lab == "[ Escribir un nuevo laboratorio ]":
            laboratorio = st.text_input(
                "Nombre del Nuevo Laboratorio", 
                placeholder="Escriba el nombre del nuevo laboratorio aquí...", 
                key=f"lab_{k}", 
                disabled=disabled_mode
            )
        else:
            laboratorio = opcion_lab
        
    col_pres_vence1, col_pres_vence2 = st.columns(2)
    with col_pres_vence1:
        opciones_p = ["Tableta / Cápsula", "Jarabe / Líquido", "Crema / Pomada", "Frasco / Pote", "Inyectable", "Unidad"]
        idx_p = opciones_p.index(val_pres_defecto) if val_pres_defecto in opciones_p else 0
        presentacion = st.selectbox(
            "Presentación del Producto",
            opciones_p,
            index=idx_p,
            key=f"pres_{k}",
            disabled=disabled_mode
        )
    with col_pres_vence2:
        fecha_venc = st.date_input(
            "Fecha de Vencimiento del Producto", 
            value=None, 
            format="DD/MM/YYYY",
            help="Selecciona la fecha de caducidad de este lote inicial.",
            key=f"venc_{k}",
            disabled=disabled_mode
        )
        
    st.subheader("Estructura de Empaque y Costos")
    
    metodo_compra = st.radio(
        "¿Cómo adquiriste este producto?",
        ["Compré por Caja / Docena / Pack", "Compré por Unidades Sueltas"],
        horizontal=True,
        key=f"metodo_{k}",
        disabled=disabled_mode
    )
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        unidades_por_caja = st.number_input("Unidades totales que trae la caja / pack", min_value=1, value=1, step=1, key=f"u_caja_{k}", disabled=disabled_mode)
    with col_c2:
        if metodo_compra == "Compré por Caja / Docena / Pack":
            costo_caja = st.number_input("Costo Total de la Caja / Pack (S/.)", min_value=0.0, value=0.0, step=0.10, key=f"c_caja_{k}", disabled=disabled_mode)
        else:
            costo_caja = 0.0
            st.info("💡 Compras por unidades sueltas: ingresa el costo directamente abajo.")
        
    if metodo_compra == "Compré por Caja / Docena / Pack":
        if unidades_por_caja > 0 and costo_caja > 0:
            costo_calculado = round(costo_caja / unidades_por_caja, 4)
        else:
            costo_calculado = 0.0
        
        costo_unidad = st.number_input(
            "Precio Costo por Unidad (S/.) [Calculado Automáticamente]", 
            min_value=0.0, 
            value=float(costo_calculado), 
            format="%.4f",
            disabled=True,
            help="Se calcula automáticamente dividiendo el costo de la caja entre las unidades totales."
        )
    else:
        costo_unidad = st.number_input("Precio Costo por Unidad (S/.)", min_value=0.0, value=0.0, step=0.10, key=f"c_u_{k}", disabled=disabled_mode)

    if presentacion == "Tableta / Cápsula":
        st.subheader("Configuración de Blísters (Solo para Tabletas)")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            unidades_por_blister = st.number_input("Unidades por Blíster", min_value=1, value=10, step=1, key=f"u_blis_{k}", disabled=disabled_mode)
        with col_b2:
            precio_venta_blister = st.number_input("Precio de Venta por Blíster (S/.)", min_value=0.0, value=0.0, step=0.1, key=f"pv_blis_{k}", disabled=disabled_mode)
    else:
        unidades_por_blister = 1
        precio_venta_blister = 0.0
        
    st.subheader("Precios de Venta al Público")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        precio_venta_unidad = st.number_input("Precio de Venta por Unidad / Frasco (S/.)", min_value=0.0, value=0.0, step=0.1, key=f"pv_u_{k}", disabled=disabled_mode)
    with col_v2:
        precio_venta_caja = st.number_input("Precio de Venta por Caja entera (S/.) (Opcional)", min_value=0.0, value=0.0, step=0.1, key=f"pv_c_{k}", disabled=disabled_mode)
        
    st.subheader("Inventario Inicial")
    col_inv1, col_inv2 = st.columns(2)
    with col_inv1:
        stock_inicial_cajas = st.number_input("Stock Inicial en Cajas / Packs", min_value=0, value=0, step=1, key=f"st_c_{k}", disabled=disabled_mode)
    with col_inv2:
        stock_inicial_unidades_sueltas = st.number_input("Stock Inicial en Unidades Sueltas", min_value=0, value=0, step=1, key=f"st_u_{k}", disabled=disabled_mode)
        
    guardar_producto = st.button("💾 Registrar Producto en el Sistema", type="primary", disabled=disabled_mode)
    
    if "mensaje_exito" in st.session_state:
        st.success(st.session_state.mensaje_exito)
        st.balloons()
        del st.session_state.mensaje_exito

    if guardar_producto and st.session_state.usuario_rol == "admin":
        if not nombre_prod:
            st.error("⚠️ El nombre del producto es obligatorio.")
        elif costo_unidad <= 0:
            st.error("⚠️ El precio costo por unidad debe ser mayor a 0. Ingresa un precio o costo válido.")
        elif precio_venta_unidad <= 0:
            st.error("⚠️ El precio de venta unitario debe ser mayor a 0.")
        else:
            stock_total_unidades = (stock_inicial_cajas * unidades_por_caja) + stock_inicial_unidades_sueltas
            
            query_insert = """
                INSERT INTO productos (
                    nombre, marca, laboratorio, presentacion, 
                    unidades_por_caja, unidades_por_blister, 
                    precio_costo_unidad, precio_venta_unidad, 
                    precio_venta_blister, precio_venta_caja, 
                    stock_actual_unidades, fecha_vencimiento
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            valores = (
                nombre_prod, 
                marca_principio, 
                laboratorio, 
                presentacion, 
                unidades_por_caja, 
                unidades_por_blister, 
                costo_unidad, 
                precio_venta_unidad, 
                precio_venta_blister, 
                precio_venta_caja, 
                stock_total_unidades,
                fecha_venc
            )
            
            db.ejecutar_query(query_insert, valores, commit=True)
            st.cache_data.clear()
            
            st.session_state.mensaje_exito = f"🎉 ¡Producto '{nombre_prod}' registrado con éxito!"
            st.session_state.reset_id += 1
            st.rerun()

# =====================================================================
# PANTALLA 3: REGISTRAR VENTA
# =====================================================================
elif menu_url == "venta":
    st.header("🧾 Registrar Boleta / Salida Multiproducto")
    
    todos_productos = db.cargar_productos_db()
    
    if "reset_boleta_item" not in st.session_state:
        st.session_state.reset_boleta_item = 0

    key_fecha_dinamica = f"input_fecha_venta_v{st.session_state.reset_fecha_version}"
    
    col_fecha, col_info_boleta = st.columns([2, 3])
    with col_fecha:
        fecha_salida = st.date_input(
            "📅 Fecha del Movimiento / Registro:", 
            value=obtener_fecha_hoy(), 
            format="DD/MM/YYYY",
            help="Puedes cambiar esta fecha si vas a ingresar ventas pasadas. Al procesar la boleta, volverá a la fecha de hoy.",
            key=key_fecha_dinamica
        )

    with col_info_boleta:
        st.info(f"📆 Registrando salidas para la fecha: **{fecha_salida.strftime('%d/%m/%Y')}**")

    st.markdown("---")
    
    k_b = st.session_state.reset_boleta_item
    
    col_tit_b, col_btn_reload = st.columns([3, 1])
    with col_tit_b:
        st.subheader("➕ Agregar Producto a la Boleta")
    with col_btn_reload:
        if st.button("🔄", help="actualizar productos", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    
    if not todos_productos:
        st.warning("⚠️ No hay productos disponibles en la base de datos.")
    else:
        opciones_prod = {f"{p['nombre']} ({p['marca']}) - [{p['presentacion']}] - Stock: {p['stock_actual_unidades']} u.": p for p in todos_productos}
        lista_desplegable = [""] + list(opciones_prod.keys())
        
        prod_sel_key = st.selectbox(
            "Buscar Producto (Haz clic y escribe las primeras letras):", 
            lista_desplegable, 
            index=0,
            placeholder="Escribe o selecciona un producto...",
            key=f"prod_sel_boleta_{k_b}"
        )
        
        p_info = opciones_prod[prod_sel_key] if prod_sel_key != "" else None
        
        col_t, col_u, col_b = st.columns([2, 1, 1])
        with col_t:
            tipo_salida_item = st.radio(
                "Tipo de Operación:", 
                ["Venta al Público", "Consumo Interno (Gasto/Uso)"], 
                horizontal=True,
                key=f"tipo_salida_radio_{k_b}"
            )
        with col_u:
            cant_unidades_item = st.number_input(
                "Unidades / Frascos", 
                min_value=0, 
                value=0, 
                step=1,
                key=f"cant_u_{k_b}"
            )
        with col_b:
            if p_info and p_info['presentacion'] == "Tableta / Cápsula":
                cant_blisters_item = st.number_input(
                    "Blísters", 
                    min_value=0, 
                    value=0, 
                    step=1,
                    key=f"cant_b_{k_b}"
                )
            else:
                cant_blisters_item = 0
                st.caption("*(No aplica blíster)*")

        if st.button("➕ Agregar a la Lista / Boleta", type="secondary", use_container_width=True):
            if not p_info:
                st.error("⚠️ Debes seleccionar o buscar un producto antes de agregar.")
            else:
                tot_unid_req = cant_unidades_item + (cant_blisters_item * p_info['unidades_por_blister'])
                tot_ya_en_boleta = sum(item['total_unidades'] for item in st.session_state.boleta_ventas if item['id_producto'] == p_info['id_producto'])
                
                if tot_unid_req <= 0:
                    st.error("⚠️ Debes indicar una cantidad de unidades o blísters mayor a 0.")
                elif (tot_ya_en_boleta + tot_unid_req) > p_info['stock_actual_unidades']:
                    st.error(f"⚠️ Stock insuficiente. Disponible en almacén: {p_info['stock_actual_unidades']} u. (Ya agregados al carrito: {tot_ya_en_boleta} u.)")
                else:
                    precio_u_float = float(p_info['precio_venta_unidad'] or 0.0)
                    precio_b_float = float(p_info['precio_venta_blister'] or 0.0)
                    precio_c_float = float(p_info['precio_costo_unidad'] or 0.0)

                    if tipo_salida_item == "Venta al Público":
                        monto_total_item = (cant_unidades_item * precio_u_float) + (cant_blisters_item * precio_b_float)
                    else:
                        monto_total_item = 0.0
                    
                    costo_total_item = tot_unid_req * precio_c_float
                    
                    st.session_state.boleta_ventas.append({
                        "id_producto": p_info['id_producto'],
                        "nombre": p_info['nombre'],
                        "marca": p_info['marca'],
                        "presentacion": p_info['presentacion'],
                        "tipo_operacion": tipo_salida_item,
                        "unidades": cant_unidades_item,
                        "blisters": cant_blisters_item,
                        "total_unidades": tot_unid_req,
                        "monto_total": monto_total_item,
                        "costo_total": costo_total_item,
                        "precio_unitario_aplicado": precio_u_float if tipo_salida_item == "Venta al Público" else 0.0
                    })
                    
                    st.toast(f"✅ Se agregó {p_info['nombre']} a la lista", icon="🛒")
                    st.session_state.reset_boleta_item += 1
                    st.rerun()

        st.markdown("---")
        st.subheader("🛒 Lista de Productos a Registrar (Boleta Actual)")
        
        if not st.session_state.boleta_ventas:
            st.info("La boleta está vacía. Selecciona un producto arriba y pulsa 'Agregar a la Lista'.")
        else:
            tabla_boleta = []
            total_cobrar_boleta = 0.0
            total_unidades_boleta = 0
            
            for idx, item in enumerate(st.session_state.boleta_ventas):
                total_cobrar_boleta += float(item['monto_total'])
                total_unidades_boleta += item['total_unidades']
                
                cant_str = f"{item['unidades']} u."
                if item['blisters'] > 0:
                    cant_str += f" + {item['blisters']} blís."
                    
                tabla_boleta.append({
                    "N°": idx + 1,
                    "Producto": f"{item['nombre']} ({item['marca']}) - [{item['presentacion']}]",
                    "Tipo de Operación": item['tipo_operacion'],
                    "Cantidad Detalle": cant_str,
                    "Total Unid.": item['total_unidades'],
                    "Subtotal (S/.)": f"S/. {item['monto_total']:.2f}" if item['tipo_operacion'] == "Venta al Público" else "S/. 0.00 (Consumo Interno)"
                })

            st.table(pd.DataFrame(tabla_boleta))
            
            col_tot1, col_tot2 = st.columns(2)
            with col_tot1:
                st.markdown(f"### 🧮 Total de Unidades a Descontar: `{total_unidades_boleta} u.`")
            with col_tot2:
                st.markdown(f"### 💵 Total a Cobrar (Ventas): `S/. {total_cobrar_boleta:,.2f}`")

            st.write("")
            col_b1, col_b2 = st.columns([1, 1])
            
            with col_b1:
                if st.button("🗑️ Vaciar Boleta / Cancelar", use_container_width=True):
                    st.session_state.boleta_ventas = []
                    st.session_state.reset_boleta_item += 1
                    st.rerun()
                    
            with col_b2:
                if st.button("🚀 Confirmar y Procesar Boleta", type="primary", use_container_width=True):
                    for item in st.session_state.boleta_ventas:
                        if item['tipo_operacion'] == "Venta al Público":
                            concepto_tipo = "VENTA"
                            ingreso_neto = item['monto_total'] - item['costo_total']
                        else:
                            concepto_tipo = "CONSUMO"
                            ingreso_neto = -item['costo_total']
                            
                        query_mov = """INSERT INTO movimientos (fecha, tipo_movimiento, id_producto, unidades, blisters, cajas, monto_total, costo_total_capital, ingreso_neto) 
                                       VALUES (%s, %s, %s, %s, %s, 0, %s, %s, %s)"""
                        db.ejecutar_query(
                            query_mov, 
                            (fecha_salida, concepto_tipo, item['id_producto'], item['unidades'], item['blisters'], item['monto_total'], item['costo_total'], ingreso_neto), 
                            commit=True
                        )
                        
                        query_stock = "UPDATE productos SET stock_actual_unidades = stock_actual_unidades - %s WHERE id_producto = %s"
                        db.ejecutar_query(query_stock, (item['total_unidades'], item['id_producto']), commit=True)
                    
                    st.session_state.boleta_ventas = []
                    st.session_state.reset_fecha_version += 1
                    st.session_state.reset_boleta_item += 1
                    st.cache_data.clear()
                    
                    st.success("🎉 ¡Todas las operaciones de la boleta fueron procesadas y la fecha regresó automáticamente a hoy!")
                    st.balloons()
                    st.rerun()

# =====================================================================
# PANTALLA 4: INGRESAR COMPRAS O STOCK
# =====================================================================
elif menu_url == "compra":
    st.header("📥 Registro de Ingreso de Compras")
    
    todos_productos = db.cargar_productos_db()
    
    if st.session_state.usuario_rol != "admin":
        st.warning("🔒 Acceso Restringido: Tu cuenta de Vendedor te permite visualizar pero no registrar o modificar datos de Compras.")
    
    disabled_mode = (st.session_state.usuario_rol != "admin")
    
    if "reset_compra_version" not in st.session_state:
        st.session_state.reset_compra_version = 0

    key_v = st.session_state.reset_compra_version

    if "mensaje_exito_compra" in st.session_state and st.session_state.mensaje_exito_compra:
        st.success(st.session_state.mensaje_exito_compra)
        st.toast("✅ ¡Compra guardada con éxito!", icon="🎉")
        st.session_state.mensaje_exito_compra = None

    if not todos_productos:
        st.warning("⚠️ No hay productos registrados en el sistema para realizar un ingreso de compra.")
    else:
        st.subheader("1. Datos del Comprobante de Compra")
        col_c_fecha, col_c_doc, col_c_num = st.columns(3)
        
        with col_c_fecha:
            fecha_ingreso_compra = st.date_input(
                "📅 Fecha de Compra", 
                value=obtener_fecha_hoy(), 
                format="DD/MM/YYYY",
                key=f"compra_fecha_{key_v}",
                disabled=disabled_mode
            )
        with col_c_doc:
            tipo_doc_compra = st.selectbox(
                "🧾 Tipo de Documento", 
                ["Boleta", "Factura", "Guía / Sin Comprobante"],
                key=f"compra_tipo_doc_{key_v}",
                disabled=disabled_mode
            )
        with col_c_num:
            num_doc_compra = st.text_input(
                "🔢 Número de Documento", 
                placeholder="Ej: B001-0001234",
                key=f"compra_num_doc_{key_v}",
                disabled=disabled_mode
            )
            
        st.markdown("---")
        
        st.subheader("2. Detalle de Productos Comprados")
        
        col_btn_add, col_vacio_comp = st.columns([1, 2])
        with col_btn_add:
            if st.button("➕ Añadir Producto a la Compra", type="secondary", use_container_width=True, disabled=disabled_mode):
                modal_agregar_producto_compra(todos_productos)
                
        if not st.session_state.carrito_compras:
            st.info("La lista de productos está vacía. Presiona el botón **'➕ Añadir Producto a la Compra'** para comenzar a agregar tus ítems.")
        else:
            tabla_compra = []
            total_global_compra = 0.0
            total_unidades_compradas = 0
            
            for idx, item in enumerate(st.session_state.carrito_compras):
                total_global_compra += float(item['costo_total'])
                total_unidades_compradas += item['total_unidades']
                
                det_cant = f"{item['cajas']:.2f}".rstrip('0').rstrip('.') + " Cajas" if item['cajas'] > 0 else ""
                if item['blisters'] > 0:
                    det_cant += f" + {item['blisters']} Blís." if det_cant else f"{item['blisters']} Blís."
                if item['unidades'] > 0:
                    det_cant += f" + {item['unidades']} Unid." if det_cant else f"{item['unidades']} Unid."
                    
                vence_str = item['fecha_vencimiento'].strftime("%d/%m/%Y") if item['fecha_vencimiento'] else "-"
                
                tabla_compra.append({
                    "N°": idx + 1,
                    "Producto": f"{item['nombre']} ({item['marca']}) - [{item['presentacion']}]",
                    "Cantidad Detalle": det_cant,
                    "Total Unid.": item['total_unidades'],
                    "Vencimiento Lote": vence_str,
                    "Costo Total Pagado (S/.)": f"S/. {item['costo_total']:.2f}"
                })
                
            st.table(pd.DataFrame(tabla_compra))
            
            col_tot_c1, col_tot_c2 = st.columns(2)
            with col_tot_c1:
                st.markdown(f"### 📦 Total Unidades Ingresadas: `{total_unidades_compradas} u.`")
            with col_tot_c2:
                st.markdown(f"### 💵 Total General Compra: `S/. {total_global_compra:,.2f}`")

            st.write("")
            col_acc_c1, col_acc_c2 = st.columns(2)
            
            with col_acc_c1:
                if st.button("🗑️ Vaciar Lista / Cancelar Compra", use_container_width=True, disabled=disabled_mode):
                    st.session_state.carrito_compras = []
                    st.rerun()
                    
            with col_acc_c2:
                if st.button("🚀 Registrar Todo e Ingestar al Inventario", type="primary", use_container_width=True, disabled=disabled_mode):
                    detalle_tipo = f"INGRESO ({tipo_doc_compra}: {num_doc_compra if num_doc_compra else 'S/N'})"
                    
                    for item in st.session_state.carrito_compras:
                        costo_u_nuevo = item['costo_total'] / item['total_unidades'] if item['total_unidades'] > 0 else 0.0
                        
                        query_mov = """INSERT INTO movimientos (fecha, tipo_movimiento, id_producto, unidades, blisters, cajas, monto_total, costo_total_capital, ingreso_neto) 
                                       VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 0)"""
                        db.ejecutar_query(
                            query_mov, 
                            (fecha_ingreso_compra, detalle_tipo, item['id_producto'], item['unidades'], item['blisters'], item['cajas'], item['costo_total'], item['costo_total']), 
                            commit=True
                        )
                        
                        query_stock = """UPDATE productos 
                                         SET stock_actual_unidades = stock_actual_unidades + %s,
                                             precio_costo_unidad = CASE WHEN %s > 0 THEN %s ELSE precio_costo_unidad END,
                                             fecha_vencimiento = %s
                                         WHERE id_producto = %s"""
                        db.ejecutar_query(
                            query_stock, 
                            (item['total_unidades'], item['costo_total'], costo_u_nuevo, item['fecha_vencimiento'], item['id_producto']), 
                            commit=True
                        )
                    
                    st.session_state.carrito_compras = []
                    st.session_state.reset_compra_version += 1
                    st.cache_data.clear()
                    st.session_state.mensaje_exito_compra = f"🎉 **¡Compra registrada con éxito!** Se procesó el comprobante **{tipo_doc_compra} - {num_doc_compra if num_doc_compra else 'S/N'}** y el inventario se actualizó."
                    st.rerun()

# =====================================================================
# PANTALLA 5: REPORTES Y CONTABILIDAD
# =====================================================================
elif menu_url == "reportes":
    col_tit_rep, col_btn_mod_rep = st.columns([0.92, 0.08])
    with col_tit_rep:
        st.header("📊 Centro de Reportes y Analítica")
    with col_btn_mod_rep:
        if st.button("✏️", type="primary", help="Corregir o Modificar Ventas", key="btn_mod_ventas_global"):
            modal_editar_ventas()
    
    tipo_reporte_sel = st.selectbox(
        "Selecciona el tipo de reporte a visualizar:",
        [
            "📊 Reporte General y Métricas",
            "🔍 Reporte por Producto",
            "💵 Reporte de Ventas",
            "🛒 Reporte de Compras"
        ]
    )
    
    st.markdown("---")

    # -----------------------------------------------------------------
    # MÓDULO: REPORTE GENERAL (PRODUCTOS MÁS/MENOS VENDIDOS)
    # -----------------------------------------------------------------
    if tipo_reporte_sel == "📊 Reporte General y Métricas":
        st.subheader("📊 Métricas Generales y Movimientos de Inventario")
        
        col_fg1, col_fg2 = st.columns(2)
        with col_fg1:
            filtro_gen_tiempo = st.selectbox("📅 Periodo de Análisis:", ["Este Mes", "Esta Semana", "Hoy", "Todo el Historial"], key="f_gen_tiempo")
        
        f_inicio_g = obtener_fecha_hoy()
        f_fin_g = obtener_fecha_hoy()
        
        if filtro_gen_tiempo == "Hoy":
            f_inicio_g = obtener_fecha_hoy()
        elif filtro_gen_tiempo == "Esta Semana":
            f_inicio_g = obtener_fecha_hoy() - timedelta(days=obtener_fecha_hoy().weekday())
        elif filtro_gen_tiempo == "Este Mes":
            f_inicio_g = obtener_fecha_hoy().replace(day=1)
        elif filtro_gen_tiempo == "Todo el Historial":
            f_inicio_g = datetime(2020, 1, 1).date()
            f_fin_g = obtener_fecha_hoy() + timedelta(days=365)

        # Consulta ranking más vendidos
        query_top = """
            SELECT p.nombre, p.marca, p.presentacion, 
                   SUM((m.unidades) + (m.blisters * p.unidades_por_blister)) AS total_unid_vendidas,
                   SUM(m.monto_total) AS total_recaudado,
                   SUM(m.ingreso_neto) AS ganancia_generada
            FROM movimientos m
            JOIN productos p ON m.id_producto = p.id_producto
            WHERE m.fecha >= %s AND m.fecha <= %s AND m.tipo_movimiento = 'VENTA'
            GROUP BY p.id_producto, p.nombre, p.marca, p.presentacion
            ORDER BY total_unid_vendidas DESC
            LIMIT 10
        """
        top_vendidos = db.ejecutar_query(query_top, (f_inicio_g, f_fin_g), fetch=True)
        
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("### 🔥 Top 10 Productos Más Vendidos")
            if not top_vendidos:
                st.info("No hay registro de ventas en el período seleccionado.")
            else:
                df_top = pd.DataFrame(top_vendidos, columns=["Producto", "Principio Activo", "Presentación", "Unid. Vendidas", "Monto Total (S/.)", "Ganancia (S/.)"])
                st.dataframe(df_top, use_container_width=True)

        with col_g2:
            st.markdown("### 💤 Productos Sin Movimiento / Menos Vendidos")
            query_sin_mov = """
                SELECT p.nombre, p.marca, p.presentacion, p.stock_actual_unidades
                FROM productos p
                WHERE p.id_producto NOT IN (
                    SELECT DISTINCT id_producto 
                    FROM movimientos 
                    WHERE fecha >= %s AND fecha <= %s AND tipo_movimiento = 'VENTA'
                )
                ORDER BY p.nombre ASC
                LIMIT 10
            """
            sin_mov = db.ejecutar_query(query_sin_mov, (f_inicio_g, f_fin_g), fetch=True)
            if not sin_mov:
                st.success("¡Todos tus productos registraron al menos una venta en este rango!")
            else:
                df_sin_mov = pd.DataFrame(sin_mov, columns=["Producto", "Principio Activo", "Presentación", "Stock Físico (Unid.)"])
                st.dataframe(df_sin_mov, use_container_width=True)

    # -----------------------------------------------------------------
    # MÓDULO: REPORTE PERSONALIZADO POR PRODUCTO (KÁRDEX)
    # -----------------------------------------------------------------
    elif tipo_reporte_sel == "🔍 Reporte por Producto":
        st.subheader("🔍 Reporte Detallado y Kárdex por Producto")
        
        todos_prods = db.cargar_productos_db()
        
        if not todos_prods:
            st.info("No hay productos registrados en la base de datos.")
        else:
            dict_prods_rep = {f"{p['nombre']} ({p['marca']}) - [{p['presentacion']}]": p for p in todos_prods}
            
            col_p1, col_p2 = st.columns([3, 2])
            with col_p1:
                prod_sel_rep = st.selectbox(
                    "📦 Selecciona el Producto:",
                    [""] + list(dict_prods_rep.keys()),
                    index=0,
                    placeholder="Haz clic para seleccionar o buscar un producto...",
                    key="sel_prod_reporte_ind"
                )
            with col_p2:
                filtro_tiempo_p = st.selectbox(
                    "📅 Periodo de Fecha:", 
                    ["Este Mes", "Esta Semana", "Hoy", "Todo el Historial", "Rango Personalizado"], 
                    key="f_prod_tiempo"
                )
                
            fecha_inicio_p = obtener_fecha_hoy()
            fecha_fin_p = obtener_fecha_hoy()
            
            if filtro_tiempo_p == "Hoy":
                fecha_inicio_p = obtener_fecha_hoy()
                fecha_fin_p = obtener_fecha_hoy()
            elif filtro_tiempo_p == "Esta Semana":
                fecha_inicio_p = obtener_fecha_hoy() - timedelta(days=obtener_fecha_hoy().weekday())
                fecha_fin_p = obtener_fecha_hoy()
            elif filtro_tiempo_p == "Este Mes":
                fecha_inicio_p = obtener_fecha_hoy().replace(day=1)
                fecha_fin_p = obtener_fecha_hoy()
            elif filtro_tiempo_p == "Todo el Historial":
                fecha_inicio_p = datetime(2020, 1, 1).date()
                fecha_fin_p = obtener_fecha_hoy() + timedelta(days=365)
            elif filtro_tiempo_p == "Rango Personalizado":
                col_dp1, col_dp2 = st.columns(2)
                with col_dp1:
                    fecha_inicio_p = st.date_input("Desde:", obtener_fecha_hoy() - timedelta(days=30), format="DD/MM/YYYY", key="rep_prod_desde")
                with col_dp2:
                    fecha_fin_p = st.date_input("Hasta:", obtener_fecha_hoy(), format="DD/MM/YYYY", key="rep_prod_hasta")

            if prod_sel_rep != "":
                p_objeto = dict_prods_rep[prod_sel_rep]
                id_p_sel = p_objeto['id_producto']
                
                # Consulta adaptada de movimientos compatibles con PostgreSQL / Neon
                query_movs_p = """
                    SELECT m.id_movimiento, m.fecha, m.tipo_movimiento, m.unidades, m.blisters, m.cajas, 
                           m.monto_total, m.costo_total_capital, m.ingreso_neto,
                           p.unidades_por_blister, p.unidades_por_caja
                    FROM movimientos m
                    JOIN productos p ON m.id_producto = p.id_producto
                    WHERE m.id_producto = %s AND m.fecha >= %s AND m.fecha <= %s
                    ORDER BY m.fecha ASC, m.id_movimiento ASC
                """
                movs_p_data = db.ejecutar_query(query_movs_p, (id_p_sel, fecha_inicio_p, fecha_fin_p), fetch=True)
                
                # Ficha del producto encabezado
                st.markdown(f"### 📋 Ficha Actual: `{p_objeto['nombre']}`")
                col_k1, col_k2, col_k3, col_k4 = st.columns(4)
                with col_k1:
                    st.markdown(f"**Principio Activo:** {p_objeto['marca']}")
                    st.markdown(f"**Presentación:** {p_objeto['presentacion']}")
                with col_k2:
                    st.markdown(f"**Laboratorio:** {p_objeto['laboratorio'] or '-'}")
                    venc_s = p_objeto['fecha_vencimiento'].strftime('%d/%m/%Y') if p_objeto['fecha_vencimiento'] else '-'
                    st.markdown(f"**Fecha Vencimiento:** {venc_s}")
                with col_k3:
                    st.markdown(f"**Costo Unitario:** S/. {p_objeto['precio_costo_unidad']:.2f}")
                    st.markdown(f"**Venta Unitario:** S/. {p_objeto['precio_venta_unidad']:.2f}")
                with col_k4:
                    st.markdown(f"📦 **Stock Físico Actual:** `{p_objeto['stock_actual_unidades']} u.`")

                st.markdown("---")

                if not movs_p_data:
                    st.info(f"ℹ️ No se registraron movimientos (ventas o compras) para **{p_objeto['nombre']}** en el período seleccionado.")
                else:
                    filas_kardex = []
                    total_unidades_ingresadas = 0
                    total_unidades_vendidas = 0
                    total_monto_ventas = 0.0
                    total_monto_compras = 0.0
                    total_ganancia_neto = 0.0

                    for m in movs_p_data:
                        f_mov = m[1]
                        if isinstance(f_mov, str):
                            f_mov = datetime.strptime(f_mov, "%Y-%m-%d").date()
                        elif isinstance(f_mov, datetime):
                            f_mov = f_mov.date()

                        tipo_m = str(m[2])
                        u_sueltas = int(m[3] or 0)
                        b_cant = int(m[4] or 0)
                        c_cant = float(m[5] or 0)
                        monto_tot = float(m[6] or 0.0)
                        costo_tot = float(m[7] or 0.0)
                        ing_neto = float(m[8] or 0.0)
                        u_por_b = int(m[9] or 1)
                        u_por_c = int(m[10] or 1)

                        tot_u_mov = u_sueltas + (b_cant * u_por_b) + int(round(c_cant * u_por_c))

                        if tipo_m == "VENTA":
                            entradas_u = 0
                            salidas_u = tot_u_mov
                            total_unidades_vendidas += tot_u_mov
                            total_monto_ventas += monto_tot
                            total_ganancia_neto += ing_neto
                            monto_str = f"S/. {monto_tot:.2f}"
                        elif tipo_m == "CONSUMO":
                            entradas_u = 0
                            salidas_u = tot_u_mov
                            monto_str = "S/. 0.00 (Consumo)"
                        else: # INGRESO / COMPRA
                            entradas_u = tot_u_mov
                            salidas_u = 0
                            total_unidades_ingresadas += tot_u_mov
                            total_monto_compras += costo_tot
                            monto_str = f"S/. {costo_tot:.2f}"

                        filas_kardex.append({
                            "Fecha": f_mov.strftime("%d/%m/%Y"),
                            "Tipo Movimiento": tipo_m,
                            "Entradas (Unid.)": entradas_u,
                            "Salidas (Unid.)": salidas_u,
                            "Monto Total (S/.)": monto_str,
                            "Ganancia Neta (S/.)": f"S/. {ing_neto:.2f}" if tipo_m == "VENTA" else "-"
                        })

                    # Métricas destacadas
                    col_m_p1, col_m_p2, col_m_p3, col_m_p4, col_m_p5 = st.columns(5)
                    with col_m_p1:
                        st.metric("📦 Entradas Totales", f"{total_unidades_ingresadas} u.")
                    with col_m_p2:
                        st.metric("🛒 Salidas / Ventas", f"{total_unidades_vendidas} u.")
                    with col_m_p3:
                        st.metric("💵 Total Recaudado Ventas", f"S/. {total_monto_ventas:,.2f}")
                    with col_m_p4:
                        st.metric("📥 Total Invertido Compras", f"S/. {total_monto_compras:,.2f}")
                    with col_m_p5:
                        st.metric("📈 Ganancia Neta Real", f"S/. {total_ganancia_neto:,.2f}")

                    st.markdown("---")
                    st.subheader(f"📜 Kárdex Histórico de '{p_objeto['nombre']}'")
                    
                    df_kardex = pd.DataFrame(filas_kardex)
                    st.dataframe(df_kardex, use_container_width=True)

                    # Exportación
                    df_kardex_excel = df_kardex.copy()
                    df_kardex_excel.insert(0, "N° Item", range(1, len(df_kardex_excel) + 1))
                    
                    output_kp = io.BytesIO()
                    with pd.ExcelWriter(output_kp, engine='openpyxl') as writer:
                        df_kardex_excel.to_excel(writer, sheet_name="Kardex Producto", index=False, startrow=4)
                        workbook = writer.book
                        worksheet = writer.sheets["Kardex Producto"]
                        
                        worksheet["A1"] = f"REPORTE KÁRDEX DE PRODUCTO: {p_objeto['nombre'].upper()}"
                        worksheet["A2"] = f"Empresa: {st.session_state.nombre_negocio.upper()} | Rango: {fecha_inicio_p.strftime('%d/%m/%Y')} al {fecha_fin_p.strftime('%d/%m/%Y')}"
                        worksheet["A3"] = f"Stock Actual Físico: {p_objeto['stock_actual_unidades']} u. | Generado: {obtener_fecha_hoy().strftime('%d/%m/%Y')} {datetime.now().strftime('%H:%M')}"
                        
                        for col in worksheet.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            col_letter = col[0].column_letter
                            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                            
                    excel_kp_bytes = output_kp.getvalue()
                    pdf_kp_bytes = generar_pdf_bonito(
                        df_kardex_excel,
                        titulo_reporte=f"Kárdex: {p_objeto['nombre']}",
                        subti_reporte=f"Empresa: {st.session_state.nombre_negocio} | Rango: {fecha_inicio_p.strftime('%d/%m/%Y')} al {fecha_fin_p.strftime('%d/%m/%Y')}",
                        es_inventario=False
                    )

                    st.write("")
                    col_kp_d1, col_kp_d2, col_kp_d3 = st.columns([2, 1, 2])
                    with col_kp_d2:
                        with st.popover("📥 Descargar Kárdex", use_container_width=True):
                            st.download_button(
                                label="Excel (.xlsx)",
                                data=excel_kp_bytes,
                                file_name=f"Kardex_{p_objeto['nombre'].replace(' ', '_')}_{fecha_inicio_p}_al_{fecha_fin_p}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            st.download_button(
                                label="PDF (.pdf)",
                                data=pdf_kp_bytes,
                                file_name=f"Kardex_{p_objeto['nombre'].replace(' ', '_')}_{fecha_inicio_p}_al_{fecha_fin_p}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )

    elif tipo_reporte_sel == "💵 Reporte de Ventas":
        st.subheader("💵 Reporte Exclusivo de Ventas y Ganancias")
        
        filtro_tiempo = st.selectbox("Selecciona Periodo:", ["Hoy", "Esta Semana", "Este Mes", "Rango Personalizado"], key="f_ventas_exc")
        fecha_inicio = obtener_fecha_hoy()
        fecha_fin = obtener_fecha_hoy()
        
        if filtro_tiempo == "Hoy":
            fecha_inicio = obtener_fecha_hoy()
        elif filtro_tiempo == "Esta Semana":
            fecha_inicio = obtener_fecha_hoy() - timedelta(days=obtener_fecha_hoy().weekday())
        elif filtro_tiempo == "Este Mes":
            fecha_inicio = obtener_fecha_hoy().replace(day=1)
        elif filtro_tiempo == "Rango Personalizado":
            col1, col2 = st.columns(2)
            with col1:
                fecha_inicio = st.date_input("Desde", obtener_fecha_hoy() - timedelta(days=30), format="DD/MM/YYYY", key="ve_desde")
            with col2:
                fecha_fin = st.date_input("Hasta", obtener_fecha_hoy(), format="DD/MM/YYYY", key="ve_hasta")

        query_v = """
            SELECT m.fecha, p.nombre, p.marca, p.presentacion, 
                   ((m.unidades) + (m.blisters * p.unidades_por_blister)) AS total_unidades,
                   m.monto_total, m.costo_total_capital, m.ingreso_neto
            FROM movimientos m
            JOIN productos p ON m.id_producto = p.id_producto
            WHERE m.fecha >= %s AND m.fecha <= %s AND m.tipo_movimiento = 'VENTA'
            ORDER BY m.fecha DESC, p.nombre ASC
        """
        ventas_data = db.ejecutar_query(query_v, (fecha_inicio, fecha_fin), fetch=True)
        
        if not ventas_data:
            st.info("No se registran ventas en este rango de fechas.")
        else:
            df_ventas = pd.DataFrame(ventas_data, columns=[
                "Fecha", "Producto", "Principio Activo", "Presentación", "Unidades Vendidas", "Venta Total (S/.)", "Costo Capital (S/.)", "Ganancia Neta (S/.)"
            ])
            
            total_v = df_ventas["Venta Total (S/.)"].sum()
            total_c = df_ventas["Costo Capital (S/.)"].sum()
            total_g = df_ventas["Ganancia Neta (S/.)"].sum()
            
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.metric(label="Total Recaudado por Ventas", value=f"S/. {total_v:,.2f}")
            with col_v2:
                st.metric(label="Capital de Costo Recuperado", value=f"S/. {total_c:,.2f}")
            with col_v3:
                st.metric(label="Ganancia Neta Real", value=f"S/. {total_g:,.2f}")
                
            df_v_disp = df_ventas.copy()
            df_v_disp["Fecha"] = pd.to_datetime(df_v_disp["Fecha"]).dt.strftime("%d/%m/%Y")
            
            st.dataframe(df_v_disp, use_container_width=True)
            
            df_v_excel = df_v_disp.copy()
            df_v_excel.insert(0, "N° Item", range(1, len(df_v_excel) + 1))
            
            output_v_e = io.BytesIO()
            with pd.ExcelWriter(output_v_e, engine='openpyxl') as writer:
                df_v_excel.to_excel(writer, sheet_name="Ventas", index=False, startrow=4)
                workbook = writer.book
                worksheet = writer.sheets["Ventas"]
                
                worksheet["A1"] = f"REPORTE DE VENTAS - {st.session_state.nombre_negocio.upper()}"
                worksheet["A2"] = f"Periodo: {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}"
                worksheet["A3"] = f"Generado el: {obtener_fecha_hoy().strftime('%d/%m/%Y')} {datetime.now().strftime('%H:%M')}"
                
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                    
            excel_v_bytes = output_v_e.getvalue()
            pdf_v_bytes = generar_pdf_bonito(
                df_v_excel,
                titulo_reporte="Reporte Exclusivo de Ventas",
                subti_reporte=f"Establecimiento: {st.session_state.nombre_negocio} | Periodo: {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}",
                es_inventario=False
            )
            
            st.write("")
            col_v_d1, col_v_d2, col_v_d3 = st.columns([2, 1, 2])
            with col_v_d2:
                with st.popover("📥 Descargar", use_container_width=True):
                    st.download_button(
                        label="Excel (.xlsx)",
                        data=excel_v_bytes,
                        file_name=f"Reporte_Ventas_{st.session_state.nombre_negocio.replace(' ', '_')}_{fecha_inicio}_al_{fecha_fin}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.download_button(
                        label="PDF (.pdf)",
                        data=pdf_v_bytes,
                        file_name=f"Reporte_Ventas_{st.session_state.nombre_negocio.replace(' ', '_')}_{fecha_inicio}_al_{fecha_fin}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

    elif tipo_reporte_sel == "🛒 Reporte de Compras":
        st.subheader("🛒 Reporte Detallado de Compras e Inversión")
        
        col_c_p1, col_c_p2 = st.columns(2)
        with col_c_p1:
            periodo_compra = st.radio("Filtro de tiempo para compras:", ["Esta Semana", "Este Mes", "Este Año", "Personalizado"], horizontal=True)
        
        hoy = obtener_fecha_hoy()
        if periodo_compra == "Esta Semana":
            f_inicio_c = hoy - timedelta(days=hoy.weekday())
            f_fin_c = hoy
        elif periodo_compra == "Este Mes":
            f_inicio_c = hoy.replace(day=1)
            f_fin_c = hoy
        elif periodo_compra == "Este Año":
            f_inicio_c = hoy.replace(month=1, day=1)
            f_fin_c = hoy
        else:
            with col_c_p2:
                col_c_d1, col_c_d2 = st.columns(2)
                with col_c_d1:
                    f_inicio_c = st.date_input("Desde:", hoy - timedelta(days=30), format="DD/MM/YYYY", key="c_desde")
                with col_c_d2:
                    f_fin_c = st.date_input("Hasta:", hoy, format="DD/MM/YYYY", key="c_hasta")

        query_compras = """
            SELECT m.fecha, m.tipo_movimiento, p.nombre, p.marca, p.presentacion,
                   m.cajas, m.blisters, m.unidades,
                   (m.unidades + (m.blisters * p.unidades_por_blister) + ROUND(m.cajas * p.unidades_por_caja)) AS total_unidades_compradas,
                   m.costo_total_capital
            FROM movimientos m
            JOIN productos p ON m.id_producto = p.id_producto
            WHERE m.fecha >= %s AND m.fecha <= %s AND m.tipo_movimiento LIKE 'INGRESO%%'
            ORDER BY m.fecha DESC, m.id_movimiento DESC
        """
        compras_data = db.ejecutar_query(query_compras, (f_inicio_c, f_fin_c), fetch=True)

        if not compras_data:
            st.info("No se registraron ingresos de mercancía/compras en el rango seleccionado.")
        else:
            filas_c = []
            tot_inversion = 0.0
            tot_unid_c = 0

            for c in compras_data:
                f_c = c[0]
                if isinstance(f_c, str):
                    f_c = datetime.strptime(f_c, "%Y-%m-%d").date()
                elif isinstance(f_c, datetime):
                    f_c = f_c.date()

                doc_info = str(c[1]).replace("INGRESO (", "").replace(")", "")
                n_prod = c[2]
                m_prod = c[3]
                p_prod = c[4]
                c_cajas = float(c[5] or 0)
                c_blis = int(c[6] or 0)
                c_unid = int(c[7] or 0)
                tot_u = int(c[8] or 0)
                costo_tot = float(c[9] or 0.0)

                tot_inversion += costo_tot
                tot_unid_c += tot_u

                det_str = f"{c_cajas:.2f}".rstrip('0').rstrip('.') + " Cajas" if c_cajas > 0 else ""
                if c_blis > 0:
                    det_str += f" + {c_blis} Blís." if det_str else f"{c_blis} Blís."
                if c_unid > 0:
                    det_str += f" + {c_unid} Unid." if det_str else f"{c_unid} Unid."

                filas_c.append({
                    "Fecha": f_c.strftime("%d/%m/%Y"),
                    "Comprobante": doc_info,
                    "Producto": f"{n_prod} ({m_prod})",
                    "Presentación": p_prod,
                    "Detalle Cantidad": det_str if det_str else f"{tot_u} u.",
                    "Unidades Compradas": tot_u,
                    "Inversión Total (S/.)": f"S/. {costo_tot:.2f}"
                })

            col_mc1, col_mc2 = st.columns(2)
            with col_mc1:
                st.metric("📦 Total Unidades Adquiridas", f"{tot_unid_c:,} u.")
            with col_mc2:
                st.metric("💵 Total Inversión en Compras", f"S/. {tot_inversion:,.2f}")

            st.markdown("---")
            df_compras = pd.DataFrame(filas_c)
            st.dataframe(df_compras, use_container_width=True)

            df_c_excel = df_compras.copy()
            df_c_excel.insert(0, "N° Item", range(1, len(df_c_excel) + 1))

            output_c_e = io.BytesIO()
            with pd.ExcelWriter(output_c_e, engine='openpyxl') as writer:
                df_c_excel.to_excel(writer, sheet_name="Compras", index=False, startrow=4)
                workbook = writer.book
                worksheet = writer.sheets["Compras"]

                worksheet["A1"] = f"REPORTE DE COMPRAS E INGRESO DE STOCK - {st.session_state.nombre_negocio.upper()}"
                worksheet["A2"] = f"Periodo: {f_inicio_c.strftime('%d/%m/%Y')} al {f_fin_c.strftime('%d/%m/%Y')}"
                worksheet["A3"] = f"Generado el: {obtener_fecha_hoy().strftime('%d/%m/%Y')} {datetime.now().strftime('%H:%M')}"

                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

            excel_c_bytes = output_c_e.getvalue()
            pdf_c_bytes = generar_pdf_bonito(
                df_c_excel,
                titulo_reporte="Reporte de Compras e Inversión",
                subti_reporte=f"Establecimiento: {st.session_state.nombre_negocio} | Periodo: {f_inicio_c.strftime('%d/%m/%Y')} al {f_fin_c.strftime('%d/%m/%Y')}",
                es_inventario=False
            )

            st.write("")
            col_c_d1, col_c_d2, col_c_d3 = st.columns([2, 1, 2])
            with col_c_d2:
                with st.popover("📥 Descargar Reporte", use_container_width=True):
                    st.download_button(
                        label="Excel (.xlsx)",
                        data=excel_c_bytes,
                        file_name=f"Reporte_Compras_{st.session_state.nombre_negocio.replace(' ', '_')}_{f_inicio_c}_al_{f_fin_c}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.download_button(
                        label="PDF (.pdf)",
                        data=pdf_c_bytes,
                        file_name=f"Reporte_Compras_{st.session_state.nombre_negocio.replace(' ', '_')}_{f_inicio_c}_al_{f_fin_c}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

# =====================================================================
# PANTALLA 6: CONFIGURACIÓN
# =====================================================================
elif menu_url == "config":
    st.header("⚙️ Configuración del Sistema")

    if st.session_state.usuario_rol != "admin":
        st.warning("🔒 Acceso Restringido: Solo administradores pueden realizar cambios en la configuración global.")
    else:
        st.subheader("🎨 Apariencia y Personalización")

        nuevo_nombre = st.text_input("Nombre de la Botica / Farmacia:", value=st.session_state.nombre_negocio)

        temas_opciones = ["Celeste Pastel", "Gris Elegante", "Azul Profesional", "Blanco Puro", "Oscuro Clásico"]
        idx_t = temas_opciones.index(st.session_state.tema_color) if st.session_state.tema_color in temas_opciones else 0
        nuevo_tema = st.selectbox("Tema de Color de Interfaz:", temas_opciones, index=idx_t)

        st.markdown("---")
        st.subheader("🖼️ Logotipo y Fondo")

        col_img1, col_img2 = st.columns(2)

        with col_img1:
            st.markdown("**Logo de la empresa**")
            up_logo = st.file_uploader("Subir nuevo Logo (PNG/JPG):", type=["png", "jpg", "jpeg"], key="up_logo_cfg")
            if up_logo:
                st.session_state.logo_bytes = base64.b64encode(up_logo.read()).decode()
                st.image(base64.b64decode(st.session_state.logo_bytes), width=120)

        with col_img2:
            st.markdown("**Imagen de Fondo (Opcional)**")
            up_fondo = st.file_uploader("Subir nueva Imagen de Fondo:", type=["png", "jpg", "jpeg"], key="up_fondo_cfg")
            if up_fondo:
                st.session_state.fondo_bytes = base64.b64encode(up_fondo.read()).decode()

        if st.button("💾 Guardar Configuración", type="primary"):
            st.session_state.nombre_negocio = nuevo_nombre
            st.session_state.tema_color = nuevo_tema

            query_up_cfg = """
                UPDATE configuracion 
                SET nombre_negocio = %s, tema_color = %s, logo_bytes = %s, fondo_bytes = %s 
                WHERE id = 1
            """
            db.ejecutar_query(query_up_cfg, (nuevo_nombre, nuevo_tema, st.session_state.logo_bytes, st.session_state.fondo_bytes), commit=True)
            st.cache_data.clear()
            st.success("🎉 ¡Configuración actualizada correctamente!")
            st.rerun()

# =====================================================================
# PANTALLA 7: GESTIÓN DE USUARIOS
# =====================================================================
elif menu_url == "usuarios":
    st.header("👥 Gestión de Usuarios y Permisos")

    if st.session_state.usuario_rol != "admin":
        st.error("🔒 Acceso Denegado: Esta pantalla es exclusiva para el administrador.")
    else:
        st.subheader("➕ Crear Nuevo Usuario")

        with st.form("form_nuevo_usr", clear_on_submit=True):
            col_u1, col_u2, col_u3 = st.columns(3)
            with col_u1:
                n_usr = st.text_input("Nombre de Usuario:")
            with col_u2:
                n_pwd = st.text_input("Contraseña:", type="password")
            with col_u3:
                n_rol = st.selectbox("Rol de Acceso:", ["vendedor", "admin"])

            btn_crear_u = st.form_submit_button("💾 Crear Usuario")

            if btn_crear_u:
                if not n_usr or not n_pwd:
                    st.error("⚠️ Debes completar todos los campos.")
                else:
                    db.ejecutar_query(
                        "INSERT INTO usuarios (usuario, password, rol) VALUES (%s, %s, %s)",
                        (n_usr.strip(), n_pwd.strip(), n_rol),
                        commit=True
                    )
                    st.success(f"🎉 ¡Usuario '{n_usr}' creado correctamente con el rol de {n_rol}!")
                    st.rerun()

        st.markdown("---")
        st.subheader("📋 Usuarios Registrados")

        usr_list = db.ejecutar_query("SELECT id_usuario, usuario, rol FROM usuarios ORDER BY id_usuario ASC", fetch=True)

        if usr_list:
            df_usr = pd.DataFrame(usr_list, columns=["ID", "Usuario", "Rol"])
            st.dataframe(df_usr, use_container_width=True)

            st.markdown("### 🗑️ Eliminar Usuario")
            u_dict = {f"{u[1]} ({u[2]})": u[0] for u in usr_list if u[1] != st.session_state.usuario_nombre}

            if u_dict:
                u_del_sel = st.selectbox("Selecciona un usuario para eliminar:", [""] + list(u_dict.keys()))
                if u_del_sel != "":
                    id_u_del = u_dict[u_del_sel]
                    if st.button("❌ Confirmar Eliminación", type="secondary"):
                        db.ejecutar_query("DELETE FROM usuarios WHERE id_usuario = %s", (id_u_del,), commit=True)
                        st.success("¡Usuario eliminado correctamente!")
                        st.rerun()
            else:
                st.caption("*(No hay otros usuarios disponibles para eliminar)*")
