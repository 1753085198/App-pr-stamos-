import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import time
import io
from datetime import datetime
import requests
import base64
from PIL import Image
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="SISTEMA DE PRÉSTAMOS PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ DE ALTO IMPACTO (CARDS Y BOTONES GIGANTES)
st.markdown("""
    <style>
    /* Estilo de Pestañas */
    button[data-baseweb="tab"] { font-size: 26px !important; font-weight: 800 !important; }
    
    /* Fuentes y Etiquetas */
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { 
        font-size: 22px !important; font-weight: 600 !important; 
    }
    
    /* BOTONES DE ACCIÓN GIGANTES */
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 28px !important; font-weight: 900 !important; height: 5.5rem !important; 
        border-radius: 20px !important; background-color: #28a745 !important; color: white !important; 
        box-shadow: 0px 5px 15px rgba(0,0,0,0.3) !important;
    }

    /* Diseño de Tarjetas (Cards) */
    .client-card {
        background-color: #ffffff;
        border-left: 10px solid #007bff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
        margin-bottom: 15px;
    }
    
    [data-testid="stMetricValue"] { font-size: 60px !important; font-weight: 800 !important; }
    </style>
""", unsafe_allow_html=True)

# 3. CONEXIÓN
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_prestamos():
    try:
        df = conn.read(worksheet="Prestamos", ttl="0") 
        if not df.empty:
            df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df
    except: return pd.DataFrame(columns=["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"])

def cargar_historial_pagos():
    try:
        df = conn.read(worksheet="Pagos", ttl="0")
        if not df.empty: df["ID_Prestamo"] = df["ID_Prestamo"].astype(str)
        return df
    except: return pd.DataFrame(columns=["ID_Prestamo", "Fecha_Pago", "Cuotas_Pagadas", "Monto_Pagado", "URL_Comprobante"])

def subir_a_imgbb_comprimido(archivo_bytes):
    try:
        imagen = Image.open(io.BytesIO(archivo_bytes))
        if imagen.width > 800:
            imagen = imagen.resize((800, int(imagen.height * (800 / float(imagen.width)))), Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        imagen.convert('RGB').save(buffer, format="JPEG", quality=70)
        img_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
        api_key = st.secrets["IMGBB_API_KEY"]
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": api_key, "image": img_data})
        return res.json()["data"]["url"] if res.status_code == 200 else ""
    except: return ""

def generar_excel_estado_cuenta(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ws_name = "ESTADO DE CUENTA"
        pd.DataFrame().to_excel(writer, sheet_name=ws_name)
        workbook = writer.book; ws = workbook[ws_name]
        f_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ws["A1"] = "ESTADO DE CUENTA PROFESIONAL"; ws["A1"].font = Font(bold=True, size=14)
        ws["A3"] = "CLIENTE:"; ws["B3"] = datos_c['Nombre'].upper()
        ws["D3"] = "SALDO:"; ws["E3"] = f"${datos_c['Saldo_Restante']}"
        h = ["Cuota #", "Descripción", "Monto", "Estado"]
        for c_idx, text in enumerate(h, 1):
            cell = ws.cell(row=6, column=c_idx, value=text)
            cell.fill = f_azul; cell.font = Font(color="FFFFFF", bold=True)
        pag = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            r = 6 + i
            ws.cell(row=r, column=1, value=i); ws.cell(row=r, column=2, value=f"Cuota {i}")
            ws.cell(row=r, column=3, value=datos_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws.cell(row=r, column=4, value=est)
            if i <= pag:
                for c in range(1, 5): ws.cell(row=r, column=c).fill = f_verde
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20
    return output.getvalue()

# --- INTERFAZ ---
st.title("🏦 GESTIÓN FINANCIERA")
df_p = cargar_prestamos(); df_h = cargar_historial_pagos()

t_lista, t_pago, t_reporte, t_nuevo = st.tabs(["📋 LISTA", "💵 PAGO", "📊 REPORTE", "➕ NUEVO"])

# TAB 1: LISTA (Con Ayudas Visuales y Cards)
with t_lista:
    st.write("### 👥 Clientes Activos")
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if activos.empty:
        st.warning("No hay préstamos activos actualmente.")
    else:
        for index, row in activos.iterrows():
            with st.container():
                # Simulamos una CARD visual
                st.markdown(f"""
                <div style="background-color:#f8f9fa; padding:20px; border-radius:15px; border-left: 10px solid #28a745; margin-bottom:10px;">
                    <span style="font-size:24px; font-weight:bold; color:#1f4e78;">👤 {row['Nombre'].upper()}</span><br>
                    <span style="font-size:18px; color:#555;">🪪 Cédula: {row['Cedula']} | 📅 Inicio: {row['Fecha']}</span><br>
                    <span style="font-size:22px; font-weight:bold; color:#28a745;">💰 Saldo Pendiente: ${row['Saldo_Restante']}</span>
                </div>
                """, unsafe_allow_html=True)

# TAB 2: PAGO (Fácil y Rápido)
with t_pago:
    act = df_p[df_p["Estado"] == "ACTIVO"]
    if not act.empty:
        if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
        with st.form(key=f"form_p_{st.session_state.pago_key}", clear_on_submit=True):
            st.write("### 💸 Registrar Abono")
            sel = st.selectbox("Seleccione Cliente:", act["Nombre"].astype(str) + " (" + act["ID"].astype(str) + ")")
            idp = sel.split("(")[1].replace(")", ""); cl = act[act["ID"] == idp].iloc[0]
            
            c1, c2 = st.columns(2)
            n_c = c1.number_input("Número de cuotas a pagar:", min_value=1, value=1)
            total_pagar = round(cl['Cuota_Mensual'] * n_c, 2)
            st.success(f"## TOTAL A COBRAR: ${total_pagar}")
            fot = c2.file_uploader("📸 Foto del Comprobante:", type=["jpg","png","jpeg"])
            
            if st.form_submit_button("✅ PROCESAR PAGO", use_container_width=True, type="primary"):
                with st.spinner('Guardando...'):
                    link = subir_a_imgbb_comprimido(fot.getvalue()) if fot else ""
                    np = pd.DataFrame([{"ID_Prestamo": idp, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": n_c, "Monto_Pagado": total_pagar, "URL_Comprobante": link}])
                    conn.update(worksheet="Pagos", data=pd.concat([df_h, np], ignore_index=True))
                    
                    idx = df_p[df_p["ID"] == idp].index[0]
                    df_p.at[idx, "Pagos_Realizados"] += n_c
                    # Cálculo proporcional del saldo restante
                    df_p.at[idx, "Saldo_Restante"] = round(max(0, cl["Saldo_Restante"] - (cl["Monto_Inicial"]/cl["Meses_Totales"])*n_c), 2)
                    if df_p.at[idx, "Pagos_Realizados"] >= cl["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                    conn.update(worksheet="Prestamos", data=df_p)
                    
                    st.session_state.pago_key += 1
                    st.balloons(); time.sleep(1); st.rerun()

# TAB 3: REPORTE (Detallado)
with t_reporte:
    if not df_p.empty:
        st.write("### 🔍 Consulta de Estado")
        opc = df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")"
        busq = st.selectbox("Buscar por nombre:", opc)
        ids = busq.split("(")[1].replace(")", ""); d = df_p[df_p["ID"] == ids].iloc[0]; h = df_h[df_h["ID_Prestamo"] == ids]
        
        col1, col2, col3 = st.columns(3)
        col1.metric("DEUDA", f"${d['Saldo_Restante']}")
        col2.metric("PAGADO", f"{d['Pagos_Realizados']}/{d['Meses_Totales']}")
        col3.metric("CUOTA", f"${d['Cuota_Mensual']}")
        
        st.download_button(f"📥 GENERAR EXCEL: {d['Nombre']}", data=generar_excel_estado_cuenta(d, h), file_name=f"Estado_{d['Nombre']}.xlsx", use_container_width=True)
        
        with st.expander("📸 Ver fotos de recibos previos"):
            st.dataframe(h, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("Ver Foto")})

        if st.button("Eliminar Registro Permanentemente", type="secondary"):
            conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != ids])
            conn.update(worksheet="Pagos", data=df_h[df_h["ID_Prestamo"] != ids])
            st.rerun()

# TAB 4: NUEVO (Configuración Inicial)
with t_nuevo:
    with st.form("form_nuevo", clear_on_submit=True):
        st.write("### 📝 Datos del Nuevo Crédito")
        n = st.text_input("Nombre Completo:"); c_id = st.text_input("Número de Cédula:")
        col1, col2, col3 = st.columns(3)
        m = col1.number_input("Monto Prestado ($)", value=500.0); t = col3.number_input("Tasa Interés (%)", value=15.0); p = col2.number_input("Plazo (Meses)", value=12)
        
        if st.form_submit_button("💾 CREAR PRÉSTAMO", use_container_width=True, type="primary"):
            tm = (t/100)/12
            cuota = m * (tm * (1+tm)**p) / ((1+tm)**p - 1) if tm > 0 else m/p
            nuevo = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": n, "Cedula": c_id, "Monto_Inicial": m, "Saldo_Restante": m, "Cuota_Mensual": round(cuota, 2), "Meses_Totales": int(p), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": t}])
            conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
            st.balloons(); time.sleep(1); st.rerun()
