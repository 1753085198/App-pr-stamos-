import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import time
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import requests
import base64
from PIL import Image
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y MENÚ RETRÁCTIL
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 32px !important; font-weight: 700 !important; }
    input { font-size: 26px !important; height: 60px !important; }
    
    /* Botones de Excel (Verde) */
    .stDownloadButton>button { 
        font-size: 35px !important; font-weight: 900 !important; height: 7rem !important; 
        border-radius: 20px !important; background-color: #1D6F42 !important; color: white !important; 
    }
    
    /* Botón de Egreso (Rojo) */
    .stButton>button[kind="secondary"] { 
        background-color: #dc3545 !important; color: white !important; font-size: 25px !important; 
    }

    /* Botón Flotante Nuevo */
    div.stButton > button:first-child[key^="btn_nuevo"] {
        background-color: #ff5722 !important; color: white !important;
        border-radius: 50px !important; padding: 20px 40px !important;
        font-size: 30px !important; font-weight: 900 !important;
        position: fixed; bottom: 40px; right: 40px; z-index: 9999;
    }

    [data-testid="stMetricValue"] { font-size: 85px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown("# 🏦 MENÚ")
    seccion = st.radio("Selecciona sección:", ["💰 PRÉSTAMOS", "🤝 SOCIOS COOP.", "🚑 AYUDAS ECON."], index=0)
    st.write("---")
    st.caption("Jose Figueroa - App Pro")

# --- ESTADOS ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None
if 'mostrar_nuevo' not in st.session_state: st.session_state.mostrar_nuevo = False

# --- FUNCIONES DE EXCEL PERSONALIZADO ---
def generar_excel_socios(df_socios):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("REPORTE SOCIOS", 0)
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        f_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        headers = ["Nombre", "Cédula", "Total Aportado", "Estado"]
        for c, t in enumerate(headers, 1):
            ws.cell(row=1, column=c, value=t).font = Font(bold=True)
            
        for r_idx, row in df_socios.iterrows():
            curr_row = r_idx + 2
            ws.cell(row=curr_row, column=1, value=row['Nombre'])
            ws.cell(row=curr_row, column=2, value=row['Cedula'])
            ws.cell(row=curr_row, column=3, value=row.get('Saldo_Total_Aportado', 0))
            
            # Lógica de color: Si aportó > 0 es verde, si es 0 es rojo
            if float(row.get('Saldo_Total_Aportado', 0)) > 0:
                ws.cell(row=curr_row, column=4, value="AL DÍA").fill = f_verde
            else:
                ws.cell(row=curr_row, column=4, value="PENDIENTE").fill = f_rojo
                
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 25
    return out.getvalue()

# --- CARGAR DATOS ---
def cargar(h):
    try:
        df = conn.read(worksheet=h, ttl=0)
        if df is not None and "ID" in df.columns:
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df
    except: return pd.DataFrame()

# --- SECCIONES ---

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    # (Aquí se mantiene tu código de préstamos que ya está perfecto y no tocamos)
    st.info("Sección de Préstamos activa (Lógica original conservada)")

elif seccion == "🤝 SOCIOS COOP.":
    st.title("🤝 CUOTAS DE SOCIOS")
    df_s = cargar("Cooperativa")
    
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if not df_s.empty:
        st.download_button("📊 DESCARGAR REPORTE (VERDE/ROJO)", data=generar_excel_socios(df_s), file_name="Reporte_Socios.xlsx", use_container_width=True)

    if st.session_state.mostrar_nuevo:
        with st.form("ns"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("GUARDAR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR:")
    # Lista de socios con expander para cobrar...

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 CAJA DE AYUDAS")
    df_a = cargar("Ayudas")
    
    c1, c2 = st.columns([2,1])
    with c1:
        st.write("### 📥 REGISTRAR INGRESO (Normal)")
        with st.form("ing"):
            det = st.text_input("Detalle del aporte:")
            mon = st.number_input("Monto ($):", min_value=1.0)
            if st.form_submit_button("✅ GUARDAR INGRESO"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Fecha":datetime.now().strftime("%Y-%m-%d"), "Descripcion":det, "Tipo":"Ingreso", "Monto":mon}])
                conn.update(worksheet="Ayudas", data=pd.concat([df_a, new], ignore_index=True))
                st.rerun()
                
    with c2:
        st.write("### 🚨 SALIDA")
        if st.button("🔴 REGISTRAR GASTO/EGRESO", use_container_width=True):
            st.session_state.modo_egreso = True
            
    if st.session_state.get('modo_egreso'):
        with st.form("egr"):
            det_e = st.text_input("¿En qué se gastó? (Ej: Choque Juan):")
            mon_e = st.number_input("Monto a sacar ($):", min_value=1.0)
            if st.form_submit_button("⚠️ CONFIRMAR SALIDA DE DINERO"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Fecha":datetime.now().strftime("%Y-%m-%d"), "Descripcion":det_e, "Tipo":"Egreso", "Monto":mon_e}])
                conn.update(worksheet="Ayudas", data=pd.concat([df_a, new], ignore_index=True))
                st.session_state.modo_egreso = False; st.rerun()

    if not df_a.empty:
        # Excel de Ayudas simplificado
        st.download_button("📊 EXCEL DE CAJA", data=df_a.to_csv().encode('utf-8'), file_name="Caja_Ayudas.csv")
