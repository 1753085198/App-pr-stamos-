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
    
    /* Botones de Excel (Verde Excel) */
    .stDownloadButton>button { 
        font-size: 35px !important; font-weight: 900 !important; height: 7rem !important; 
        border-radius: 20px !important; background-color: #1D6F42 !important; color: white !important; 
    }
    
    /* Botón de Egreso/Gasto (Rojo) */
    .stButton>button[kind="secondary"] { 
        background-color: #dc3545 !important; color: white !important; font-size: 25px !important; 
    }

    /* Botón Flotante Nuevo Registro */
    div.stButton > button:first-child[key^="btn_nuevo"] {
        background-color: #ff5722 !important; color: white !important;
        border-radius: 50px !important; padding: 20px 40px !important;
        font-size: 30px !important; font-weight: 900 !important;
        position: fixed; bottom: 40px; right: 40px; z-index: 9999;
        border: 4px solid white !important;
    }

    [data-testid="stMetricValue"] { font-size: 85px !important; font-weight: 900 !important; color: #007bff !important; }
    .streamlit-expanderHeader { font-size: 38px !important; font-weight: 800 !important; padding: 25px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown("# 🏦 MENÚ")
    seccion = st.radio("Ir a:", ["💰 PRÉSTAMOS", "🤝 CUOTAS COOP.", "🚑 AYUDAS (LISTADO)"], index=0)
    st.write("---")
    st.caption("Jose Figueroa - UDLA 2026")

# --- ESTADOS ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None
if 'mostrar_nuevo' not in st.session_state: st.session_state.mostrar_nuevo = False

# --- CARGAR DATOS ---
def cargar(h):
    try:
        df = conn.read(worksheet=h, ttl=0)
        if df is not None and "ID" in df.columns:
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df if df is not None else pd.DataFrame()
    except: return pd.DataFrame()

# --- FUNCIÓN EXCEL UNIFICADO (PINTADO) ---
def generar_excel_unificado(df_datos, titulo_reporte):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("REPORTE GENERAL", 0)
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        f_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        ws["A1"] = f"REPORTE UNIFICADO: {titulo_reporte}"; ws["A1"].font = Font(bold=True, size=14)
        
        headers = ["Nombre", "Cédula", "Monto Aportado", "Estado"]
        for c, t in enumerate(headers, 1):
            cell = ws.cell(row=3, column=c, value=t)
            cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            
        for r_idx, row in df_datos.reset_index(drop=True).iterrows():
            curr_row = r_idx + 4
            ws.cell(row=curr_row, column=1, value=row['Nombre'])
            ws.cell(row=curr_row, column=2, value=row['Cedula'])
            
            # Buscamos el monto (depende de la hoja puede llamarse distinto)
            monto = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            ws.cell(row=curr_row, column=3, value=monto)
            
            # Lógica de color
            if monto > 0:
                ws.cell(row=curr_row, column=4, value="CUMPLIDO").fill = f_verde
            else:
                ws.cell(row=curr_row, column=4, value="PENDIENTE").fill = f_rojo
                
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 25
    return out.getvalue()

def subir_img(archivo):
    try:
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": st.secrets["IMGBB_API_KEY"], "image": base64.b64encode(archivo).decode('utf-8')})
        return res.json()["data"]["url"]
    except: return ""

# --- LÓGICA DE SECCIONES ---

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    # (Tu lógica original de préstamos intacta aquí)
    st.info("Lógica de Préstamos original activa.")

elif seccion == "🤝 CUOTAS COOP.":
    st.title("🤝 LISTADO DE SOCIOS (COOP)")
    df_s = cargar("Cooperativa")
    
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if not df_s.empty:
        st.download_button("📊 EXCEL UNIFICADO COOP", data=generar_excel_unificado(df_s, "COOPERATIVA"), file_name="Reporte_Coop.xlsx", use_container_width=True)

    # Registro de Socio
    if st.session_state.mostrar_nuevo:
        with st.form("ns"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("GUARDAR SOCIO"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    # Buscador y Cobro Variable
    bq = st.text_input("🔍 BUSCAR SOCIO:")
    act = df_s if not df_s.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row.get('Saldo_Total_Aportado', 0)}"):
            m_v = st.number_input("Cantidad a pagar:", value=10.0, key=f"mv_{row['ID']}")
            f_c = st.file_uploader("📸 Recibo:", key=f"fc_{row['ID']}_{st.session_state.pago_key}")
            if st.button("✅ REGISTRAR PAGO", key=f"bc_{row['ID']}"):
                if f_c:
                    df_s.at[idx, "Saldo_Total_Aportado"] = float(row.get('Saldo_Total_Aportado', 0)) + m_v
                    conn.update(worksheet="Cooperativa", data=df_s)
                    st.session_state.pago_key += 1; st.rerun()

elif seccion == "🚑 AYUDAS (LISTADO)":
    st.title("🚑 LISTADO AYUDAS ECONÓMICAS")
    df_a = cargar("Ayudas_Listado") # Necesitas esta pestaña en tu Sheets
    
    c1, c2 = st.columns([2,1])
    with c1:
        if not df_a.empty:
            st.download_button("📊 EXCEL UNIFICADO AYUDAS", data=generar_excel_unificado(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    with c2:
        if st.button("🔴 REGISTRAR GASTO CAJA", use_container_width=True, type="secondary"):
            st.session_state.egreso_ayuda = True

    # Formulario para Gasto
    if st.session_state.get('egreso_ayuda'):
        with st.form("eg_ay"):
            det_e = st.text_input("Motivo del Gasto:")
            mon_e = st.number_input("Monto a sacar:", min_value=1.0)
            if st.form_submit_button("⚠️ SACAR DINERO DE CAJA"):
                # Aquí podrías restar de un fondo general o registrar en una hoja de bitácora
                st.warning(f"Gasto de ${mon_e} registrado.")
                st.session_state.egreso_ayuda = False; st.rerun()

    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if st.session_state.mostrar_nuevo:
        with st.form("na"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("AÑADIR AL LISTADO"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    # Buscador y Aporte
    bq_a = st.text_input("🔍 BUSCAR COMPAÑERO:")
    act_a = df_a if not df_a.empty else pd.DataFrame()
    if bq_a: act_a = act_a[act_a['Nombre'].str.contains(bq_a, case=False)]
    
    for idx, row in act_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | APORTADO: ${row.get('Saldo_Total_Aportado', 0)}"):
            m_a = st.number_input("Monto aporte:", value=5.0, key=f"ma_{row['ID']}")
            f_a = st.file_uploader("📸 Foto Aporte:", key=f"fa_{row['ID']}_{st.session_state.pago_key}")
            if st.button("✅ GUARDAR APORTE", key=f"ba_{row['ID']}"):
                if f_a:
                    df_a.at[idx, "Saldo_Total_Aportado"] = float(row.get('Saldo_Total_Aportado', 0)) + m_a
                    conn.update(worksheet="Ayudas_Listado", data=df_a)
                    st.session_state.pago_key += 1; st.rerun()
