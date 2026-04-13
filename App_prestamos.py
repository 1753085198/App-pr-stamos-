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

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    [data-testid="stSidebar"] { background-color: #1e1e1e !important; width: 400px !important; }
    [data-testid="stSidebar"] .stMarkdown p { font-size: 30px !important; font-weight: 800; color: #ffffff !important; }
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 32px !important; font-weight: 700 !important; }
    input { font-size: 26px !important; height: 60px !important; }
    .stButton>button[kind="primary"] { font-size: 35px !important; font-weight: 900 !important; height: 7rem !important; border-radius: 20px !important; background-color: #28a745 !important; color: white !important; }
    .stDownloadButton>button { font-size: 35px !important; height: 7rem !important; border-radius: 20px !important; background-color: #1D6F42 !important; color: white !important; }
    div.stButton > button:first-child[key^="btn_nuevo_circular"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 35px !important; font-weight: 900 !important; box-shadow: 0px 10px 25px rgba(255, 87, 34, 0.6) !important; border: 4px solid white !important; position: fixed; bottom: 40px; right: 40px; z-index: 9999; }
    [data-testid="stMetricValue"] { font-size: 85px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown("## 🏦 GESTIÓN")
    seccion = st.radio("Ir a:", ["💰 PRÉSTAMOS", "🤝 SOCIOS COOP.", "🚑 AYUDAS ECON."], index=0)
    st.write("---")
    st.info("UDLA - Economía 2026")

# --- ESTADOS ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None
if 'mostrar_nuevo' not in st.session_state: st.session_state.mostrar_nuevo = False

# --- FUNCIONES CORE ---
def cargar_datos(hoja):
    try:
        df = conn.read(worksheet=hoja, ttl=0)
        if df is not None:
            for c in ["ID", "Cedula"]:
                if c in df.columns: df[c] = df[c].astype(str).str.replace(".0", "", regex=False)
        return df
    except:
        st.warning(f"⚠️ La hoja '{hoja}' no existe en tu Google Sheets. Por favor, créala.")
        return pd.DataFrame()

def subir_img(archivo):
    try:
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": st.secrets["IMGBB_API_KEY"], "image": base64.b64encode(archivo).decode('utf-8')})
        return res.json()["data"]["url"]
    except: return ""

# --- SECCIONES ---
if seccion == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    df_p = cargar_datos("Prestamos")
    df_h = cargar_datos("Pagos")
    
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo_circular"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if st.session_state.mostrar_nuevo:
        with st.form("n_p"):
            nm, cd, ml = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Correo:")
            m, t, p = st.number_input("Monto:"), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("💾 GUARDAR"):
                i = (t/100)/12
                cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                nuevo = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":nm, "Cedula":cd, "Email":ml, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO", "Tasa":t, "Fecha":datetime.now().strftime("%Y-%m-%d")}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR:")
    act = df_p[df_p["Estado"]=="ACTIVO"] if not df_p.empty else pd.DataFrame()
    if bq and not act.empty: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}", expanded=(st.session_state.id_abierto == row['ID'])):
            c1, c2 = st.columns(2)
            with c1:
                st.metric("CUOTA", f"${row['Cuota_Mensual']}")
                st.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
            with c2:
                with st.form(key=f"f_{row['ID']}_{st.session_state.pago_key}"):
                    ft = st.file_uploader("📸 RECIBO:", key=f"foto_{row['ID']}_{st.session_state.pago_key}")
                    if st.form_submit_button("✅ CONFIRMAR"):
                        if ft:
                            st.session_state.id_abierto = row['ID']
                            url = subir_img(ft.getvalue())
                            new_p = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Monto_Pagado": row['Cuota_Mensual'], "URL_Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_p], ignore_index=True))
                            row["Pagos_Realizados"] += 1; row["Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                            if row["Pagos_Realizados"] >= row["Meses_Totales"]: row["Estado"] = "PAGADO"
                            df_p.loc[idx] = row; conn.update(worksheet="Prestamos", data=df_p)
                            st.session_state.pago_key += 1; st.rerun()

elif seccion == "🤝 SOCIOS COOP.":
    st.title("🤝 CUOTAS DE SOCIOS")
    df_s = cargar_datos("Cooperativa")
    df_hc = cargar_datos("Pagos_Coop")
    
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo_circular"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if st.session_state.mostrar_nuevo:
        with st.form("n_s"):
            nm, cd = st.text_input("Nombre Socio:"), st.text_input("Cédula:")
            if st.form_submit_button("💾 REGISTRAR"):
                nuevo = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":nm, "Cedula":cd, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, nuevo], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR SOCIO:")
    act = df_s if not df_s.empty else pd.DataFrame()
    if bq and not act.empty: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row.get('Saldo_Total_Aportado',0)}"):
            monto = st.number_input("Valor hoy:", min_value=1.0, key=f"v_{row['ID']}")
            ft = st.file_uploader("📸 Recibo:", key=f"f_{row['ID']}_{st.session_state.pago_key}")
            if st.button("✅ GUARDAR APORTE", key=f"b_{row['ID']}"):
                if ft:
                    url = subir_img(ft.getvalue())
                    new_hc = pd.DataFrame([{"ID_Socio":row['ID'], "Monto":monto, "Fecha":datetime.now().strftime("%Y-%m-%d"), "URL":url}])
                    conn.update(worksheet="Pagos_Coop", data=pd.concat([df_hc, new_hc], ignore_index=True))
                    df_s.at[idx, "Saldo_Total_Aportado"] = float(row.get('Saldo_Total_Aportado',0)) + monto
                    conn.update(worksheet="Cooperativa", data=df_s)
                    st.success("Aporte Guardado!"); time.sleep(1); st.rerun()

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 CAJA DE AYUDAS")
    df_a = cargar_datos("Ayudas")
    if not df_a.empty:
        ing = df_a[df_a['Tipo']=='Ingreso']['Monto'].sum()
        egr = df_a[df_a['Tipo']=='Egreso']['Monto'].sum()
        st.metric("💰 FONDO TOTAL", f"${round(ing-egr, 2)}")
    
    with st.form("ay"):
        desc, tipo, monto = st.text_input("Detalle:"), st.selectbox("Tipo:", ["Ingreso", "Egreso"]), st.number_input("Monto:")
        if st.form_submit_button("🚀 REGISTRAR"):
            nuevo = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Fecha":datetime.now().strftime("%Y-%m-%d"), "Descripcion":desc, "Tipo":tipo, "Monto":monto}])
            conn.update(worksheet="Ayudas", data=pd.concat([df_a, nuevo], ignore_index=True))
            st.rerun()
