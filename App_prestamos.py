import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import requests
import base64
from openpyxl.styles import PatternFill, Font

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 28px !important; font-weight: 900 !important; position: fixed; bottom: 30px; right: 30px; z-index: 9999; border: 3px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 65px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNCIONES CORE ---
def cargar(h):
    try:
        df = conn.read(worksheet=h, ttl=0)
        if df is not None and "ID" in df.columns:
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df if df is not None else pd.DataFrame()
    except: return pd.DataFrame()

def subir_img(archivo):
    try:
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": st.secrets["IMGBB_API_KEY"], "image": base64.b64encode(archivo).decode('utf-8')})
        return res.json()["data"]["url"]
    except: return ""

def generar_excel_personal(socio_row, df_historial, tipo_label):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("HISTORIAL", 0)
        ws.append([f"COMPROBANTE DETALLADO - {tipo_label}"])
        ws.append([f"NOMBRE: {socio_row['Nombre']}"]); ws.append([f"CÉDULA: {socio_row['Cedula']}"])
        ws.append([""]); ws.append(["Fecha", "Monto", "Link Recibo"])
        pagos = df_historial[df_historial['ID_Socio'] == socio_row['ID']] if not df_historial.empty else pd.DataFrame()
        for _, p in pagos.iterrows(): ws.append([p['Fecha'], p['Monto'], p.get('Comprobante', 'N/A')])
        ws.append([""]); ws.append(["TOTAL ACUMULADO:", socio_row.get('Saldo_Total_Aportado', socio_row.get('Saldo_Restante', 0))])
    return out.getvalue()

def enviar_mail(dest, nom, exc, url, tipo):
    try:
        msg = MIMEMultipart(); msg['From'] = st.secrets["EMAIL_USER"]; msg['To'] = dest; msg['Subject'] = f"✅ Recibo de Pago - {tipo} - {nom}"
        msg.attach(MIMEText(f"Hola {nom}, se ha registrado tu pago con éxito.\n\nVer Comprobante: {url}", 'plain'))
        p = MIMEBase('application', 'octet-stream'); p.set_payload(exc); encoders.encode_base64(p)
        p.add_header('Content-Disposition', f"attachment; filename=Recibo_{tipo}.xlsx"); msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"]); s.send_message(msg); s.quit()
    except: pass

# --- SECCIONES ---
with st.sidebar:
    st.markdown("# 🏦 PANEL")
    seccion = st.radio("Ir a:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    # (Lógica original de préstamos conservada...)

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    v_x = st.number_input("💵 VALOR CUOTA (X):", value=10.0)
    
    for idx, row in df_s.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | ACUMULADO: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fc_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_x)
                    ft = st.file_uploader("📸 RECIBO:", key=f"ic_{row['ID']}")
                    if st.form_submit_button("✅ CONFIRMAR PAGO"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Coop", data=pd.concat([df_ph, new_h], ignore_index=True))
                            df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Cooperativa", data=df_s)
                            # ENVÍO DE CORREO
                            excel_p = generar_excel_personal(df_s.loc[idx], pd.concat([df_ph, new_h]), "COOPERATIVA")
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], excel_p, url, "Cooperativa")
                            st.success("¡Pago y Correo Enviados!"); time.sleep(1); st.rerun()

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 APORTE ESTÁNDAR (X):", value=5.0)
    
    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo_a"): st.session_state.n_a = True
    if st.session_state.get('n_a'):
        with st.form("na"):
            n, c, e = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Email:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Email":e, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True)); st.session_state.n_a = False; st.rerun()

    for idx, row in df_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fa_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_y)
                    ft = st.file_uploader("📸 RECIBO:", key=f"ia_{row['ID']}")
                    if st.form_submit_button("✅ REGISTRAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_ah, new_h], ignore_index=True))
                            df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Ayudas_Listado", data=df_a)
                            # ENVÍO DE CORREO
                            excel_ayuda = generar_excel_personal(df_a.loc[idx], pd.concat([df_ah, new_h]), "AYUDAS")
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], excel_ayuda, url, "Ayudas")
                            st.success("¡Aporte y Correo Enviados!"); time.sleep(1); st.rerun()
