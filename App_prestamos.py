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
from PIL import Image
from openpyxl.styles import PatternFill, Font

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 24px !important; font-weight: 700 !important; border-radius: 12px !important; }
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

def generar_excel_grupal(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("GENERAL", 0)
        f_v, f_r = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE GENERAL - {titulo}"]); ws.append(["Nombre", "Cédula", "Monto Total", "Estado"])
        for _, row in df.iterrows():
            m = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            ws.append([row['Nombre'], row['Cedula'], m, "AL DÍA" if m > 0 else "PENDIENTE"])
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = f_v if m > 0 else f_r
    return out.getvalue()

def generar_excel_personal(row, historial, tipo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("HISTORIAL", 0)
        ws.append([f"COMPROBANTE DE {tipo}"]); ws.append([f"NOMBRE: {row['Nombre']}"]); ws.append([f"CÉDULA: {row['Cedula']}"])
        ws.append([""]); ws.append(["Fecha", "Monto", "Recibo"])
        pagos = historial[historial['ID_Socio'] == row['ID']] if not historial.empty else pd.DataFrame()
        for _, p in pagos.iterrows(): ws.append([p['Fecha'], p['Monto'], p.get('Comprobante', 'N/A')])
        ws.append([""]); ws.append(["TOTAL ACUMULADO:", row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0))])
    return out.getvalue()

def enviar_mail(dest, nom, exc, url, tipo):
    try:
        msg = MIMEMultipart(); msg['From'] = st.secrets["EMAIL_USER"]; msg['To'] = dest; msg['Subject'] = f"✅ Recibo {tipo} - {nom}"
        msg.attach(MIMEText(f"Hola {nom}, se registró tu pago.\n\nRecibo: {url}", 'plain'))
        p = MIMEBase('application', 'octet-stream'); p.set_payload(exc); encoders.encode_base64(p)
        p.add_header('Content-Disposition', f"attachment; filename=Recibo_{tipo}.xlsx"); msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"]); s.send_message(msg); s.quit()
    except: pass

# --- MENÚ ---
with st.sidebar:
    st.markdown("# 🏦 SISTEMA")
    sec = st.radio("Ir a:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)

# --- ESTADOS ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0

# --- LÓGICA PRÉSTAMOS ---
if sec == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    if not df_p.empty: st.download_button("📊 EXCEL GRUPAL PRÉSTAMOS", data=generar_excel_grupal(df_p, "PRESTAMOS"), file_name="Reporte_Prestamos.xlsx", use_container_width=True)
    
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo_p"): st.session_state.n_p = True
    if st.session_state.get('n_p'):
        with st.form("np"):
            n, c, e = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Email:")
            m, t, p = st.number_input("Monto:"), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("GUARDAR"):
                i = (t/100)/12; cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":n, "Cedula":c, "Email":e, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO"}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True)); st.rerun()

    for idx, row in df_p[df_p["Estado"]=="ACTIVO"].iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}"):
            c1, c2 = st.columns(2)
            with c1: st.metric("CUOTA", f"${row['Cuota_Mensual']}"); st.metric("PROGRESO", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
            with c2:
                with st.form(key=f"fp_{row['ID']}"):
                    ft = st.file_uploader("📸 RECIBO:", key=f"ip_{row['ID']}")
                    if st.form_submit_button("✅ COBRAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": row['Cuota_Mensual'], "Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_h], ignore_index=True))
                            df_p.at[idx, "Pagos_Realizados"] += 1; df_p.at[idx, "Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                            if df_p.at[idx, "Pagos_Realizados"] >= row["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                            conn.update(worksheet="Prestamos", data=df_p)
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], generar_excel_personal(df_p.loc[idx], pd.concat([df_h, new_h]), "PRESTAMO"), url, "Prestamos")
                            st.rerun()

# --- LÓGICA COOPERATIVA ---
elif sec == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    v_x = st.number_input("💵 VALOR CUOTA ESTÁNDAR:", value=10.0)
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL COOP", data=generar_excel_grupal(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    
    for idx, row in df_s.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fc_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_x); ft = st.file_uploader("📸 RECIBO:", key=f"ic_{row['ID']}")
                    if st.form_submit_button("✅ CONFIRMAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Coop", data=pd.concat([df_ph, new_h], ignore_index=True))
                            df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Cooperativa", data=df_s)
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], generar_excel_personal(df_s.loc[idx], pd.concat([df_ph, new_h]), "COOPERATIVA"), url, "Cooperativa")
                            st.rerun()
            with c2: st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_personal(row, df_ph, "COOP"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dlc_{row['ID']}")

# --- LÓGICA AYUDAS ---
elif sec == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 APORTE ESTÁNDAR:", value=5.0)
    if not df_a.empty: st.download_button("📊 EXCEL GRUPAL AYUDAS", data=generar_excel_grupal(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    
    for idx, row in df_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fa_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_y); ft = st.file_uploader("📸 RECIBO:", key=f"ia_{row['ID']}")
                    if st.form_submit_button("✅ REGISTRAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_ah, new_h], ignore_index=True))
                            df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Ayudas_Listado", data=df_a)
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], generar_excel_personal(df_a.loc[idx], pd.concat([df_ah, new_h]), "AYUDAS"), url, "Ayudas")
                            st.rerun()
            with c2: st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_personal(row, df_ah, "AYUDAS"), file_name=f"Ayuda_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
