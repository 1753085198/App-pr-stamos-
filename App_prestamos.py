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
st.set_page_config(page_title="SISTEMA FINANCIERO", page_icon="🏦", layout="wide")

# 2. CSS PARA EL EXAMEN (LIMPIO Y GIGANTE)
st.markdown("""
    <style>
    button[data-baseweb="tab"] { font-size: 40px !important; font-weight: 900 !important; height: 100px !important; }
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { font-size: 32px !important; font-weight: 700 !important; }
    input { font-size: 26px !important; height: 60px !important; }
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 35px !important; font-weight: 900 !important; height: 7rem !important; 
        border-radius: 20px !important; background-color: #28a745 !important; color: white !important; 
    }
    [data-testid="stMetricValue"] { font-size: 85px !important; font-weight: 900 !important; color: #007bff !important; }
    .streamlit-expanderHeader { font-size: 38px !important; font-weight: 800 !important; padding: 25px !important; }
    </style>
""", unsafe_allow_html=True)

# Conexión con reintento simple
conn = st.connection("gsheets", type=GSheetsConnection)

if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None

def cargar_datos():
    try:
        df_p = conn.read(worksheet="Prestamos", ttl=0)
        df_h = conn.read(worksheet="Pagos", ttl=0)
        if df_p is not None:
            df_p["Cedula"] = df_p["Cedula"].astype(str).str.replace(".0", "", regex=False)
            df_p["ID"] = df_p["ID"].astype(str).str.replace(".0", "", regex=False)
        return df_p, df_h
    except Exception as e:
        st.error("⚠️ Error de conexión con Google Sheets. Revisa el link en Secrets o espera 1 minuto.")
        return None, None

def subir_img(archivo):
    try:
        img = Image.open(io.BytesIO(archivo))
        if img.width > 800: img = img.resize((800, int(img.height * (800 / float(img.width)))), Image.Resampling.LANCZOS)
        buf = io.BytesIO(); img.convert('RGB').save(buf, format="JPEG", quality=70)
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": st.secrets["IMGBB_API_KEY"], "image": base64.b64encode(buf.getvalue()).decode('utf-8')})
        return res.json()["data"]["url"]
    except: return ""

def generar_excel(d_c, h_c):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("ESTADO DE CUENTA", 0)
        f_a = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        f_v = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ws["A1"] = "ESTADO DE CUENTA BANCARIO"; ws["A1"].font = Font(bold=True, size=16)
        ws["A3"], ws["B3"] = "NOMBRE:", str(d_c['Nombre']).upper()
        ws["A4"], ws["B4"] = "CÉDULA:", str(d_c['Cedula'])
        ws["D3"], ws["E3"] = "MONTO TOTAL:", f"${d_c['Monto_Inicial']}"
        ws["D4"], ws["E4"] = "SALDO ACTUAL:", f"${d_c['Saldo_Restante']}"
        h = ["N°", "Descripción", "Valor", "Estado"]
        for c, t in enumerate(h, 1):
            cell = ws.cell(row=7, column=c, value=t)
            cell.fill = f_a; cell.font = Font(color="FFFFFF", bold=True)
        pag = int(d_c['Pagos_Realizados'])
        for i in range(1, int(d_c['Meses_Totales']) + 1):
            r = 7 + i
            ws.cell(row=r, column=1, value=i); ws.cell(row=r, column=3, value=d_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws.cell(row=r, column=4, value=est)
            if i <= pag:
                for col in range(1, 5): ws.cell(row=r, column=col).fill = f_v
        ws2 = writer.book.create_sheet("HISTORIAL", 1)
        if h_c is not None and not h_c.empty:
            for r_idx, row in enumerate(dataframe_to_rows(h_c, index=False, header=True), 1):
                for c_idx, val in enumerate(row, 1): ws2.cell(row=r_idx, column=c_idx, value=val)
        for w in [ws, ws2]:
            for col in w.columns: w.column_dimensions[col[0].column_letter].width = 25
    return out.getvalue()

def enviar_mail(dest, nom, exc, url):
    try:
        msg = MIMEMultipart(); msg['From'] = st.secrets["EMAIL_USER"]; msg['To'] = dest; msg['Subject'] = f"✅ Pago Recibido - {nom}"
        msg.attach(MIMEText(f"Hola {nom}, adjunto tu estado actualizado.\nRecibo: {url}", 'plain'))
        p = MIMEBase('application', 'octet-stream'); p.set_payload(exc); encoders.encode_base64(p)
        p.add_header('Content-Disposition', f"attachment; filename=Estado_{nom}.xlsx"); msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"]); s.send_message(msg); s.quit()
    except: pass

# --- INTERFAZ ---
st.title("🏦 PANEL DE CONTROL")
df_p, df_h = cargar_datos()

if df_p is not None:
    t_g, t_n = st.tabs(["📋 GESTIÓN", "➕ NUEVO"])

    with t_g:
        bq = st.text_input("🔍 BUSCAR:", placeholder="Escribe nombre...")
        act = df_p[df_p["Estado"] == "ACTIVO"]
        if bq: act = act[act['Nombre'].str.contains(bq, case=False) | act['Cedula'].str.contains(bq)]
        
        for idx, row in act.iterrows():
            is_open = st.session_state.id_abierto == row['ID']
            with st.expander(f"👤 {row['Nombre'].upper()} | 💰 SALDO: ${row['Saldo_Restante']}", expanded=is_open):
                c1, c2 = st.columns(2)
                with c1:
                    st.metric("CUOTA", f"${row['Cuota_Mensual']}")
                    st.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
                    h_c = df_h[df_h["ID_Prestamo"] == row['ID']] if df_h is not None else pd.DataFrame()
                    st.download_button("📥 EXCEL", data=generar_excel(row, h_c), file_name=f"Estado_{row['Nombre']}.xlsx", key=f"ex_{row['ID']}", use_container_width=True)
                with c2:
                    with st.form(key=f"f_{row['ID']}_{st.session_state.pago_key}"):
                        ml = st.text_input("Correo:", value=row.get('Email', ""))
                        nc = st.number_input("Cuotas:", min_value=1, value=1)
                        ft = st.file_uploader("📸 RECIBO (OBLIGATORIO):", type=["jpg","png","jpeg"], key=f"foto_{row['ID']}_{st.session_state.pago_key}")
                        if st.form_submit_button("✅ CONFIRMAR", use_container_width=True, type="primary"):
                            if ft is None:
                                st.error("❌ ERROR: Carga la foto para procesar.")
                            else:
                                st.session_state.id_abierto = row['ID']
                                url = subir_img(ft.getvalue())
                                new_pg = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": nc, "Monto_Pagado": round(row['Cuota_Mensual']*nc, 2), "URL_Comprobante": url}])
                                conn.update(worksheet="Pagos", data=pd.concat([df_h, new_pg], ignore_index=True))
                                r_u = row.copy(); r_u["Pagos_Realizados"] += nc
                                r_u["Saldo_Restante"] = round(max(0, row["Saldo_Restante"] - (row["Monto_Inicial"]/row["Meses_Totales"])*nc), 2)
                                if r_u["Pagos_Realizados"] >= row["Meses_Totales"]: r_u["Estado"] = "PAGADO"
                                df_p.loc[idx] = r_u; conn.update(worksheet="Prestamos", data=df_p)
                                if ml: enviar_mail(ml, row['Nombre'], generar_excel(r_u, pd.concat([df_h, new_pg])), url)
                                st.session_state.pago_key += 1
                                st.balloons(); time.sleep(0.5); st.rerun()

    with t_n:
        with st.form("n", clear_on_submit=True):
            nm, cd, ml = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Correo:")
            c1, c2, c3 = st.columns(3)
            mn, ts, pz = c1.number_input("Monto:"), c3.number_input("Tasa %:"), c2.number_input("Meses:")
            if st.form_submit_button("💾 GUARDAR", use_container_width=True, type="primary"):
                tm = (ts/100)/12; cu = mn * (tm * (1+tm)**pz) / ((1+tm)**pz - 1) if tm > 0 else mn/pz
                new = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nm, "Cedula": cd, "Email": ml, "Monto_Inicial": mn, "Saldo_Restante": mn, "Cuota_Mensual": round(cu, 2), "Meses_Totales": int(pz), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": ts}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
                st.balloons(); time.sleep(0.5); st.rerun()
