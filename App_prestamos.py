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

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="CONTROL DE PRESTAMOS PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ ULTRA-GIGANTE
st.markdown("""
    <style>
    button[data-baseweb="tab"] { font-size: 40px !important; font-weight: 900 !important; height: 100px !important; }
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { font-size: 32px !important; font-weight: 700 !important; }
    
    /* Buscador tamaño ajustado */
    input { font-size: 26px !important; height: 60px !important; }
    
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 35px !important; font-weight: 900 !important; height: 7rem !important; 
        border-radius: 20px !important; background-color: #28a745 !important; color: white !important; 
        box-shadow: 0px 8px 16px rgba(0,0,0,0.3) !important;
    }
    .stButton>button[kind="secondary"] { font-size: 22px !important; height: 4rem !important; }
    [data-testid="stMetricValue"] { font-size: 85px !important; font-weight: 900 !important; color: #007bff !important; }
    .streamlit-expanderHeader { font-size: 38px !important; font-weight: 800 !important; padding: 25px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- ESTADOS DE SESIÓN ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'cliente_abierto' not in st.session_state: st.session_state.cliente_abierto = None

# --- FUNCIÓN ENVÍO DE CORREO ---
def enviar_correo_pago(email_destino, nombre_completo, excel_data, url_comprobante):
    try:
        remitente = st.secrets["EMAIL_USER"]
        password = st.secrets["EMAIL_PASS"]
        
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = email_destino
        msg['Subject'] = f"✅ Comprobante de Pago - {nombre_completo}"
        
        cuerpo = f"""
        Hola {nombre_completo},
        
        ¡Muchas gracias por tu pago! Hemos registrado tu abono correctamente.
        
        Adjunto encontrarás tu Estado de Cuenta actualizado.
        Puedes ver tu comprobante aquí: {url_comprobante}
        
        Gracias por tu confianza.
        """
        msg.attach(MIMEText(cuerpo, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= Estado_Cuenta_{nombre_completo}.xlsx")
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        return True
    except:
        return False

# --- CARGA DE DATOS ---
def cargar_datos():
    df_p = conn.read(worksheet="Prestamos", ttl="0")
    df_h = conn.read(worksheet="Pagos", ttl="0")
    if df_p is not None:
        df_p["Cedula"] = df_p["Cedula"].astype(str).str.replace(".0", "", regex=False)
        df_p["ID"] = df_p["ID"].astype(str).str.replace(".0", "", regex=False)
    return df_p, df_h

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

def generar_excel_de_gala(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ws1 = writer.book.create_sheet("ESTADO DE CUENTA", 0)
        fill_h = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        fill_v = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        f_w = Font(color="FFFFFF", bold=True); f_g = Font(color="006100", bold=True)
        
        ws1["A1"] = "ESTADO DE CUENTA"; ws1["A1"].font = Font(bold=True, size=16)
        ws1["A3"] = "NOMBRE:"; ws1["B3"] = str(datos_c['Nombre']).upper()
        ws1["A4"] = "CÉDULA:"; ws1["B4"] = str(datos_c['Cedula']).replace(".0", "")
        ws1["D3"] = "VALOR TOTAL:"; ws1["E3"] = f"${datos_c['Monto_Inicial']}"
        ws1["D4"] = "SALDO ACTUAL:"; ws1["E4"] = f"${datos_c['Saldo_Restante']}"
        
        headers = ["N° Cuota", "Descripción", "Valor Cuota ($)", "Estado Actual"]
        for c, t in enumerate(headers, 1):
            cell = ws1.cell(row=7, column=c, value=t)
            cell.fill = fill_h; cell.font = f_w; cell.alignment = Alignment(horizontal="center")
            
        pag = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            r = 7 + i
            ws1.cell(row=r, column=1, value=i); ws1.cell(row=r, column=2, value=f"Cuota mes {i}")
            ws1.cell(row=r, column=3, value=datos_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws1.cell(row=r, column=4, value=est)
            if i <= pag:
                for col in range(1, 5): ws1.cell(row=r, column=col).fill = fill_v; ws1.cell(row=r, column=col).font = f_g
        
        ws2 = writer.book.create_sheet("HISTORIAL PAGOS", 1)
        if not historial_c.empty:
            for r_idx, row in enumerate(dataframe_to_rows(historial_c, index=False, header=True), 1):
                for c_idx, val in enumerate(row, 1): ws2.cell(row=r_idx, column=c_idx, value=val)
        for ws in [ws1, ws2]:
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 25
    return output.getvalue()

# --- INTERFAZ ---
st.title("🏦 PANEL DE CONTROL")
df_p, df_h = cargar_datos()

t_g, t_n = st.tabs(["📋 GESTIÓN", "➕ NUEVO"])

with t_g:
    bq = st.text_input("🔍 BUSCAR POR NOMBRE O CÉDULA:", placeholder="Escribe aquí...")
    act = df_p[df_p["Estado"] == "ACTIVO"] if df_p is not None else pd.DataFrame()
    if bq and not act.empty:
        act = act[act['Nombre'].str.contains(bq, case=False) | act['Cedula'].str.contains(bq)]

    if not act.empty:
        for idx, row in act.iterrows():
            is_open = st.session_state.cliente_abierto == row['ID']
            with st.expander(f"👤 {row['Nombre'].upper()} | 💰 SALDO: ${row['Saldo_Restante']}", expanded=is_open):
                st.session_state.cliente_abierto = row['ID']
                c1, c2 = st.columns(2)
                with c1:
                    st.write("### ℹ️ RESUMEN")
                    st.metric("CUOTA", f"${row['Cuota_Mensual']}")
                    st.metric("AVANCE", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
                    h_c = df_h[df_h["ID_Prestamo"] == row['ID']] if df_h is not None else pd.DataFrame()
                    exc_f = generar_excel_de_gala(row, h_c)
                    st.download_button(f"📥 DESCARGAR EXCEL", data=exc_f, file_name=f"Estado_{row['Nombre']}.xlsx", key=f"ex_{row['ID']}", use_container_width=True)
                with c2:
                    st.write("### 💵 COBRAR PAGO")
                    with st.form(key=f"f_{row['ID']}_{st.session_state.pago_key}"):
                        ml = st.text_input("Enviar a este Correo:", value=row.get('Email', ""))
                        nc = st.number_input("Cuotas:", min_value=1, value=1)
                        ft = st.file_uploader("📸 RECIBO:", type=["jpg","png","jpeg"], key=f"f_{row['ID']}_{st.session_state.pago_key}")
                        if st.form_submit_button("✅ CONFIRMAR Y ENVIAR", use_container_width=True, type="primary"):
                            with st.spinner('Procesando...'):
                                url = subir_a_imgbb_comprimido(ft.getvalue()) if ft else ""
                                np = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": nc, "Monto_Pagado": round(row['Cuota_Mensual']*nc, 2), "URL_Comprobante": url}])
                                conn.update(worksheet="Pagos", data=pd.concat([df_h, np], ignore_index=True))
                                df_p.at[idx, "Pagos_Realizados"] += nc
                                df_p.at[idx, "Saldo_Restante"] = round(max(0, row["Saldo_Restante"] - (row["Monto_Inicial"]/row["Meses_Totales"])*nc), 2)
                                if df_p.at[idx, "Pagos_Realizados"] >= row["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                                conn.update(worksheet="Prestamos", data=df_p)
                                if ml: enviar_correo_pago(ml, row['Nombre'], exc_f, url)
                                st.session_state.pago_key += 1
                                st.balloons(); time.sleep(1); st.rerun()
                if st.button(f"🗑️ ELIMINAR NOMBRE", key=f"del_{row['ID']}", use_container_width=True, type="secondary"):
                    conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != row['ID']]); st.rerun()

with t_n:
    with st.form("n", clear_on_submit=True):
        st.write("### 📝 REGISTRO DE NOMBRE")
        nm = st.text_input("Nombre Completo:"); cd = st.text_input("Cédula:"); em = st.text_input("Correo Electrónico:")
        c1, c2, c3 = st.columns(3)
        mn = c1.number_input("Monto:", value=500.0); ts = c3.number_input("Tasa %:", value=15.0); pz = c2.number_input("Meses:", value=12)
        if st.form_submit_button("💾 GUARDAR", use_container_width=True, type="primary"):
            tm = (ts/100)/12; cu = mn * (tm * (1+tm)**pz) / ((1+tm)**pz - 1) if tm > 0 else mn/pz
            new = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nm, "Cedula": cd, "Email": em, "Monto_Inicial": mn, "Saldo_Restante": mn, "Cuota_Mensual": round(cu, 2), "Meses_Totales": int(pz), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": ts}])
            conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
            st.balloons(); time.sleep(1); st.rerun()
