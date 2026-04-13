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
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="CONTROL DE PRESTAMOS PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ ULTRA-GIGANTE
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

conn = st.connection("gsheets", type=GSheetsConnection)

# --- ESTADOS DE SESIÓN (CRUCIAL PARA LA PERSISTENCIA) ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'cliente_abierto' not in st.session_state: st.session_state.cliente_abierto = None

# --- FUNCION CORREO ---
def enviar_correo_pago(email_destino, nombre, excel_data, url_comprobante):
    try:
        remitente = st.secrets["EMAIL_USER"]
        password = st.secrets["EMAIL_PASS"]
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = email_destino
        msg['Subject'] = f"✅ Comprobante de Pago - {nombre}"
        cuerpo = f"Hola {nombre},\n\nConfirmamos tu pago. Adjunto tu estado de cuenta actualizado.\nRecibo: {url_comprobante}"
        msg.attach(MIMEText(cuerpo, 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= Estado_{nombre}.xlsx")
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

# --- GENERADOR DE EXCEL (ESTILO IMAGEN) ---
def generar_excel_premium(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ws1 = writer.book.create_sheet("ESTADO DE CUENTA", 0)
        f_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        font_b = Font(bold=True); font_w = Font(color="FFFFFF", bold=True); font_g = Font(color="006100", bold=True)
        
        ws1["A1"] = "ESTADO DE CUENTA BANCARIO"; ws1["A1"].font = Font(bold=True, size=16)
        ws1["A3"] = "NOMBRE:"; ws1["B3"] = str(datos_c['Nombre']).upper()
        ws1["A4"] = "CÉDULA:"; ws1["B4"] = str(datos_c['Cedula']).replace(".0", "")
        ws1["A5"] = "FECHA DE INICIO:"; ws1["B5"] = str(datos_c['Fecha'])
        ws1["D3"] = "MONTO TOTAL PRESTADO:"; ws1["E3"] = f"${datos_c['Monto_Inicial']}"
        ws1["D4"] = "TASA DE INTERÉS:"; ws1["E4"] = f"{datos_c['Tasa']}% Anual"
        ws1["D5"] = "PLAZO TOTAL:"; ws1["E5"] = f"{datos_c['Meses_Totales']} Meses"
        ws1["D7"] = "PAGOS REALIZADOS:"; ws1["E7"] = f"{datos_c['Pagos_Realizados']} Cuotas"
        ws1["D8"] = "VALOR FALTANTE (SALDO):"; ws1["E8"] = f"${datos_c['Saldo_Restante']}"
        ws1["E8"].font = Font(bold=True, color="FF0000")
        
        for r in [3, 4, 5, 7, 8]: ws1[f"A{r}"].font = font_b; ws1[f"D{r}"].font = font_b

        ws1["A10"] = "PLAN DE PAGOS"; ws1["A10"].font = font_b
        h = ["N° Cuota", "Descripción del Pago", "Valor Cuota ($)", "Estado Actual"]
        for c, t in enumerate(h, 1):
            cell = ws1.cell(row=11, column=c, value=t)
            cell.fill = f_azul; cell.font = font_w; cell.alignment = Alignment(horizontal="center")
            
        pag = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            r = 11 + i
            ws1.cell(row=r, column=1, value=i); ws1.cell(row=r, column=2, value=f"Cuota del mes número {i}")
            ws1.cell(row=r, column=3, value=datos_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws1.cell(row=r, column=4, value=est)
            if i <= pag:
                for col in range(1, 5): ws1.cell(row=r, column=col).fill = f_verde; ws1.cell(row=r, column=col).font = font_g
        
        ws2 = writer.book.create_sheet("HISTORIAL", 1)
        if not historial_c.empty:
            for r_idx, row in enumerate(dataframe_to_rows(historial_c, index=False, header=True), 1):
                for c_idx, val in enumerate(row, 1): ws2.cell(row=r_idx, column=c_idx, value=val)
        for ws in [ws1, ws2]:
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 25
    return output.getvalue()

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

# --- INTERFAZ ---
st.title("🏦 PANEL DE CONTROL")
df_p, df_h = cargar_datos()
t_g, t_n = st.tabs(["📋 GESTIÓN", "➕ NUEVO"])

with t_g:
    bq = st.text_input("🔍 BUSCAR NOMBRE O CÉDULA:", placeholder="Escribe aquí...")
    act = df_p[df_p["Estado"] == "ACTIVO"] if df_p is not None else pd.DataFrame()
    if bq and not act.empty:
        act = act[act['Nombre'].str.contains(bq, case=False) | act['Cedula'].str.contains(bq)]
    
    if not act.empty:
        for idx, row in act.iterrows():
            # ESTA ES LA CLAVE: expanded=True si el ID coincide
            abierto_manual = st.session_state.cliente_abierto == row['ID']
            with st.expander(f"👤 {row['Nombre'].upper()} | 💰 SALDO: ${row['Saldo_Restante']}", expanded=abierto_manual):
                # Si el usuario hace clic en el expander, guardamos su ID
                if st.button(f"Fijar vista de {row['Nombre'].split()[0]}", key=f"fijar_{row['ID']}"):
                    st.session_state.cliente_abierto = row['ID']
                
                c1, c2 = st.columns(2)
                with c1:
                    st.metric("CUOTA", f"${row['Cuota_Mensual']}")
                    st.metric("AVANCE", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
                    h_c = df_h[df_h["ID_Prestamo"] == row['ID']] if df_h is not None else pd.DataFrame()
                    st.download_button(f"📥 EXCEL ACTUAL", data=generar_excel_premium(row, h_c), file_name=f"Estado_{row['Nombre']}.xlsx", key=f"ex_{row['ID']}", use_container_width=True)
                with c2:
                    # Formulario con llave dinámica para limpiar la foto
                    with st.form(key=f"f_{row['ID']}_{st.session_state.pago_key}"):
                        st.session_state.cliente_abierto = row['ID'] # Asegurar que se guarde al interactuar
                        mail = st.text_input("Enviar reporte a:", value=row.get('Email', ""))
                        nc = st.number_input("Cuotas:", min_value=1, value=1)
                        ft = st.file_uploader("📸 RECIBO:", type=["jpg","png","jpeg"], key=f"foto_{row['ID']}_{st.session_state.pago_key}")
                        if st.form_submit_button("✅ CONFIRMAR Y ENVIAR", use_container_width=True, type="primary"):
                            url = subir_a_imgbb_comprimido(ft.getvalue()) if ft else ""
                            # Actualización lógica
                            new_pg = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": nc, "Monto_Pagado": round(row['Cuota_Mensual']*nc, 2), "URL_Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_pg], ignore_index=True))
                            r_u = row.copy()
                            r_u["Pagos_Realizados"] += nc
                            r_u["Saldo_Restante"] = round(max(0, row["Saldo_Restante"] - (row["Monto_Inicial"]/row["Meses_Totales"])*nc), 2)
                            if r_u["Pagos_Realizados"] >= row["Meses_Totales"]: r_u["Estado"] = "PAGADO"
                            df_p.loc[idx] = r_u
                            conn.update(worksheet="Prestamos", data=df_p)
                            # Enviar correo con datos nuevos
                            h_act = pd.concat([df_h, new_pg], ignore_index=True)
                            exc_act = generar_excel_premium(r_u, h_act[h_act["ID_Prestamo"] == row['ID']])
                            if mail: enviar_correo_pago(mail, row['Nombre'], exc_act, url)
                            
                            st.session_state.pago_key += 1
                            st.balloons(); time.sleep(1); st.rerun()
                
                if st.button(f"🗑️ ELIMINAR", key=f"del_{row['ID']}", use_container_width=True, type="secondary"):
                    conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != row['ID']]); st.rerun()

with t_n:
    with st.form("n", clear_on_submit=True):
        st.write("### 📝 REGISTRO")
        nm = st.text_input("Nombre Completo:"); cd = st.text_input("Cédula:"); ml = st.text_input("Correo:"); mnt = st.number_input("Monto:", value=500.0); ts = st.number_input("Tasa %:", value=15.0); pz = st.number_input("Meses:", value=12)
        if st.form_submit_button("💾 GUARDAR", use_container_width=True, type="primary"):
            tm = (ts/100)/12; cu = mnt * (tm * (1+tm)**pz) / ((1+tm)**pz - 1) if tm > 0 else mnt/pz
            new = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nm, "Cedula": cd, "Email": ml, "Monto_Inicial": mnt, "Saldo_Restante": mnt, "Cuota_Mensual": round(cu, 2), "Meses_Totales": int(pz), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": ts}])
            conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
            st.balloons(); time.sleep(1); st.rerun()
