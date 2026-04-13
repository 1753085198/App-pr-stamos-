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
st.set_page_config(page_title="SISTEMA PRO", page_icon="🏦", layout="wide")

# 2. CSS PERSONALIZADO
st.markdown("""
    <style>
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { font-size: 32px !important; font-weight: 700 !important; }
    input { font-size: 26px !important; height: 60px !important; }
    .stButton>button[kind="primary"] { 
        font-size: 35px !important; font-weight: 900 !important; height: 7rem !important; 
        border-radius: 20px !important; background-color: #28a745 !important; color: white !important; 
    }
    .stDownloadButton>button {
        font-size: 35px !important; font-weight: 900 !important; height: 7rem !important;
        border-radius: 20px !important; background-color: #1D6F42 !important; color: white !important;
    }
    div.stButton > button:first-child[key^="btn_nuevo_circular"] {
        background-color: #ff5722 !important; color: white !important;
        border-radius: 50% !important; width: 120px !important; height: 120px !important;
        font-size: 60px !important; position: fixed; bottom: 40px; right: 40px; z-index: 9999;
    }
    [data-testid="stMetricValue"] { font-size: 85px !important; font-weight: 900 !important; color: #007bff !important; }
    .streamlit-expanderHeader { font-size: 38px !important; font-weight: 800 !important; padding: 25px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None
if 'mostrar_nuevo' not in st.session_state: st.session_state.mostrar_nuevo = False

def cargar_datos():
    try:
        df_p = conn.read(worksheet="Prestamos", ttl=0)
        df_h = conn.read(worksheet="Pagos", ttl=0)
        if df_p is not None:
            df_p["Cedula"] = df_p["Cedula"].astype(str).str.replace(".0", "", regex=False)
            df_p["ID"] = df_p["ID"].astype(str).str.replace(".0", "", regex=False)
        return df_p, df_h
    except: return None, None

def subir_img(archivo):
    try:
        img = Image.open(io.BytesIO(archivo))
        if img.width > 800: img = img.resize((800, int(img.height * (800 / float(img.width)))), Image.Resampling.LANCZOS)
        buf = io.BytesIO(); img.convert('RGB').save(buf, format="JPEG", quality=70)
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": st.secrets["IMGBB_API_KEY"], "image": base64.b64encode(buf.getvalue()).decode('utf-8')})
        return res.json()["data"]["url"]
    except: return ""

def generar_excel_pro(d_c, h_c):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("ESTADO DE CUENTA", 0)
        f_azul, f_verde = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid"), PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ws["A1"] = "ESTADO DE CUENTA BANCARIO"; ws["A1"].font = Font(bold=True, size=16)
        ws["A3"], ws["B3"] = "NOMBRE:", str(d_c['Nombre']).upper()
        ws["A4"], ws["B4"] = "CÉDULA:", str(d_c['Cedula'])
        ws["D3"], ws["E3"] = "MONTO TOTAL:", f"${d_c['Monto_Inicial']}"
        ws["D4"], ws["E4"] = "SALDO ACTUAL:", f"${d_c['Saldo_Restante']}"
        
        headers = ["N°", "Descripción", "Valor Cuota", "Estado"]
        for c, t in enumerate(headers, 1):
            cell = ws.cell(row=11, column=c, value=t)
            cell.fill = f_azul; cell.font = Font(color="FFFFFF", bold=True)
            
        pag = int(d_c['Pagos_Realizados'])
        for i in range(1, int(d_c['Meses_Totales']) + 1):
            r = 11 + i
            ws.cell(row=r, column=1, value=i); ws.cell(row=r, column=3, value=d_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws.cell(row=r, column=4, value=est)
            if i <= pag:
                for col in range(1, 5): ws.cell(row=r, column=col).fill = f_verde

        ws_comp = writer.book.create_sheet("COMPROBANTES", 1)
        ws_comp["A1"] = "REGISTRO DE RECIBOS"; ws_comp["A1"].font = Font(bold=True, size=14)
        h_comp = ["Fecha", "Monto", "Cuotas", "Link"]
        for c, t in enumerate(h_comp, 1):
            cell = ws_comp.cell(row=3, column=c, value=t)
            cell.fill = f_azul; cell.font = Font(color="FFFFFF", bold=True)
        
        if h_c is not None and not h_c.empty:
            # FILTRAR SOLO LOS PAGOS DE ESTA PERSONA
            h_c_filtrado = h_c[h_c["ID_Prestamo"] == d_c['ID']].reset_index(drop=True)
            for r_idx, fila_h in h_c_filtrado.iterrows():
                r_num = r_idx + 4
                ws_comp.cell(row=r_num, column=1, value=str(fila_h['Fecha_Pago']))
                ws_comp.cell(row=r_num, column=2, value=fila_h['Monto_Pagado'])
                ws_comp.cell(row=r_num, column=3, value=fila_h['Cuotas_Pagadas'])
                url_val = fila_h.get('URL_Comprobante', "")
                if pd.notnull(url_val) and str(url_val).startswith("http"):
                    link_c = ws_comp.cell(row=r_num, column=4, value="VER RECIBO")
                    link_c.hyperlink = str(url_val)
                    link_c.font = Font(color="0000FF", underline="single")
        
        for w in writer.book.worksheets:
            for col in w.columns: w.column_dimensions[col[0].column_letter].width = 25
    return out.getvalue()

def enviar_mail(dest, nom, exc, url):
    try:
        msg = MIMEMultipart(); msg['From'] = st.secrets["EMAIL_USER"]; msg['To'] = dest; msg['Subject'] = f"✅ Comprobante - {nom}"
        msg.attach(MIMEText(f"Hola {nom}, adjunto tu estado. Recibo: {url}", 'plain'))
        p = MIMEBase('application', 'octet-stream'); p.set_payload(exc); encoders.encode_base64(p)
        p.add_header('Content-Disposition', f"attachment; filename=Estado_{nom}.xlsx"); msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"]); s.send_message(msg); s.quit()
    except: pass

# --- INTERFAZ ---
st.title("🏦 PANEL DE CONTROL")
df_p, df_h = cargar_datos()

if df_p is not None:
    if st.button("👤➕", key="btn_nuevo_circular"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if st.session_state.mostrar_nuevo:
        with st.form("n_form", clear_on_submit=True):
            st.markdown("### 👤➕ REGISTRAR NUEVO")
            nm, ced, ml = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Correo:")
            c1, c2, c3 = st.columns(3)
            mn, ts, pz = c1.number_input("Monto:"), c3.number_input("Tasa %:"), c2.number_input("Meses:")
            if st.form_submit_button("💾 GUARDAR"):
                tm = (ts/100)/12; cu = mn * (tm * (1+tm)**pz) / ((1+tm)**pz - 1) if tm > 0 else mn/pz
                nuevo = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nm, "Cedula": ced, "Email": ml, "Monto_Inicial": mn, "Saldo_Restante": mn, "Cuota_Mensual": round(cu, 2), "Meses_Totales": int(pz), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": ts}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR:", placeholder="Escribe nombre...")
    act = df_p[df_p["Estado"] == "ACTIVO"]
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        is_open = st.session_state.id_abierto == row['ID']
        with st.expander(f"👤 {row['Nombre'].upper()} | 💰 SALDO: ${row['Saldo_Restante']}", expanded=is_open):
            c1, c2 = st.columns(2)
            with c1:
                st.metric("CUOTA", f"${row['Cuota_Mensual']}")
                st.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
                h_c_local = df_h[df_h["ID_Prestamo"] == row['ID']] if df_h is not None else pd.DataFrame()
                st.download_button("📊 DESCARGAR EXCEL", data=generar_excel_pro(row, h_c_local), file_name=f"Estado_{row['Nombre']}.xlsx", key=f"ex_{row['ID']}", use_container_width=True)
            with c2:
                with st.form(key=f"f_{row['ID']}_{st.session_state.pago_key}"):
                    correo, n_cuotas = st.text_input("Correo:", value=row.get('Email', "")), st.number_input("Cuotas:", min_value=1, value=1)
                    ft = st.file_uploader("📸 RECIBO:", type=["jpg","png","jpeg"], key=f"foto_{row['ID']}")
                    if st.form_submit_button("✅ CONFIRMAR"):
                        if ft:
                            st.session_state.id_abierto = row['ID']
                            url = subir_img(ft.getvalue())
                            new_p = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": n_cuotas, "Monto_Pagado": round(row['Cuota_Mensual']*n_cuotas, 2), "URL_Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_p], ignore_index=True))
                            
                            row_upd = row.copy()
                            row_upd["Pagos_Realizados"] += n_cuotas
                            row_upd["Saldo_Restante"] = round(max(0, row["Saldo_Restante"] - (row["Monto_Inicial"]/row["Meses_Totales"])*n_cuotas), 2)
                            if row_upd["Pagos_Realizados"] >= row["Meses_Totales"]: row_upd["Estado"] = "PAGADO"
                            df_p.loc[idx] = row_upd; conn.update(worksheet="Prestamos", data=df_p)
                            
                            # GENERAR EXCEL FILTRADO SOLO PARA ESTE CORREO
                            h_c_actualizado = pd.concat([h_c_local, new_p], ignore_index=True)
                            exc_final = generar_excel_pro(row_upd, h_c_actualizado)
                            
                            if correo: enviar_mail(correo, row['Nombre'], exc_final, url)
                            st.session_state.pago_key += 1; st.rerun()
            
            if st.button(f"🗑️ ELIMINAR {row['Nombre'].split()[0]}", key=f"del_{row['ID']}", use_container_width=True):
                conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != row['ID']]); st.rerun()
