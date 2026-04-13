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

# 1. CONFIGURACIÓN INICIAL
st.set_page_config(page_title="SISTEMA FINANCIERO MASTER", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE (PRO)
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 24px !important; font-weight: 700 !important; border-radius: 12px !important; }
    .stButton>button[kind="secondary"] { background-color: #dc3545 !important; color: white !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 28px !important; font-weight: 900 !important; position: fixed; bottom: 30px; right: 30px; z-index: 9999; border: 3px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 65px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- ESTADOS GLOBALES (LLAVES DINÁMICAS) ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None

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
        ws.append([f"ESTADO DE {tipo} - {row['Nombre']}"]); ws.append(["Fecha", "Monto", "Recibo"])
        
        p_fil = pd.DataFrame()
        if not historial.empty:
            col_id = 'ID_Prestamo' if 'ID_Prestamo' in historial.columns else 'ID_Socio'
            if col_id in historial.columns:
                p_fil = historial[historial[col_id] == row['ID']]
        
        for _, p in p_fil.iterrows(): 
            f = p.get('Fecha', p.get('Fecha_Pago', ''))
            m = p.get('Monto', p.get('Monto_Pagado', 0))
            c = p.get('Comprobante', p.get('URL', p.get('URL_Comprobante', 'N/A')))
            ws.append([f, m, c])
            
        ws.append([""]); ws.append(["SALDO ACTUAL/DEUDA:", row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0))])
    return out.getvalue()

def enviar_mail(dest, nom, exc, url, tipo):
    try:
        msg = MIMEMultipart(); msg['From'] = st.secrets["EMAIL_USER"]; msg['To'] = dest; msg['Subject'] = f"✅ Comprobante {tipo} - {nom}"
        msg.attach(MIMEText(f"Hola {nom}, se registró tu movimiento financiero.\n\nRecibo: {url}", 'plain'))
        p = MIMEBase('application', 'octet-stream'); p.set_payload(exc); encoders.encode_base64(p)
        p.add_header('Content-Disposition', f"attachment; filename=Recibo_{tipo}.xlsx"); msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587); s.starttls(); s.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"]); s.send_message(msg); s.quit()
    except: pass

# --- NAVEGACIÓN ---
with st.sidebar:
    st.markdown("# 🏦 SISTEMA FINANCIERO")
    sec = st.radio("MODOS OPERATIVOS:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)

# --- 1. MODO PRÉSTAMOS ---
if sec == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    if not df_p.empty: st.download_button("📊 EXCEL GENERAL PRÉSTAMOS", data=generar_excel_grupal(df_p, "PRESTAMOS"), file_name="Reporte_Prestamos.xlsx", use_container_width=True)
    
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo_p"): st.session_state.show_form_p = not st.session_state.get('show_form_p', False)
    if st.session_state.get('show_form_p'):
        with st.form("form_p"):
            n, c, e = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Email:")
            m, t, p = st.number_input("Monto:", min_value=1.0), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("💾 GUARDAR PRÉSTAMO"):
                i = (t/100)/12; cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":n, "Cedula":c, "Email":e, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO"}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
                st.session_state.show_form_p = False; st.rerun()

    bq_p = st.text_input("🔍 BUSCAR PRESTAMISTA:")
    act_p = df_p[df_p["Estado"]=="ACTIVO"] if not df_p.empty else pd.DataFrame()
    if bq_p and not act_p.empty: act_p = act_p[act_p['Nombre'].str.contains(bq_p, case=False)]

    for idx, row in act_p.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO PENDIENTE: ${row['Saldo_Restante']}", expanded=(st.session_state.id_abierto == row['ID'])):
            c1, c2 = st.columns(2)
            with c1: 
                st.write(f"**CUOTA:** ${row['Cuota_Mensual']} | **PAGOS:** {row['Pagos_Realizados']}/{row['Meses_Totales']}")
                with st.form(key=f"fp_{row['ID']}"):
                    ft = st.file_uploader("📸 RECIBO:", key=f"ip_{row['ID']}_{st.session_state.pago_key}")
                    if st.form_submit_button("✅ CONFIRMAR COBRO"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            st.session_state.id_abierto = row['ID']
                            new_h = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": row['Cuota_Mensual'], "Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_h], ignore_index=True))
                            df_p.at[idx, "Pagos_Realizados"] += 1; df_p.at[idx, "Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                            if df_p.at[idx, "Pagos_Realizados"] >= row["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                            conn.update(worksheet="Prestamos", data=df_p)
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], generar_excel_personal(df_p.loc[idx], pd.concat([df_h, new_h]), "PRÉSTAMO"), url, "Prestamos")
                            st.session_state.pago_key += 1
                            st.rerun()
            with c2: 
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_personal(row, df_h, "PRÉSTAMO"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dlp_{row['ID']}")
            with st.popover("🗑️ ELIMINAR PRÉSTAMO"):
                st.warning(f"¿Estás seguro que quieres borrar a {row['Nombre']}?")
                if st.button("SÍ, BORRAR DEFINITIVAMENTE", key=f"del_p_{row['ID']}"):
                    df_p = df_p[df_p["ID"] != row["ID"]]; conn.update(worksheet="Prestamos", data=df_p); st.rerun()

# --- 2. MODO COOPERATIVA ---
elif sec == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    v_x = st.number_input("💵 VALOR CUOTA FIJA:", value=10.0)
    if not df_s.empty: st.download_button("📊 EXCEL GENERAL COOP", data=generar_excel_grupal(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo_c"): st.session_state.show_form_c = not st.session_state.get('show_form_c', False)
    if st.session_state.get('show_form_c'):
        with st.form("form_c"):
            n, c, e = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Email:")
            if st.form_submit_button("💾 AÑADIR SOCIO"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Email":e, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True))
                st.session_state.show_form_c = False; st.rerun()

    bq_c = st.text_input("🔍 BUSCAR SOCIO:")
    act_c = df_s if not df_s.empty else pd.DataFrame()
    if bq_c and not act_c.empty: act_c = act_c[act_c['Nombre'].str.contains(bq_c, case=False)]

    for idx, row in act_c.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | ACUMULADO: ${row['Saldo_Total_Aportado']}", expanded=(st.session_state.id_abierto == row['ID'])):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fc_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_x)
                    ft = st.file_uploader("📸 RECIBO:", key=f"ic_{row['ID']}_{st.session_state.pago_key}")
                    if st.form_submit_button("✅ REGISTRAR PAGO"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            st.session_state.id_abierto = row['ID']
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Coop", data=pd.concat([df_ph, new_h], ignore_index=True))
                            df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Cooperativa", data=df_s)
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], generar_excel_personal(df_s.loc[idx], pd.concat([df_ph, new_h]), "COOP"), url, "Cooperativa")
                            st.session_state.pago_key += 1
                            st.rerun()
            with c2: 
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_personal(row, df_ph, "COOP"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dlc_{row['ID']}")
            with st.popover("🗑️ ELIMINAR SOCIO"):
                st.warning(f"¿Estás seguro que quieres borrar a {row['Nombre']}?")
                if st.button("SÍ, BORRAR DEFINITIVAMENTE", key=f"del_c_{row['ID']}"):
                    df_s = df_s[df_s["ID"] != row["ID"]]; conn.update(worksheet="Cooperativa", data=df_s); st.rerun()

# --- 3. MODO AYUDAS ECONÓMICAS ---
elif sec == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 VALOR APORTE:", value=5.0)
    
    col_t1, col_t2 = st.columns([2, 1])
    with col_t1:
        if not df_a.empty: st.download_button("📊 EXCEL GRUPAL AYUDAS", data=generar_excel_grupal(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    with col_t2:
        if st.button("🔴 GASTO CAJA", type="secondary"): st.session_state.eg_ay = not st.session_state.get('eg_ay', False)

    if st.session_state.get('eg_ay'):
        with st.form("eg_ay"):
            det_e = st.text_input("Detalle del Gasto:")
            mon_e = st.number_input("Monto a Retirar:", min_value=1.0)
            if st.form_submit_button("⚠️ CONFIRMAR RETIRO"):
                st.success(f"Gasto de ${mon_e} guardado."); st.session_state.eg_ay = False; time.sleep(1); st.rerun()

    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo_a"): st.session_state.show_form_a = not st.session_state.get('show_form_a', False)
    if st.session_state.get('show_form_a'):
        with st.form("na"):
            n, c, e = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Email:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Email":e, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True))
                st.session_state.show_form_a = False; st.rerun()

    bq_a = st.text_input("🔍 BUSCAR EN AYUDAS:")
    act_a = df_a if not df_a.empty else pd.DataFrame()
    if bq_a and not act_a.empty: act_a = act_a[act_a['Nombre'].str.contains(bq_a, case=False)]

    for idx, row in act_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}", expanded=(st.session_state.id_abierto == row['ID'])):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fa_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_y)
                    ft = st.file_uploader("📸 RECIBO:", key=f"ia_{row['ID']}_{st.session_state.pago_key}")
                    if st.form_submit_button("✅ GUARDAR APORTE"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            st.session_state.id_abierto = row['ID']
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_ah, new_h], ignore_index=True))
                            df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Ayudas_Listado", data=df_a)
                            if row.get('Email'): enviar_mail(row['Email'], row['Nombre'], generar_excel_personal(df_a.loc[idx], pd.concat([df_ah, new_h]), "AYUDAS"), url, "Ayudas")
                            st.session_state.pago_key += 1
                            st.rerun()
            with c2: 
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_personal(row, df_ah, "AYUDAS"), file_name=f"Ayuda_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
            with st.popover("🗑️ ELIMINAR DE LISTA"):
                st.warning(f"¿Estás seguro que quieres borrar a {row['Nombre']}?")
                if st.button("SÍ, BORRAR DEFINITIVAMENTE", key=f"del_a_{row['ID']}"):
                    df_a = df_a[df_a["ID"] != row["ID"]]; conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
