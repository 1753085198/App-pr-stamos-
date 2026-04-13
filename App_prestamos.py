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

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y BOTÓN FLOTANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 30px !important; font-weight: 700 !important; }
    input { font-size: 26px !important; height: 60px !important; }
    .stDownloadButton>button { font-size: 32px !important; font-weight: 900 !important; height: 6rem !important; border-radius: 20px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button[kind="primary"] { font-size: 32px !important; font-weight: 900 !important; height: 6rem !important; border-radius: 20px !important; background-color: #28a745 !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 30px !important; font-weight: 900 !important; position: fixed; bottom: 40px; right: 40px; z-index: 9999; border: 4px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 80px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown("# 🏦 NAVEGACIÓN")
    seccion = st.radio("Selecciona:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)
    st.write("---")
    st.info("Jose Figueroa - UDLA")

# --- ESTADOS ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None
if 'mostrar_nuevo' not in st.session_state: st.session_state.mostrar_nuevo = False

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

def generar_excel_pintado(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("REPORTE", 0)
        f_verde, f_rojo = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE: {titulo}"]); ws.append(["Nombre", "Cédula", "Monto", "Estado"])
        for r_idx, row in df.iterrows():
            monto = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            estado = "PAGADO" if monto > 0 else "PENDIENTE"
            color = f_verde if monto > 0 else f_rojo
            ws.append([row['Nombre'], row['Cedula'], monto, estado])
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = color
    return out.getvalue()

# --- SECCIONES ---

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS (FUNCIONAL)")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if st.session_state.mostrar_nuevo:
        with st.form("np"):
            nm, cd = st.text_input("Nombre:"), st.text_input("Cédula:")
            m, t, p = st.number_input("Monto:"), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("GUARDAR"):
                i = (t/100)/12
                cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                nuevo = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":nm, "Cedula":cd, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO"}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR PRESTAMISTA:")
    act = df_p[df_p["Estado"]=="ACTIVO"] if not df_p.empty else pd.DataFrame()
    if bq and not act.empty: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}"):
            c1, c2 = st.columns(2)
            with c1:
                st.metric("CUOTA", f"${row['Cuota_Mensual']}")
                st.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
            with c2:
                with st.form(key=f"f_{row['ID']}_{st.session_state.pago_key}"):
                    ft = st.file_uploader("📸 RECIBO:", key=f"foto_{row['ID']}")
                    if st.form_submit_button("✅ CONFIRMAR COBRO"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_p = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d"), "Monto_Pagado": row['Cuota_Mensual'], "URL_Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_p], ignore_index=True))
                            row["Pagos_Realizados"] += 1; row["Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                            if row["Pagos_Realizados"] >= row["Meses_Totales"]: row["Estado"] = "PAGADO"
                            df_p.loc[idx] = row; conn.update(worksheet="Prestamos", data=df_p)
                            st.session_state.pago_key += 1; st.rerun()

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 LISTADO COOPERATIVA")
    df_s = cargar("Cooperativa")
    if not df_s.empty:
        st.download_button("📊 EXCEL UNIFICADO COOP", data=generar_excel_pintado(df_s, "COOPERATIVA"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    # (Lógica de registro de socios igual que antes...)

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 LISTADO AYUDAS (INDIVIDUAL Y GRUPAL)")
    df_a = cargar("Ayudas_Listado")
    
    if not df_a.empty:
        st.download_button("📊 EXCEL UNIFICADO AYUDAS (GRUPAL)", data=generar_excel_pintado(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)

    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo"):
        st.session_state.mostrar_nuevo = not st.session_state.mostrar_nuevo

    if st.session_state.mostrar_nuevo:
        with st.form("na"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq_a = st.text_input("🔍 BUSCAR COMPAÑERO:")
    act_a = df_a if not df_a.empty else pd.DataFrame()
    if bq_a: act_a = act_a[act_a['Nombre'].str.contains(bq_a, case=False)]
    
    for idx, row in act_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row.get('Saldo_Total_Aportado', 0)}"):
            c1, c2 = st.columns(2)
            with c1:
                st.write("### 💵 REGISTRAR APORTE")
                m_a = st.number_input("Monto:", value=5.0, key=f"ma_{row['ID']}")
                if st.button("✅ GUARDAR", key=f"ba_{row['ID']}"):
                    df_a.at[idx, "Saldo_Total_Aportado"] = float(row.get('Saldo_Total_Aportado', 0)) + m_a
                    conn.update(worksheet="Ayudas_Listado", data=df_a)
                    st.rerun()
            with c2:
                st.write("### 📄 COMPROBANTE PERSONAL")
                # Botón de Excel Personal solo para este socio
                excel_personal = generar_excel_pintado(pd.DataFrame([row]), f"Socio {row['Nombre']}")
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=excel_personal, file_name=f"Comprobante_{row['Nombre']}.xlsx", key=f"dl_{row['ID']}")
