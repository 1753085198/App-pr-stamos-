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

# 1. CONFIGURACIÓN MASTER
st.set_page_config(page_title="SISTEMA FINANCIERO PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y BOTÓN FLOTANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 24px !important; font-weight: 700 !important; border-radius: 12px !important; }
    .stButton>button[kind="secondary"] { background-color: #dc3545 !important; color: white !important; }
    /* Botón Nuevo Naranja Flotante */
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 28px !important; font-weight: 900 !important; position: fixed; bottom: 30px; right: 30px; z-index: 9999; border: 3px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 65px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNCIONES CORE ---
def cargar(h):
    try:
        df = conn.read(worksheet=h, ttl=0)
        if df is not None:
            if "ID" in df.columns: df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
            return df
        return pd.DataFrame()
    except: return pd.DataFrame()

def generar_excel_grupal(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("GENERAL", 0)
        f_v, f_r = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE GENERAL - {titulo}"]); ws.append(["Nombre", "Cédula", "Monto", "Estado"])
        for _, row in df.iterrows():
            m = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            ws.append([row['Nombre'], row['Cedula'], m, "ACTIVO" if m > 0 else "PENDIENTE"])
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = f_v if m > 0 else f_r
    return out.getvalue()

def generar_excel_personal(row, historial, tipo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("HISTORIAL", 0)
        ws.append([f"COMPROBANTE {tipo} - {row['Nombre']}"]); ws.append(["Fecha", "Monto", "Recibo"])
        p_fil = historial[historial['ID_Socio'] == row['ID']] if not historial.empty else pd.DataFrame()
        for _, p in p_fil.iterrows(): ws.append([p['Fecha'], p['Monto'], p.get('Comprobante', 'N/A')])
        ws.append([""]); ws.append(["SALDO ACTUAL:", row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0))])
    return out.getvalue()

# --- NAVEGACIÓN ---
with st.sidebar:
    st.markdown("# 🏦 SISTEMA")
    sec = st.radio("SECCIONES:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)

# --- 1. PRÉSTAMOS (REPARADO) ---
if sec == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo_p"): st.session_state.show_form_p = True
    
    if st.session_state.get('show_form_p'):
        with st.form("form_p"):
            n, c, e = st.text_input("Nombre:"), st.text_input("Cédula:"), st.text_input("Email:")
            m, t, p = st.number_input("Monto:"), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("💾 GUARDAR"):
                i = (t/100)/12; cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":n, "Cedula":c, "Email":e, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO"}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
                st.session_state.show_form_p = False; st.rerun()

    if not df_p.empty:
        st.download_button("📊 EXCEL GENERAL", data=generar_excel_grupal(df_p, "PRESTAMOS"), file_name="Reporte_Préstamos.xlsx", use_container_width=True)
        # Filtramos los que están activos para mostrarlos
        act = df_p[df_p["Estado"]=="ACTIVO"]
        for idx, row in act.iterrows():
            with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}"):
                c1, c2 = st.columns(2)
                with c1: st.metric("CUOTA", f"${row['Cuota_Mensual']}"); st.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
                with c2:
                    if st.button("✅ REGISTRAR COBRO CUOTA", key=f"pay_{row['ID']}"):
                        df_p.at[idx, "Pagos_Realizados"] += 1
                        df_p.at[idx, "Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                        if df_p.at[idx, "Pagos_Realizados"] >= row["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                        conn.update(worksheet="Prestamos", data=df_p); st.rerun()

# --- 2. COOPERATIVA ---
elif sec == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    v_x = st.number_input("💵 VALOR CUOTA FIJA:", value=10.0)
    
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo_c"): st.session_state.show_form_c = True
    if st.session_state.get('show_form_c'):
        with st.form("form_c"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("💾 AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True))
                st.session_state.show_form_c = False; st.rerun()

    if not df_s.empty:
        st.download_button("📊 EXCEL GENERAL", data=generar_excel_grupal(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
        for idx, row in df_s.iterrows():
            with st.expander(f"👤 {row['Nombre'].upper()} | ACUMULADO: ${row['Saldo_Total_Aportado']}"):
                m = st.number_input("Monto:", value=v_x, key=f"val_{row['ID']}")
                if st.button("✅ REGISTRAR PAGO", key=f"btn_{row['ID']}"):
                    df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                    conn.update(worksheet="Cooperativa", data=df_s); st.rerun()

# --- 3. AYUDAS ECONÓMICAS ---
elif sec == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 APORTE ESTÁNDAR:", value=5.0)
    
    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo_a"): st.session_state.show_form_a = True
    if st.session_state.get('show_form_a'):
        with st.form("form_a"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("💾 AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True))
                st.session_state.show_form_a = False; st.rerun()

    if not df_a.empty:
        st.download_button("📊 EXCEL GENERAL", data=generar_excel_grupal(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
        for idx, row in df_a.iterrows():
            with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
                c1, c2 = st.columns(2)
                with c1:
                    m_a = st.number_input("Monto:", value=v_y, key=f"ay_{row['ID']}")
                    col_b1, col_b2 = st.columns(2)
                    if col_b1.button("✅ APORTE", key=f"ba1_{row['ID']}"):
                        df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m_a
                        conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
                    if col_b2.button("🔴 EGRESO", key=f"ba2_{row['ID']}"):
                        df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) - m_a
                        conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
                with c2:
                    st.download_button("📊 MI EXCEL", data=generar_excel_personal(row, df_ah, "AYUDAS"), file_name=f"Ayuda_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
