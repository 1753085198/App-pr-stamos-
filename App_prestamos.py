import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
import time
from datetime import datetime
import requests
import base64
from openpyxl.styles import PatternFill, Font

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 24px !important; font-weight: 700 !important; border-radius: 12px !important; }
    /* Botón Nuevo Naranja */
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

def generar_excel_pintado(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("REPORTE", 0)
        f_v, f_r = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE: {titulo}"]); ws.append(["Nombre", "Cédula", "Monto", "Estado"])
        for _, row in df.iterrows():
            m = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            ws.append([row['Nombre'], row['Cedula'], m, "AL DÍA" if m > 0 else "PENDIENTE"])
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = f_v if m > 0 else f_r
    return out.getvalue()

def generar_excel_individual(row, historial, tipo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("HISTORIAL", 0)
        ws.append([f"HISTORIAL DE {tipo} - {row['Nombre']}"]); ws.append(["Fecha", "Monto", "Recibo"])
        pagos = historial[historial['ID_Socio'] == row['ID']] if not historial.empty else pd.DataFrame()
        for _, p in pagos.iterrows(): ws.append([p['Fecha'], p['Monto'], p.get('Comprobante', 'N/A')])
        ws.append([""]); ws.append(["TOTAL ACUMULADO:", row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0))])
    return out.getvalue()

# --- MENÚ ---
with st.sidebar:
    st.markdown("# 🏦 PANEL")
    sec = st.radio("Ir a:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=1)

# --- SECCIONES ---
if sec == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    if not df_p.empty: st.download_button("📊 EXCEL GRUPAL PRÉSTAMOS", data=generar_excel_pintado(df_p, "PRESTAMOS"), file_name="Reporte_Prestamos.xlsx", use_container_width=True)
    
    for idx, row in df_p[df_p["Estado"]=="ACTIVO"].iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}"):
            c1, c2 = st.columns(2)
            with c1: st.metric("CUOTA", f"${row['Cuota_Mensual']}"); st.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
            with c2:
                with st.form(key=f"fp_{row['ID']}"):
                    if st.form_submit_button("✅ CONFIRMAR COBRO"):
                        df_p.at[idx, "Pagos_Realizados"] += 1
                        df_p.at[idx, "Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                        if df_p.at[idx, "Pagos_Realizados"] >= row["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                        conn.update(worksheet="Prestamos", data=df_p); st.rerun()

elif sec == "🤝 COOPERATIVA":
    st.title("🤝 GESTIÓN COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    v_x = st.number_input("💵 VALOR CUOTA ESTÁNDAR:", value=10.0)
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL COOP", data=generar_excel_pintado(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    
    for idx, row in df_s.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                m = st.number_input("Monto:", value=v_x, key=f"mc_{row['ID']}")
                if st.button("✅ REGISTRAR PAGO", key=f"bc_{row['ID']}"):
                    df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                    conn.update(worksheet="Cooperativa", data=df_s); st.rerun()
            with c2:
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_individual(row, df_ph, "COOP"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dlc_{row['ID']}")

elif sec == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 VALOR APORTE (X):", value=5.0)
    if not df_a.empty: st.download_button("📊 EXCEL GRUPAL AYUDAS", data=generar_excel_pintado(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    
    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo_a"): st.session_state.n_a = True
    if st.session_state.get('n_a'):
        with st.form("na"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True)); st.session_state.n_a = False; st.rerun()

    for idx, row in df_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                m_a = st.number_input("Monto:", value=v_y, key=f"ma_{row['ID']}")
                col_b1, col_b2 = st.columns(2)
                if col_b1.button("✅ APORTE", key=f"ba1_{row['ID']}"):
                    df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m_a
                    conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
                if col_b2.button("🔴 EGRESO", key=f"ba2_{row['ID']}"):
                    df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) - m_a
                    conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
            with c2:
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_individual(row, df_ah, "AYUDAS"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
