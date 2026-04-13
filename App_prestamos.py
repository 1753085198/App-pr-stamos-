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
st.set_page_config(page_title="SISTEMA FINANCIERO FINAL", page_icon="🏦", layout="wide")

# 2. CSS PERSONALIZADO
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
        ws.append([""]); ws.append(["SALDO TOTAL:", row.get('Saldo_Total_Aportado', 0)])
    return out.getvalue()

# --- SECCIONES ---
with st.sidebar:
    st.markdown("# 🏦 PANEL")
    sec = st.radio("Sección:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=2)

if sec == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    # Lógica de préstamos intacta...

elif sec == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL COOP", data=generar_excel_pintado(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    # Lista de socios...

elif sec == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 VALOR APORTE (X):", value=5.0)
    
    # 1. EXCEL GRUPAL
    if not df_a.empty: 
        st.download_button("📊 DESCARGAR REPORTE GRUPAL AYUDAS", data=generar_excel_pintado(df_a, "AYUDAS"), file_name="Reporte_Grupal_Ayudas.xlsx", use_container_width=True)
    
    # 2. LISTADO Y EXCEL INDIVIDUAL
    for idx, row in df_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fa_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_y)
                    col_b1, col_b2 = st.columns(2)
                    pago = col_b1.form_submit_button("✅ APORTE")
                    retiro = col_b2.form_submit_button("🔴 EGRESO")
                    if pago or retiro:
                        m_final = m if pago else -m
                        new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m_final}])
                        conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_ah, new_h], ignore_index=True))
                        df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m_final
                        conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
            with c2:
                # 3. EXCEL INDIVIDUAL
                st.write("### 📄 MI REPORTE")
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_individual(row, df_ah, "AYUDAS"), file_name=f"Ayuda_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
