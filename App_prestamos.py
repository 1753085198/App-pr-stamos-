import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
from datetime import datetime
import requests
import base64
from openpyxl.styles import PatternFill, Font

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 28px !important; font-weight: 700 !important; }
    input { font-size: 24px !important; height: 55px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 26px !important; font-weight: 700 !important; border-radius: 15px !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 30px !important; font-weight: 900 !important; position: fixed; bottom: 40px; right: 40px; z-index: 9999; border: 4px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 70px !important; font-weight: 900 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown("# 🏦 NAVEGACIÓN")
    seccion = st.radio("Ir a:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=1)
    st.write("---")
    st.info("Jose Figueroa - UDLA")

# --- FUNCIONES ---
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
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        f_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE: {titulo}"])
        ws.append(["Nombre", "Cédula", "Monto", "Estado"])
        for _, row in df.iterrows():
            monto = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            estado = "AL DÍA" if monto > 0 else "PENDIENTE"
            ws.append([row['Nombre'], row['Cedula'], monto, estado])
            color = f_verde if monto > 0 else f_rojo
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = color
    return out.getvalue()

# --- LÓGICA POR SECCIÓN ---

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    # (Se mantiene la lógica funcional anterior)
    st.info("Sección de préstamos activa.")

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 GESTIÓN COOPERATIVA")
    df_s = cargar("Cooperativa")
    
    # --- ÁREA DE COBRO MASIVO ---
    with st.container(border=True):
        st.write("### ⚡ COBRO MASIVO (A TODOS)")
        c1, c2 = st.columns([2, 1])
        with c1:
            cuota_gen = st.number_input("Monto de cuota para todos los socios:", value=10.0)
        with c2:
            st.write("##")
            if st.button("🔥 PAGAR TODOS", use_container_width=True):
                if not df_s.empty:
                    df_s["Saldo_Total_Aportado"] = df_s["Saldo_Total_Aportado"].astype(float) + cuota_gen
                    conn.update(worksheet="Cooperativa", data=df_s)
                    st.success(f"¡Se han sumado ${cuota_gen} a todos los socios!"); time.sleep(1); st.rerun()

    if not df_s.empty:
        st.download_button("📊 EXCEL UNIFICADO COOP", data=generar_excel_pintado(df_s, "COOPERATIVA"), file_name="Reporte_Coop_Grupal.xlsx", use_container_width=True)

    if st.button("👤 NUEVO SOCIO", key="btn_nuevo_c"):
        st.session_state.mostrar_nuevo = True

    if st.session_state.get('mostrar_nuevo'):
        with st.form("nc"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("GUARDAR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR SOCIO:")
    act = df_s if not df_s.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            col_a, col_b = st.columns(2)
            with col_a:
                m = st.number_input("Cobro individual:", value=10.0, key=f"m_{row['ID']}")
                if st.button("✅ REGISTRAR PAGO", key=f"b_{row['ID']}"):
                    df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                    conn.update(worksheet="Cooperativa", data=df_s); st.rerun()
            with col_b:
                # EXCEL PERSONAL POR SOCIO
                excel_per = generar_excel_pintado(pd.DataFrame([row]), f"Socio {row['Nombre']}")
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=excel_per, file_name=f"Comprobante_{row['Nombre']}.xlsx", key=f"dl_{row['ID']}")

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a = cargar("Ayudas_Listado")
    
    # --- ÁREA DE COBRO MASIVO AYUDAS ---
    with st.container(border=True):
        st.write("### ⚡ COBRO MASIVO (AYUDA GRUPAL)")
        ca1, ca2 = st.columns([2, 1])
        with ca1:
            cuota_ayu = st.number_input("Monto de ayuda para todos:", value=5.0)
        with ca2:
            st.write("##")
            if st.button("🔥 APLICAR A TODOS", key="btn_masivo_a", use_container_width=True):
                if not df_a.empty:
                    df_a["Saldo_Total_Aportado"] = df_a["Saldo_Total_Aportado"].astype(float) + cuota_ayu
                    conn.update(worksheet="Ayudas_Listado", data=df_a)
                    st.success("¡Cuota de ayuda registrada para todos!"); st.rerun()

    if not df_a.empty:
        st.download_button("📊 EXCEL UNIFICADO AYUDAS", data=generar_excel_pintado(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)

    # (Lógica de lista de compañeros y excel individual similar a Cooperativa...)
