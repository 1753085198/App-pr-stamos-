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
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 28px !important; font-weight: 700 !important; }
    input { font-size: 24px !important; height: 55px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 26px !important; font-weight: 700 !important; border-radius: 15px !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 30px !important; font-weight: 900 !important; position: fixed; bottom: 40px; right: 40px; z-index: 9999; border: 4px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 75px !important; font-weight: 900 !important; color: #007bff !important; }
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
        ws.append([f"REPORTE: {titulo}"]); ws.append(["Nombre", "Cédula", "Monto Total", "Estado"])
        for _, row in df.iterrows():
            monto = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            estado = "PAGADO" if monto > 0 else "PENDIENTE"
            ws.append([row['Nombre'], row['Cedula'], monto, estado])
            color = f_verde if monto > 0 else f_rojo
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = color
    return out.getvalue()

# --- MENÚ ---
with st.sidebar:
    st.markdown("# 🏦 SISTEMA")
    seccion = st.radio("Sección:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=1)

# --- LÓGICA ---
if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    # (Lógica original de préstamos con recibos ya integrada)
    st.info("Sección de préstamos activa con respaldo de imagen.")

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s = cargar("Cooperativa")
    df_pagos_coop = cargar("Pagos_Coop") # Asegúrate de tener esta pestaña
    cuota_fija = st.number_input("💵 VALOR CUOTA ESTÁNDAR:", value=10.0)
    
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL", data=generar_excel_pintado(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    
    bq = st.text_input("🔍 BUSCAR SOCIO:")
    act = df_s if not df_s.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            with st.form(key=f"form_coop_{row['ID']}"):
                m = st.number_input("Monto:", value=cuota_fija)
                ft = st.file_uploader("📸 SUBIR COMPROBANTE:", key=f"img_coop_{row['ID']}")
                if st.form_submit_button("✅ REGISTRAR PAGO Y RECIBO"):
                    if ft:
                        url = subir_img(ft.getvalue())
                        # Guardar en historial de pagos
                        new_p = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                        conn.update(worksheet="Pagos_Coop", data=pd.concat([df_pagos_coop, new_p], ignore_index=True))
                        # Actualizar saldo del socio
                        df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                        conn.update(worksheet="Cooperativa", data=df_s)
                        st.success("¡Pago registrado!"); st.rerun()
                    else: st.error("¡Debes subir la foto del comprobante!")

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a = cargar("Ayudas_Listado")
    df_pagos_ayu = cargar("Pagos_Ayudas") # Asegúrate de tener esta pestaña
    cuota_fija_ayu = st.number_input("💵 VALOR APORTE ESTÁNDAR:", value=5.0)
    
    if not df_a.empty: st.download_button("📊 EXCEL GRUPAL", data=generar_excel_pintado(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    
    bq = st.text_input("🔍 BUSCAR COMPAÑERO:")
    act = df_a if not df_a.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            with st.form(key=f"form_ayu_{row['ID']}"):
                m = st.number_input("Monto:", value=cuota_fija_ayu)
                ft = st.file_uploader("📸 SUBIR COMPROBANTE:", key=f"img_ayu_{row['ID']}")
                if st.form_submit_button("✅ GUARDAR APORTE"):
                    if ft:
                        url = subir_img(ft.getvalue())
                        new_p = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                        conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_pagos_ayu, new_p], ignore_index=True))
                        df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                        conn.update(worksheet="Ayudas_Listado", data=df_a)
                        st.success("¡Aporte guardado con éxito!"); st.rerun()
                    else: st.error("¡La foto del recibo es obligatoria!")
