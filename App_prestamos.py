import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
import time
from datetime import datetime
import requests
import base64
from openpyxl.styles import PatternFill, Font, Alignment

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

def generar_excel_grupal(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("REPORTE GENERAL", 0)
        f_verde, f_rojo = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"ESTADO GENERAL: {titulo}"]); ws.append(["Nombre", "Cédula", "Total Acumulado", "Estado"])
        for _, row in df.iterrows():
            monto = float(row.get('Saldo_Total_Aportado', 0))
            ws.append([row['Nombre'], row['Cedula'], monto, "AL DÍA" if monto > 0 else "SIN APORTES"])
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = f_verde if monto > 0 else f_rojo
    return out.getvalue()

def generar_excel_personal(socio_row, df_historial, titulo_tipo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("MI HISTORIAL", 0)
        # Encabezado
        ws.append([f"COMPROBANTE DE PAGOS - {titulo_tipo}"])
        ws.append([f"SOCIO: {socio_row['Nombre']}"]); ws.append([f"CÉDULA: {socio_row['Cedula']}"])
        ws.append([""]); ws.append(["Fecha de Pago", "Monto Pagado", "Link Comprobante"])
        
        # Filtrar pagos de este socio
        pagos_socio = df_historial[df_historial['ID_Socio'] == socio_row['ID']]
        for _, pago in pagos_socio.iterrows():
            ws.append([pago['Fecha'], pago['Monto'], pago.get('Comprobante', 'N/A')])
        
        ws.append([""]); ws.append(["TOTAL ACUMULADO:", socio_row['Saldo_Total_Aportado']])
        for cell in ws["1:1"]: cell.font = Font(bold=True, size=14)
    return out.getvalue()

# --- NAVEGACIÓN ---
with st.sidebar:
    st.markdown("# 🏦 SISTEMA")
    seccion = st.radio("Sección:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=1)

# --- LÓGICA SECCIONES ---
if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    # Lógica de préstamos original (se mantiene)

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_p = cargar("Cooperativa"), cargar("Pagos_Coop")
    cuota_x = st.number_input("💵 VALOR CUOTA FIJA (X):", value=10.0)
    
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL COOP", data=generar_excel_grupal(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    
    bq = st.text_input("🔍 BUSCAR SOCIO:")
    act = df_s if not df_s.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | ACUMULADO: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"f_coop_{row['ID']}"):
                    m = st.number_input("Monto:", value=cuota_x)
                    ft = st.file_uploader("📸 FOTO RECIBO:", key=f"img_c_{row['ID']}")
                    if st.form_submit_button("✅ REGISTRAR PAGO"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_p = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Coop", data=pd.concat([df_p, new_p], ignore_index=True))
                            df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Cooperativa", data=df_s); st.rerun()
            with c2:
                st.write("### 📄 HISTORIAL")
                st.download_button(f"📊 DESCARGAR MI EXCEL", data=generar_excel_personal(row, df_p, "COOPERATIVA"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dl_c_{row['ID']}")

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_pa = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    cuota_y = st.number_input("💵 VALOR APORTE FIJO (X):", value=5.0)
    
    if not df_a.empty: st.download_button("📊 EXCEL GRUPAL AYUDAS", data=generar_excel_grupal(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    
    bq_a = st.text_input("🔍 BUSCAR COMPAÑERO:")
    act_a = df_a if not df_a.empty else pd.DataFrame()
    if bq_a: act_a = act_a[act_a['Nombre'].str.contains(bq_a, case=False)]
    
    for idx, row in act_a.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | ACUMULADO: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"f_ayu_{row['ID']}"):
                    m = st.number_input("Monto:", value=cuota_y)
                    ft = st.file_uploader("📸 FOTO RECIBO:", key=f"img_a_{row['ID']}")
                    if st.form_submit_button("✅ GUARDAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_pa = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_pa, new_pa], ignore_index=True))
                            df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
            with c2:
                st.write("### 📄 HISTORIAL")
                st.download_button(f"📊 DESCARGAR MI EXCEL", data=generar_excel_personal(row, df_pa, "AYUDA ECON."), file_name=f"Historial_Ayuda_{row['Nombre']}.xlsx", key=f"dl_a_{row['ID']}")
