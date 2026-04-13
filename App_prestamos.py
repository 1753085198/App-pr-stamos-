import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import time
import io
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
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { 
        font-size: 32px !important; font-weight: 700 !important; line-height: 1.5 !important;
    }
    input { font-size: 30px !important; height: 70px !important; }
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 40px !important; font-weight: 900 !important; height: 8rem !important; 
        border-radius: 25px !important; background-color: #28a745 !important; color: white !important; 
        box-shadow: 0px 10px 20px rgba(0,0,0,0.4) !important;
    }
    .stButton>button[kind="secondary"] { font-size: 25px !important; height: 5rem !important; background-color: #ff4b4b !important; color: white !important; }
    [data-testid="stMetricValue"] { font-size: 90px !important; font-weight: 900 !important; color: #007bff !important; }
    .streamlit-expanderHeader { font-size: 38px !important; font-weight: 800 !important; padding: 25px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNCIONES DE CARGA ---
def cargar_prestamos():
    cols = ["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"]
    try:
        df = conn.read(worksheet="Prestamos", ttl="0") 
        if df is None or df.empty: return pd.DataFrame(columns=cols)
        for c in cols: 
            if c not in df.columns: df[c] = ""
        return df
    except: return pd.DataFrame(columns=cols)

def cargar_historial_pagos():
    cols = ["ID_Prestamo", "Fecha_Pago", "Cuotas_Pagadas", "Monto_Pagado", "URL_Comprobante"]
    try:
        df = conn.read(worksheet="Pagos", ttl="0")
        if df is None or df.empty: return pd.DataFrame(columns=cols)
        return df
    except: return pd.DataFrame(columns=cols)

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

# --- GENERADOR DE EXCEL PREMIUM ---
def generar_excel_de_gala(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # HOJA 1: ESTADO DE CUENTA (AMORTIZACIÓN)
        ws1 = writer.book.create_sheet("ESTADO DE CUENTA", 0)
        
        # Estilos
        azul_oscuro = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        verde_claro = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        font_blanca = Font(color="FFFFFF", bold=True)
        font_verde = Font(color="006100", bold=True)
        
        # Encabezado de Identidad
        ws1["A1"] = "REPORTE DETALLADO DE PRÉSTAMO"; ws1["A1"].font = Font(bold=True, size=16)
        ws1["A3"] = "CLIENTE:"; ws1["B3"] = str(datos_c['Nombre']).upper()
        ws1["A4"] = "CÉDULA:"; ws1["B4"] = str(datos_c['Cedula']).replace(".0", "")
        ws1["D3"] = "VALOR TOTAL:"; ws1["E3"] = f"${datos_c['Monto_Inicial']}"
        ws1["D4"] = "TASA:"; ws1["E4"] = f"{datos_c['Tasa']}% Anual"
        ws1["D5"] = "SALDO ACTUAL:"; ws1["E5"] = f"${datos_c['Saldo_Restante']}"
        
        # Tabla de Amortización
        headers_plan = ["N° Cuota", "Descripción", "Valor Cuota", "Estado"]
        for c_idx, val in enumerate(headers_plan, 1):
            cell = ws1.cell(row=7, column=c_idx, value=val)
            cell.fill = azul_oscuro; cell.font = font_blanca; cell.alignment = Alignment(horizontal="center")
            
        pagados = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            r = 7 + i
            ws1.cell(row=r, column=1, value=i)
            ws1.cell(row=r, column=2, value=f"Cuota mes {i}")
            ws1.cell(row=r, column=3, value=datos_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pagados else "PENDIENTE"
            ws1.cell(row=r, column=4, value=est)
            if i <= pagados:
                for col in range(1, 5):
                    ws1.cell(row=r, column=col).fill = verde_claro; ws1.cell(row=r, column=col).font = font_verde
        
        # HOJA 2: HISTORIAL DE PAGOS (BITÁCORA)
        ws2 = writer.book.create_sheet("HISTORIAL DE PAGOS", 1)
        headers_pagos = ["Fecha Pago", "Cuotas Pagadas", "Monto Pagado", "Link Comprobante"]
        for c_idx, val in enumerate(headers_pagos, 1):
            cell = ws2.cell(row=1, column=c_idx, value=val)
            cell.fill = azul_oscuro; cell.font = font_blanca
            
        if not historial_c.empty:
            for r_idx, row in enumerate(historial_c.values, 2):
                ws2.cell(row=r_idx, column=1, value=row[1]) # Fecha
                ws2.cell(row=r_idx, column=2, value=row[2]) # Cuotas
                ws2.cell(row=r_idx, column=3, value=row[3]) # Monto
                link_cell = ws2.cell(row=r_idx, column=4, value=row[4]) # URL
                link_cell.font = Font(color="0000FF", underline="single")
                
        # Ajuste de Columnas
        for ws in [ws1, ws2]:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 30
                
    return output.getvalue()

# --- INTERFAZ ---
st.title("🏦 PANEL DE CONTROL")
df_p = cargar_prestamos(); df_h = cargar_historial_pagos()
t_gestion, t_nuevo = st.tabs(["📋 GESTIÓN", "➕ NUEVO"])

if 'foto_key' not in st.session_state: st.session_state.foto_key = 0

with t_gestion:
    busq = st.text_input("🔍 BUSCADOR:", placeholder="NOMBRE O CÉDULA...")
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if busq: activos = activos[activos['Nombre'].str.contains(busq, case=False) | activos['Cedula'].str.contains(busq)]

    for idx, r_data in activos.iterrows():
        with st.expander(f"👤 {r_data['Nombre'].upper()}  |  💰 SALDO: ${r_data['Saldo_Restante']}"):
            col1, col2 = st.columns(2)
            with col1:
                st.write("### ℹ️ DETALLES")
                st.metric("CUOTA", f"${r_data['Cuota_Mensual']}")
                st.metric("PAGOS", f"{r_data['Pagos_Realizados']}/{r_data['Meses_Totales']}")
                h_c = df_h[df_h["ID_Prestamo"] == r_data['ID']]
                # EXCEL PREMIUM
                st.download_button(f"📥 EXCEL ESTADO DE CUENTA", data=generar_excel_de_gala(r_data, h_c), file_name=f"Reporte_{r_data['Nombre']}.xlsx", key=f"ex_{r_data['ID']}", use_container_width=True)

            with col2:
                st.write("### 💵 PAGAR")
                # FORMULARIO CON AUTO-RESET
                with st.form(key=f"f_{r_data['ID']}_{st.session_state.foto_key}", clear_on_submit=True):
                    n = st.number_input("Cuotas:", min_value=1, value=1, key=f"n_{r_data['ID']}")
                    st.success(f"TOTAL: ${round(r_data['Cuota_Mensual'] * n, 2)}")
                    f = st.file_uploader("📸 SUBIR RECIBO:", type=["jpg","png","jpeg"], key=f"foto_{r_data['ID']}_{st.session_state.foto_key}")
                    
                    if st.form_submit_button("✅ CONFIRMAR", use_container_width=True, type="primary"):
                        with st.spinner('Procesando...'):
                            url = subir_a_imgbb_comprimido(f.getvalue()) if f else ""
                            new_p = pd.DataFrame([{"ID_Prestamo": r_data['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": n, "Monto_Pagado": round(r_data['Cuota_Mensual']*n, 2), "URL_Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_p], ignore_index=True))
                            df_p.at[idx, "Pagos_Realizados"] += n
                            df_p.at[idx, "Saldo_Restante"] = round(max(0, r_data["Saldo_Restante"] - (r_data["Monto_Inicial"]/r_data["Meses_Totales"])*n), 2)
                            if df_p.at[idx, "Pagos_Realizados"] >= r_data["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                            conn.update(worksheet="Prestamos", data=df_p)
                            
                            st.session_state.foto_key += 1 # RESET DE IMAGEN
                            st.balloons(); time.sleep(1); st.rerun()
            
            if st.button(f"🗑️ BORRAR CLIENTE", key=f"del_{r_data['ID']}", type="secondary", use_container_width=True):
                conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != r_data['ID']])
                conn.update(worksheet="Pagos", data=df_h[df_h["ID_Prestamo"] != r_data['ID']])
                st.warning("Borrado."); time.sleep(1); st.rerun()

with t_nuevo:
    with st.form("nuevo", clear_on_submit=True):
        st.write("### 📝 NUEVO PRÉSTAMO")
        nom = st.text_input("Nombre:"); cid = st.text_input("Cédula:")
        c1, c2, c3 = st.columns(3)
        m = c1.number_input("Monto:", value=500.0); t = c3.number_input("Tasa %:", value=15.0); p = c2.number_input("Meses:", value=12)
        if st.form_submit_button("💾 GUARDAR", use_container_width=True, type="primary"):
            tm = (t/100)/12
            cuo = m * (tm * (1+tm)**p) / ((1+tm)**p - 1) if tm > 0 else m/p
            new_r = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nom, "Cedula": cid, "Monto_Inicial": m, "Saldo_Restante": m, "Cuota_Mensual": round(cuo, 2), "Meses_Totales": int(p), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": t}])
            conn.update(worksheet="Prestamos", data=pd.concat([df_p, new_r], ignore_index=True))
            st.balloons(); time.sleep(1); st.rerun()
