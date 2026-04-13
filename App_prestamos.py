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
st.set_page_config(page_title="CONTROL DE PRESTAMOS", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y PROFESIONAL
st.markdown("""
    <style>
    /* Pestañas Gigantes */
    button[data-baseweb="tab"] { font-size: 28px !important; font-weight: 700 !important; padding: 1.2rem !important; }
    
    /* Etiquetas de texto y campos */
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { 
        font-size: 24px !important; font-weight: 600 !important; 
    }
    
    /* BOTONES DE ACCIÓN (Verde Dinero) */
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 30px !important; font-weight: 900 !important; height: 6rem !important; 
        border-radius: 18px !important; background-color: #28a745 !important; color: white !important; 
        box-shadow: 0px 6px 15px rgba(0,0,0,0.2) !important;
    }

    /* Botón Eliminar (Discreto y Elegante) */
    .stButton>button[kind="secondary"] { 
        font-size: 18px !important; background-color: transparent !important; 
        color: #dc3545 !important; border: 1px solid #dc3545 !important; opacity: 0.5;
    }
    
    /* Métricas de Saldo */
    [data-testid="stMetricValue"] { font-size: 65px !important; font-weight: 800 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

# 3. CONEXIÓN A GOOGLE SHEETS
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_prestamos():
    try:
        df = conn.read(worksheet="Prestamos", ttl="0") 
        if not df.empty:
            df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
            df["Nombre"] = df["Nombre"].astype(str)
        return df
    except: return pd.DataFrame(columns=["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"])

def cargar_historial_pagos():
    try:
        df = conn.read(worksheet="Pagos", ttl="0")
        if not df.empty: df["ID_Prestamo"] = df["ID_Prestamo"].astype(str)
        return df
    except: return pd.DataFrame(columns=["ID_Prestamo", "Fecha_Pago", "Cuotas_Pagadas", "Monto_Pagado", "URL_Comprobante"])

# 4. MOTOR TURBO: COMPRESIÓN DE IMAGEN
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

# 5. GENERADOR DE EXCEL PROFESIONAL
def generar_excel_estado_cuenta(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ws_name = "ESTADO DE CUENTA"
        pd.DataFrame().to_excel(writer, sheet_name=ws_name)
        workbook = writer.book; ws = workbook[ws_name]
        
        # Estilos
        fill_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True); fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        font_verde = Font(color="006100", bold=True)

        # Encabezado de identidad
        ws["A1"] = "ESTADO DE CUENTA"; ws["A1"].font = Font(bold=True, size=16)
        ws["A3"] = "CLIENTE:"; ws["B3"] = datos_c['Nombre'].upper()
        ws["A4"] = "CÉDULA:"; ws["B4"] = datos_c['Cedula']
        ws["D3"] = "MONTO:"; ws["E3"] = f"${datos_c['Monto_Inicial']}"
        ws["D4"] = "SALDO ACTUAL:"; ws["E4"] = f"${datos_c['Saldo_Restante']}"

        # Tabla de Amortización
        headers = ["Cuota #", "Detalle", "Valor Cuota", "Estado"]
        for col_num, h in enumerate(headers, 1):
            cell = ws.cell(row=7, column=col_num); cell.value = h; cell.fill = fill_azul; cell.font = font_header; cell.alignment = Alignment(horizontal="center")
        
        pag = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            r = 7 + i
            ws.cell(row=r, column=1, value=i); ws.cell(row=r, column=2, value=f"Cuota mes {i}"); ws.cell(row=r, column=3, value=datos_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws.cell(row=r, column=4, value=est)
            if i <= pag:
                for col in range(1, 5): ws.cell(row=r, column=col).fill = fill_verde; ws.cell(row=r, column=col).font = font_verde
        
        ws.column_dimensions['B'].width = 30
        if not historial_c.empty:
            h_ws = workbook.create_sheet("HISTORIAL DE ABONOS")
            for r in dataframe_to_rows(historial_c, index=False, header=True): h_ws.append(r)
    return output.getvalue()

# --- INTERFAZ ---
st.title("🏦 CONTROL DE PRESTAMOS")

df_p = cargar_prestamos(); df_h = cargar_historial_pagos()

t1, t2, t3, t4 = st.tabs(["💵 REGISTRAR PAGO", "➕ NUEVO CLIENTE", "📋 LISTA", "📊 REPORTE"])

# TAB 1: PAGAR
with t1:
    act = df_p[df_p["Estado"] == "ACTIVO"]
    if not act.empty:
        if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
        with st.form(key=f"form_pago_{st.session_state.pago_key}", clear_on_submit=True):
            ps = st.selectbox("🎯 SELECCIONA CLIENTE:", act["Nombre"].astype(str) + " (" + act["ID"].astype(str) + ")")
            idp = ps.split("(")[1].replace(")", ""); cl = act[act["ID"] == idp].iloc[0]
            
            st.info(f"### Saldo Actual: ${cl['Saldo_Restante']}")
            
            c1, c2 = st.columns(2)
            n_c = c1.number_input("¿Cuántas cuotas cobra?", min_value=1, value=1)
            st.success(f"## TOTAL: ${round(cl['Cuota_Mensual'] * n_c, 2)}")
            fot = c2.file_uploader("📸 Foto del Recibo:", type=["jpg","png","jpeg"])
            
            if st.form_submit_button("✅ CONFIRMAR PAGO", use_container_width=True, type="primary"):
                with st.spinner('Procesando...'):
                    link = subir_a_imgbb_comprimido(fot.getvalue()) if fot else ""
                    # Registro en historial
                    np = pd.DataFrame([{"ID_Prestamo": idp, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": n_c, "Monto_Pagado": round(cl['Cuota_Mensual'] * n_c, 2), "URL_Comprobante": link}])
                    conn.update(worksheet="Pagos", data=pd.concat([df_h, np], ignore_index=True))
                    # Actualización de saldo
                    idx = df_p[df_p["ID"] == idp].index[0]
                    df_p.at[idx, "Pagos_Realizados"] += n_c
                    df_p.at[idx, "Saldo_Restante"] = round(max(0, cl["Saldo_Restante"] - (cl["Monto_Inicial"]/cl["Meses_Totales"])*n_c), 2)
                    if df_p.at[idx, "Pagos_Realizados"] >= cl["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                    conn.update(worksheet="Prestamos", data=df_p)
                    
                    st.session_state.pago_key += 1 # Reset visual
                    st.balloons(); time.sleep(1); st.rerun()

# TAB 2: NUEVO CLIENTE
with t2:
    with st.form("f_nuevo", clear_on_submit=True):
        n = st.text_input("👤 Nombre:"); c = st.text_input("🪪 Cédula:")
        col1, col2, col3 = st.columns(3)
        m = col1.number_input("Monto ($)", value=500.0); ms = col2.number_input("Plazo (Meses)", value=12); ts = col3.number_input("Tasa %", value=15.0)
        if st.form_submit_button("💾 GUARDAR NUEVO PRESTAMO", use_container_width=True, type="primary"):
            tm = (ts/100)/12
            cuo = m * (tm * (1+tm)**ms) / ((1+tm)**ms - 1) if tm > 0 else m/ms
            nuevo = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": n, "Cedula": c, "Monto_Inicial": m, "Saldo_Restante": m, "Cuota_Mensual": round(cuo, 2), "Meses_Totales": int(ms), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": ts}])
            conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
            st.balloons(); time.sleep(1); st.rerun()

# TAB 3: LISTA
with t3:
    st.dataframe(df_p[df_p["Estado"] == "ACTIVO"], use_container_width=True, height=500)

# TAB 4: REPORTES Y BORRADO
with t4:
    if not df_p.empty:
        op = df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")"
        sel = st.selectbox("🔍 BUSCAR CLIENTE:", op)
        ids = sel.split("(")[1].replace(")", ""); d = df_p[df_p["ID"] == ids].iloc[0]; h = df_h[df_h["ID_Prestamo"] == ids]
        
        c1, c2, c3 = st.columns(3)
        c1.metric("DEUDA ACTUAL", f"${d['Saldo_Restante']}")
        c2.metric("PROGRESO", f"{d['Pagos_Realizados']}/{d['Meses_Totales']}")
        c3.metric("CUOTA", f"${d['Cuota_Mensual']}")
        
        st.write("---")
        st.download_button(f"📥 DESCARGAR ESTADO DE CUENTA: {d['Nombre']}", data=generar_excel_estado_cuenta(d, h), file_name=f"Reporte_{d['Nombre']}.xlsx", use_container_width=True)
        
        st.write("### Historial de Fotos")
        st.dataframe(h, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("📸 Ver Recibo")})
        
        st.write("---")
        if st.button(f"Eliminar a {d['Nombre']} del sistema", type="secondary"):
            conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != ids])
            conn.update(worksheet="Pagos", data=df_h[df_h["ID_Prestamo"] != ids])
            st.rerun()
