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

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="CONTROL DE PRESTAMOS", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y BOTONES LLAMATIVOS
st.markdown("""
    <style>
    button[data-baseweb="tab"] { font-size: 30px !important; font-weight: 700 !important; padding: 1.5rem !important; }
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 600 !important; }
    
    /* BOTONES ACCIÓN (Verde) */
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 30px !important; font-weight: 900 !important; height: 6rem !important; 
        border-radius: 15px !important; background-color: #28a745 !important; color: white !important; 
        border: none !important; box-shadow: 0px 5px 15px rgba(0,0,0,0.3) !important;
    }

    /* BOTÓN DISCRETO (Eliminar) */
    .stButton>button[kind="secondary"] { 
        font-size: 18px !important; background-color: transparent !important; 
        color: #dc3545 !important; border: 1px solid #dc3545 !important; opacity: 0.6;
    }

    /* Métricas Gigantes */
    [data-testid="stMetricValue"] { font-size: 70px !important; color: #007bff !important; font-weight: 800 !important; }
    [data-testid="stMetricLabel"] { font-size: 26px !important; font-weight: bold !important; }
    .stDataFrame { font-size: 24px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_prestamos():
    try:
        df = conn.read(worksheet="Prestamos", ttl="0") 
        if not df.empty:
            df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df
    except: return pd.DataFrame(columns=["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"])

def cargar_historial_pagos():
    try:
        df = conn.read(worksheet="Pagos", ttl="0")
        if not df.empty: df["ID_Prestamo"] = df["ID_Prestamo"].astype(str)
        return df
    except: return pd.DataFrame(columns=["ID_Prestamo", "Fecha_Pago", "Cuotas_Pagadas", "Monto_Pagado", "URL_Comprobante"])

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

# 5. GENERADOR DE EXCEL CON ENCABEZADO Y TABLA SUBRAYADA
def generar_excel_estado_cuenta(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # PESTAÑA: ESTADO DE CUENTA
        # Crear DataFrame vacío para estructurar el diseño manualmente
        ws_name = "ESTADO DE CUENTA"
        pd.DataFrame().to_excel(writer, sheet_name=ws_name)
        workbook = writer.book
        ws = workbook[ws_name]

        # Estilos de diseño
        font_titulo = Font(bold=True, size=14, color="1F4E78")
        font_header = Font(bold=True, color="FFFFFF")
        fill_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        font_verde = Font(color="006100", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # --- 1. ENCABEZADO ---
        ws["A1"] = "ESTADO DE CUENTA DETALLADO"
        ws["A1"].font = Font(bold=True, size=16)
        
        ws["A3"] = "CLIENTE:"; ws["B3"] = datos_c['Nombre'].upper()
        ws["A4"] = "CÉDULA:"; ws["B4"] = datos_c['Cedula']
        ws["A5"] = "FECHA PRESTAMO:"; ws["B5"] = datos_c['Fecha']
        ws["A6"] = "ID CRÉDITO:"; ws["B6"] = datos_c['ID']

        ws["D3"] = "MONTO TOTAL:"; ws["E3"] = f"${datos_c['Monto_Inicial']}"
        ws["D4"] = "TASA INTERÉS:"; ws["E4"] = f"{datos_c['Tasa']}% Anual"
        ws["D5"] = "CUOTA MENSUAL:"; ws["E5"] = f"${datos_c['Cuota_Mensual']}"
        ws["D6"] = "SALDO ACTUAL:"; ws["E6"] = f"${datos_c['Saldo_Restante']}"

        for row in range(3, 7):
            ws[f"A{row}"].font = Font(bold=True)
            ws[f"D{row}"].font = Font(bold=True)

        # --- 2. TABLA DE AMORTIZACIÓN ---
        ws["A8"] = "PLAN DE PAGOS Y SEGUIMIENTO"
        ws["A8"].font = font_titulo
        
        headers = ["N° Cuota", "Descripción", "Valor Cuota", "Estado de Pago"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=9, column=col_num)
            cell.value = header
            cell.fill = fill_azul
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center")

        pagados = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            row_num = 9 + i
            ws.cell(row=row_num, column=1, value=i)
            ws.cell(row=row_num, column=2, value=f"Cuota correspondiente al mes {i}")
            ws.cell(row=row_num, column=3, value=datos_c['Cuota_Mensual'])
            estado = "PAGADO" if i <= pagados else "PENDIENTE"
            ws.cell(row=row_num, column=4, value=estado)

            # Subrayado automático en verde si está pagado
            if i <= pagados:
                for col in range(1, 5):
                    ws.cell(row=row_num, column=col).fill = fill_verde
                    ws.cell(row=row_num, column=col).font = font_verde

        # Ajustar ancho de columnas
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 15

        # Pestaña de historial de fotos
        if not historial_c.empty:
            hist_ws = workbook.create_sheet("HISTORIAL DE ABONOS")
            for r in dataframe_to_rows(historial_c, index=False, header=True):
                hist_ws.append(r)
                
    return output.getvalue()

from openpyxl.utils.dataframe import dataframe_to_rows

# --- INTERFAZ ---
st.title("🏦 CONTROL DE PRESTAMOS")
df_p = cargar_prestamos(); df_h = cargar_historial_pagos()
t1, t2, t3, t4 = st.tabs(["➕ NUEVO", "📋 LISTA", "🔍 REPORTE", "💵 PAGAR"])

with t1:
    with st.form("f"):
        nom = st.text_input("👤 NOMBRE DEL CLIENTE:"); ced = st.text_input("🪪 CÉDULA:")
        c1, c2, c3 = st.columns(3)
        mon = c1.number_input("💵 MONTO ($):", value=500.0); mes = c2.number_input("📅 MESES:", value=12); tas = c3.number_input("📈 TASA %:", value=15.0)
        if st.form_submit_button("💾 GUARDAR PRÉSTAMO", use_container_width=True, type="primary"):
            if nom and ced:
                tm = (tas/100)/12
                cuo = mon * (tm * (1+tm)**mes) / ((1+tm)**mes - 1) if tm > 0 else mon/mes
                nuevo = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nom, "Cedula": str(ced), "Monto_Inicial": mon, "Saldo_Restante": mon, "Cuota_Mensual": round(cuo, 2), "Meses_Totales": mes, "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tas}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                st.balloons(); time.sleep(1); st.rerun()

with t2:
    st.dataframe(df_p[df_p["Estado"] == "ACTIVO"], use_container_width=True, height=500)

with t3:
    if not df_p.empty:
        sel = st.selectbox("SELECCIONE CLIENTE:", df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")")
        id_s = sel.split("(")[1].replace(")", ""); dat = df_p[df_p["ID"] == id_s].iloc[0]; his = df_h[df_h["ID_Prestamo"] == id_s]
        c1, c2, c3 = st.columns(3)
        c1.metric("DEUDA", f"${dat['Saldo_Restante']}"); c2.metric("PROGRESO", f"{dat['Pagos_Realizados']}/{dat['Meses_Totales']}"); c3.metric("CUOTA", f"${dat['Cuota_Mensual']}")
        st.dataframe(his, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("📸 VER FOTO")})
        st.download_button(f"📥 GENERAR REPORTE PARA {dat['Nombre'].upper()}", data=generar_excel_estado_cuenta(dat, his), file_name=f"ESTADO_CUENTA_{dat['Nombre'].replace(' ','_')}.xlsx", use_container_width=True)
        if st.button("Eliminar este registro", type="secondary"):
            conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != id_s])
            conn.update(worksheet="Pagos", data=df_h[df_h["ID_Prestamo"] != id_s])
            st.rerun()

with t4:
    act = df_p[df_p["Estado"] == "ACTIVO"]
    if not act.empty:
        ps = st.selectbox("¿QUIÉN PAGA?", act["Nombre"].astype(str) + " (" + act["ID"].astype(str) + ")")
        idp = ps.split("(")[1].replace(")", ""); cl = act[act[ "ID" ] == idp].iloc[0]
        cp1, cp2 = st.columns(2)
        ncuo = cp1.number_input("NÚMERO DE CUOTAS:", min_value=1, value=1)
        st.success(f"### TOTAL A COBRAR: ${round(cl['Cuota_Mensual'] * ncuo, 2)}")
        fot = cp2.file_uploader("📸 FOTO DEL RECIBO:", type=["jpg","png","jpeg"])
        if st.button("✅ CONFIRMAR PAGO AHORA", use_container_width=True, type="primary"):
            link = subir_a_imgbb_comprimido(fot.getvalue()) if fot else ""
            np = pd.DataFrame([{"ID_Prestamo": idp, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": ncuo, "Monto_Pagado": round(cl['Cuota_Mensual'] * ncuo, 2), "URL_Comprobante": link}])
            conn.update(worksheet="Pagos", data=pd.concat([df_h, np], ignore_index=True))
            idx = df_p[df_p["ID"] == idp].index[0]
            df_p.at[idx, "Pagos_Realizados"] += ncuo
            df_p.at[idx, "Saldo_Restante"] = round(max(0, cl["Saldo_Restante"] - (cl["Monto_Inicial"]/cl["Meses_Totales"])*ncuo), 2)
            if df_p.at[idx, "Pagos_Realizados"] >= cl["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
            conn.update(worksheet="Prestamos", data=df_p)
            st.balloons(); time.sleep(1); st.rerun()
