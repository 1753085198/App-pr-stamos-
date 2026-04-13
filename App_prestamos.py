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
from openpyxl.styles import PatternFill, Font, Alignment

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="CONTROL DE PRESTAMOS", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    button[data-baseweb="tab"] { font-size: 30px !important; font-weight: 700 !important; padding: 1.5rem !important; }
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 600 !important; }
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 30px !important; font-weight: 900 !important; height: 6rem !important; 
        border-radius: 15px !important; background-color: #28a745 !important; color: white !important; 
        border: none !important; box-shadow: 0px 5px 15px rgba(0,0,0,0.3) !important;
    }
    .stButton>button[kind="secondary"] { 
        font-size: 18px !important; background-color: transparent !important; 
        color: #dc3545 !important; border: 1px solid #dc3545 !important; opacity: 0.6;
    }
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

# 5. GENERADOR DE EXCEL CORREGIDO (Sin PIL.Image.open en la lógica de celdas)
def generar_excel_con_estilo(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja de Resumen
        resumen = pd.DataFrame([datos_c])
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        
        # Hoja de Historial
        if not historial_c.empty:
            historial_c.to_excel(writer, sheet_name="PAGOS_DETALLE", index=False)
        
        workbook = writer.book
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        center_align = Alignment(horizontal="center")
        
        for sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
            # Estilizar encabezados
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            # Auto-ajustar ancho de columnas
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[column].width = max_length + 4
    return output.getvalue()

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
        st.download_button(f"📥 DESCARGAR EXCEL DE {dat['Nombre'].upper()}", data=generar_excel_con_estilo(dat, his), file_name=f"REPORTE_{dat['Nombre'].replace(' ','_')}.xlsx", use_container_width=True)
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
