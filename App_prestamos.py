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

# 2. CSS PARA INTERFAZ DE ALTO IMPACTO
st.markdown("""
    <style>
    button[data-baseweb="tab"] { font-size: 26px !important; font-weight: 800 !important; }
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { 
        font-size: 22px !important; font-weight: 600 !important; 
    }
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 26px !important; font-weight: 900 !important; height: 5rem !important; 
        border-radius: 15px !important; background-color: #28a745 !important; color: white !important; 
    }
    .stButton>button[kind="secondary"] { 
        font-size: 16px !important; opacity: 0.6;
    }
    [data-testid="stMetricValue"] { font-size: 55px !important; font-weight: 800 !important; color: #007bff !important; }
    </style>
""", unsafe_allow_html=True)

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

def generar_excel_estado_cuenta(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ws_name = "ESTADO DE CUENTA"
        pd.DataFrame().to_excel(writer, sheet_name=ws_name)
        workbook = writer.book; ws = workbook[ws_name]
        f_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        # Encabezado
        ws["A1"] = "ESTADO DE CUENTA"; ws["A1"].font = Font(bold=True, size=14)
        ws["A3"] = "CLIENTE:"; ws["B3"] = str(datos_c['Nombre']).upper()
        ws["A4"] = "CÉDULA:"; ws["B4"] = str(datos_c['Cedula'])
        ws["D3"] = "SALDO:"; ws["E3"] = f"${datos_c['Saldo_Restante']}"
        
        headers = ["Cuota #", "Descripción", "Monto", "Estado"]
        for c_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=7, column=c_idx, value=text)
            cell.fill = f_azul; cell.font = Font(color="FFFFFF", bold=True)
            
        pag = int(datos_c['Pagos_Realizados'])
        for i in range(1, int(datos_c['Meses_Totales']) + 1):
            r_idx = 7 + i
            ws.cell(row=r_idx, column=1, value=i)
            ws.cell(row=r_idx, column=2, value=f"Cuota mes {i}")
            ws.cell(row=r_idx, column=3, value=datos_c['Cuota_Mensual'])
            est = "PAGADO" if i <= pag else "PENDIENTE"
            ws.cell(row=r_idx, column=4, value=est)
            if i <= pag:
                for c in range(1, 5): ws.cell(row=r_idx, column=c).fill = f_verde
        
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 22
    return output.getvalue()

# --- INTERFAZ ---
st.title("🏦 PANEL DE CONTROL")
df_p = cargar_prestamos(); df_h = cargar_historial_pagos()

t_gestion, t_nuevo = st.tabs(["📋 GESTIÓN DE CLIENTES", "➕ NUEVO PRÉSTAMO"])

with t_gestion:
    busqueda = st.text_input("🔍 BUSCAR CLIENTE POR NOMBRE O CÉDULA:", placeholder="Escribe aquí...")
    
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if busqueda:
        activos = activos[activos['Nombre'].str.contains(busqueda, case=False) | activos['Cedula'].str.contains(busqueda)]

    if activos.empty:
        st.warning("No se encontraron clientes activos.")
    else:
        for index, row_data in activos.iterrows():
            # CARD VISUAL DENTRO DE UN EXPANDER
            with st.expander(f"👤 {row_data['Nombre'].upper()}  |  💰 Saldo: ${row_data['Saldo_Restante']}"):
                col_a, col_b = st.columns([1, 1])
                
                with col_a:
                    st.write("### ℹ️ Información")
                    st.write(f"**Cédula:** {row_data['Cedula']}")
                    st.metric("Cuota Mensual", f"${row_data['Cuota_Mensual']}")
                    st.metric("Progreso", f"{row_data['Pagos_Realizados']}/{row_data['Meses_Totales']}")
                    
                    h_cliente = df_h[df_h["ID_Prestamo"] == row_data['ID']]
                    st.download_button(
                        f"📥 EXCEL DE {str(row_data['Nombre']).split()[0]}", 
                        data=generar_excel_estado_cuenta(row_data, h_cliente), 
                        file_name=f"Estado_{str(row_data['Nombre']).replace(' ','_')}.xlsx", 
                        key=f"btn_ex_{row_data['ID']}",
                        use_container_width=True
                    )

                with col_b:
                    st.write("### 💵 Registrar Pago")
                    with st.form(key=f"f_pago_{row_data['ID']}"):
                        n_c = st.number_input("Cuotas:", min_value=1, value=1, key=f"nc_{row_data['ID']}")
                        st.success(f"Cobrar: ${round(row_data['Cuota_Mensual'] * n_c, 2)}")
                        foto = st.file_uploader("📸 Recibo:", type=["jpg","png","jpeg"], key=f"foto_{row_data['ID']}")
                        
                        if st.form_submit_button("✅ CONFIRMAR PAGO", use_container_width=True, type="primary"):
                            with st.spinner('Procesando...'):
                                link = subir_a_imgbb_comprimido(foto.getvalue()) if foto else ""
                                np = pd.DataFrame([{"ID_Prestamo": row_data['ID'], "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": n_c, "Monto_Pagado": round(row_data['Cuota_Mensual'] * n_c, 2), "URL_Comprobante": link}])
                                conn.update(worksheet="Pagos", data=pd.concat([df_h, np], ignore_index=True))
                                
                                df_p.at[index, "Pagos_Realizados"] += n_c
                                abono_saldo = (row_data["Monto_Inicial"] / row_data["Meses_Totales"]) * n_c
                                df_p.at[index, "Saldo_Restante"] = round(max(0, row_data["Saldo_Restante"] - abono_saldo), 2)
                                if df_p.at[index, "Pagos_Realizados"] >= row_data["Meses_Totales"]: df_p.at[index, "Estado"] = "PAGADO"
                                
                                conn.update(worksheet="Prestamos", data=df_p)
                                st.balloons(); time.sleep(1); st.rerun()
                
                if st.button(f"Eliminar a {str(row_data['Nombre']).split()[0]}", type="secondary", key=f"del_{row_data['ID']}"):
                    conn.update(worksheet="Prestamos", data=df_p[df_p["ID"] != row_data['ID']])
                    conn.update(worksheet="Pagos", data=df_h[df_h["ID_Prestamo"] != row_data['ID']])
                    st.rerun()

with t_nuevo:
    with st.form("form_nuevo", clear_on_submit=True):
        st.write("### 📝 Datos del Préstamo")
        nom = st.text_input("Nombre Completo:"); cid = st.text_input("Cédula:")
        c1, c2, c3 = st.columns(3)
        mon = c1.number_input("Monto ($)", value=500.0); tas = c3.number_input("Tasa % Anual", value=15.0); pla = c2.number_input("Meses", value=12)
        
        if st.form_submit_button("💾 CREAR PRÉSTAMO", use_container_width=True, type="primary"):
            if nom and cid:
                tm = (tas/100)/12
                cuota = mon * (tm * (1+tm)**pla) / ((1+tm)**pla - 1) if tm > 0 else mon/pla
                nuevo = pd.DataFrame([{"ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Nombre": nom, "Cedula": cid, "Monto_Inicial": mon, "Saldo_Restante": mon, "Cuota_Mensual": round(cuota, 2), "Meses_Totales": int(pla), "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tas}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                st.balloons(); time.sleep(1); st.rerun()
