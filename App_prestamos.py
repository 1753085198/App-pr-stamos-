import streamlit as st
from streamlit_gsheets import GSheetsConnection
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError # <-- NUEVA LIBRERÍA PARA LEER EL ERROR
import pandas as pd
import uuid
import io
import time
from datetime import datetime

# --- CONFIGURACIÓN DE SEGURIDAD Y NUBE ---
ID_CARPETA_DRIVE = "1TwmKxziawFk5qWTCy1De12adpIxEnOha"

def obtener_servicio_drive():
    info = st.secrets["connections"]["gsheets"]
    SCOPES = ['https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def subir_a_drive(archivo_bytes, nombre_archivo, mime_type):
    try:
        servicio = obtener_servicio_drive()
        metadatos = {'name': nombre_archivo, 'parents': [ID_CARPETA_DRIVE]}
        media = MediaIoBaseUpload(io.BytesIO(archivo_bytes), mimetype=mime_type)
        
        # Agregamos supportsAllDrives por si tu carpeta está en una unidad compartida
        archivo = servicio.files().create(
            body=metadatos, 
            media_body=media, 
            fields='id',
            supportsAllDrives=True 
        ).execute()
        
        return archivo.get('id')
    except HttpError as error:
        # AQUÍ OBLIGAMOS A MOSTRAR LA VERDAD
        st.error(f"🔍 GOOGLE DICE EXACTAMENTE ESTO: {error}")
        st.stop()

# --- CONFIGURACIÓN DE INTERFAZ PANORÁMICA ---
st.set_page_config(page_title="Sistema Pro de Préstamos", page_icon="🏦", layout="wide")

# Conexión a Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_datos():
    df = conn.read(ttl="0")
    if not df.empty:
        df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
        df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
    return df

# --- MATEMÁTICA FINANCIERA ---
def calcular_cuota(capital, meses, tasa_anual):
    tasa_mensual = (tasa_anual / 100) / 12
    if tasa_mensual > 0:
        return capital * (tasa_mensual * (1 + tasa_mensual)**meses) / ((1 + tasa_mensual)**meses - 1)
    return capital / meses

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 Sistema Integral de Préstamos y Cobranza")
st.write("Bienvenido al centro de mando. Gestiona tu cartera con total seguridad en la nube.")

try:
    df_prestamos = cargar_datos()
except Exception as e:
    st.error("Esperando conexión con Google Sheets. Verifica tus Secrets.")
    st.stop()

if df_prestamos.empty:
    columnas = ["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"]
    df_prestamos = pd.DataFrame(columns=columnas)

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo Préstamo", "📋 Lista de Clientes", "🔍 Auditoría", "💵 Registrar Pagos"])

with tab1:
    st.subheader("Registrar nuevo crédito")
    with st.form("form_nuevo"):
        nombre = st.text_input("Nombre Completo:")
        cedula = st.text_input("Cédula de Identidad:")
        col_a, col_b, col_c = st.columns(3)
        monto = col_a.number_input("Monto del Préstamo ($):", min_value=0.0, value=500.0)
        plazo = col_b.number_input("Plazo (Meses):", min_value=1, value=12)
        tasa = col_c.number_input("Interés Anual (%):", value=15.0)
        
        if st.form_submit_button("Sincronizar con Google Sheets"):
            if nombre and cedula:
                with st.spinner('Conectando con la nube...'):
                    cuota = calcular_cuota(monto, plazo, tasa)
                    nuevo_registro = pd.DataFrame([{
                        "ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"),
                        "Nombre": nombre, "Cedula": str(cedula), "Monto_Inicial": monto, "Saldo_Restante": monto,
                        "Cuota_Mensual": round(cuota, 2), "Meses_Totales": plazo, "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tasa
                    }])
                    df_actualizado = pd.concat([df_prestamos, nuevo_registro], ignore_index=True)
                    conn.update(data=df_actualizado)
                    
                    st.balloons()
                    st.success(f"Préstamo para {nombre} registrado exitosamente.")
                    time.sleep(2)
                    st.rerun()

with tab2:
    st.subheader("Estado actual de la cartera")
    if not df_prestamos.empty:
        activos = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
        st.write(f"### 🟢 Préstamos Activos: {len(activos)}")
        st.dataframe(activos, use_container_width=True, height=450)
    else:
        st.info("No hay datos en la nube de Google.")

with tab3:
    st.subheader("Resumen general")
    if not df_prestamos.empty:
        st.dataframe(df_prestamos, use_container_width=True, height=500)
    else:
        st.info("Base de datos vacía.")

with tab4:
    st.subheader("Módulo de Cobranza y Comprobantes")
    activos_pago = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
    
    if not activos_pago.empty:
        seleccion = st.selectbox("Seleccionar Cliente:", activos_pago["Nombre"] + " (ID: " + activos_pago["ID"] + ")")
        id_cliente = seleccion.split("ID: ")[1].replace(")", "")
        cliente_info = activos_pago[activos_pago["ID"] == id_cliente].iloc[0]
        
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"Cuota mensual: ${cliente_info['Cuota_Mensual']}")
            cuotas_pagar = st.number_input("Número de cuotas a pagar hoy:", min_value=1, max_value=int(cliente_info['Meses_Totales'] - cliente_info['Pagos_Realizados']), value=1)
        with c2:
            st.write(f"### Total a recibir: ${round(float(cliente_info['Cuota_Mensual']) * cuotas_pagar, 2)}")
            comprobante = st.file_uploader("Subir foto del pago a Google Drive:", type=["jpg", "jpeg", "png"])

        if st.button("✅ Confirmar Pago Permanente", use_container_width=True):
            with st.spinner('Subiendo evidencia a Drive y actualizando Sheets...'):
                if comprobante:
                    ext = comprobante.name.split('.')[-1] if '.' in comprobante.name else 'png'
                    tipo_mime = comprobante.type
                    nombre_archivo_drive = f"Recibo_{id_cliente}_{datetime.now().strftime('%Y%m%d_%H%M')}.{ext}"
                    subir_a_drive(comprobante.getvalue(), nombre_archivo_drive, tipo_mime)
                
                idx = df_prestamos[df_prestamos["ID"] == id_cliente].index[0]
                df_prestamos.at[idx, "Pagos_Realizados"] += cuotas_pagar
                abono = (float(cliente_info['Monto_Inicial']) / int(cliente_info['Meses_Totales'])) * cuotas_pagar
                df_prestamos.at[idx, "Saldo_Restante"] = round(max(0, float(cliente_info['Saldo_Restante']) - abono), 2)
                
                if int(df_prestamos.at[idx, "Pagos_Realizados"]) >= int(cliente_info['Meses_Totales']):
                    df_prestamos.at[idx, "Estado"] = "PAGADO"
                
                conn.update(data=df_prestamos)
                
                st.balloons()
                st.success("¡Pago registrado y asegurado en la nube!")
                time.sleep(3)
                st.rerun()
