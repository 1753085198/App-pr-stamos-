import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import time
from datetime import datetime
import requests
import base64

# --- CONFIGURACIÓN PANORÁMICA ---
st.set_page_config(page_title="Sistema Pro de Préstamos", page_icon="🏦", layout="wide")

# --- CONEXIÓN A GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_datos():
    df = conn.read(ttl="0")
    if not df.empty:
        # Forzamos todo a texto para que no haya TypeErrors en el celular
        df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
        df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        df["Nombre"] = df["Nombre"].astype(str)
        
        if "Comprobante_URL" not in df.columns:
            df["Comprobante_URL"] = ""
    return df

# --- FUNCIÓN: SUBIR FOTO A IMGBB ---
def subir_a_imgbb(archivo_bytes):
    # Esto busca la llave que pusiste hasta arriba en los Secrets
    api_key = st.secrets["IMGBB_API_KEY"] 
    url = "https://api.imgbb.com/1/upload"
    payload = {
        "key": api_key,
        "image": base64.b64encode(archivo_bytes).decode('utf-8')
    }
    try:
        respuesta = requests.post(url, data=payload)
        if respuesta.status_code == 200:
            return respuesta.json()["data"]["url"]
        return "Error al subir"
    except:
        return "Error de conexión"

# --- MATEMÁTICA FINANCIERA ---
def calcular_cuota(capital, meses, tasa_anual):
    tasa_mensual = (tasa_anual / 100) / 12
    if tasa_mensual > 0:
        return capital * (tasa_mensual * (1 + tasa_mensual)**meses) / ((1 + tasa_mensual)**meses - 1)
    return capital / meses

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 Sistema Integral de Préstamos y Cobranza")
st.write("Base de datos conectada: Google Sheets + ImgBB para recibos.")

try:
    df_prestamos = cargar_datos()
except Exception as e:
    st.error("Esperando conexión con Google Sheets. Verifica tus Secrets.")
    st.stop()

if df_prestamos.empty:
    columnas = ["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa", "Comprobante_URL"]
    df_prestamos = pd.DataFrame(columns=columnas)

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo Préstamo", "📋 Lista de Clientes", "🔍 Auditoría", "💵 Registrar Pagos"])

# PESTAÑA 1: CREACIÓN
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
                        "Cuota_Mensual": round(cuota, 2), "Meses_Totales": plazo, "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tasa, "Comprobante_URL": ""
                    }])
                    df_actualizado = pd.concat([df_prestamos, nuevo_registro], ignore_index=True)
                    conn.update(data=df_actualizado)
                    
                    st.balloons()
                    st.success(f"Préstamo para {nombre} registrado exitosamente.")
                    time.sleep(2)
                    st.rerun()

# PESTAÑA 2: LISTA GENERAL
with tab2:
    st.subheader("Estado actual de la cartera")
    if not df_prestamos.empty:
        activos = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
        st.write(f"### 🟢 Préstamos Activos: {len(activos)}")
        st.dataframe(activos, use_container_width=True, height=450)
    else:
        st.info("No hay datos en la nube de Google.")

# PESTAÑA 3: AUDITORÍA BÁSICA
with tab3:
    st.subheader("Resumen general")
    if not df_prestamos.empty:
        st.dataframe(
            df_prestamos, 
            use_container_width=True, 
            height=500,
            column_config={
                "Comprobante_URL": st.column_config.LinkColumn("Ver Último Comprobante")
            }
        )
    else:
        st.info("Base de datos vacía.")

# PESTAÑA 4: COBRANZA
with tab4:
    st.subheader("Módulo de Cobranza y Comprobantes")
    activos_pago = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
    
    if not activos_pago.empty:
        # BLINDAJE: Todo convertido a texto explícitamente para evitar TypeError
        opciones_pago = activos_pago["Nombre"].astype(str) + " (ID: " + activos_pago["ID"].astype(str) + ")"
        seleccion = st.selectbox("Seleccionar Cliente:", opciones_pago)
        
        id_cliente = seleccion.split("ID: ")[1].replace(")", "")
        cliente_info = activos_pago[activos_pago["ID"] == id_cliente].iloc[0]
        
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"Cuota mensual: ${cliente_info['Cuota_Mensual']}")
            cuotas_pagar = st.number_input("Número de cuotas a pagar hoy:", min_value=1, max_value=int(cliente_info['Meses_Totales'] - cliente_info['Pagos_Realizados']), value=1)
        with c2:
            st.write(f"### Total a recibir: ${round(float(cliente_info['Cuota_Mensual']) * cuotas_pagar, 2)}")
            comprobante = st.file_uploader("Subir foto del pago:", type=["jpg", "jpeg", "png"])

        if st.button("✅ Confirmar Pago", use_container_width=True):
            with st.spinner('Procesando pago y subiendo evidencia a ImgBB...'):
                url_foto = ""
                if comprobante:
                    url_foto = subir_a_imgbb(comprobante.getvalue())
                
                idx = df_prestamos[df_prestamos["ID"] == id_cliente].index[0]
                df_prestamos.at[idx, "Pagos_Realizados"] += cuotas_pagar
                abono = (float(cliente_info['Monto_Inicial']) / int(cliente_info['Meses_Totales'])) * cuotas_pagar
                df_prestamos.at[idx, "Saldo_Restante"] = round(max(0, float(cliente_info['Saldo_Restante']) - abono), 2)
                
                if url_foto and url_foto != "Error al subir" and url_foto != "Error de conexión":
                    df_prestamos.at[idx, "Comprobante_URL"] = url_foto
                
                if int(df_prestamos.at[idx, "Pagos_Realizados"]) >= int(cliente_info['Meses_Totales']):
                    df_prestamos.at[idx, "Estado"] = "PAGADO"
                
                conn.update(data=df_prestamos)
                
                st.balloons()
                st.success("¡Pago registrado y link de foto asegurado en el Excel!")
                time.sleep(3)
                st.rerun()
