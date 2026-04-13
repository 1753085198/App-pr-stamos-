import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import time
import io
from datetime import datetime
import requests
import base64

# --- CONFIGURACIÓN PANORÁMICA ---
st.set_page_config(page_title="Sistema Pro Jose Figueroa", page_icon="🏦", layout="wide")

# --- CONEXIÓN A GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_prestamos():
    df = conn.read(worksheet="Sheet1", ttl="0") # Tu hoja principal
    if not df.empty:
        df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
        df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
    return df

def cargar_historial_pagos():
    try:
        df = conn.read(worksheet="Pagos", ttl="0")
        return df
    except:
        return pd.DataFrame(columns=["ID_Prestamo", "Fecha_Pago", "Cuotas_Pagadas", "Monto_Pagado", "URL_Comprobante"])

# --- FUNCIÓN: SUBIR FOTO A IMGBB ---
def subir_a_imgbb(archivo_bytes):
    try:
        api_key = st.secrets["IMGBB_API_KEY"]
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": api_key, "image": base64.b64encode(archivo_bytes).decode('utf-8')}
        res = requests.post(url, data=payload)
        return res.json()["data"]["url"] if res.status_code == 200 else ""
    except:
        return ""

# --- MOTOR DE EXCEL ---
def generar_excel_pro(cliente, historial_cliente):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: Resumen del Préstamo
        resumen = pd.DataFrame([cliente])
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        
        # Hoja 2: Historial de Pagos Reales
        if not historial_cliente.empty:
            historial_cliente.to_excel(writer, sheet_name="Historial_Pagos", index=False)
        
    return output.getvalue()

# --- INTERFAZ ---
st.title("🏦 Gestión de Préstamos Pro - Jose Figueroa")

df_p = cargar_prestamos()
df_h = cargar_historial_pagos()

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo", "📋 Lista", "🔍 Historial & Excel", "💵 Pagar"])

with tab1:
    with st.form("f1"):
        nombre = st.text_input("Nombre:")
        ced = st.text_input("Cédula:")
        monto = st.number_input("Monto:", value=100.0)
        meses = st.number_input("Meses:", value=6)
        tasa = st.number_input("Tasa Anual %:", value=15.0)
        if st.form_submit_button("Guardar"):
            t_m = (tasa/100)/12
            cuota = monto * (t_m * (1+t_m)**meses) / ((1+t_m)**meses - 1) if t_m > 0 else monto/meses
            nuevo = pd.DataFrame([{
                "ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"),
                "Nombre": nombre, "Cedula": str(ced), "Monto_Inicial": monto, "Saldo_Restante": monto,
                "Cuota_Mensual": round(cuota, 2), "Meses_Totales": meses, "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tasa
            }])
            conn.update(worksheet="Sheet1", data=pd.concat([df_p, nuevo], ignore_index=True))
            st.balloons()
            st.rerun()

with tab2:
    st.dataframe(df_p, use_container_width=True, height=400)

with tab3:
    if not df_p.empty:
        sel = st.selectbox("Ver historial de:", df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")")
        id_sel = sel.split("(")[1].replace(")", "")
        
        historial_especifico = df_h[df_h["ID_Prestamo"] == id_sel]
        cliente_datos = df_p[df_p["ID"] == id_sel].iloc[0]
        
        col1, col2 = st.columns(2)
        col1.metric("Saldo Pendiente", f"${cliente_datos['Saldo_Restante']}")
        col2.metric("Pagos hechos", f"{cliente_datos['Pagos_Realizados']} de {cliente_datos['Meses_Totales']}")
        
        st.write("### 📖 Bitácora de Pagos")
        st.dataframe(historial_especifico, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("Ver Foto")})
        
        excel_data = generar_excel_pro(cliente_datos, historial_especifico)
        st.download_button("📥 Descargar Reporte Excel", data=excel_data, file_name=f"Reporte_{id_sel}.xlsx", use_container_width=True)

with tab4:
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if not activos.empty:
        p_sel = st.selectbox("Registrar pago para:", activos["Nombre"].astype(str) + " (" + activos["ID"].astype(str) + ")", key="pago_sel")
        id_pago = p_sel.split("(")[1].replace(")", "")
        c_pago = activos[activos["ID"] == id_pago].iloc[0]
        
        foto = st.file_uploader("Subir Recibo:", type=["jpg", "png", "jpeg"])
        n_cuotas = st.number_input("Cuotas a pagar:", min_value=1, value=1)
        
        if st.button("✅ Confirmar Pago"):
            with st.spinner("Sincronizando..."):
                link = subir_a_imgbb(foto.getvalue()) if foto else ""
                
                # 1. Registrar en Historial
                nuevo_pago = pd.DataFrame([{
                    "ID_Prestamo": id_pago, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Cuotas_Pagadas": n_cuotas, "Monto_Pagado": round(c_pago["Cuota_Mensual"] * n_cuotas, 2),
                    "URL_Comprobante": link
                }])
                conn.update(worksheet="Pagos", data=pd.concat([df_h, nuevo_pago], ignore_index=True))
                
                # 2. Actualizar Saldo en Sheet1
                idx = df_p[df_p["ID"] == id_pago].index[0]
                df_p.at[idx, "Pagos_Realizados"] += n_cuotas
                abono = (c_pago["Monto_Inicial"] / c_pago["Meses_Totales"]) * n_cuotas
                df_p.at[idx, "Saldo_Restante"] = round(max(0, c_pago["Saldo_Restante"] - abono), 2)
                if df_p.at[idx, "Pagos_Realizados"] >= c_pago["Meses_Totales"]:
                    df_p.at[idx, "Estado"] = "PAGADO"
                
                conn.update(worksheet="Sheet1", data=df_p)
                st.balloons()
                time.sleep(2)
                st.rerun()
