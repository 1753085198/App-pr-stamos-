import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import time
import io
from datetime import datetime
import requests
import base64

# 1. CONFIGURACIÓN DE PÁGINA PANORÁMICA
st.set_page_config(page_title="Sistema Pro Jose Figueroa", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE (Botones y letras más grandes)
st.markdown("""
    <style>
    /* Agrandar Pestañas */
    button[data-baseweb="tab"] {
        font-size: 22px !important;
        font-weight: 600 !important;
        padding: 1rem !important;
    }
    /* Letras de formularios y etiquetas */
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label {
        font-size: 20px !important;
        font-weight: 500 !important;
    }
    /* Botones de acción */
    .stButton>button {
        font-size: 22px !important;
        font-weight: bold !important;
        height: 3.5rem !important;
        border-radius: 10px !important;
    }
    /* Números de métricas */
    [data-testid="stMetricValue"] {
        font-size: 40px !important;
    }
    /* Tablas de datos */
    .stDataFrame {
        font-size: 18px !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. CONEXIÓN A GOOGLE SHEETS
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_prestamos():
    try:
        df = conn.read(worksheet="Prestamos", ttl="0") 
        if not df.empty:
            # Blindaje para evitar errores de tipo en móviles
            df["Cedula"] = df["Cedula"].astype(str).str.replace(".0", "", regex=False)
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
            df["Nombre"] = df["Nombre"].astype(str)
        return df
    except:
        return pd.DataFrame(columns=["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"])

def cargar_historial_pagos():
    try:
        df = conn.read(worksheet="Pagos", ttl="0")
        if not df.empty:
            df["ID_Prestamo"] = df["ID_Prestamo"].astype(str)
        return df
    except:
        return pd.DataFrame(columns=["ID_Prestamo", "Fecha_Pago", "Cuotas_Pagadas", "Monto_Pagado", "URL_Comprobante"])

# 4. FUNCIÓN: SUBIR FOTO A IMGBB
def subir_a_imgbb(archivo_bytes):
    try:
        api_key = st.secrets["IMGBB_API_KEY"]
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": api_key, "image": base64.b64encode(archivo_bytes).decode('utf-8')}
        res = requests.post(url, data=payload)
        return res.json()["data"]["url"] if res.status_code == 200 else ""
    except:
        return ""

# 5. MOTOR DE GENERACIÓN DE EXCEL
def generar_excel_reporte(cliente, historial):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame([cliente]).to_excel(writer, sheet_name="Resumen_Cliente", index=False)
        if not historial.empty:
            historial.to_excel(writer, sheet_name="Historial_de_Pagos", index=False)
    return output.getvalue()

# --- LÓGICA DE INTERFAZ ---
st.title("🏦 Sistema Maestro de Préstamos")

df_p = cargar_prestamos()
df_h = cargar_historial_pagos()

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo Préstamo", "📋 Lista General", "🔍 Historial & Excel", "💵 Registrar Pago"])

# PESTAÑA 1: NUEVO
with tab1:
    st.subheader("Datos del nuevo crédito")
    with st.form("f_nuevo"):
        nombre = st.text_input("Nombre del Cliente:")
        ced = st.text_input("Cédula:")
        c1, c2, c3 = st.columns(3)
        monto = c1.number_input("Monto ($):", value=500.0)
        meses = c2.number_input("Plazo (Meses):", value=12)
        tasa = c3.number_input("Tasa Anual %:", value=15.0)
        
        if st.form_submit_button("💾 GUARDAR EN GOOGLE CLOUD", use_container_width=True):
            if nombre and ced:
                with st.spinner('Sincronizando...'):
                    t_m = (tasa/100)/12
                    cuota = monto * (t_m * (1+t_m)**meses) / ((1+t_m)**meses - 1) if t_m > 0 else monto/meses
                    nuevo = pd.DataFrame([{
                        "ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"),
                        "Nombre": nombre, "Cedula": str(ced), "Monto_Inicial": monto, "Saldo_Restante": monto,
                        "Cuota_Mensual": round(cuota, 2), "Meses_Totales": meses, "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tasa
                    }])
                    conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                    st.balloons()
                    time.sleep(1)
                    st.rerun()

# PESTAÑA 2: LISTA
with tab2:
    st.subheader("Cartera de Clientes")
    st.dataframe(df_p, use_container_width=True, height=500)

# PESTAÑA 3: HISTORIAL Y EXCEL
with tab3:
    if not df_p.empty:
        opciones = df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")"
        sel = st.selectbox("Seleccionar para ver bitácora:", opciones)
        id_sel = sel.split("(")[1].replace(")", "")
        
        c_datos = df_p[df_p["ID"] == id_sel].iloc[0]
        c_historial = df_h[df_h["ID_Prestamo"] == id_sel]
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Saldo Actual", f"${c_datos['Saldo_Restante']}")
        col2.metric("Cuotas Pagadas", f"{c_datos['Pagos_Realizados']}/{c_datos['Meses_Totales']}")
        col3.metric("Cuota Fija", f"${c_datos['Cuota_Mensual']}")
        
        st.write("### 📜 Movimientos Registrados")
        st.dataframe(c_historial, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("📸 Ver Foto")})
        
        data_ex = generar_excel_reporte(c_datos, c_historial)
        st.download_button(f"📥 Descargar Excel de {c_datos['Nombre']}", data=data_ex, file_name=f"Reporte_{id_sel}.xlsx", use_container_width=True)

# PESTAÑA 4: PAGAR
with tab4:
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if not activos.empty:
        op_pagos = activos["Nombre"].astype(str) + " (" + activos["ID"].astype(str) + ")"
        p_sel = st.selectbox("¿Quién va a pagar?", op_pagos)
        id_p = p_sel.split("(")[1].replace(")", "")
        cliente = activos[activos["ID"] == id_p].iloc[0]
        
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            n_cuotas = st.number_input("Cantidad de cuotas a cobrar:", min_value=1, value=1)
            st.write(f"## Total: ${round(cliente['Cuota_Mensual'] * n_cuotas, 2)}")
        with col_p2:
            foto = st.file_uploader("📸 Subir Recibo/Captura:", type=["jpg", "png", "jpeg"])
        
        if st.button("🚀 CONFIRMAR PAGO Y SUBIR EVIDENCIA", use_container_width=True):
            with st.spinner('Subiendo a ImgBB y actualizando historial...'):
                link_foto = subir_a_imgbb(foto.getvalue()) if foto else ""
                
                # 1. Guardar en Hoja "Pagos" (Historial)
                nuevo_p = pd.DataFrame([{
                    "ID_Prestamo": id_p, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Cuotas_Pagadas": n_cuotas, "Monto_Pagado": round(cliente["Cuota_Mensual"] * n_cuotas, 2),
                    "URL_Comprobante": link_foto
                }])
                df_h_nuevo = pd.concat([df_h, nuevo_p], ignore_index=True)
                conn.update(worksheet="Pagos", data=df_h_nuevo)
                
                # 2. Actualizar Hoja "Prestamos" (Saldos)
                idx = df_p[df_p["ID"] == id_p].index[0]
                df_p.at[idx, "Pagos_Realizados"] += n_cuotas
                descuento = (cliente["Monto_Inicial"] / cliente["Meses_Totales"]) * n_cuotas
                df_p.at[idx, "Saldo_Restante"] = round(max(0, cliente["Saldo_Restante"] - descuento), 2)
                
                if df_p.at[idx, "Pagos_Realizados"] >= cliente["Meses_Totales"]:
                    df_p.at[idx, "Estado"] = "PAGADO"
                
                conn.update(worksheet="Prestamos", data=df_p)
                st.balloons()
                time.sleep(2)
                st.rerun()
    else:
        st.success("🎉 ¡No hay préstamos activos pendientes de cobro!")
