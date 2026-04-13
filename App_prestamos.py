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

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="CONTROL DE PRESTAMOS", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y CÓMODA
st.markdown("""
    <style>
    /* Agrandar Pestañas Superiores */
    button[data-baseweb="tab"] {
        font-size: 24px !important;
        font-weight: 600 !important;
        padding: 1.2rem !important;
    }
    /* Letras de formularios y etiquetas */
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label {
        font-size: 21px !important;
        font-weight: 500 !important;
    }
    /* Botones de acción principales */
    .stButton>button {
        font-size: 24px !important;
        font-weight: bold !important;
        height: 4rem !important;
        border-radius: 12px !important;
        background-color: #007bff !important;
        color: white !important;
    }
    /* Métricas (Saldos) */
    [data-testid="stMetricValue"] {
        font-size: 48px !important;
    }
    /* Tablas */
    .stDataFrame {
        font-size: 19px !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. CONEXIONES A LA NUBE
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_prestamos():
    try:
        df = conn.read(worksheet="Prestamos", ttl="0") 
        if not df.empty:
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

# 4. FUNCIÓN MODO TURBO: COMPRESIÓN Y SUBIDA A IMGBB
def subir_a_imgbb_comprimido(archivo_bytes):
    try:
        # Comprimir imagen para que suba instantáneo
        imagen = Image.open(io.BytesIO(archivo_bytes))
        ancho_max = 800
        if imagen.width > ancho_max:
            prop = ancho_max / float(imagen.width)
            imagen = imagen.resize((ancho_max, int(imagen.height * prop)), Image.Resampling.LANCZOS)
        
        buffer = io.BytesIO()
        imagen.convert('RGB').save(buffer, format="JPEG", quality=70)
        img_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        api_key = st.secrets["IMGBB_API_KEY"]
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": api_key, "image": img_data})
        return res.json()["data"]["url"] if res.status_code == 200 else ""
    except:
        return ""

# 5. GENERADOR DE REPORTES EXCEL
def generar_excel_cliente(datos_c, historial_c):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame([datos_c]).to_excel(writer, sheet_name="Estado_Cuenta", index=False)
        if not historial_c.empty:
            historial_c.to_excel(writer, sheet_name="Detalle_Pagos", index=False)
    return output.getvalue()

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 CONTROL DE PRESTAMOS")

df_p = cargar_prestamos()
df_h = cargar_historial_pagos()

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo Préstamo", "📋 Lista General", "🔍 Historial & Excel", "💵 Registrar Pago"])

with tab1:
    st.subheader("Crear un nuevo registro")
    with st.form("f_nuevo"):
        nombre = st.text_input("Nombre del Cliente:")
        ced = st.text_input("Cédula:")
        c1, c2, c3 = st.columns(3)
        monto = c1.number_input("Monto ($):", value=500.0)
        meses = c2.number_input("Plazo (Meses):", value=12)
        tasa = c3.number_input("Tasa Anual %:", value=15.0)
        
        if st.form_submit_button("💾 GUARDAR PRESTAMO", use_container_width=True):
            if nombre and ced:
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

with tab2:
    st.subheader("Clientes Activos")
    st.dataframe(df_p[df_p["Estado"] == "ACTIVO"], use_container_width=True, height=500)

with tab3:
    if not df_p.empty:
        opciones = df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")"
        sel = st.selectbox("Seleccionar Cliente:", opciones)
        id_sel = sel.split("(")[1].replace(")", "")
        
        c_datos = df_p[df_p["ID"] == id_sel].iloc[0]
        c_historial = df_h[df_h["ID_Prestamo"] == id_sel]
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Saldo Deudor", f"${c_datos['Saldo_Restante']}")
        col2.metric("Progreso", f"{c_datos['Pagos_Realizados']}/{c_datos['Meses_Totales']}")
        col3.metric("Cuota", f"${c_datos['Cuota_Mensual']}")
        
        st.write("### 📝 Historial de Abonos")
        st.dataframe(c_historial, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("📸 Ver Foto")})
        
        data_ex = generar_excel_cliente(c_datos, c_historial)
        st.download_button(f"📥 Descargar Excel de {c_datos['Nombre']}", data=data_ex, file_name=f"Control_{id_sel}.xlsx", use_container_width=True)

with tab4:
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if not activos.empty:
        op_pagos = activos["Nombre"].astype(str) + " (" + activos["ID"].astype(str) + ")"
        p_sel = st.selectbox("Registrar pago de:", op_pagos)
        id_p = p_sel.split("(")[1].replace(")", "")
        cliente = activos[activos["ID"] == id_p].iloc[0]
        
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            n_cuotas = st.number_input("Cuotas a cobrar:", min_value=1, value=1)
            st.write(f"## Total: ${round(cliente['Cuota_Mensual'] * n_cuotas, 2)}")
        with c_p2:
            foto = st.file_uploader("📸 Recibo (JPG/PNG):", type=["jpg", "png", "jpeg"])
        
        if st.button("🚀 CONFIRMAR PAGO", use_container_width=True):
            with st.spinner('Procesando pago...'):
                link_foto = subir_a_imgbb_comprimido(foto.getvalue()) if foto else ""
                
                # 1. Registrar en Hoja "Pagos"
                nuevo_p = pd.DataFrame([{
                    "ID_Prestamo": id_p, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Cuotas_Pagadas": n_cuotas, "Monto_Pagado": round(cliente["Cuota_Mensual"] * n_cuotas, 2),
                    "URL_Comprobante": link_foto
                }])
                conn.update(worksheet="Pagos", data=pd.concat([df_h, nuevo_p], ignore_index=True))
                
                # 2. Actualizar Hoja "Prestamos"
                idx = df_p[df_p["ID"] == id_p].index[0]
                df_p.at[idx, "Pagos_Realizados"] += n_cuotas
                abono = (cliente["Monto_Inicial"] / cliente["Meses_Totales"]) * n_cuotas
                df_p.at[idx, "Saldo_Restante"] = round(max(0, cliente["Saldo_Restante"] - abono), 2)
                
                if df_p.at[idx, "Pagos_Realizados"] >= cliente["Meses_Totales"]:
                    df_p.at[idx, "Estado"] = "PAGADO"
                
                conn.update(worksheet="Prestamos", data=df_p)
                st.balloons()
                time.sleep(2)
                st.rerun()
