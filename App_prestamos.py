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

# 2. CSS PARA INTERFAZ SÚPER GIGANTE Y JERARQUÍA VISUAL
st.markdown("""
    <style>
    /* --- Pestañas Superiores Extra Grandes --- */
    button[data-baseweb="tab"] { 
        font-size: 28px !important; 
        font-weight: 700 !important; 
        padding: 1.5rem !important; 
    }
    
    /* --- Textos normales y etiquetas Súper Visibles --- */
    .stMarkdown p, label, .stSelectbox p, .stNumberInput label, .stTextInput label { 
        font-size: 24px !important; 
        font-weight: 600 !important; 
    }
    
    /* --- BOTONES LLAMATIVOS (Acciones principales: Guardar, Confirmar, Excel) --- */
    .stButton>button[kind="primary"], .stDownloadButton>button { 
        font-size: 28px !important; 
        font-weight: 900 !important; 
        height: 5.5rem !important; 
        border-radius: 15px !important; 
        background-color: #00C851 !important; /* Verde Brillante */
        color: white !important; 
        border: none !important;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.2) !important;
        transition: all 0.3s ease;
    }
    .stButton>button[kind="primary"]:hover, .stDownloadButton>button:hover {
        background-color: #007E33 !important;
        transform: translateY(-2px);
    }

    /* --- BOTÓN DISCRETO (Eliminar) --- */
    .stButton>button[kind="secondary"] { 
        font-size: 16px !important; 
        font-weight: normal !important; 
        height: 2.5rem !important; 
        background-color: transparent !important; 
        color: #ff4b4b !important; 
        border: 1px solid #ff4b4b !important; 
        opacity: 0.7;
    }
    .stButton>button[kind="secondary"]:hover {
        opacity: 1;
        background-color: #ffeaea !important;
    }

    /* --- Métricas (Saldos) Gigantes --- */
    [data-testid="stMetricValue"] { 
        font-size: 60px !important; 
        color: #007bff !important;
    }
    [data-testid="stMetricLabel"] { 
        font-size: 24px !important; 
        font-weight: bold !important;
    }
    
    /* --- Tablas más legibles --- */
    .stDataFrame { 
        font-size: 22px !important; 
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

# 4. MODO TURBO: COMPRESIÓN DE FOTOS
def subir_a_imgbb_comprimido(archivo_bytes):
    try:
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
        if not historial_c.empty: historial_c.to_excel(writer, sheet_name="Detalle_Pagos", index=False)
    return output.getvalue()

# --- INTERFAZ PRINCIPAL ---
st.title("🏦 CONTROL DE PRESTAMOS")

df_p = cargar_prestamos()
df_h = cargar_historial_pagos()

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo Préstamo", "📋 Lista de Clientes", "🔍 Historial & Excel", "💵 Registrar Pago"])

# --- PESTAÑA 1: NUEVO PRÉSTAMO ---
with tab1:
    st.info("💡 Llena los datos del cliente para registrar un nuevo crédito en el sistema.")
    with st.form("f_nuevo"):
        nombre = st.text_input("👤 Nombre Completo del Cliente:")
        ced = st.text_input("🪪 Cédula de Identidad:")
        c1, c2, c3 = st.columns(3)
        monto = c1.number_input("💵 Monto ($):", value=500.0)
        meses = c2.number_input("📅 Plazo (Meses):", value=12)
        tasa = c3.number_input("📈 Tasa Anual %:", value=15.0)
        
        st.write("") # Espacio
        if st.form_submit_button("💾 GUARDAR NUEVO PRÉSTAMO", use_container_width=True, type="primary"):
            if nombre and ced:
                t_m = (tasa/100)/12
                cuota = monto * (t_m * (1+t_m)**meses) / ((1+t_m)**meses - 1) if t_m > 0 else monto/meses
                nuevo = pd.DataFrame([{
                    "ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"),
                    "Nombre": nombre, "Cedula": str(ced), "Monto_Inicial": monto, "Saldo_Restante": monto,
                    "Cuota_Mensual": round(cuota, 2), "Meses_Totales": meses, "Pagos_Realizados": 0, "Estado": "ACTIVO", "Tasa": tasa
                }])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, nuevo], ignore_index=True))
                st.balloons(); time.sleep(1); st.rerun()

# --- PESTAÑA 2: LISTA GENERAL ---
with tab2:
    st.success("🟢 Clientes con préstamos activos actualmente:")
    st.dataframe(df_p[df_p["Estado"] == "ACTIVO"], use_container_width=True, height=600)

# --- PESTAÑA 3: HISTORIAL, EXCEL Y BORRADO ---
with tab3:
    if not df_p.empty:
        opciones = df_p["Nombre"].astype(str) + " (" + df_p["ID"].astype(str) + ")"
        sel = st.selectbox("🔎 Seleccionar Cliente para ver detalles:", opciones)
        id_sel = sel.split("(")[1].replace(")", "")
        
        c_datos = df_p[df_p["ID"] == id_sel].iloc[0]
        c_historial = df_h[df_h["ID_Prestamo"] == id_sel]
        
        st.write("---")
        col1, col2, col3 = st.columns(3)
        col1.metric("Deuda Restante", f"${c_datos['Saldo_Restante']}")
        col2.metric("Meses Pagados", f"{c_datos['Pagos_Realizados']} de {c_datos['Meses_Totales']}")
        col3.metric("Cuota Mensual", f"${c_datos['Cuota_Mensual']}")
        st.write("---")
        
        st.write("### 📝 Bitácora de Pagos del Cliente")
        st.dataframe(c_historial, use_container_width=True, column_config={"URL_Comprobante": st.column_config.LinkColumn("📸 Ver Foto")})
        
        st.write("")
        st.download_button(
            label=f"📥 GENERAR EXCEL DE {c_datos['Nombre'].upper()}", 
            data=generar_excel_cliente(c_datos, c_historial), 
            file_name=f"Control_{id_sel}.xlsx", 
            use_container_width=True
        )

        st.write("---")
        st.write("###### ⚠️ Opciones de Administrador")
        # Este botón es secundario, pequeño, con texto rojo y contorno (muy discreto)
        if st.button(f"Eliminar registro de {c_datos['Nombre']}", use_container_width=False, type="secondary"):
            with st.spinner('Eliminando...'):
                df_p_nuevo = df_p[df_p["ID"] != id_sel]
                conn.update(worksheet="Prestamos", data=df_p_nuevo)
                df_h_nuevo = df_h[df_h["ID_Prestamo"] != id_sel]
                conn.update(worksheet="Pagos", data=df_h_nuevo)
                time.sleep(1); st.rerun()

# --- PESTAÑA 4: REGISTRO DE PAGOS ---
with tab4:
    activos = df_p[df_p["Estado"] == "ACTIVO"]
    if not activos.empty:
        st.warning("💰 Módulo de Cobranza")
        op_pagos = activos["Nombre"].astype(str) + " (" + activos["ID"].astype(str) + ")"
        p_sel = st.selectbox("¿De quién vas a registrar el pago?", op_pagos)
        id_p = p_sel.split("(")[1].replace(")", ""); cliente = activos[activos["ID"] == id_p].iloc[0]
        
        st.write("---")
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            n_cuotas = st.number_input("🔢 Número de cuotas que está pagando:", min_value=1, value=1)
            st.info(f"### Dinero a recibir hoy: ${round(cliente['Cuota_Mensual'] * n_cuotas, 2)}")
        with c_p2: 
            foto = st.file_uploader("📸 Subir Foto/Captura del Recibo:", type=["jpg", "png", "jpeg"])
        
        st.write("---")
        if st.button("✅ CONFIRMAR Y REGISTRAR PAGO", use_container_width=True, type="primary"):
            with st.spinner('Procesando pago y foto a máxima velocidad...'):
                link_foto = subir_a_imgbb_comprimido(foto.getvalue()) if foto else ""
                
                nuevo_p = pd.DataFrame([{"ID_Prestamo": id_p, "Fecha_Pago": datetime.now().strftime("%Y-%m-%d %H:%M"), "Cuotas_Pagadas": n_cuotas, "Monto_Pagado": round(cliente["Cuota_Mensual"] * n_cuotas, 2), "URL_Comprobante": link_foto}])
                conn.update(worksheet="Pagos", data=pd.concat([df_h, nuevo_p], ignore_index=True))
                
                idx = df_p[df_p["ID"] == id_p].index[0]
                df_p.at[idx, "Pagos_Realizados"] += n_cuotas
                descuento = (cliente["Monto_Inicial"] / cliente["Meses_Totales"]) * n_cuotas
                df_p.at[idx, "Saldo_Restante"] = round(max(0, cliente["Saldo_Restante"] - descuento), 2)
                if df_p.at[idx, "Pagos_Realizados"] >= cliente["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                
                conn.update(worksheet="Prestamos", data=df_p)
                st.balloons(); time.sleep(1); st.rerun()
