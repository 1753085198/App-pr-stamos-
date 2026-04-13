import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
import time
from datetime import datetime
import requests
import base64
from openpyxl.styles import PatternFill, Font, Alignment

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL PRO", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE Y BOTONES DE ALTO IMPACTO
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 24px !important; font-weight: 700 !important; border-radius: 12px !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 28px !important; font-weight: 900 !important; position: fixed; bottom: 30px; right: 30px; z-index: 9999; border: 3px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 65px !important; font-weight: 900 !important; color: #007bff !important; }
    .stExpander { border: 2px solid #e6e9ef !important; border-radius: 10px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNCIONES DE SOPORTE ---
def cargar(h):
    try:
        df = conn.read(worksheet=h, ttl=0)
        if df is not None and "ID" in df.columns:
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df if df is not None else pd.DataFrame()
    except: return pd.DataFrame()

def subir_img(archivo):
    try:
        res = requests.post("https://api.imgbb.com/1/upload", data={"key": st.secrets["IMGBB_API_KEY"], "image": base64.b64encode(archivo).decode('utf-8')})
        return res.json()["data"]["url"]
    except: return ""

def generar_excel_grupal(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("GENERAL", 0)
        f_v, f_r = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE GENERAL - {titulo}"]); ws.append(["Nombre", "Cédula", "Saldo/Aporte Total", "Estado"])
        for _, row in df.iterrows():
            monto = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            ws.append([row['Nombre'], row['Cedula'], monto, "ACTIVO" if monto > 0 else "PENDIENTE"])
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = f_v if monto > 0 else f_r
    return out.getvalue()

def generar_excel_personal(socio_row, df_historial, tipo_label):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("HISTORIAL", 0)
        ws.append([f"HISTORIAL DETALLADO DE {tipo_label}"]); ws.append([f"NOMBRE: {socio_row['Nombre']}"]); ws.append([f"CÉDULA: {socio_row['Cedula']}"])
        ws.append([""]); ws.append(["Fecha", "Monto", "Link Recibo"])
        pagos = df_historial[df_historial['ID_Socio'] == socio_row['ID']] if not df_historial.empty else pd.DataFrame()
        for _, p in pagos.iterrows(): ws.append([p['Fecha'], p['Monto'], p.get('Comprobante', 'N/A')])
        ws.append([""]); ws.append(["TOTAL ACUMULADO:", socio_row.get('Saldo_Total_Aportado', socio_row.get('Monto_Inicial', 0))])
    return out.getvalue()

# --- MENÚ PRINCIPAL ---
with st.sidebar:
    st.markdown("# 🏦 PANEL DE CONTROL")
    seccion = st.radio("Seleccione Operación:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)
    st.write("---")
    st.caption("Jose Figueroa - UDLA 2026")

# --- LÓGICA DE SECCIONES ---

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo_p"): st.session_state.n_p = True
    if st.session_state.get('n_p'):
        with st.form("np"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            m, t, p = st.number_input("Monto:"), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("CREAR"):
                i = (t/100)/12
                cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":n, "Cedula":c, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO"}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
                st.session_state.n_p = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR PRESTAMISTA:")
    act = df_p[df_p["Estado"]=="ACTIVO"] if not df_p.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | DEUDA: ${row['Saldo_Restante']}"):
            c1, c2 = st.columns(2)
            with c1:
                st.metric("CUOTA MENSUAL", f"${row['Cuota_Mensual']}")
                st.metric("PROGRESO", f"{row['Pagos_Realizados']}/{row['Meses_Totales']} Meses")
            with c2:
                with st.form(key=f"fp_{row['ID']}"):
                    ft = st.file_uploader("📸 FOTO RECIBO:", key=f"ip_{row['ID']}")
                    if st.form_submit_button("✅ REGISTRAR COBRO"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_p = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": row['Cuota_Mensual'], "Comprobante": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_p], ignore_index=True))
                            df_p.at[idx, "Pagos_Realizados"] += 1
                            df_p.at[idx, "Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                            if df_p.at[idx, "Pagos_Realizados"] >= row["Meses_Totales"]: df_p.at[idx, "Estado"] = "PAGADO"
                            conn.update(worksheet="Prestamos", data=df_p); st.rerun()

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s, df_ph = cargar("Cooperativa"), cargar("Pagos_Coop")
    v_x = st.number_input("💵 CUOTA MENSUAL ESTÁNDAR (X):", value=10.0)
    
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL", data=generar_excel_grupal(df_s, "COOP"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo_c"): st.session_state.n_c = True
    if st.session_state.get('n_c'):
        with st.form("nc"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("GUARDAR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True)); st.session_state.n_c = False; st.rerun()

    act = df_s if not df_s.empty else pd.DataFrame()
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fc_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_x)
                    ft = st.file_uploader("📸 RECIBO:", key=f"ic_{row['ID']}")
                    if st.form_submit_button("✅ COBRAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Coop", data=pd.concat([df_ph, new_h], ignore_index=True))
                            df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Cooperativa", data=df_s); st.rerun()
            with c2:
                st.download_button("📊 DESCARGAR MI HISTORIAL", data=generar_excel_personal(row, df_ph, "COOPERATIVA"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dlc_{row['ID']}")

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a, df_ah = cargar("Ayudas_Listado"), cargar("Pagos_Ayudas")
    v_y = st.number_input("💵 APORTE ESTÁNDAR (X):", value=5.0)
    
    if not df_a.empty: st.download_button("📊 EXCEL GRUPAL", data=generar_excel_grupal(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo_a"): st.session_state.n_a = True
    if st.session_state.get('n_a'):
        with st.form("na"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True)); st.session_state.n_a = False; st.rerun()

    act = df_a if not df_a.empty else pd.DataFrame()
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                with st.form(key=f"fa_{row['ID']}"):
                    m = st.number_input("Monto:", value=v_y)
                    ft = st.file_uploader("📸 RECIBO:", key=f"ia_{row['ID']}")
                    if st.form_submit_button("✅ REGISTRAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Socio": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": m, "Comprobante": url}])
                            conn.update(worksheet="Pagos_Ayudas", data=pd.concat([df_ah, new_h], ignore_index=True))
                            df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                            conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
            with c2:
                st.download_button("📊 DESCARGAR MI HISTORIAL", data=generar_excel_personal(row, df_ah, "AYUDAS"), file_name=f"Historial_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
