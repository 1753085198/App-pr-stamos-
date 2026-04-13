import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
import time
from datetime import datetime
import requests
import base64
from openpyxl.styles import PatternFill, Font

# 1. CONFIGURACIÓN
st.set_page_config(page_title="SISTEMA FINANCIERO TOTAL", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 28px !important; font-weight: 700 !important; }
    input { font-size: 24px !important; height: 55px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 26px !important; font-weight: 700 !important; border-radius: 15px !important; }
    div.stButton > button:first-child[key^="btn_nuevo"] { background-color: #ff5722 !important; color: white !important; border-radius: 50px !important; padding: 20px 40px !important; font-size: 30px !important; font-weight: 900 !important; position: fixed; bottom: 40px; right: 40px; z-index: 9999; border: 4px solid white !important; }
    [data-testid="stMetricValue"] { font-size: 75px !important; font-weight: 900 !important; color: #007bff !important; }
    .stExpander { border: 2px solid #f0f2f6 !important; margin-bottom: 10px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- MENÚ LATERAL ---
with st.sidebar:
    st.markdown("# 🏦 NAVEGACIÓN")
    seccion = st.radio("Sección:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=1)
    st.write("---")
    st.info("Jose Figueroa - UDLA")

# --- FUNCIONES CORE ---
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

def generar_excel_pintado(df, titulo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        ws = writer.book.create_sheet("REPORTE", 0)
        f_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        f_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.append([f"REPORTE: {titulo}"])
        ws.append(["Nombre", "Cédula", "Monto Total", "Estado"])
        for _, row in df.iterrows():
            monto = float(row.get('Saldo_Total_Aportado', row.get('Saldo_Restante', 0)))
            estado = "PAGADO" if monto > 0 else "PENDIENTE"
            ws.append([row['Nombre'], row['Cedula'], monto, estado])
            color = f_verde if monto > 0 else f_rojo
            for col in range(1, 5): ws.cell(row=ws.max_row, column=col).fill = color
    return out.getvalue()

# --- LÓGICA POR SECCIÓN ---

if seccion == "💰 PRÉSTAMOS":
    st.title("💰 PRÉSTAMOS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    if st.button("👤 NUEVO PRÉSTAMO", key="btn_nuevo_p"): st.session_state.mostrar_nuevo = True
    if st.session_state.get('mostrar_nuevo'):
        with st.form("np"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            m, t, p = st.number_input("Monto:"), st.number_input("Tasa %:", value=15.0), st.number_input("Meses:", value=12)
            if st.form_submit_button("GUARDAR"):
                i = (t/100)/12
                cuo = m * (i*(1+i)**p)/((1+i)**p-1) if i>0 else m/p
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:8], "Nombre":n, "Cedula":c, "Monto_Inicial":m, "Saldo_Restante":round(cuo*p,2), "Cuota_Mensual":round(cuo,2), "Meses_Totales":p, "Pagos_Realizados":0, "Estado":"ACTIVO"}])
                conn.update(worksheet="Prestamos", data=pd.concat([df_p, new], ignore_index=True))
                st.session_state.mostrar_nuevo = False; st.rerun()
    
    bq = st.text_input("🔍 BUSCAR:")
    act = df_p[df_p["Estado"]=="ACTIVO"] if not df_p.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    for _, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}"):
            c1, c2 = st.columns(2); c1.metric("CUOTA", f"${row['Cuota_Mensual']}"); c1.metric("PAGOS", f"{row['Pagos_Realizados']}/{row['Meses_Totales']}")
            with c2:
                with st.form(key=f"f_{row['ID']}"):
                    ft = st.file_uploader("📸 RECIBO:", key=f"foto_{row['ID']}")
                    if st.form_submit_button("✅ COBRAR"):
                        if ft:
                            url = subir_img(ft.getvalue())
                            new_h = pd.DataFrame([{"ID_Prestamo": row['ID'], "Fecha": datetime.now().strftime("%Y-%m-%d"), "Monto": row['Cuota_Mensual'], "URL": url}])
                            conn.update(worksheet="Pagos", data=pd.concat([df_h, new_h], ignore_index=True))
                            row["Pagos_Realizados"] += 1; row["Saldo_Restante"] = round(row["Saldo_Restante"] - row["Cuota_Mensual"], 2)
                            if row["Pagos_Realizados"] >= row["Meses_Totales"]: row["Estado"] = "PAGADO"
                            df_p.loc[df_p['ID']==row['ID']] = row; conn.update(worksheet="Prestamos", data=df_p); st.rerun()

elif seccion == "🤝 COOPERATIVA":
    st.title("🤝 COOPERATIVA")
    df_s = cargar("Cooperativa")
    cuota_fija = st.number_input("💵 VALOR DE CUOTA ESTÁNDAR (X):", value=10.0)
    
    if not df_s.empty: st.download_button("📊 EXCEL GRUPAL COOP", data=generar_excel_pintado(df_s, "COOPERATIVA"), file_name="Reporte_Coop.xlsx", use_container_width=True)
    if st.button("👤 NUEVO SOCIO", key="btn_nuevo_c"): st.session_state.mostrar_nuevo = True
    if st.session_state.get('mostrar_nuevo'):
        with st.form("nc"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Cooperativa", data=pd.concat([df_s, new], ignore_index=True)); st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR SOCIO:")
    act = df_s if not df_s.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                m = st.number_input("Monto a pagar:", value=cuota_fija, key=f"mc_{row['ID']}")
                if st.button("✅ REGISTRAR", key=f"bc_{row['ID']}"):
                    df_s.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                    conn.update(worksheet="Cooperativa", data=df_s); st.rerun()
            with c2:
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_pintado(pd.DataFrame([row]), f"Socio {row['Nombre']}"), file_name=f"Comprobante_{row['Nombre']}.xlsx", key=f"dlc_{row['ID']}")

elif seccion == "🚑 AYUDAS ECON.":
    st.title("🚑 AYUDAS ECONÓMICAS")
    df_a = cargar("Ayudas_Listado")
    cuota_ayuda_fija = st.number_input("💵 VALOR DE APORTE ESTÁNDAR (X):", value=5.0)
    
    if not df_a.empty: st.download_button("📊 EXCEL GRUPAL AYUDAS", data=generar_excel_pintado(df_a, "AYUDAS"), file_name="Reporte_Ayudas.xlsx", use_container_width=True)
    if st.button("👤 NUEVO COMPAÑERO", key="btn_nuevo_a"): st.session_state.mostrar_nuevo = True
    if st.session_state.get('mostrar_nuevo'):
        with st.form("na"):
            n, c = st.text_input("Nombre:"), st.text_input("Cédula:")
            if st.form_submit_button("AÑADIR"):
                new = pd.DataFrame([{"ID":str(uuid.uuid4())[:5], "Nombre":n, "Cedula":c, "Saldo_Total_Aportado":0}])
                conn.update(worksheet="Ayudas_Listado", data=pd.concat([df_a, new], ignore_index=True)); st.session_state.mostrar_nuevo = False; st.rerun()

    bq = st.text_input("🔍 BUSCAR EN AYUDAS:")
    act = df_a if not df_a.empty else pd.DataFrame()
    if bq: act = act[act['Nombre'].str.contains(bq, case=False)]
    for idx, row in act.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | TOTAL APORTADO: ${row['Saldo_Total_Aportado']}"):
            c1, c2 = st.columns(2)
            with c1:
                m = st.number_input("Monto aporte:", value=cuota_ayuda_fija, key=f"ma_{row['ID']}")
                if st.button("✅ GUARDAR APORTE", key=f"ba_{row['ID']}"):
                    df_a.at[idx, "Saldo_Total_Aportado"] = float(row['Saldo_Total_Aportado']) + m
                    conn.update(worksheet="Ayudas_Listado", data=df_a); st.rerun()
            with c2:
                st.download_button(f"📊 EXCEL {row['Nombre'].split()[0]}", data=generar_excel_pintado(pd.DataFrame([row]), f"Compañero {row['Nombre']}"), file_name=f"Ayuda_{row['Nombre']}.xlsx", key=f"dla_{row['ID']}")
