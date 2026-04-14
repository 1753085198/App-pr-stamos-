import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import uuid
import io
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from dateutil.relativedelta import relativedelta
import requests
import base64
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# 1. CONFIGURACIÓN INICIAL
st.set_page_config(page_title="SISTEMA FINANCIERO COOP SAN BLAS", page_icon="🏦", layout="wide")

# 2. CSS PARA INTERFAZ GIGANTE
st.markdown("""
    <style>
    .stMarkdown p, label, .stNumberInput label, .stTextInput label { font-size: 26px !important; font-weight: 700 !important; }
    input { font-size: 22px !important; height: 50px !important; }
    .stDownloadButton>button { font-size: 28px !important; font-weight: 800 !important; height: 5rem !important; border-radius: 15px !important; background-color: #1D6F42 !important; color: white !important; }
    .stButton>button { font-size: 24px !important; font-weight: 700 !important; border-radius: 12px !important; }
    </style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- ESTADOS GLOBALES ---
if 'pago_key' not in st.session_state: st.session_state.pago_key = 0
if 'id_abierto' not in st.session_state: st.session_state.id_abierto = None

# --- FUNCIONES CORE ---
def cargar(h):
    try:
        df = conn.read(worksheet=h, ttl=0)
        if df is not None and "ID" in df.columns:
            df["ID"] = df["ID"].astype(str).str.replace(".0", "", regex=False)
        return df if df is not None else pd.DataFrame()
    except: return pd.DataFrame()

def generar_excel_personal(row, historial, tipo):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        header_font = Font(bold=True)
        right_align = Alignment(horizontal="right")
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        p_fil = pd.DataFrame()
        if not historial.empty:
            col_id = 'ID_Socio' if 'ID_Socio' in historial.columns else 'ID_Prestamo'
            if col_id in historial.columns: p_fil = historial[historial[col_id] == row['ID']]

        if tipo == "PRÉSTAMO":
            ws1 = writer.book.create_sheet("AMORTIZACION_SAN_BLAS", 0)
            
            # Lógica de fechas
            f_inicio_str = row.get('Fecha_Inicio', datetime.now().strftime("%Y-%m-%d"))
            f_inicio = datetime.strptime(f_inicio_str, "%Y-%m-%d")
            meses_plazo = int(row.get('Meses_Totales', 1))
            f_vencimiento = f_inicio + relativedelta(months=meses_plazo)
            
            # Cálculo de Mora / Estado al día
            hoy = datetime.now()
            diff = relativedelta(hoy, f_inicio)
            meses_transcurridos = diff.years * 12 + diff.months
            meses_a_pagar = min(meses_transcurridos, meses_plazo)
            
            pagos_reales = int(row.get('Pagos_Realizados', 0))
            cuota_val = float(row.get('Cuota_Mensual', 0))
            debe_estar_pagado = meses_a_pagar * cuota_val
            lo_que_ha_pagado = pagos_reales * cuota_val
            faltante_dia = max(0, debe_estar_pagado - lo_que_ha_pagado)

            info = [
                ("COOP SAN BLAS - ESTADO DE AYUDA", ""),
                ("NOMBRE", row['Nombre']),
                ("CEDULA", row['Cedula']),
                ("FECHA INICIAL", f_inicio.strftime("%d/%m/%Y")),
                ("FECHA VENCIMIENTO", f_vencimiento.strftime("%d/%m/%Y")),
                ("PLAZO MESES", meses_plazo),
                ("CONTRIBUCION ANUAL", "8%"),
                ("VALOR AYUDA REEMBOLSABLE", round(float(row.get('Monto_Inicial', 0)), 2)),
                ("SUMA TOTAL CON INTERÉS", round(cuota_val * meses_plazo, 2)),
                ("--- ESTADO DE CUENTA A LA FECHA ---", ""),
                ("MESES TRANSCURRIDOS", meses_transcurridos),
                ("CUOTAS QUE DEBERÍA TENER", meses_a_pagar),
                ("CUOTAS PAGADAS REALES", pagos_reales),
                ("VALOR PARA ESTAR AL DÍA", f"$ {round(faltante_dia, 2)}")
            ]
            
            for idx, (label, val) in enumerate(info, start=1):
                ws1.cell(row=idx, column=2, value=label).font = header_font
                cell_val = ws1.cell(row=idx, column=3, value=val)
                if label == "CONTRIBUCION ANUAL": cell_val.alignment = right_align # 8% a la derecha
                ws1.cell(row=idx, column=2).border = thin_border
                ws1.cell(row=idx, column=3).border = thin_border

            ws1.append([])
            headers = ["N°", "FECHAS", "SALDO", "CAPITAL", "CONTRIB", "CUOTA", "PAGADO", "DIFERENCIA"]
            ws1.append(headers)
            for cell in ws1[ws1.max_row]: 
                cell.fill = header_fill; cell.font = header_font; cell.alignment = center_align; cell.border = thin_border

            # Tabla amortización simplificada tasa plana
            monto_ini = float(row.get('Monto_Inicial', 0))
            cap_m = monto_ini / meses_plazo
            int_m = (cuota_val * meses_plazo - monto_ini) / meses_plazo
            
            for i in range(1, meses_plazo + 1):
                ws1.append([i, "", "---", round(cap_m, 2), round(int_m, 2), round(cuota_val, 2), 
                            (round(cuota_val, 2) if i <= pagos_reales else 0), "---"])
                for cell in ws1[ws1.max_row]: cell.border = thin_border; cell.alignment = center_align

            for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']: ws1.column_dimensions[col].width = 22
            
    return out.getvalue()

# --- NAVEGACIÓN ---
with st.sidebar:
    st.markdown("# 🏦 COOP SAN BLAS")
    sec = st.radio("SECCIONES:", ["💰 PRÉSTAMOS", "🤝 COOPERATIVA", "🚑 AYUDAS ECON."], index=0)
    
    st.write("---")
    st.write("### 📂 IMPORTAR DATOS 2024-2026")
    archivo_antiguo = st.file_uploader("Sube tu Excel antiguo aquí", type=["xlsx", "csv"])
    if archivo_antiguo:
        st.info("Archivo cargado. Puedes copiar estos datos a tu Google Sheets para integrarlos.")

# --- MODO PRÉSTAMOS ---
if sec == "💰 PRÉSTAMOS":
    st.title("💰 GESTIÓN DE PRÉSTAMOS - SAN BLAS")
    df_p, df_h = cargar("Prestamos"), cargar("Pagos")
    
    bq_p = st.text_input("🔍 BUSCAR PRESTAMISTA (Cédula o Nombre):")
    act_p = df_p[df_p["Estado"]=="ACTIVO"] if not df_p.empty else pd.DataFrame()
    if bq_p: act_p = act_p[act_p['Nombre'].str.contains(bq_p, case=False)]

    for idx, row in act_p.iterrows():
        with st.expander(f"👤 {row['Nombre'].upper()} | SALDO: ${row['Saldo_Restante']}", expanded=(st.session_state.id_abierto == row['ID'])):
            c1, c2 = st.columns(2)
            with c1:
                st.write(f"**Cuota:** ${row['Cuota_Mensual']}")
                with st.form(key=f"fp_{row['ID']}"):
                    ft = st.file_uploader("📸 RECIBO:", key=f"ip_{row['ID']}_{st.session_state.pago_key}")
                    if st.form_submit_button("✅ REGISTRAR"):
                        if ft:
                            st.session_state.id_abierto = row['ID']
                            # Lógica de guardado...
                            st.session_state.pago_key += 1; st.rerun()
            with c2:
                st.download_button(f"📊 DESCARGAR EXCEL COMPLETO", 
                                   data=generar_excel_personal(row, df_h, "PRÉSTAMO"), 
                                   file_name=f"Reporte_SanBlas_{row['Nombre']}.xlsx")
