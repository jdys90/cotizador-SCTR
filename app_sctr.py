import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import re
import os
from io import BytesIO
from datetime import timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from google.oauth2.service_account import Credentials
import requests
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y CSS PREMIUM
# ==========================================
st.set_page_config(page_title="Cotizador SCTR Digital | YQ Corredores", layout="centered", page_icon="🏗️")

st.markdown("""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
        footer, .stApp > footer, [data-testid="stBottom"], [data-testid="stBottomBar"], [data-testid="stViewerBadge"] { display: none !important; visibility: hidden !important; }
        .block-container { padding-bottom: 0rem !important; }

        /* Botones Corporativos */
        div.stButton > button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; width: 100% !important; font-size: 16px !important; font-weight: 700 !important;}
        div.stButton > button:hover, div.stButton > button:active, div.stButton > button:focus { background-color: #022266 !important; border-color: #022266 !important; color: white !important; box-shadow: none !important; }
        div[data-testid="stDownloadButton"] button { background-color: #03318e !important; color: white !important; border: 2px solid #03318e !important; border-radius: 8px !important; height: 48px !important; width: 100% !important; font-size: 16px !important; font-weight: 700 !important;}
        div[data-testid="stDownloadButton"] button:hover { background-color: #022266 !important; border-color: #022266 !important; }
    </style>
""", unsafe_allow_html=True)

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ImageRL
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
except ImportError:
    st.error("❌ Falta la librería 'reportlab'.")
    st.stop()

# --- CONSTANTES DE ACCESO Y NEGOCIO ---
CODIGO_ADMIN = "ADMIN2026"
CODIGOS_ASESORES = ["ASE01", "ASE02", "ASE03", "VENTAS2026"]
RMV_VIGENTE = 1230.0

def obtener_hora_peru():
    return datetime.datetime.utcnow() - timedelta(hours=5)

# ==========================================
# 2. FUNCIONES GOOGLE SHEETS Y ZOHO CRM
# ==========================================
def get_gspread_client():
    try:
        if "gcp_service_account" not in st.secrets: return None
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(credentials)
    except Exception: return None

def guardar_en_sheets(datos_fila):
    try:
        client = get_gspread_client()
        if client: client.open("historial_cotizador_sctr").sheet1.append_row(datos_fila)
    except Exception: pass

def descargar_historial_sheets():
    try:
        client = get_gspread_client()
        if client: return pd.DataFrame(client.open("historial_cotizador_sctr").sheet1.get_all_records())
    except Exception: return None

def enviar_notificacion_sctr(cliente, ruc, celular, email, rubro, actividad, trabajadores, planilla, mejor_cia, mejor_total):
    try:
        if "ZOHO_REFRESH_TOKEN" not in st.secrets: return False
        url_auth = "https://accounts.zoho.com/oauth/v2/token"
        datos_auth = {"refresh_token": st.secrets["ZOHO_REFRESH_TOKEN"], "client_id": st.secrets["ZOHO_CLIENT_ID"], "client_secret": st.secrets["ZOHO_CLIENT_SECRET"], "grant_type": "refresh_token"}
        access_token = requests.post(url_auth, data=datos_auth).json().get("access_token")
        
        if access_token:
            desc = f"Rubro: {rubro} | Actividad: {actividad} | Trabajadores: {trabajadores} | Planilla: S/ {planilla:,.2f} | Mejor Opción: {mejor_cia} (S/ {mejor_total:,.2f})"
            payload = {"data": [{"Company": cliente, "Last_Name": "Contacto Comercial", "RUC": str(ruc), "Email": email, "Mobile": str(celular), "Lead_Source": "Cotizador SCTR Web", "Seguro_de_Inter_s": "SCTR", "Description": desc}]}
            requests.post("https://www.zohoapis.com/crm/v2/Leads", headers={"Authorization": f"Zoho-oauthtoken {access_token}"}, json=payload)
    except Exception: pass

# ==========================================
# 3. MOTOR DE CÁLCULO SCTR
# ==========================================
@st.cache_data
def cargar_tarifario_sctr():
    try:
        archivos_posibles = ['Tarifario_sctr_3.xlsx', 'tarifario_sctr.xlsx', 'Tarifario_sctr.xlsx']
        ruta_activa = next((f for f in archivos_posibles if os.path.exists(f)), None)
        if not ruta_activa: return None, None
        
        xls = pd.ExcelFile(ruta_activa)
        if len(xls.sheet_names) >= 2:
            df_tarifas = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
            df_imp = pd.read_excel(xls, sheet_name=xls.sheet_names[1])
        else:
            df_tarifas = pd.read_excel(xls, sheet_name=0)
            df_imp = pd.DataFrame()
            
        df_tarifas.columns = df_tarifas.columns.str.strip()
        df_imp.columns = df_imp.columns.str.strip()
        
        renombres = {
            'Rubro Comercial (Nombre amigable para la web)': 'Rubro', 'Actividades / CIIU Principales que agrupa': 'Actividad',
            'Rímac Salud': 'Tasa_Salud_Rimac', 'Rímac Pensión': 'Tasa_Pension_Rimac',
            'Pacífico Salud': 'Tasa_Salud_Pacifico', 'Pacífico Pensión': 'Tasa_Pension_Pacifico',
            'Mapfre Salud': 'Tasa_Salud_Mapfre', 'Mapfre Pensión': 'Tasa_Pension_Mapfre',
            'Sanitas Salud': 'Tasa_Salud_Sanitas', 'Protecta Pensión': 'Tasa_Pension_Protecta'
        }
        df_tarifas = df_tarifas.rename(columns=renombres)
        df_tarifas['Rubro'] = df_tarifas['Rubro'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_tarifas['Actividad'] = df_tarifas['Actividad'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        return df_tarifas, df_imp
    except Exception as e:
        return None, None

def cotizar_sctr(df_t, df_i, actividad_elegida, planilla):
    if df_t is None or df_t.empty: return pd.DataFrame()
    fila = df_t[df_t['Actividad'] == actividad_elegida]
    if fila.empty: return pd.DataFrame()
    row = fila.iloc[0]

    def get_val(df_m, concepto, col_name, default=0.0):
        if df_m is None or df_m.empty: return default
        try:
            col_match = next((c for c in df_m.columns if str(col_name) in str(c)), None)
            if not col_match: return default
            val = df_m.loc[df_m['Concepto'].str.contains(concepto, case=False, na=False), col_match].values[0]
            if isinstance(val, str) and '%' in val: return float(val.replace('%', '').replace(',', '.')) / 100.0
            return float(str(val).replace(',', '.'))
        except: return default

    cias = [
        ("Rímac", "Tasa_Salud_Rimac", "Rímac Salud", "Tasa_Pension_Rimac", "Rímac Pens"),
        ("Pacífico", "Tasa_Salud_Pacifico", "Pacífico Salu", "Tasa_Pension_Pacifico", "Pacífico Pens"),
        ("Mapfre", "Tasa_Salud_Mapfre", "Mapfre Salu", "Tasa_Pension_Mapfre", "Mapfre Pensi"),
        ("Sanitas - Protecta", "Tasa_Salud_Sanitas", "Sanitas Salu", "Tasa_Pension_Protecta", "Protecta Pens")
    ]
    
    resultados = []
    for nombre, col_ts, col_hoja2_s, col_tp, col_hoja2_p in cias:
        if col_ts in row and pd.notna(row[col_ts]) and col_tp in row and pd.notna(row[col_tp]):
            t_sal, t_pen = float(str(row[col_ts]).replace(',', '.')), float(str(row[col_tp]).replace(',', '.'))
            factor_s, factor_p = (t_sal / 100.0) if t_sal > 0.1 else t_sal, (t_pen / 100.0) if t_pen > 0.1 else t_pen
            
            min_s, min_p = get_val(df_i, 'Prima Minima Emision', col_hoja2_s, 60.0), get_val(df_i, 'Prima Minima Emision', col_hoja2_p, 60.0)
            imp_s, imp_p = get_val(df_i, 'Impuestos', col_hoja2_s, 0.18), get_val(df_i, 'Impuestos', col_hoja2_p, 0.2154)
            if imp_s > 1: imp_s = imp_s / 100.0
            if imp_p > 1: imp_p = imp_p / 100.0
            
            neto_sal, neto_pen = max(planilla * factor_s, min_s), max(planilla * factor_p, min_p)
            bruto_sal, bruto_pen = round(neto_sal * (1 + imp_s), 2), round(neto_pen * (1 + imp_p), 2)
            
            resultados.append({
                "Aseguradora": nombre, "Tasa_Salud": t_sal, "Tasa_Pension": t_pen,
                "Bruto_Salud": bruto_sal, "Aplico_Min_Salud": (planilla * factor_s) < min_s,
                "Bruto_Pension": bruto_pen, "Aplico_Min_Pension": (planilla * factor_p) < min_p,
                "Total_Bruto": round(bruto_sal + bruto_pen, 2)
            })
            
    df_out = pd.DataFrame(resultados)
    if not df_out.empty: df_out = df_out.sort_values(by="Total_Bruto", ascending=True).reset_index(drop=True)
    return df_out

# ==========================================
# 4. MOTOR GENERADOR DE PDF SCTR
# ==========================================
def generar_pdf_sctr(datos, df_res, folio, razon_experto=""):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=16, bottomMargin=16)
    estilos = getSampleStyleSheet()

    AZUL_CORP, GRIS_TEXTO, GRIS_FONDO, BORDE_SUAVE = colors.HexColor("#03318e"), colors.HexColor("#3c3c3c"), colors.HexColor("#F8F9FA"), colors.HexColor("#DEE2E6")
    st_tit = ParagraphStyle('T', parent=estilos['Heading1'], fontName='Helvetica-Bold', fontSize=13, textColor=AZUL_CORP, leading=15)
    st_sub = ParagraphStyle('S', parent=estilos['Normal'], fontName='Helvetica-Bold', fontSize=9.5, textColor=AZUL_CORP, spaceAfter=4)
    st_norm = ParagraphStyle('N', parent=estilos['Normal'], fontSize=8, textColor=GRIS_TEXTO, leading=10.5)
    st_bold = ParagraphStyle('B', parent=st_norm, fontName='Helvetica-Bold', textColor=AZUL_CORP)
    st_th = ParagraphStyle('TH', parent=estilos['Normal'], fontSize=7.5, fontName='Helvetica-Bold', textColor=colors.white, alignment=1, leading=9)
    st_td = ParagraphStyle('TD', parent=estilos['Normal'], fontSize=7.5, textColor=GRIS_TEXTO, leading=9.5)
    st_td_c = ParagraphStyle('TDC', parent=st_td, alignment=1)
    st_rec = ParagraphStyle('REC', parent=st_td_c, fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor("#bf8d1b"))

    elements = []

    # --- 1. CABECERA ---
    img_yq = ImageRL("logo.png", width=3.8*cm, height=1.9*cm, kind='proportional') if os.path.exists("logo.png") else Paragraph("<b>YQ CORREDORES</b>", st_tit)
    txt_header = "<b>YQ CORREDORES DE SEGUROS</b><br/>Slip de Cotización Seguro SCTR Salud y Pensión<br/><font size='7.5' color='#666666'>Ley N° 26790 / Decreto Legislativo N° 688</font>"
    
    p_header = Paragraph(txt_header, st_tit)
    fecha_str = obtener_hora_peru().strftime('%d/%m/%Y %H:%M')
    p_folio = Paragraph(f"<font color='#3c3c3c'><b>Cotización:</b> {folio}<br/><b>Fecha:</b> {fecha_str}<br/><b>Validez:</b> 15 días</font>", ParagraphStyle('F', parent=st_norm, alignment=2))

    t_head = Table([[img_yq, p_header, p_folio]], colWidths=[3.8*cm, 9.7*cm, 4.5*cm])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1, AZUL_CORP), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    elements.append(t_head)
    elements.append(Spacer(1, 8))

    # --- 2. DATOS DEL CONTRATANTE ---
    elements.append(Paragraph("1. INFORMACIÓN DEL CONTRATANTE Y RIESGO ASEGURABLE", st_sub))
    data_cliente = [
        [Paragraph("<b>Contratante:</b>", st_bold), Paragraph(datos['razon_social'], st_norm), Paragraph("<b>RUC / DNI:</b>", st_bold), Paragraph(datos['ruc'], st_norm)],
        [Paragraph("<b>Rubro:</b>", st_bold), Paragraph(datos['rubro'], st_norm), Paragraph("<b>Trabajadores:</b>", st_bold), Paragraph(f"{datos['trabajadores']} asegurados", st_norm)],
        [Paragraph("<b>Actividad / CIIU:</b>", st_bold), Paragraph(datos['actividad'], st_norm), Paragraph("<b>Planilla Mensual:</b>", st_bold), Paragraph(f"<b>S/ {datos['planilla']:,.2f}</b>", st_norm)]
    ]
    t_cli = Table(data_cliente, colWidths=[3.0*cm, 7.5*cm, 3.0*cm, 4.5*cm])
    t_cli.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 4)]))
    elements.extend([t_cli, Spacer(1, 10)])

    # --- 3. PROPUESTA ECONÓMICA ---
    elements.append(Paragraph("2. PROPUESTA ECONÓMICA SCTR (MONTOS MENSUALES FINALES CON IGV)", st_sub))
    headers = ['Aseguradora', 'SCTR Salud c/IGV', 'SCTR Pensión c/IGV', 'Total Mensual\n(Salud + Pensión)', 'Proyección Anual\n(12 meses)', 'Acción / Emisión']
    
    anchos = [3.4*cm, 3.0*cm, 3.0*cm, 3.0*cm, 3.0*cm, 2.6*cm]
    data_tabla = [[Paragraph(h, st_th) for h in headers]]

    for idx, row in df_res.iterrows():
        es_mejor = (idx == 0)
        cia = str(row['Aseguradora'])
        celda_cia = []
        if es_mejor: celda_cia.append(Paragraph("<b>► MEJOR OPCIÓN</b>", st_rec))

        mapa_logos = {"Rímac": "logo_rimac.png", "Pacífico": "logo_pacifico.png", "Mapfre": "logo_mapfre.png", "Sanitas - Protecta": "logo_sanitas_protecta.png"}
        ruta_logo = mapa_logos.get(cia, "")
        if ruta_logo and os.path.exists(ruta_logo):
            t_lg = Table([[ImageRL(ruta_logo, width=2.3*cm, height=0.75*cm, kind='proportional')]], colWidths=[3.0*cm])
            t_lg.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            celda_cia.append(t_lg)
        else:
            celda_cia.append(Paragraph(f"<font color='#03318e'><b>{cia}</b></font>", st_td_c))

        tag_min_s = "<br/><font size='6' color='#666666'>(Prima Minima)</font>" if row.get('Aplico_Min_Salud') else ""
        txt_m = f"<b>S/ {row['Bruto_Salud']:,.2f}</b>{tag_min_s}"
        tag_min_p = "<br/><font size='6' color='#666666'>(Prima Minima)</font>" if row.get('Aplico_Min_Pension') else ""
        txt_t = f"<b>S/ {row['Bruto_Pension']:,.2f}</b>{tag_min_p}"
        txt_s = f"<font color='#089685' size='9'><b>S/ {row['Total_Bruto']:,.2f}</b></font>"
        txt_a = f"<b>S/ {row.get('Total_Anual', row['Total_Bruto']*12):,.2f}</b>"
        msg_wa = f"Hola YQ, revisé mi cotización {folio} de SCTR para {datos['razon_social']} y deseo emitir con {cia} por S/ {row['Total_Bruto']:,.2f} al mes."

        link_wa = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_wa)}"
        txt_accion = f"<a href='{link_wa}' color='#089685'><font size='7.5'><b>► EMITIR<br/>PÓLIZA</b></font></a>"
        data_tabla.append([celda_cia, Paragraph(txt_m, st_td_c), Paragraph(txt_t, st_td_c), Paragraph(txt_s, st_td_c), Paragraph(txt_a, st_td_c), Paragraph(txt_accion, st_td_c)])
    
    t_ofertas = Table(data_tabla, colWidths=anchos, repeatRows=1)
    estilos_tabla = [('BACKGROUND', (0,0), (-1,0), AZUL_CORP), ('GRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 5)]
    if len(df_res) > 0: estilos_tabla.extend([('BACKGROUND', (0,1), (-1,1), AZUL_CLARO), ('BOX', (0,1), (-1,1), 1.5, AZUL_CORP)])
    t_ofertas.setStyle(TableStyle(estilos_tabla))
    elements.extend([t_ofertas, Spacer(1, 8)])

    # --- 4. RECOMENDACIÓN DEL EXPERTO (PUNTO 3) ---
    if razon_experto:
        elements.append(Paragraph("3. ANÁLISIS Y RECOMENDACIÓN DEL EXPERTO", st_sub))
        elements.append(Spacer(1, 3))
        t_box = Table([[Paragraph(f"<b>ESTRATEGIA RECOMENDADA:</b><br/>{razon_experto}", st_norm)]], colWidths=[18.0*cm])
        t_box.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 6)]))
        elements.extend([t_box, Spacer(1, 8)])

    algun_minimo = any(row.get('Aplico_Min_Salud') or row.get('Aplico_Min_Pension') for _, row in df_res.iterrows())
    if algun_minimo:
        nota_impuestos = "<b>* Nota Comercial:</b> Los montos mostrados son finales e incluyen el IGV (18%) en SCTR Salud, y el Derecho de Emisión (3%) en Pensión. Cuando un importe indica <i>(Prima Mínima)</i>, significa que la planilla declarada activa la prima mínima mensual obligatoria."
        elements.extend([Paragraph(nota_impuestos, ParagraphStyle('NM', parent=st_norm, fontSize=7, textColor=colors.HexColor("#555555"))), Spacer(1, 8)])

    # --- 5. COBERTURAS Y REGLAS (PUNTO 4) ---
    elements.append(Paragraph("4. BENEFICIOS Y COBERTURAS PRINCIPALES", st_sub))
    cob_salud_txt = "<b>SCTR SALUD (EPS):</b><br/>• <b>Atención Médica al 100%:</b> Cobertura integral sin límite de dinero ni complejidad hasta la recuperación total.<br/>• <b>Rehabilitación:</b> Terapias, prótesis y aparatos ortopédicos necesarios.<br/>• <b>Gestión:</b> Programas de prevención y salud ocupacional."
    cob_pension_txt = "<b>SCTR PENSIÓN (ASEGURADORA):</b><br/>• <b>Pensión por Invalidez:</b> Ingreso mensual vitalicio para el trabajador ante accidentes de trabajo o enfermedades profesionales.<br/>• <b>Pensión de Sobrevivencia:</b> Respaldo económico asegurado para cónyuge e hijos en caso de fallecimiento.<br/>• <b>Gastos de Sepelio:</b> Reembolso de costos funerarios según los límites vigentes de la ley."
    t_cob = Table([[Paragraph(cob_salud_txt, st_td), Paragraph(cob_pension_txt, st_td)]], colWidths=[9.0*cm, 9.0*cm])
    t_cob.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 6)]))
    elements.extend([t_cob, Spacer(1, 8)])

    excl_txt = "<b>EXCLUSIONES Y CONSIDERACIONES IMPORTANTES:</b><br/>• No cubre accidentes ocurridos fuera de la jornada laboral o en el trayecto habitual (casa-trabajo), salvo movilidad contratada por la empresa.<br/>• Quedan excluidos siniestros por imprudencia temeraria, embriaguez o trabajadores no declarados formalmente en la planilla.<br/>• <b>Vigencia:</b> La cobertura inicia con la aceptación formal de la aseguradora y el pago de la prima correspondiente."
    t_excl = Table([[Paragraph(excl_txt, st_td)]], colWidths=[18.0*cm])
    t_excl.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F0F4F8")), ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP), ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE), ('PADDING', (0,0), (-1,-1), 6)]))
    elements.extend([t_excl, Spacer(1, 8)])

    footer_txt = "<b>YQ CORREDORES DE SEGUROS (Código SBS N5410 | RUC 10421076184):</b> <i>Te acompañamos en la emisión inmediata de constancias, inclusiones de personal y defensa ante siniestros.</i>"
    elements.append(Paragraph(footer_txt, ParagraphStyle('FT', parent=st_norm, fontSize=7.5, textColor=AZUL_CORP, alignment=1)))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 6. INTERFAZ DE USUARIO (STREAMLIT)
# ==========================================
df_t, df_i = cargar_tarifario_sctr()
origen_url = st.query_params.get("origen", "")

if origen_url != "web":
    if os.path.exists("logo_web.png"):
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2: st.image("logo_web.png", use_container_width=True)
    st.markdown("""
        <div style="text-align: center; padding: 10px 0px;">
            <h1 style="color: #03318e; font-size: 2.3rem; font-weight: 800; margin-bottom: 5px;">Cotiza tu SCTR al instante.</h1>
            <p style="color: #3c3c3c; font-size: 1.1rem;">Compara Pacífico, Rímac, Mapfre y Protecta.</p>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

codigo_actual = st.session_state.get('codigo_secreto', '')
es_admin = (codigo_actual == CODIGO_ADMIN or codigo_actual in CODIGOS_ASESORES)

if df_t is None:
    st.error("❌ No se encontró el archivo 'Tarifario_sctr.xlsx'.")
else:
    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-bottom: 10px;'>1. Datos de la Empresa</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: razon_social, ruc = st.text_input("Razón Social", placeholder="Constructora SAC"), st.text_input("RUC o DNI", max_chars=11)
    with col2: celular, email = st.text_input("Celular", max_chars=9), st.text_input("Correo", placeholder="mail@empresa.com")

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-top: 15px;'>2. Actividad de Riesgo y Planilla</h3>", unsafe_allow_html=True)
    rubro_sel = st.selectbox("Rubro Comercial", df_t['Rubro'].dropna().unique().tolist())
    actividad_sel = st.selectbox("Actividad Económica Específica (CIIU)", df_t[df_t['Rubro'] == rubro_sel]['Actividad'].dropna().unique().tolist())
    
    col_t1, col_t2 = st.columns(2)
    with col_t1: n_trabajadores = st.number_input("N° Trabajadores", min_value=1, max_value=5000, value=1)
    with col_t2: planilla_mensual = st.number_input("Planilla Total Mensual (S/)", min_value=float(n_trabajadores*RMV_VIGENTE), value=float(n_trabajadores*RMV_VIGENTE), step=100.0)
    st.caption(f"*Calculado sobre el sueldo mínimo (S/ {n_trabajadores*RMV_VIGENTE:,.2f}).*")

    st.markdown("<br>", unsafe_allow_html=True)
    
    # ---------------------------------------------------------
    # PASO 1 UX: BOTÓN PRINCIPAL (SOLO SCTR)
    # ---------------------------------------------------------
    if st.button("Generar Propuesta SCTR", type="primary", use_container_width=True):
        cel_cl, ruc_cl, mail_cl = "".join(filter(str.isdigit, str(celular))), ruc.strip(), email.strip()
        es_ruc, es_dni = bool(re.match(r'^(10|20)\d{9}$', ruc_cl)), bool(re.match(r'^\d{8}$', ruc_cl))

        if not razon_social.strip(): st.error("⚠️ Ingresa la Razón Social.")
        elif not es_admin and (not ruc_cl or not (es_ruc or es_dni)): st.error("🏢 RUC/DNI inválido.")
        elif not es_admin and (len(cel_cl) != 9 or not cel_cl.startswith('9')): st.error("📱 Número de celular inválido.")
        elif not es_admin and (not mail_cl or not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", mail_cl)): st.error("📧 Correo inválido.")
        else:
            df_res = cotizar_sctr(df_t, df_i, actividad_sel, planilla_mensual)
            st.session_state['res_sctr'], st.session_state['id_sctr'] = df_res, f"SCTR-{obtener_hora_peru().strftime('%m%d-%H%M')}"
            st.session_state['datos_sctr'] = {"razon_social": razon_social.strip(), "ruc": ruc_cl, "celular": cel_cl, "email": mail_cl, "rubro": rubro_sel, "actividad": actividad_sel, "trabajadores": n_trabajadores, "planilla": planilla_mensual}
            
            if not df_res.empty and not es_admin:
                enviar_notificacion_sctr(razon_social.strip(), ruc_cl, cel_cl, mail_cl, rubro_sel, actividad_sel, n_trabajadores, planilla_mensual, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Total_Bruto'])
                guardar_en_sheets([obtener_hora_peru().strftime('%Y-%m-%d'), obtener_hora_peru().strftime('%H:%M:%S'), st.session_state['id_sctr'], razon_social.strip(), ruc_cl, cel_cl, mail_cl, rubro_sel, actividad_sel, n_trabajadores, planilla_mensual, df_res.iloc[0]['Aseguradora'], df_res.iloc[0]['Total_Bruto']])

    if st.session_state.get('res_sctr') is not None:
        df_res = st.session_state['res_sctr'].copy()
        datos = st.session_state['datos_sctr']
        folio = st.session_state['id_sctr']
        
        if df_res.empty: st.warning("No hay tasas para esta actividad.")
        else:
            opciones_cias = df_res['Aseguradora'].tolist()
            
            if es_admin:
                st.success("🔓 MODO CORREDOR ACTIVADO - Configura la recomendación para el cliente:")
                mejor_opcion_elegida = st.selectbox("Selecciona la Aseguradora a destacar:", opciones_cias, index=0)
                df_res['es_top'] = df_res['Aseguradora'] == mejor_opcion_elegida
                df_res = df_res.sort_values(by=['es_top', 'Total_Bruto'], ascending=[False, True]).drop(columns=['es_top']).reset_index(drop=True)
                razon_default = f"Este plan de {mejor_opcion_elegida} ofrece la tarifa más competitiva del mercado para la actividad de {datos['rubro']}, optimizando los costos de planilla y cumpliendo con las exigencias de ley."
                razon_admin = st.text_area("Motivo (Análisis del Experto para el PDF):", value=razon_default)
            else:
                mejor_opcion_elegida = opciones_cias[0]
                razon_admin = f"Nuestro algoritmo ha analizado el mercado SCTR filtrando las tasas de todas las aseguradoras disponibles. Destacamos a {mejor_opcion_elegida} como la opción de mejor relación costo-beneficio para proteger a tus trabajadores, garantizando cobertura inmediata ante la autoridad laboral y optimizando el presupuesto de tu empresa."

            st.markdown("---")
            st.success(f"¡Comparativa SCTR lista! Opción recomendada destacada: **{mejor_opcion_elegida}**")
            df_res['Total_Anual'] = df_res['Total_Bruto'] * 12
            df_vista = df_res[['Aseguradora', 'Bruto_Salud', 'Bruto_Pension', 'Total_Bruto', 'Total_Anual']].copy()
            df_vista.columns = ['Aseguradora', 'Salud c/IGV', 'Pensión c/IGV', 'Total Mensual', 'Total Anual (Proyección)']
            for c in df_vista.columns[1:]: df_vista[c] = df_vista[c].apply(lambda x: f"S/ {x:,.2f}")
            st.dataframe(df_vista, hide_index=True, use_container_width=True)

            pdf_buffer = generar_pdf_sctr(datos, df_res, folio, razon_admin)
            nombre_clean = re.sub(r'[^a-zA-Z0-9]', '', datos['razon_social'])[:15] or "Empresa"

            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.download_button(
                    label="Descargar Slip SCTR Oficial (PDF)",
                    data=pdf_buffer, file_name=f"SLIP_SCTR_{nombre_clean}_{obtener_hora_peru().strftime('%d%m%y')}.pdf",
                    mime="application/pdf", use_container_width=True, key="btn_down_solo_sctr"
                )
            with col_b2:
                mejor_total = df_res.iloc[0]['Total_Bruto']
                msg_wa = f"Hola YQ, acabo de cotizar el SCTR ({folio}) para {datos['razon_social']} y deseo emitir la póliza con {mejor_opcion_elegida} por S/ {mejor_total:,.2f}."
                st.markdown(f"<a href='https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_wa)}' target='_blank' style='display: flex; align-items: center; justify-content: center; width: 100%; height: 48px; background-color: #25D366; color: white; border-radius: 8px; text-decoration: none; font-weight: 700; font-size: 16px;'>Emitir SCTR por WhatsApp</a>", unsafe_allow_html=True)

            # ---------------------------------------------------------
            # PASO 2 UX MOBILE: BUMP COMERCIAL (SOLICITUD AL ASESOR)
            # ---------------------------------------------------------
            st.markdown(f"""
                <div style="background: linear-gradient(135deg, #F0F6FF 0%, #E6EEFF 100%); border: 2px solid #03318e; border-radius: 12px; padding: 20px; margin-top: 28px; margin-bottom: 12px;">
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <span style="font-size: 28px;">🛡️</span>
                        <div>
                            <h3 style="color: #03318e; margin: 0; font-weight: 800; font-size: 1.15rem;">¿Tu empresa también necesita el Seguro Vida Ley obligatorio?</h3>
                            <p style="color: #3c3c3c; font-size: 0.92rem; margin: 4px 0 0 0;">Solicítalo a tu asesor experto al momento de emitir tu SCTR por WhatsApp y obtén las mejores tarifas en tu <b>Pack Corporativo</b>.</p>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            msg_wa_pack = f"Hola YQ, acabo de cotizar mi SCTR ({folio}) para {datos['razon_social']} y también me interesa conocer el precio para agregar el Seguro de Vida Ley al paquete."
            st.markdown(f"<a href='https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_wa_pack)}' target='_blank' style='display: flex; align-items: center; justify-content: center; width: 100%; height: 48px; background-color: #089685; color: white; border-radius: 8px; text-decoration: none; font-weight: 700; font-size: 16px;'>📲 Consultar Pack (Vida Ley + SCTR) por WhatsApp</a>", unsafe_allow_html=True)

with st.expander("🛡️ Acceso Interno YQ (Solo Empleados)"):
    st.text_input("Código de Autorización", type="password", key="codigo_secreto")
    if es_admin and st.button("📥 Descargar Historial SCTR (CSV)"):
        df_hist = descargar_historial_sheets()
        if df_hist is not None and not df_hist.empty: st.download_button("💾 Guardar CSV", df_hist.to_csv(index=False).encode('utf-8-sig'), f"Historial_SCTR.csv", "text/csv")
