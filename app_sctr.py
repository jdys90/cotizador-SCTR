import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import re
import os
from io import BytesIO
from datetime import timedelta
import smtplib
import gspread
from google.oauth2.service_account import Credentials
import requests
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y CSS PREMIUM
# ==========================================
st.set_page_config(page_title="Cotizador SCTR Digital | YQ Corredores", layout="centered", page_icon="🛡️")

st.markdown("""
    <style>
        header, [data-testid="stHeader"], [data-testid="stToolbar"] {
            display: none !important; 
            visibility: hidden !important; 
        }
        footer {visibility: hidden !important; display: none !important;}
        .stApp > footer {display: none !important;}
        [data-testid="stBottom"], [data-testid="stBottomBar"] {display: none !important;}
        [data-testid="stViewerBadge"] {display: none !important;}
        [class*="embeddedAppMetaInfoBar"], [class*="viewerBadge"] {
            display: none !important; visibility: hidden !important; opacity: 0 !important; pointer-events: none !important; height: 0px !important;
        }
        a[href*="streamlit"] {display: none !important; pointer-events: none !important;}
        button[title*="fullscreen"], button[title*="Fullscreen"], [data-testid="StyledFullScreenButton"] {display: none !important;}
        .block-container { padding-bottom: 0rem !important; }

        div.stButton > button {
            background-color: #03318e !important;
            color: white !important;
            border: 2px solid #03318e !important;
            border-radius: 8px !important;
            height: 48px !important;
        }
        div.stButton > button:hover, div.stButton > button:active, div.stButton > button:focus {
            background-color: #022266 !important;
            border-color: #022266 !important;
            color: white !important;
            box-shadow: none !important;
        }
        div.stButton > button, div.stButton > button * {
            font-size: 16px !important;
            font-weight: 700 !important;
            font-family: sans-serif !important;
            color: white !important;
        }

        div[data-testid="stDownloadButton"] button {
            background-color: #03318e !important; 
            color: white !important; 
            border: 2px solid #03318e !important;
            border-radius: 8px !important;
            height: 48px !important;
            width: 100% !important;
        }
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #022266 !important;
            border-color: #022266 !important;
        }
        div[data-testid="stDownloadButton"] button, div[data-testid="stDownloadButton"] button * {
            font-size: 16px !important;
            font-weight: 700 !important;
            font-family: sans-serif !important;
            color: white !important;
        }
    </style>
""", unsafe_allow_html=True)

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ImageRL
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
except ImportError:
    st.error("❌ Falta la librería 'reportlab'. Asegúrate de incluirla en tu requirements.txt")
    st.stop()

# --- CONSTANTES DE ACCESO ---
CODIGO_ADMIN = "ADMIN2026"
CODIGOS_ASESORES = ["ASE01", "ASE02", "ASE03", "VENTAS2026"]

def obtener_hora_peru():
    return datetime.datetime.utcnow() - timedelta(hours=5)

# ==========================================
# 2. FUNCIONES GOOGLE SHEETS Y ZOHO CRM
# ==========================================
def get_gspread_client():
    try:
        if "gcp_service_account" not in st.secrets:
            return None
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(credentials)
    except Exception as e:
        print(f"❌ Error conexión Google Sheets: {e}")
        return None

def guardar_en_sheets(datos_fila):
    try:
        client = get_gspread_client()
        if not client: return
        sheet = client.open("historial_cotizador_sctr").sheet1 
        sheet.append_row(datos_fila)
    except Exception as e:
        print(f"❌ Error al guardar en Sheets SCTR: {e}")

def descargar_historial_sheets():
    try:
        client = get_gspread_client()
        if not client: return None
        sheet = client.open("historial_cotizador_sctr").sheet1
        return pd.DataFrame(sheet.get_all_records())
    except Exception as e:
        st.error(f"Error descargando historial SCTR: {e}")
        return None

def enviar_notificacion_sctr(cliente, ruc, celular, email, rubro, actividad, trabajadores, planilla, mejor_cia, mejor_total):
    try:
        if "ZOHO_REFRESH_TOKEN" not in st.secrets: return False, "Sin secretos."
        url_auth = "https://accounts.zoho.com/oauth/v2/token"
        datos_auth = {
            "refresh_token": st.secrets["ZOHO_REFRESH_TOKEN"],
            "client_id": st.secrets["ZOHO_CLIENT_ID"],
            "client_secret": st.secrets["ZOHO_CLIENT_SECRET"],
            "grant_type": "refresh_token"
        }
        res_auth = requests.post(url_auth, data=datos_auth).json()
        access_token = res_auth.get("access_token")
        if not access_token: raise Exception("No access token")

        url_crm = "https://www.zohoapis.com/crm/v2/Leads"
        headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
        
        # Resumen técnico que se inyectará en la descripción del CRM
        desc = f"Rubro: {rubro} | Actividad: {actividad} | Trabajadores: {trabajadores} | Planilla: S/ {planilla:,.2f} | Mejor Opción: {mejor_cia} (S/ {mejor_total:,.2f})"
        
        # Mapeo exacto con los campos reales de tu CRM (Incluyendo el campo 'RUC')
        payload = {
            "data": [{
                "Company": cliente,                 # Va a la casilla "Empresa" en Zoho
                "Last_Name": "Contacto Comercial",  # Apellido obligatorio para Leads en Zoho
                "RUC": str(ruc),                    # <--- ¡Va directo al campo RUC de tu CRM!
                "Email": email,
                "Mobile": str(celular),
                "Lead_Source": "Cotizador SCTR Web",
                "Seguro_de_Inter_s": "SCTR",
                "Description": desc
            }]
        }
        
        res_crm = requests.post(url_crm, headers=headers, json=payload)
        if res_crm.status_code in [200, 201]: return True, "Éxito CRM"
        else: raise Exception(res_crm.text)
    except Exception as e:
        print(f"⚠️ Error CRM SCTR: {e}. Activando respaldo por correo...")
        try:
            if "EMAIL_PASSWORD" not in st.secrets: return False, "Fallo"
            msg = MIMEMultipart()
            msg['From'] = "administracion@yqcorredores.com"
            msg['To'] = "administracion@yqcorredores.com"
            msg['Subject'] = f"🔔 [RESPALDO] Nuevo Lead SCTR: {cliente} (RUC: {ruc})"
            cuerpo = f"Nuevo Lead SCTR Generado:\n\nContratante: {cliente}\nRUC: {ruc}\nCelular: {celular}\nCorreo: {email}\nRubro: {rubro}\nActividad: {actividad}\nTrabajadores: {trabajadores}\nPlanilla: S/ {planilla:,.2f}\nMejor Opción: {mejor_cia} (S/ {mejor_total:,.2f})"
            msg.attach(MIMEText(cuerpo, 'plain'))
            server = smtplib.SMTP("smtppro.zoho.com", 587)
            server.starttls()
            server.login("administracion@yqcorredores.com", st.secrets["EMAIL_PASSWORD"])
            server.send_message(msg)
            server.quit()
            return True, "Enviado por respaldo correo."
        except Exception as ex:
            return False, str(ex)

# ==========================================
# 3. MOTOR GENERADOR DE PDF (SLIP SCTR CON ANÁLISIS DE EXPERTO)
# ==========================================
def generar_pdf_sctr(datos, df_res, folio, razon_experto=""):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=18, bottomMargin=18)
    estilos = getSampleStyleSheet()

    AZUL_CORP = colors.HexColor("#03318e")
    VERDE_EXITO = colors.HexColor("#089685")
    DORADO_VIP = colors.HexColor("#bf8d1b")
    AZUL_CLARO = colors.HexColor("#F0F6FF")
    GRIS_TEXTO = colors.HexColor("#3c3c3c")
    GRIS_FONDO = colors.HexColor("#F8F9FA")
    BORDE_SUAVE = colors.HexColor("#DEE2E6")

    st_tit = ParagraphStyle('T', parent=estilos['Heading1'], fontName='Helvetica-Bold', fontSize=13, textColor=AZUL_CORP, leading=15)
    st_sub = ParagraphStyle('S', parent=estilos['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=AZUL_CORP, spaceAfter=4)
    st_norm = ParagraphStyle('N', parent=estilos['Normal'], fontSize=8, textColor=GRIS_TEXTO, leading=11)
    st_bold = ParagraphStyle('B', parent=st_norm, fontName='Helvetica-Bold', textColor=AZUL_CORP)
    st_th = ParagraphStyle('TH', parent=estilos['Normal'], fontSize=8, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)
    st_td = ParagraphStyle('TD', parent=estilos['Normal'], fontSize=7.5, textColor=GRIS_TEXTO, leading=10)
    st_td_c = ParagraphStyle('TDC', parent=st_td, alignment=1)
    st_rec = ParagraphStyle('REC', parent=st_td_c, fontName='Helvetica-Bold', fontSize=7.5, textColor=DORADO_VIP)

    elements = []

    img_yq = ImageRL("logo.png", width=4.0*cm, height=2.0*cm, kind='proportional') if os.path.exists("logo.png") else Paragraph("<b>YQ CORREDORES</b>", st_tit)
    p_header = Paragraph("<b>YQ CORREDORES DE SEGUROS</b><br/>Slip de Cotización SCTR (Salud y Pensión)<br/><font size='7.5' color='#666666'>Seguro Complementario de Trabajo de Riesgo - Ley N° 26790</font>", st_tit)
    fecha_str = obtener_hora_peru().strftime('%d/%m/%Y %H:%M')
    p_folio = Paragraph(f"<font color='#3c3c3c'><b>Cotización:</b> {folio}<br/><b>Fecha:</b> {fecha_str}<br/><b>Validez:</b> 15 días</font>", ParagraphStyle('F', parent=st_norm, alignment=2))

    t_head = Table([[img_yq, p_header, p_folio]], colWidths=[4.0*cm, 9.5*cm, 4.5*cm])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1, AZUL_CORP), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    elements.append(t_head)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("1. INFORMACIÓN DEL CONTRATANTE Y RIESGO COTIZADO", st_sub))
    data_cliente = [
        [Paragraph("<b>Contratante:</b>", st_bold), Paragraph(datos['razon_social'], st_norm),
         Paragraph("<b>RUC / DNI:</b>", st_bold), Paragraph(datos['ruc'], st_norm)],
        [Paragraph("<b>Rubro:</b>", st_bold), Paragraph(datos['rubro'], st_norm),
         Paragraph("<b>N° Trabajadores:</b>", st_bold), Paragraph(f"{datos['trabajadores']} asegurados", st_norm)],
        [Paragraph("<b>Actividad / CIIU:</b>", st_bold), Paragraph(datos['actividad'], st_norm),
         Paragraph("<b>Planilla Mensual:</b>", st_bold), Paragraph(f"<b>S/ {datos['planilla']:,.2f}</b>", st_norm)]
    ]
    t_cli = Table(data_cliente, colWidths=[2.8*cm, 7.7*cm, 3.0*cm, 4.5*cm])
    t_cli.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO),
        ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
        ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    elements.append(t_cli)
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("2. PROPUESTA ECONÓMICA (MONTOS FINALES CON IGV INCLUIDO)", st_sub))
    
    headers = ['Aseguradora', 'SCTR Salud c/IGV', 'SCTR Pensión c/IGV', 'Total Mensual\n(Salud + Pensión)', 'Acción / Emisión']
    anchos = [3.8*cm, 4.2*cm, 4.2*cm, 3.0*cm, 2.8*cm]
    data_tabla = [[Paragraph(h, st_th) for h in headers]]

    for idx, row in df_res.iterrows():
        es_mejor = (idx == 0)
        cia = str(row['Aseguradora'])

        celda_cia = []
        if es_mejor:
            celda_cia.append(Paragraph("<b>► MEJOR OPCIÓN</b>", st_rec))

        if "Sanitas" in cia and "Protecta" in cia:
            logos_dupla = []
            if os.path.exists("logo_sanitas.png"): logos_dupla.append(ImageRL("logo_sanitas.png", width=1.4*cm, height=0.6*cm, kind='proportional'))
            if os.path.exists("logo_protecta.png"): logos_dupla.append(ImageRL("logo_protecta.png", width=1.4*cm, height=0.6*cm, kind='proportional'))
            if logos_dupla:
                t_dupla = Table([logos_dupla], colWidths=[1.5*cm]*len(logos_dupla))
                t_dupla.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                celda_cia.append(t_dupla)
            else:
                # Solo si NO hay logos de Sanitas/Protecta, pone el texto
                celda_cia.append(Paragraph("<b>Sanitas + Protecta</b>", st_td_c))
        else:
            mapa_logos = {"Rímac": "logo_rimac.png", "Pacífico": "logo_pacifico.png", "Mapfre": "logo_mapfre.png"}
            ruta_logo = mapa_logos.get(cia, "")
            if ruta_logo and os.path.exists(ruta_logo):
                t_lg = Table([[ImageRL(ruta_logo, width=2.2*cm, height=0.7*cm, kind='proportional')]], colWidths=[3.2*cm])
                t_lg.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
                celda_cia.append(t_lg)
            else:
                # Solo si NO hay logo, pone el texto
                celda_cia.append(Paragraph(f"<font color='#03318e'><b>{cia}</b></font>", st_td_c))

        tag_min_s = " <font size='6.5' color='#666666'>(Prima Minima)</font>" if row['Aplico_Min_Salud'] else ""
        txt_salud = f"<b>S/ {row['Bruto_Salud']:,.2f}</b>{tag_min_s}"
        
        tag_min_p = " <font size='6.5' color='#666666'>(Prima Minima)</font>" if row['Aplico_Min_Pension'] else ""
        txt_pension = f"<b>S/ {row['Bruto_Pension']:,.2f}</b>{tag_min_p}"

        txt_total = f"<font color='#089685' size='9.5'><b>S/ {row['Total_Bruto']:,.2f}</b></font>"

        msg_wa = f"Hola YQ, revisé mi cotización {folio} y deseo emitir el SCTR con {cia} por un total mensual de S/ {row['Total_Bruto']:,.2f} para {datos['razon_social']}."
        link_wa = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_wa)}"
        txt_accion = f"<a href='{link_wa}' color='#089685'><font size='8'><b>► EMITIR<br/>AHORA</b></font></a>"

        data_tabla.append([
            celda_cia,
            Paragraph(txt_salud, st_td_c),
            Paragraph(txt_pension, st_td_c),
            Paragraph(txt_total, st_td_c),
            Paragraph(txt_accion, st_td_c)
        ])

    t_ofertas = Table(data_tabla, colWidths=anchos, repeatRows=1)
    estilos_tabla = [
        ('BACKGROUND', (0,0), (-1,0), AZUL_CORP),
        ('GRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 6)
    ]
    if len(df_res) > 0:
        estilos_tabla.extend([
            ('BACKGROUND', (0,1), (-1,1), AZUL_CLARO),
            ('BOX', (0,1), (-1,1), 1.5, AZUL_CORP)
        ])
    t_ofertas.setStyle(TableStyle(estilos_tabla))
    elements.append(t_ofertas)
    elements.append(Spacer(1, 6))

    algun_minimo = any(row['Aplico_Min_Salud'] or row['Aplico_Min_Pension'] for _, row in df_res.iterrows())
    if algun_minimo:
        nota_impuestos = (
            "<b>* Nota Comercial:</b> Los montos mostrados son finales e incluyen el IGV (18%) en SCTR Salud, y el Derecho de Emisión (3%) según corresponda. "
            "Cuando un importe indica <i>(Prima Mínima)</i>, significa que la planilla declarada activa la prima mínima mensual obligatoria de la aseguradora."
        )
        elements.append(Paragraph(nota_impuestos, ParagraphStyle('NI', parent=st_norm, fontSize=7, textColor=colors.HexColor("#555555"))))
        elements.append(Spacer(1, 10))

    elements.append(Paragraph("3. BENEFICIOS Y COBERTURAS PRINCIPALES", st_sub))
    cob_salud_txt = (
        "<b>SCTR SALUD (EPS):</b><br/>"
        "• <b>Atención Médica al 100%:</b> Cobertura integral sin límite de dinero ni complejidad (clínicas, medicamentos, cirugías) hasta la recuperación total.<br/>"
        "• <b>Rehabilitación:</b> Terapias, prótesis y aparatos ortopédicos necesarios.<br/>"
        "• <b>Gestión:</b> Programas de prevención y salud ocupacional."
    )
    cob_pension_txt = (
        "<b>SCTR PENSIÓN (ASEGURADORA):</b><br/>"
        "• <b>Pensión por Invalidez:</b> Ingreso mensual vitalicio para el trabajador ante accidentes de trabajo o enfermedades profesionales.<br/>"
        "• <b>Pensión de Sobrevivencia:</b> Respaldo económico asegurado para cónyuge e hijos en caso de fallecimiento.<br/>"
        "• <b>Gastos de Sepelio:</b> Reembolso de costos funerarios según los límites vigentes de la ley."
    )
    t_cob = Table([[Paragraph(cob_salud_txt, st_td), Paragraph(cob_pension_txt, st_td)]], colWidths=[9.0*cm, 9.0*cm])
    t_cob.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO),
        ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
        ('INNERGRID', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6)
    ]))
    elements.append(t_cob)
    elements.append(Spacer(1, 10))

    excl_txt = (
        "<b>EXCLUSIONES Y CONSIDERACIONES IMPORTANTES:</b><br/>"
        "• No cubre accidentes ocurridos fuera de la jornada laboral o en el trayecto habitual (casa-trabajo / <i>in itinere</i>), salvo movilidad contratada por la empresa.<br/>"
        "• Quedan excluidos siniestros por imprudencia temeraria, embriaguez o trabajadores no declarados formalmente en la planilla/trama.<br/>"
        "• <b>Vigencia:</b> La cobertura inicia con la aceptación formal de la aseguradora y el pago de la prima correspondiente (Art. 329 Ley 26702)."
    )
    t_excl = Table([[Paragraph(excl_txt, st_td)]], colWidths=[18.0*cm])
    t_excl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F0F4F8")),
        ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP),
        ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
        ('PADDING', (0,0), (-1,-1), 6)
    ]))
    elements.append(t_excl)
    elements.append(Spacer(1, 10))

    if razon_experto:
        elements.append(Paragraph("4. ANÁLISIS Y RECOMENDACIÓN DEL EXPERTO", st_sub))
        elements.append(Spacer(1, 4))
        t_box = Table([[Paragraph(f"<b>JUSTIFICACIÓN COMERCIAL:</b><br/><br/>{razon_experto}", st_norm)]], colWidths=[18.0*cm])
        t_box.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), GRIS_FONDO),
            ('LINELEFT', (0,0), (-1,-1), 2, AZUL_CORP),
            ('BOX', (0,0), (-1,-1), 0.5, BORDE_SUAVE),
            ('PADDING', (0,0), (-1,-1), 6)
        ]))
        elements.append(t_box)
        elements.append(Spacer(1, 8))

    footer_txt = "<b>YQ CORREDORES DE SEGUROS (Código SBS N5410 | RUC 10421076184):</b> <i>Te acompañamos en la emisión inmediata de constancias, inclusiones de personal en minutos y defensa técnica en siniestros laborales.</i>"
    elements.append(Paragraph(footer_txt, ParagraphStyle('FT', parent=st_norm, fontSize=7.5, textColor=AZUL_CORP, alignment=1)))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 4. MOTOR DE CÁLCULO SCTR
# ==========================================
@st.cache_data
def cargar_tarifario_sctr():
    archivos_posibles = ['Tarifario_sctr_2.xlsx', 'Tarifario_sctr.xlsx']
    ruta_activa = next((f for f in archivos_posibles if os.path.exists(f)), None)
    if not ruta_activa: return None, None
    try:
        xls = pd.ExcelFile(ruta_activa)
        df_tarifas = pd.read_excel(xls, sheet_name='Tarifario_sctr')
        df_imp = pd.read_excel(xls, sheet_name='impuestos_minimos')
        df_tarifas.columns = df_tarifas.columns.str.strip()
        df_imp.columns = df_imp.columns.str.strip()
        return df_tarifas, df_imp
    except Exception as e:
        st.error(f"Error leyendo Excel SCTR: {e}")
        return None, None

def cotizar_sctr(actividad_elegida, trabajadores, planilla):
    df_t, df_i = cargar_tarifario_sctr()
    if df_t is None or df_i is None: return pd.DataFrame()

    col_act = 'Actividades / CIIU Principales que agrupa'
    fila = df_t[df_t[col_act] == actividad_elegida]
    if fila.empty: return pd.DataFrame()
    
    fila = fila.iloc[0]
    resultados = []

    aseguradoras_map = [
        ("Sanitas - Protecta", "Sanitas Salud", "Protecta Pensión"),
        ("Pacífico", "Pacífico Salud", "Pacífico Pensión"),
        ("Rímac", "Rímac Salud", "Rímac Pensión"),
        ("Mapfre", "Mapfre Salud", "Mapfre Pensión")
    ]

    for cia_nombre, col_salud, col_pension in aseguradoras_map:
        val_s = fila.get(col_salud, None)
        val_p = fila.get(col_pension, None)
        if pd.isna(val_s) or pd.isna(val_p): continue

        tasa_salud = float(str(val_s).replace(',', '.'))
        tasa_pension = float(str(val_p).replace(',', '.'))

        imp_salud = float(df_i[col_salud].iloc[0])
        min_emision_salud = float(df_i[col_salud].iloc[1])
        min_inc_salud = float(df_i[col_salud].iloc[2])

        imp_pension = float(df_i[col_pension].iloc[0])
        min_emision_pension = float(df_i[col_pension].iloc[1])
        min_inc_pension = float(df_i[col_pension].iloc[2])

        neto_calc_salud = planilla * (tasa_salud / 100.0)
        aplico_min_salud = neto_calc_salud < min_emision_salud
        neto_salud = max(neto_calc_salud, min_emision_salud)
        bruto_salud = neto_salud * (1.0 + imp_salud)

        neto_calc_pension = planilla * (tasa_pension / 100.0)
        aplico_min_pension = neto_calc_pension < min_emision_pension
        neto_pension = max(neto_calc_pension, min_emision_pension)
        bruto_pension = neto_pension * (1.0 + imp_pension)

        total_bruto = bruto_salud + bruto_pension

        resultados.append({
            "Aseguradora": cia_nombre,
            "Tasa_Salud": tasa_salud,
            "Tasa_Pension": tasa_pension,
            "Bruto_Salud": round(bruto_salud, 2),
            "Aplico_Min_Salud": aplico_min_salud,
            "Bruto_Pension": round(bruto_pension, 2),
            "Aplico_Min_Pension": aplico_min_pension,
            "Min_Inc_Salud": round(min_inc_salud, 2),
            "Min_Inc_Pension": round(min_inc_pension, 2),
            "Total_Bruto": round(total_bruto, 2)
        })

    df_out = pd.DataFrame(resultados)
    if not df_out.empty:
        df_out = df_out.sort_values(by="Total_Bruto", ascending=True).reset_index(drop=True)
    return df_out

# ==========================================
# 5. INTERFAZ DE USUARIO Y PANEL ADMIN
# ==========================================
df_t, df_i = cargar_tarifario_sctr()
origen_url = st.query_params.get("origen", "")

if origen_url != "web":
    if os.path.exists("logo_web.png"):
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2: st.image("logo_web.png", use_container_width=True)
    st.markdown("""
        <div style="text-align: center; padding: 15px 0px;">
            <h1 style="color: #03318e; font-size: 2.3rem; font-weight: 800; margin-bottom: 8px;">
                Cotiza y emite tu SCTR en minutos.
            </h1>
            <p style="color: #3c3c3c; font-size: 1.1rem;">
                Compara tarifas oficiales de <b>Pacífico, Rímac, Mapfre y Sanitas-Protecta</b> con emisión de constancias inmediata.
            </p>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

codigo_actual = st.session_state.get('codigo_secreto', '')
es_admin = (codigo_actual == CODIGO_ADMIN or codigo_actual in CODIGOS_ASESORES)

if df_t is None:
    st.error("❌ No se encontró el archivo 'Tarifario_sctr.xlsx' en el repositorio.")
else:
    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-bottom: 10px;'>1. Datos de la Empresa / Contratante</h3>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        razon_social = st.text_input("Razón Social o Nombre Completo", placeholder="Ej: Constructora Peruana SAC")
        ruc = st.text_input("RUC o DNI", max_chars=11, placeholder="Ej: 20601234567 o 40123456")
    with col2:
        celular = st.text_input("Celular / WhatsApp", max_chars=9, placeholder="Ej: 999123456")
        email = st.text_input("Correo Electrónico", placeholder="compras@empresa.com")

    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-top: 15px; margin-bottom: 10px;'>2. Actividad de Riesgo y Planilla</h3>", unsafe_allow_html=True)
    
    col_rubro = 'Rubro Comercial (Nombre amigable para la web)'
    col_act = 'Actividades / CIIU Principales que agrupa'
    
    rubros_unicos = df_t[col_rubro].dropna().unique().tolist()
    rubro_sel = st.selectbox("Rubro Comercial", rubros_unicos)
    
    actividades_filtradas = df_t[df_t[col_rubro] == rubro_sel][col_act].dropna().unique().tolist()
    actividad_sel = st.selectbox("Actividad Económica Específica (CIIU)", actividades_filtradas)

    # --- AVISO SI NO ENCUENTRA SU ACTIVIDAD ---
    st.info("💡 ¿No encuentras la actividad exacta de tu empresa? **Contáctanos directo por WhatsApp** y te ayudaremos a cotizar al instante.")
    msg_no_encuentra = "Hola YQ, busqué mi actividad en el cotizador SCTR pero necesito asesoría para cotizar un rubro especial."
    link_no_encuentra = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_no_encuentra)}"
    st.markdown(f"<a href='{link_no_encuentra}' target='_blank' style='color: #089685; font-weight: 700; text-decoration: underline;'>👉 Clic aquí para cotizar por WhatsApp si tu actividad no figura en la lista</a>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    RMV_VIGENTE = 1230.0
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        n_trabajadores = st.number_input("Número de Trabajadores", min_value=1, max_value=5000, value=1)
    with col_t2:
        planilla_minima_legal = float(n_trabajadores * RMV_VIGENTE)
        planilla_mensual = st.number_input(
            "Planilla Total Mensual (S/)", 
            min_value=planilla_minima_legal, 
            max_value=10000000.0, 
            value=planilla_minima_legal, 
            step=100.0
        )
        st.caption(f"*Calculado sobre el sueldo mínimo vigente (S/ 1,230.00 por trabajador = S/ {planilla_minima_legal:,.2f}). Puedes digitar un monto mayor.*")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Generar Propuesta SCTR", type="primary", use_container_width=True):
        celular_limpio = "".join(filter(str.isdigit, str(celular)))
        correo_limpio = email.strip()
        ruc_limpio = ruc.strip()
        
        es_ruc = bool(re.match(r'^(10|20)\d{9}$', ruc_limpio))
        es_dni = bool(re.match(r'^\d{8}$', ruc_limpio))

        # Validaciones condicionadas: Si NO es admin, exige los datos. Si ES admin, lo deja pasar.
        if not razon_social.strip():
            st.error("⚠️ Por favor ingresa la Razón Social o Nombre del Contratante.")
        elif not es_admin and (not ruc_limpio or not (es_ruc or es_dni)):
            st.error("🏢 Ingresa un RUC válido (11 dígitos, empezando en 10 o 20) o un DNI válido (8 dígitos).")
        elif not es_admin and (len(celular_limpio) != 9 or not celular_limpio.startswith('9')):
            st.error("📱 Por favor ingresa un número de celular válido de 9 dígitos (Ej: 999123456).")
        elif not es_admin and (not correo_limpio or not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", correo_limpio)):
            st.error("📧 Por favor ingresa un correo electrónico real para enviarte tu propuesta.")
        else:
            df_res = cotizar_sctr(actividad_sel, n_trabajadores, planilla_mensual)
            st.session_state['res_sctr'] = df_res
            st.session_state['datos_sctr'] = {
                "razon_social": razon_social.strip(),
                "ruc": ruc_limpio,
                "celular": celular_limpio,
                "email": correo_limpio,
                "rubro": rubro_sel,
                "actividad": actividad_sel,
                "trabajadores": n_trabajadores,
                "planilla": planilla_mensual
            }
            st.session_state['id_sctr'] = f"SCTR-{obtener_hora_peru().strftime('%m%d-%H%M')}"
            
            if not df_res.empty:
                mejor_c = df_res.iloc[0]['Aseguradora']
                mejor_t = df_res.iloc[0]['Total_Bruto']
                enviar_notificacion_sctr(razon_social.strip(), ruc_limpio, celular_limpio, correo_limpio, rubro_sel, actividad_sel, n_trabajadores, planilla_mensual, mejor_c, mejor_t)
                
                f_log, h_log = obtener_hora_peru().strftime('%Y-%m-%d'), obtener_hora_peru().strftime('%H:%M:%S')
                guardar_en_sheets([f_log, h_log, st.session_state['id_sctr'], razon_social.strip(), ruc_limpio, celular_limpio, correo_limpio, rubro_sel, actividad_sel, n_trabajadores, planilla_mensual, mejor_c, mejor_t])

    if st.session_state.get('res_sctr') is not None:
        df_res = st.session_state['res_sctr']
        if df_res.empty:
            st.warning("No se encontraron tasas automáticas para esta actividad.")
        else:
            opciones_cias = df_res['Aseguradora'].tolist()
            
            if es_admin:
                st.success("🔓 MODO CORREDOR ACTIVADO - Configura la recomendación para el cliente:")
                mejor_opcion_elegida = st.selectbox("Selecciona la Aseguradora a destacar como 'Mejor Opción':", opciones_cias, index=0)
                razon_default = f"Este plan ofrece la tarifa más competitiva del mercado para la actividad de {st.session_state['datos_sctr']['rubro']}, optimizando los costos de planilla y cumpliendo con las exigencias de ley."
                razon_admin = st.text_area("Motivo (Análisis del Experto para el PDF):", value=razon_default)
                
                df_res['es_top'] = df_res['Aseguradora'] == mejor_opcion_elegida
                df_res = df_res.sort_values(by=['es_top', 'Total_Bruto'], ascending=[False, True]).drop(columns=['es_top']).reset_index(drop=True)
            else:
                mejor_opcion_elegida = opciones_cias[0]
                razon_admin = (
                    "Nuestro algoritmo ha analizado el mercado SCTR filtrando las tasas de todas las aseguradoras disponibles. "
                    "Esta opción representa la mejor relación costo-beneficio para proteger a tus trabajadores, garantizando "
                    "cobertura inmediata ante la autoridad laboral y optimizando el presupuesto de tu empresa."
                )

            st.markdown("---")
            st.success(f"¡Comparativa lista! Opción recomendada destacada: **{mejor_opcion_elegida}**")
            
            df_vista = df_res[['Aseguradora', 'Bruto_Salud', 'Bruto_Pension', 'Total_Bruto']].copy()
            df_vista.columns = ['Aseguradora', 'Salud c/IGV', 'Pensión c/IGV', 'Total Mensual c/IGV']
            df_vista['Salud c/IGV'] = df_vista['Salud c/IGV'].apply(lambda x: f"S/ {x:,.2f}")
            df_vista['Pensión c/IGV'] = df_vista['Pensión c/IGV'].apply(lambda x: f"S/ {x:,.2f}")
            df_vista['Total Mensual c/IGV'] = df_vista['Total Mensual c/IGV'].apply(lambda x: f"S/ {x:,.2f}")
            
            st.dataframe(df_vista, hide_index=True, use_container_width=True)

            datos = st.session_state['datos_sctr']
            folio = st.session_state['id_sctr']
            
            pdf_buffer = generar_pdf_sctr(datos, df_res, folio, razon_admin)
            nombre_clean = re.sub(r'[^a-zA-Z0-9]', '', datos['razon_social'])[:15] or "Empresa"
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.download_button(
                    label="Descargar Slip Oficial (PDF)",
                    data=pdf_buffer,
                    file_name=f"SLIP_SCTR_{nombre_clean}_{obtener_hora_peru().strftime('%d%m%y')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col_b2:
                mejor_cia = df_res.iloc[0]['Aseguradora']
                mejor_total = df_res.iloc[0]['Total_Bruto']
                msg_wa = f"Hola YQ, acabo de cotizar el SCTR ({folio}) para {datos['razon_social']} y deseo emitir la póliza con {mejor_cia} por S/ {mejor_total:,.2f}."
                enlace_wa = f"https://api.whatsapp.com/send?phone=51906462225&text={urllib.parse.quote(msg_wa)}"
                st.markdown(f"""
                    <a href='{enlace_wa}' target='_blank' style='display: flex; align-items: center; justify-content: center; width: 100%; height: 48px; background-color: #25D366; color: white; border-radius: 8px; text-decoration: none; font-weight: 700; font-size: 16px; box-sizing: border-box;'>
                        Emitir Póliza por WhatsApp
                    </a>
                """, unsafe_allow_html=True)

# ==========================================
# 6. PIE DE PÁGINA: ACCESO ADMIN
# ==========================================
st.markdown("<br><br><br><br><br><br>", unsafe_allow_html=True)

with st.expander("🛡️ Acceso Interno YQ (Solo Empleados)"):
    st.text_input("Código de Autorización", type="password", placeholder="Ingresa clave", key="codigo_secreto")
    
    if es_admin:
        st.success("✅ Modo Interno Activado. Historial de cotizaciones disponible:")
        if st.button("📥 Descargar Historial SCTR (CSV)"):
            df_hist = descargar_historial_sheets()
            if df_hist is not None and not df_hist.empty:
                csv_data = df_hist.to_csv(index=False).encode('utf-8-sig')
                st.download_button("💾 Guardar Archivo CSV", csv_data, f"Historial_SCTR_{obtener_hora_peru().strftime('%d%m%Y')}.csv", "text/csv")
            else:
                st.info("No hay registros en Google Sheets todavía.")

st.markdown("<br><br><br><br><br><br>", unsafe_allow_html=True)
