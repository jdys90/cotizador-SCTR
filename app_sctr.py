import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import re
import os
from io import BytesIO
from datetime import timedelta
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y CSS PREMIUM
# ==========================================
st.set_page_config(page_title="Cotizador SCTR Digital | YQ Corredores", layout="centered", page_icon="🛡️")

st.markdown("""
    <style>
        /* 1. Ocultar cabeceras y menú nativo */
        header, [data-testid="stHeader"], [data-testid="stToolbar"] {
            display: none !important; 
            visibility: hidden !important; 
        }
        
        /* 2. Destruir marca de agua y footers */
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

        /* 3. Botón Principal Corporativo (#03318e) */
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

        /* 4. Botón Descargar PDF (#03318e) */
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

def obtener_hora_peru():
    return datetime.datetime.utcnow() - timedelta(hours=5)

# ==========================================
# 2. MOTOR GENERADOR DE PDF (SLIP SCTR LIMPIO Y PROFESIONAL)
# ==========================================
def generar_pdf_sctr(datos, df_res, folio):
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

    # --- CABECERA ---
    img_yq = ImageRL("logo.png", width=4.0*cm, height=2.0*cm, kind='proportional') if os.path.exists("logo.png") else Paragraph("<b>YQ CORREDORES</b>", st_tit)
    p_header = Paragraph("<b>YQ CORREDORES DE SEGUROS</b><br/>Slip de Cotización SCTR (Salud y Pensión)<br/><font size='7.5' color='#666666'>Seguro Complementario de Trabajo de Riesgo - Ley N° 26790</font>", st_tit)
    fecha_str = obtener_hora_peru().strftime('%d/%m/%Y %H:%M')
    p_folio = Paragraph(f"<font color='#3c3c3c'><b>Cotización:</b> {folio}<br/><b>Fecha:</b> {fecha_str}<br/><b>Validez:</b> 15 días</font>", ParagraphStyle('F', parent=st_norm, alignment=2))

    t_head = Table([[img_yq, p_header, p_folio]], colWidths=[4.0*cm, 9.5*cm, 4.5*cm])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1, AZUL_CORP), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    elements.append(t_head)
    elements.append(Spacer(1, 10))

    # --- DATOS DEL CONTRATANTE ---
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

    # --- TABLA COMPARATIVA SIMPLIFICADA (SOLO PRECIO FINAL / BRUTO) ---
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
            celda_cia.append(Paragraph("<b>Sanitas + Protecta</b>", st_td_c))
        else:
            mapa_logos = {"Rímac": "logo_rimac.png", "Pacífico": "logo_pacifico.png", "Mapfre": "logo_mapfre.png"}
            ruta_logo = mapa_logos.get(cia, "")
            if ruta_logo and os.path.exists(ruta_logo):
                t_lg = Table([[ImageRL(ruta_logo, width=2.2*cm, height=0.7*cm, kind='proportional')]], colWidths=[3.2*cm])
                t_lg.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
                celda_cia.append(t_lg)
                celda_cia.append(Paragraph(f"<b>{cia}</b>", st_td_c))
            else:
                celda_cia.append(Paragraph(f"<font color='#03318e'><b>{cia}</b></font>", st_td_c))

        # Mostramos únicamente lo que el cliente va a pagar (Bruto c/IGV)
        tag_min_s = " <font size='6.5' color='#666666'>(Mínimo de ley)</font>" if row['Aplico_Min_Salud'] else ""
        txt_salud = f"<b>S/ {row['Bruto_Salud']:,.2f}</b>{tag_min_s}"
        
        tag_min_p = " <font size='6.5' color='#666666'>(Mínimo de ley)</font>" if row['Aplico_Min_Pension'] else ""
        txt_pension = f"<b>S/ {row['Bruto_Pension']:,.2f}</b>{tag_min_p}"

        txt_total = f"<font color='#089685' size='9.5'><b>S/ {row['Total_Bruto']:,.2f}</b></font>"

        msg_wa = f"Hola YQ, revisé mi cotización {folio} y deseo emitir el SCTR con {cia} por un total mensual de S/ {row['Total_Bruto']:,.2f} para {datos['razon_social']}."
        link_wa = f"https://wa.me/51906462225?text={urllib.parse.quote(msg_wa)}"
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

    nota_impuestos = (
        "<b>* Nota Comercial:</b> Los montos mostrados son finales e incluyen el IGV (18%) en SCTR Salud, y el Derecho de Emisión (3%) más IGV (18%) en SCTR Pensión "
        "(excepto Protecta Pensión que aplica solo 18%). Cuando un monto indica <i>(Mínimo de ley)</i>, significa que la planilla declarada activa la prima mínima mensual obligatoria de la aseguradora."
    )
    elements.append(Paragraph(nota_impuestos, ParagraphStyle('NI', parent=st_norm, fontSize=7, textColor=colors.HexColor("#555555"))))
    elements.append(Spacer(1, 10))
    # --- 3. COBERTURAS SIMPLIFICADAS Y COMERCIALES ---
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

    # --- 4. EXCLUSIONES CLAVE ---
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

    # --- 5. PIE DE PÁGINA ---
    footer_txt = "<b>YQ CORREDORES DE SEGUROS (Código SBS N5410 | RUC 10421076184):</b> <i>Te acompañamos en la emisión inmediata de constancias, inclusiones de personal en minutos y defensa técnica en caso de siniestros laborales.</i>"
    elements.append(Paragraph(footer_txt, ParagraphStyle('FT', parent=st_norm, fontSize=7.5, textColor=AZUL_CORP, alignment=1)))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 3. MOTOR DE CÁLCULO SCTR (LECTOR DE EXCEL)
# ==========================================
@st.cache_data
def cargar_tarifario_sctr():
    archivos_posibles = ['Tarifario_sctr_2.xlsx', 'Tarifario_sctr.xlsx']
    ruta_activa = next((f for f in archivos_posibles if os.path.exists(f)), None)
    if not ruta_activa:
        return None, None
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
    if df_t is None or df_i is None:
        return pd.DataFrame()

    col_act = 'Actividades / CIIU Principales que agrupa'
    fila = df_t[df_t[col_act] == actividad_elegida]
    if fila.empty:
        return pd.DataFrame()
    
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
        if pd.isna(val_s) or pd.isna(val_p):
            continue

        tasa_salud = float(str(val_s).replace(',', '.'))
        tasa_pension = float(str(val_p).replace(',', '.'))

        # Extraer Impuestos (fila 0), Prima Mínima Emisión (fila 1), Prima Mínima Inclusión (fila 2)
        imp_salud = float(df_i[col_salud].iloc[0])
        min_emision_salud = float(df_i[col_salud].iloc[1])
        min_inc_salud = float(df_i[col_salud].iloc[2])

        imp_pension = float(df_i[col_pension].iloc[0])
        min_emision_pension = float(df_i[col_pension].iloc[1])
        min_inc_pension = float(df_i[col_pension].iloc[2])

        # Cálculo Salud
        neto_calc_salud = planilla * (tasa_salud / 100.0)
        aplico_min_salud = neto_calc_salud < min_emision_salud
        neto_salud = max(neto_calc_salud, min_emision_salud)
        bruto_salud = neto_salud * (1.0 + imp_salud)

        # Cálculo Pensión
        neto_calc_pension = planilla * (tasa_pension / 100.0)
        aplico_min_pension = neto_calc_pension < min_emision_pension
        neto_pension = max(neto_calc_pension, min_emision_pension)
        bruto_pension = neto_pension * (1.0 + imp_pension)

        total_bruto = bruto_salud + bruto_pension

        resultados.append({
            "Aseguradora": cia_nombre,
            "Tasa_Salud": tasa_salud,
            "Tasa_Pension": tasa_pension,
            "Neto_Salud": round(neto_salud, 2),
            "Bruto_Salud": round(bruto_salud, 2),
            "Aplico_Min_Salud": aplico_min_salud,
            "Neto_Pension": round(neto_pension, 2),
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
# 4. INTERFAZ DE USUARIO (MODO CAMALEÓN + LOGOS)
# ==========================================
df_t, df_i = cargar_tarifario_sctr()
origen_url = st.query_params.get("origen", "")

if origen_url != "web":
    if os.path.exists("logo_web.png"):
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2:
            st.image("logo_web.png", use_container_width=True)
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

if df_t is None:
    st.error("❌ No se encontró el archivo 'Tarifario_sctr.xlsx' o 'Tarifario_sctr_2.xlsx' en el proyecto.")
else:
    st.markdown("<h3 style='color: #03318e; font-weight: 700; margin-bottom: 10px;'>1. Datos de la Empresa / Contratante</h3>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        razon_social = st.text_input("Razón Social o Nombre Completo", placeholder="Ej: Constructora Peruana SAC")
        ruc = st.text_input("RUC o DNI", max_chars=11, placeholder="Ej: 20601234567")
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

    RMV_VIGENTE = 1230.0

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        n_trabajadores = st.number_input("Número de Trabajadores", min_value=1, max_value=5000, value=1)
    with col_t2:
        # Calculamos la planilla mínima legal multiplicando N° trabajadores por S/ 1,230
        planilla_minima_legal = float(n_trabajadores * RMV_VIGENTE)
        planilla_mensual = st.number_input(
            "Planilla Total Mensual (S/)", 
            min_value=planilla_minima_legal, 
            max_value=10000000.0, 
            value=planilla_minima_legal, 
            step=100.0
        )
        st.caption(f"*Calculado sobre el sueldo mínimo vigente (S/ 1,230.00 por trabajador = S/ {planilla_minima_legal:,.2f}). Puedes digitar un monto mayor si tu planilla real es superior.*")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Generar Propuesta SCTR", type="primary", use_container_width=True):
        celular_limpio = "".join(filter(str.isdigit, str(celular)))
        correo_limpio = email.strip()
        
        if not razon_social.strip():
            st.error("⚠️ Por favor ingresa la Razón Social o Nombre del Contratante.")
        elif not ruc.strip() or len(ruc.strip()) not in [8, 11]:
            st.error("🏢 Por favor ingresa un número de RUC (11 dígitos) o DNI (8 dígitos) válido.")
        elif len(celular_limpio) != 9 or not celular_limpio.startswith('9'):
            st.error("📱 Por favor ingresa un número de celular válido de 9 dígitos (Ej: 999123456).")
        elif not correo_limpio or not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", correo_limpio):
            st.error("📧 Por favor ingresa un correo electrónico real para enviarte tu propuesta.")
        else:
            df_res = cotizar_sctr(actividad_sel, n_trabajadores, planilla_mensual)
            st.session_state['res_sctr'] = df_res
            st.session_state['datos_sctr'] = {
                "razon_social": razon_social.strip(),
                "ruc": ruc.strip() if ruc.strip() else "Por confirmar",
                "celular": celular_limpio,
                "email": email.strip(),
                "rubro": rubro_sel,
                "actividad": actividad_sel,
                "trabajadores": n_trabajadores,
                "planilla": planilla_mensual
            }
            st.session_state['id_sctr'] = f"SCTR-{obtener_hora_peru().strftime('%m%d-%H%M')}"

    if st.session_state.get('res_sctr') is not None:
        df_res = st.session_state['res_sctr']
        if df_res.empty:
            st.warning("No se encontraron tasas automáticas para esta actividad.")
        else:
            st.success(f"¡Hemos comparado el mercado! Aquí tienes las {len(df_res)} mejores opciones para tu empresa:")
            
            df_vista = df_res[['Aseguradora', 'Bruto_Salud', 'Bruto_Pension', 'Total_Bruto']].copy()
            df_vista.columns = ['Aseguradora', 'Salud c/IGV', 'Pensión c/IGV', 'Total Mensual c/IGV']
            df_vista['Salud c/IGV'] = df_vista['Salud c/IGV'].apply(lambda x: f"S/ {x:,.2f}")
            df_vista['Pensión c/IGV'] = df_vista['Pensión c/IGV'].apply(lambda x: f"S/ {x:,.2f}")
            df_vista['Total Mensual c/IGV'] = df_vista['Total Mensual c/IGV'].apply(lambda x: f"S/ {x:,.2f}")
            
            st.dataframe(df_vista, hide_index=True, use_container_width=True)

            datos = st.session_state['datos_sctr']
            folio = st.session_state['id_sctr']
            pdf_buffer = generar_pdf_sctr(datos, df_res, folio)
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
                enlace_wa = f"https://wa.me/51906462225?text={urllib.parse.quote(msg_wa)}"
                st.markdown(f"""
                    <a href='{enlace_wa}' target='_blank' style='display: flex; align-items: center; justify-content: center; width: 100%; height: 48px; background-color: #25D366; color: white; border-radius: 8px; text-decoration: none; font-weight: 700; font-size: 16px; box-sizing: border-box;'>
                        Emitir Póliza por WhatsApp
                    </a>
                """, unsafe_allow_html=True)

st.markdown("<br><br><br><br><br><br>", unsafe_allow_html=True)
