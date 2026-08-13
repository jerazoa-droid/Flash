import os
import re
import streamlit as st
import anthropic
from dotenv import load_dotenv

# Importaciones de ReportLab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from io import BytesIO

load_dotenv()
api_key = os.getenv("ANTHROPIC_API_KEY")

st.set_page_config(page_title="FlashClass - Guías Académicas", layout="centered")

# --- CSS (Ocultar Sidebar) ---
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none; }
        [data-testid="collapsedControl"] { display: none; }
    </style>
""", unsafe_allow_html=True)

# Botón volver
if st.button("⬅️ Volver al Inicio"):
    st.switch_page("app.py")

# --- MOTOR DE PDF CON REPORTLAB ---
def formatear_texto_reportlab(texto):
    # 1. Filtro Anti-LaTeX: Limpiar símbolos matemáticos residuales
    texto = texto.replace("$", "") # Elimina signos de dólar
    texto = texto.replace("\\Delta", "Δ").replace("\\eta", "η").replace("\\approx", "≈")
    texto = texto.replace("\\cdot", "·").replace("\\times", "x").replace("\\quad", " ")
    texto = texto.replace("\\text", "").replace("\\{", "(").replace("\\}", ")")
    texto = texto.replace("\\frac", "") # Evitar que imprima el comando de fracción
    texto = texto.replace("\\", "") # Elimina barras invertidas restantes
    
    # 2. Escapar caracteres XML para ReportLab
    texto = texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    
    # 3. Formato Markdown (Negritas y Cursivas)
    texto = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', texto)
    texto = re.sub(r'\*(.*?)\*', r'<i>\1</i>', texto)
    return texto

def crear_guia_pdf_reportlab(texto_guia, asignatura, tema, tipo):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=54, leftMargin=54, topMargin=54, bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    titulo_style = ParagraphStyle(
        'TituloGuia', parent=styles['Heading1'],
        fontName='Helvetica-Bold', fontSize=18, leading=22,
        alignment=1, textColor=colors.HexColor('#1A237E')
    )
    
    subtitulo_style = ParagraphStyle(
        'SubtituloGuia', parent=styles['Heading2'],
        fontName='Helvetica-Bold', fontSize=13, leading=17,
        textColor=colors.HexColor('#0D47A1'), spaceBefore=14, spaceAfter=6
    )
    
    cuerpo_style = ParagraphStyle(
        'CuerpoGuia', parent=styles['Normal'],
        fontName='Helvetica', fontSize=10, leading=15,
        textColor=colors.HexColor('#212121'), spaceAfter=6
    )
    
    lista_style = ParagraphStyle(
        'ListaGuia', parent=cuerpo_style,
        leftIndent=15, spaceAfter=3
    )

    story = []

    # Encabezado del documento limpio
    story.append(Paragraph(f"{tema}", titulo_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"<b>Asignatura:</b> {asignatura} | <i>{tipo}</i> (FlashClass)", ParagraphStyle('Sub', parent=cuerpo_style, alignment=1, fontName='Helvetica-Oblique', textColor=colors.gray)))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#BDBDBD'), spaceAfter=12))

    for linea in texto_guia.split("\n"):
        linea_limpia = linea.strip()
        
        if not linea_limpia:
            story.append(Spacer(1, 4))
            continue
        
        if linea_limpia.startswith("## ") or linea_limpia.startswith("# "):
            texto_limpio = linea_limpia.replace("#", "").strip()
            story.append(Paragraph(formatear_texto_reportlab(texto_limpio), subtitulo_style))
        elif linea_limpia.startswith("- ") or linea_limpia.startswith("* ") or re.match(r'^[a-z]\)', linea_limpia):
            story.append(Paragraph(f"• {formatear_texto_reportlab(linea_limpia[2:].strip())}", lista_style))
        else:
            story.append(Paragraph(formatear_texto_reportlab(linea_limpia), cuerpo_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def redactar_guia_investigada(asignatura, tema, nivel, tipo):
    client = anthropic.Anthropic(api_key=api_key)
    
    if "Teorico-Practico" in tipo:
        enfoque_prompt = """
        ENFOQUE: Capítulo Teórico-Práctico de Manual Universitario.
        - 70% del contenido: Explicación conceptual profunda, deducción lógica de los principios y contexto de ingeniería.
        - 30% del contenido: Un único problema de aplicación rigurosamente desarrollado paso a paso.
        """
    else:
        enfoque_prompt = """
        ENFOQUE: Guía Intensiva de Problemas Aplicados (Taller / Práctica).
        - 15% del contenido: Resumen breve de fórmulas clave y modelos analíticos.
        - 85% del contenido: Batería masiva de ejercicios. Incluye 2 problemas resueltos paso a paso y al menos 5 PROBLEMAS PROPUESTOS adicionales con sus respectivas respuestas numéricas finales.
        """

    prompt = f"""
    Actúa como un autor y profesor de libros de texto universitarios. Redacta un documento formal sobre '{tema}' 
    para la asignatura '{asignatura}'. Nivel solicitado: {nivel}.
    
    {enfoque_prompt}
    
    INSTRUCCIONES CRÍTICAS Y ESTRICTAS DE FORMATO MATEMÁTICO:
    - ESTÁ ESTRICTAMENTE PROHIBIDO utilizar código LaTeX. NO uses el símbolo de dólar bajo ninguna circunstancia.
    - NO utilices comandos con barras invertidas (como \\frac, \\Delta, \\eta).
    - Escribe todas las fórmulas en TEXTO PLANO lineal.
    - Usa palabras o símbolos de texto básicos. (Ejemplo correcto: "Eficiencia = 1 - (T_frio / T_caliente)", "Delta U = Q - W").
    
    INSTRUCCIONES GENERALES:
    - NO coloques títulos genéricos como "Marco Teórico". Inicia directamente explicando el tema de forma orgánica.
    - Utiliza negritas (**texto**) para conceptos clave.
    - Estructura con '##' para cada sección y finaliza con la '## Bibliografía Consultada (Formato APA)'.
    """

    with st.spinner("FlashClass generando capítulo académico y maquetando PDF..."):
        try:
            resp = client.messages.create(
                model="claude-sonnet-5",
                max_tokens=8192,
                messages=[{"role": "user", "content": prompt}]
            )
            texto = "".join(b.text for b in resp.content if hasattr(b, "text"))
            return crear_guia_pdf_reportlab(texto, asignatura, tema, tipo)
        except Exception as e:
            st.error(f"Error en la generación: {e}")
            return None

# --- INTERFAZ ---
st.title("⚡ FlashClass - Generador Académico")
with st.form("form_guia"):
    asig = st.text_input("Asignatura", value="Ingeniería Civil")
    tema = st.text_input("Tema", value="Leyes de la Termodinámica")
    nivel = st.selectbox("Nivel", ["Introductorio", "Intermedio", "Avanzado"])
    tipo = st.radio("Tipo", ["Capítulo Teórico-Práctico", "Guía de Problemas Aplicados"])
    
    if st.form_submit_button("Generar Capítulo / Guía"):
        pdf = redactar_guia_investigada(asig, tema, nivel, tipo)
        if pdf:
            st.session_state.pdf_final = pdf
            st.rerun()

if "pdf_final" in st.session_state:
    st.success("¡Documento estructurado con FlashClass exitosamente!")
    st.download_button(
        label="Descargar Documento FlashClass (.pdf)", 
        data=st.session_state.pdf_final, 
        file_name="FlashClass_Documento_Academico.pdf", 
        mime="application/pdf"
    )
