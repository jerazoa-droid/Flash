import streamlit as st

st.set_page_config(page_title="FlashClass", layout="centered")

# Ocultar barra lateral y el icono del menú
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none; }
        [data-testid="collapsedControl"] { display: none; }
    </style>
""", unsafe_allow_html=True)

st.title("⚡ FlashClass")
st.markdown("""
Esta es tu plataforma integral para la generación de material docente propulsada por IA.

Selecciona a continuación la herramienta que deseas utilizar para comenzar tu trabajo:
""")

st.write("") 
st.write("") 

col1, col2 = st.columns(2)

with col1:
    if st.button("Generar Presentación (PPTX)", use_container_width=True, type="primary"):
        st.switch_page("pages/1_Presentaciones.py")

with col2:
    if st.button("Generar Guía Académica (PDF)", use_container_width=True, type="primary"):
        st.switch_page("pages/2_Guias_Academicas.py")
