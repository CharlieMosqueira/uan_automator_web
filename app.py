import streamlit as st

st.set_page_config(
    page_title="UAN Automator Web",
    page_icon="🎓",
    layout="centered"
)

# ===============================
#   INTERFAZ PRINCIPAL
# ===============================

st.title("🎓 UAN AUTOMATOR – Versión Web")
st.write("Automatizador oficial de trámites UAN desarrollado por **AI Softwares**.")

st.markdown("---")

st.header("📌 Subir archivos base")
st.write("Sube aquí los dos Excels originales para iniciar.")

machote = st.file_uploader("Subir MACHOTE DE TRÁMITES.xlsx", type=["xlsx"])
gestores = st.file_uploader("Subir FORMATO PARA LOS GESTORES.xlsx", type=["xlsx"])

st.markdown("---")

st.header("🧾 Captura del alumno")
curp = st.text_input("CURP del Alumno").upper()
nombre = st.text_input("Nombre(s)")
p_ap = st.text_input("Primer Apellido")
s_ap = st.text_input("Segundo Apellido")
carrera = st.text_input("Carrera")
institucion = st.text_input("Institución de Procedencia")
terminacion = st.text_input("Fecha de Terminación (dd/mm/aaaa)")
ciclo = st.text_input("Ciclo Escolar")
promedio = st.text_input("Promedio Final")

st.markdown("---")

st.header("🗂️ Documentos del alumno")
docs = {
    "certificado": st.file_uploader("Certificado (sec/bach)", type=["pdf", "png", "jpg"]),
    "acta": st.file_uploader("Acta de Nacimiento", type=["pdf", "png", "jpg"]),
    "curp": st.file_uploader("CURP Impresa", type=["pdf", "png", "jpg"]),
    "foto_inf": st.file_uploader("Fotos Infantil", type=["pdf", "png", "jpg"]),
    "foto_cred": st.file_uploader("Fotos Credencial", type=["pdf", "png", "jpg"]),
    "foto_tit": st.file_uploader("Fotos Título", type=["pdf", "png", "jpg"]),
    "ine": st.file_uploader("INE", type=["pdf", "png", "jpg"]),
    "comp_dom": st.file_uploader("Comprobante de Domicilio", type=["pdf", "png", "jpg"]),
    "ine_tutor": st.file_uploader("INE del Tutor (si aplica)", type=["pdf", "png", "jpg"])
}

st.markdown("---")

st.header("⚙️ Generar Documentos")

if st.button("GENERAR ARCHIVOS DEL ALUMNO"):
    if not (machote and gestores):
        st.error("Debes subir los archivos Excel base.")
    elif curp == "":
        st.error("Falta la CURP.")
    else:
        st.success("🚀 Próximamente: Aquí se ejecutará el motor UAN completo.")
        st.info("En esta versión solo estamos preparando la interfaz.")

st.markdown("---")
st.write("Desarrollado por **AI Softwares®** – Arquitectos de Ideas")
