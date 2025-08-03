import streamlit as st
import os
import tempfile
import re
import pandas as pd
from collections import defaultdict
from datetime import datetime
import win32com.client as win32
import pythoncom
import zipfile

# ------------------------ UTILS ------------------------

def normalize(text):
    return re.sub(r'\s+', ' ', str(text).strip().upper())

def load_proveedores(excel_path):
    df = pd.read_excel(excel_path)
    df['Razón Social'] = df['Razón Social'].apply(normalize)
    df['Mail'] = df['Mail'].astype(str)

    proveedor_emails = defaultdict(set)
    for _, row in df.iterrows():
        razon = row['Razón Social']
        for email in re.split(r'[;,\s]+', row['Mail']):
            if re.match(r"[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$", email.strip()):
                proveedor_emails[razon].add(email.strip())

    return {razon: list(emails) for razon, emails in proveedor_emails.items()}

def detect_file_type_and_rs(file_name):
    name_upper = os.path.basename(file_name).upper()
    razon_social = None
    tipo = None

    if 'PAGO' in name_upper:
        tipo = 'pago'
        match = re.search(r'PAGO\s+(.+?)\.PDF$', name_upper)
        if match:
            razon_social = match.group(1).strip()

    elif 'OP' in name_upper:
        tipo = 'op'
        parts = name_upper.split('_')
        if len(parts) >= 3:
            razon_social = parts[2].replace('.PDF', '').strip()

    elif 'CG' in name_upper:
        tipo = 'cg'
        parts = name_upper.split('_')
        if len(parts) >= 3:
            razon_social = parts[2].replace('.PDF', '').strip()

    return tipo, normalize(razon_social) if razon_social else None

def clasificar_archivos(carpeta_pdfs, excel_proveedores):
    proveedores = load_proveedores(excel_proveedores)
    resultado = defaultdict(lambda: {'email': [], 'pago': [], 'op': [], 'cg': []})

    for root, _, files in os.walk(carpeta_pdfs):
        for file in files:
            if not file.lower().endswith(".pdf"):
                continue
            full_path = os.path.join(root, file)
            tipo, razon_social = detect_file_type_and_rs(file)
            if tipo and razon_social in proveedores:
                resultado[razon_social]['email'] = proveedores[razon_social]
                resultado[razon_social][tipo].append(full_path)

    return dict(resultado)

def extraer_fechas_desde_archivos(lista_archivos):
    fechas = []
    for archivo in lista_archivos:
        nombre = os.path.basename(archivo)
        match = re.match(r'(\d{4}-\d{2}-\d{2})', nombre)
        if match:
            try:
                fecha = datetime.strptime(match.group(1), "%Y-%m-%d")
                fechas.append(fecha)
            except:
                continue
    return sorted(fechas)

def enviar_mail_outlook(razon_social, emails, archivos_adjuntos):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    fechas = extraer_fechas_desde_archivos(archivos_adjuntos)
    fecha_str = fechas[0].strftime("%d/%m/%Y") if fechas else datetime.today().strftime("%d/%m/%Y")

    mail.Subject = f"Pago {razon_social} {fecha_str}"
    mail.To = ";".join(emails)
    mail.Body = (
        "Estimados:\n\n"
        "Se adjuntan todos los comprobantes correspondientes al último pago.\n\n"
        "Saludos!"
    )

    for adjunto in archivos_adjuntos:
        mail.Attachments.Add(os.path.abspath(adjunto))

    mail.Send()

# ------------------------ APP STREAMLIT ------------------------

st.set_page_config(page_title="Envío de Pagos", layout="wide")
st.image("logo_ctc.png", width=180)
st.title("Envio de Comprobantes de Pago")

uploaded_excel = st.file_uploader("Importá el Excel de proveedores", type=["xlsx"])
uploaded_zip = st.file_uploader("Importá el Zip con los comprobantes PDF", type=["zip"])

if uploaded_excel and uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, "proveedores.xlsx")
        with open(excel_path, "wb") as f:
            f.write(uploaded_excel.read())

        zip_path = os.path.join(tmpdir, "archivos.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        pdf_dir = os.path.join(tmpdir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(pdf_dir)

        resultado = clasificar_archivos(pdf_dir, excel_path)
        st.success(f"Se procesaron {len(resultado)} proveedores con archivos asociados.")

        for razon_social, data in resultado.items():
            st.markdown(f"### {razon_social}")
            st.markdown(f"**Emails:** {', '.join(data['email'])}")
            st.markdown(f"- Pagos: {len(data['pago'])} archivo(s)")
            st.markdown(f"- OPs: {len(data['op'])} archivo(s)")
            st.markdown(f"- CGs: {len(data['cg'])} archivo(s)")

            with st.expander("Ver archivos"):
                for tipo in ['pago', 'op', 'cg']:
                    for f in data[tipo]:
                        st.text(f"[{tipo.upper()}] {os.path.basename(f)}")

            enviar = st.radio(
                f"¿Querés enviar el mail ahora a {razon_social}?",
                options=["No", "Sí"],
                index=0,
                key=razon_social
            )

            if enviar == "Sí":
                try:
                    archivos = data['pago'] + data['op'] + data['cg']
                    enviar_mail_outlook(razon_social, data['email'], archivos)
                    st.success(f"Mail enviado a {razon_social} correctamente.")
                except Exception as e:
                    st.error(f"Error al enviar el mail: {e}")


