from flask import Flask, render_template, request, send_file
import imaplib
import email
import os
import fitz
import re
import shutil
import pandas as pd
from datetime import datetime, timedelta
from email.header import decode_header
from dotenv import load_dotenv
import os

# Cargar el archivo .env
load_dotenv()

app = Flask(__name__)

# Configuración de Gmail
# Acceder a las variables de entorno
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
IMAP_SERVER = "imap.gmail.com"
FOLDER = "INBOX"
SENDER_FILTER = "noreply@junin.gob.ar"
DOWNLOAD_FOLDER = "boletas_pdf"

# Lista de propietarios (el mismo que tienes)

PROPIETARIOS = {
    "11144": ["Abdala - Venerte", "Cte. Escribano 321", "Nosotros"],
    "16360": ["Arriola", "Cnel. Suarez 352", "Nosotros"],
    "24651": ["Arriola", "Los Perales 350", "Nosotros"],
    "54937": ["Arriola", "Chacabuco 277", "Nosotros"],
    "19576": ["Carballeira Olga", "Almafuerte 352", "Nosotros"],
    "58357": ["Carballeira Olga", "Javier Muñiz 229", "Nosotros"],
    "58109": ["Covalchi", "Saavedra 16", "Nosotros"],
    "58110": ["Covalchi", "Saavedra 16 P1D", "Nosotros"],
    "57171": ["Dalesandro", "R.E. San Martin 61", "Nosotros"],
    "8587": ["Esquinoval Fabian", "Alte. Brown 423", "Nosotros"],
    "3219": ["Gralato Dorys", "M. Lopez 380", "Nosotros"],
    "9594": ["Marcantonio", "Padre Ghio 662", "Nosotros"],
    "46837": ["Marturano", "-", "Nosotros"],
    "13979": ["Mastrogiusepe", "Narbondo 241", "Nosotros"],
    "5065": ["Mucciarone Dalia", "Winter 26", "Nosotros"],
    "55119": ["Perez Olga", "Avda. San Martin 14", "Nosotros"],
    "46472": ["Pierrard", "R.E. San Martin 28", "Nosotros"],
    "46539": ["Pierrard", "R.E. San Martin 36", "Nosotros"],
    "47855": ["Pierrard", "Avda. San Martin 239", "Nosotros"],
    "53264": ["Pierrard", "Avda. San Martin 290", "Nosotros"],
    "53297": ["Pierrard", "Avda. San Martin 290", "Nosotros"],
    "55366": ["Pierrard", "Lebensohn 19", "Nosotros"],
    "55385": ["Pierrard", "Lebensohn 19 (Cochera)", "Nosotros"],
    "56149": ["Pierrard", "Gral. Paz 314", "Nosotros"],
    "702349": ["Poggio", "Roque Vazquez 786 PB D", "Nosotros"],
    "64188": ["Santangello Isabel", "Alberdi 70", "Nosotros"],
    "25053": ["Santos Norma", "25 de mayo 8", "Nosotros"],
    "59534": ["Sanz Elida", "Ameghino 177", "Nosotros"],
    "4812": ["Tobal Federico", "Winter 273", "Nosotros"],
    "20694": ["Varela", "Pasteur 470", "Nosotros"],
    "58089": ["Abrahan Domingo", "", "Propietario"],
    "18475": ["Amigo Alberto", "", "Propietario"],
    "17255": ["Bianchelli Alfredo", "", "Propietario"],
    "702142": ["Bianchelli Alfredo", "", "Propietario"],
    "18475": ["Boselli Luis", "", "Propietario"],
    "33317": ["Dammiano Lucia", "Uruguay", "Inquilino"],
    "51596": ["De Benedetto Jose Luis", "R. Hernandez 1032", "Propietario"],
    "7970": ["Di Prinzio Alcides", "Alem 262", "Inquilino"],
    "15036": ["Espindola Daniel", "Arias 440", "Propietario"],
    "40022": ["Gas Carlos", "Saenz Peña 273", "Propietario"],
    "40027": ["Gas Carlos", "Saenz Peña 273", "Propietario"],
    "25962": ["Lima Alfredo", "Pellegrini 1080", "Propietario"],
    "28625": ["Limonta Nestor", "Gandini 914", "Propietario"],
    "62138": ["Mariani Nancy", "Saenz Peña 293", "Propietario"],
    "61205": ["Mastromauro Nestor", "Saenz Peña 249", "Propietario"],
    "51970": ["Woinilowiez", "", "Propietario"],
    "23190": ["Falasconi Carlos", "Mariano Moreno 428", "Nosotros"],
    "13671": ["Meoni", "Vicente López y Planes 349", "Nosotros"],
    "30507": ["Artori Ricardo", "Rioja 93", "Nosotros"],
}

# Funciones para manejar la conexión y el procesamiento
def connect_gmail():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    mail.select(FOLDER)
    return mail

def get_date_5_days_ago():
    today = datetime.today()
    five_days_ago = today - timedelta(days=25)
    return five_days_ago.strftime("%d-%b-%Y")

from email.header import decode_header
import os
import imaplib
import email
import shutil

# Descargar archivos PDF
def download_pdfs(mail):
    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)

    date_5_days_ago = get_date_5_days_ago()
    result, data = mail.search(None, f'(FROM "{SENDER_FILTER}" SINCE {date_5_days_ago})')
    email_ids = data[0].split()

    if not email_ids:
        print(f"No se encontraron correos de {SENDER_FILTER} en los últimos 5 días.")
        return "No se encontraron correos de {SENDER_FILTER} en los últimos 5 días."

    for email_id in email_ids:
        result, msg_data = mail.fetch(email_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get_content_subtype() != 'pdf':
                continue

            # Obtener el nombre del archivo de manera segura
            filename = part.get_filename()

            # Asegurarse de que filename no sea None ni booleano
            if filename:
                decoded_filename = decode_header(filename)
                # Decodificar el nombre del archivo correctamente
                filename = decoded_filename[0][0]
                encoding = decoded_filename[0][1]

                # Si el nombre está en bytes, lo decodificamos
                if isinstance(filename, bytes):
                    filename = filename.decode(encoding if encoding else 'utf-8')

                # Reemplazar caracteres no válidos en el nombre del archivo
                filename = filename.replace('/', '_').replace('\\', '_')

                # Guardar el archivo en el directorio de descarga
                filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Descargado: {filename}")
            else:
                print("El archivo no tiene un nombre válido, saltando...")

    return "Archivos descargados correctamente"


def extract_data_from_pdf(pdf_path):
    filename = os.path.basename(pdf_path)
    partida_match = re.search(r"(\d+)-", filename)
    numero_partida = partida_match.group(1) if partida_match else "N/A"

    doc = fitz.open(pdf_path)
    text = "\n".join([page.get_text("text") for page in doc])

    fecha_match = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    fecha_vencimiento = fecha_match.group(1) if fecha_match else "N/A"

    importe_match = re.search(r"(\d{2}/\d{2}/\d{4})\s+(\d{1,3}(?:\.\d{3})*,\d{2})", text)
    importe = importe_match.group(2) if importe_match else "N/A"

    propietario, direccion, quien_paga = PROPIETARIOS.get(numero_partida, ["Desconocido", "-", "Desconocido"])

    return {
        "Propietario": propietario,
        "Dirección": direccion,
        "Partida": numero_partida,
        "Fecha de Vencimiento": fecha_vencimiento,
        "Importe": importe,
        "Quién Paga": quien_paga

    }

def process_pdfs():
    data = []
    for filename in os.listdir(DOWNLOAD_FOLDER):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(DOWNLOAD_FOLDER, filename)
            data.append(extract_data_from_pdf(pdf_path))

    if data:
        df = pd.DataFrame(data)
        df.to_excel("boletas.xlsx", index=False)
        shutil.rmtree(DOWNLOAD_FOLDER)
        return "Datos guardados en boletas.xlsx"
    return "No se encontraron datos para procesar."

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    mail = connect_gmail()
    download_message = download_pdfs(mail)

    # Verifica si 'download_message' es una cadena antes de llamar a 'startswith'
    if isinstance(download_message, str) and download_message.startswith("No"):
        return render_template('index.html', message=download_message)
    
    process_message = process_pdfs()
    return render_template('index.html', message=process_message)


@app.route('/download', methods=['GET'])
def download_file():
    return send_file('boletas.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
