import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
from dotenv import load_dotenv
import pandas as pd
import os

# Cargar variables de entorno desde el archivo .env
load_dotenv()

# Obtener configuración del archivo .env
smtp_server = os.getenv('EMAIL_HOST')
smtp_port = int(os.getenv('EMAIL_PORT'))
smtp_user = os.getenv('EMAIL_HOST_USER')
smtp_password = os.getenv('EMAIL_HOST_PASSWORD')

# Crear el mensaje
def create_message(subject, html_body, from_email, to_email, attachment_paths=None):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Cuerpo HTML
    html_part = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(html_part)

    # Adjuntar archivos
    if attachment_paths:
        for attachment_path in attachment_paths:
            with open(attachment_path, 'rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                msg.attach(part)
    
    return msg

# Leer la plantilla HTML
def load_template(template_path):
    with open(template_path, 'r', encoding='utf-8') as file:
        return file.read()

# Enviar correos
def send_emails(subject, template_path, from_email, to_emails, attachment_paths=None):
    html_body = load_template(template_path)
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        for email in to_emails:
            try:
                msg = create_message(subject, html_body, from_email, email, attachment_paths)
                server.sendmail(from_email, email, msg.as_string())
                print(f'Correo enviado a: {email}')
            except Exception as e:
                print(f'Error al enviar correo a {email}: {e}')

# Leer correos desde un archivo Excel
def read_emails_from_excel(file_path, correos_recepcion):
    df = pd.read_excel(file_path)
    return df[correos_recepcion].dropna().tolist()
0
# configuración 
asunto = 'Acción Requerida: Acceso al Correo Institucional'
template = 'plantillas/template.html'
user_send = 'UTIC | Posgrado'

# Lista de destinatarios
""" to_emails = [
    'mijharv@gmail.com',
    'mjrojasv@unitru.edu.pe',
] """

# Lista de destinatarios por excel
excel_file = 'media/correos.xlsx'  
email_column = 'correos_recepcion'
to_emails =  read_emails_from_excel(excel_file, email_column)


# Rutas de los archivos adjuntos
attachment_paths = [    
    'media/comunicado.pdf'
]

# Enviar el correo
send_emails(asunto, template, user_send, to_emails, attachment_paths)
