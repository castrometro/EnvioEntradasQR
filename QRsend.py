import io
import pandas as pd
import qrcode
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from io import BytesIO
from PIL import Image, ImageDraw, ImageFilter, ImageFont

# Configuración del servidor SMTP de Gmail
def connect_to_gmail():
    smtp_server = null
    smtp_port = null
    gmail_user = null
    gmail_password =  null
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        server.login(gmail_user, gmail_password)
    except smtplib.SMTPAuthenticationError:
        print("Error: No se pudo autenticar con Gmail. Verifica tus credenciales.")
        return None
    return server

# Función para generar un QR
def generate_qr(rut):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(f"https://ryhascensoresapp.com/#/verificar/{rut}")
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    img_byte_arr = BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

# Funcion para colocar el qr en la imagen, (fondo,imagen a pegar)
def paste_qr(image_path, qr_image_bytes):
    # Cargar la imagen de fondo
    background = Image.open(image_path)
    #img_w, img_h = img.size
    # Cargar el QR
    # Convertir bytes en un objeto de imagen PIL
    qr_resized = Image.open(io.BytesIO(qr_image_bytes))
    #qr_resized = qr.resize((300,300), Image.Resampling.LANCZOS)  # O usa Image.Resampling.LANCZOS para Pillow >= 8.0.0
    #qr_w, qr_h = qr.size
    # Pegar el QR en la imagen, pero copiando la imagen de fondo primero
    backup_background = background.copy()
    backup_background.paste(qr_resized , (760, 20))
    #Convertir imagen en bytes para no guardarla localmente y solo enviarla.
    imagenfinal_enbytes = BytesIO()
    backup_background.save(imagenfinal_enbytes, format='JPEG')
    imagenfinal_enbytes.seek(0)  # Mover el cursor al inicio del stream?
    return imagenfinal_enbytes.getvalue()

# Enviar correo electrónico
def send_email(server, receiver_email, subject, imagen_adjunta, imagen_footer, gmail_user):
    msg = MIMEMultipart('related')
    msg['From'] = gmail_user
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Cuerpo del mensaje en HTML
    html = f"""\
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #f4f4f4;
            }}
            .container {{
                background-color: #fff;
                margin: 10px auto;
                padding: 20px;
                max-width: 600px;
                border-radius: 8px;
                box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }}
            .footer {{
                margin-top: 20px;
                text-align: center;
                font-size: 12px;
                color: #888;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>¡Tu Entrada para la Tocata está Lista!</h2>
            <p>¡Saludos!</p>
            <p>Tu entrada para <strong>Quinto Rey, Josévideo y Luke Bongcam EN VIVO</strong> está aquí. ¡Prepárate para una noche llena de buena música y energía vibrante!</p>
            <p><strong>No olvides traer esta entrada contigo junto con tu cédula de identidad.</strong> ¡Te esperamos para disfrutar juntos de una experiencia musical inolvidable!</p>
        </div>
        <footer>
            <img src="cid:imagen_footer" style="max-width: 100%; height: auto;">
        </footer>
    </body>
    </html>
    """
    msg.attach(MIMEText(html, 'html'))

    # Adjuntar la imagen con el QR
    img_with_qr = MIMEImage(imagen_adjunta)
    img_with_qr.add_header('Content-ID', '<imagen_adjunta>')
    msg.attach(img_with_qr)

    # Adjuntar la imagen adicional
    with open(imagen_footer, 'rb') as f:
        mime_type, _ = mimetypes.guess_type(imagen_footer)
        mime_type_main, mime_subtype = mime_type.split('/')
        img_additional = MIMEImage(f.read(), _subtype=mime_subtype)
        img_additional.add_header('Content-ID', '<imagen_footer>')
        msg.attach(img_additional)

    # Enviar el correo
    try:
        server.send_message(msg)
    except Exception as e:
        print(f"Error: No se pudo enviar el correo electrónico a {receiver_email}. Error: {e}")



# Procesamiento de las filas
def main():
    server = connect_to_gmail()
    if server is None:
        print("Error de conexión con Gmail. Abortando script.")
        return

    try:
        file_path = '/Users/mimac/Desktop/asistencia_test.xlsx'
        data = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error: No se pudo leer el archivo Excel. Error: {e}")
        return

    for index, row in data[data['Qr Enviado (Y/N)'] == 'N'].iterrows():
        qr_image = generate_qr(row['Rut'])
        image_with_qr_bytes = paste_qr('/Users/mimac/Documents/Mis Códigos/5to Rey/EnviarQR/Imágenes/Entrada.jpeg', qr_image)
        send_email(server, row['Mail'], 'Entrada Tocata', image_with_qr_bytes, '/Users/mimac/Documents/Mis Códigos/5to Rey/EnviarQR/Imágenes/FooterNegro.jpeg', server.user)
        data.at[index, 'Qr Enviado (Y/N)'] = 'Y'

    # Guardar los cambios en el archivo Excel original
    try:
        data.to_excel(file_path, index=False)
    except Exception as e:
        print(f"Error: No se pudo guardar el archivo Excel. Error: {e}")

    server.quit()

if __name__ == "__main__":
    main()

