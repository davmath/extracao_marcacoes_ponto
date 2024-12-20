import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações do Gmail
smtp_server = "smtp.gmail.com"
smtp_port = 587
email_user = "rpa*****@gmail.com"
email_password = "******"

# Função para enviar e-mail
def enviar_email(destinatario, assunto, mensagem_html):
    try:
        # Configurar a mensagem de e-mail
        msg = MIMEMultipart()
        msg["From"] = email_user
        msg["To"] = destinatario
        msg["Subject"] = assunto

        # Corpo do e-mail em HTML
        msg.attach(MIMEText(mensagem_html, "html"))

        # Conectar ao servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Iniciar a conexão segura
        server.login(email_user, email_password)  # Login
        server.send_message(msg)  # Enviar e-mail
        server.quit()

        print(f"E-mail enviado com sucesso para {destinatario}!")

    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")