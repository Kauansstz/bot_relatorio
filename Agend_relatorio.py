import schedule  # type: ignore
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
import os

EMAIL_REMETENTE = os.environ.get("EMAIL_REMETENTE")
PASSWORD_REMETENTE = os.environ.get("PASSWORD_REMETENTE")
EMAIL_DESTINARIO = os.environ.get("EMAIL_DESTINARIO")


def enviar_relatorio():
    # Configurações do e-mail
    remetente = EMAIL_REMETENTE
    senha = PASSWORD_REMETENTE
    destinatario = EMAIL_DESTINARIO
    assunto = "Relatório Mensal"

    # Criação do e-mail
    msg = MIMEMultipart()
    msg["From"] = remetente  # type: ignore
    msg["To"] = destinatario  # type: ignore
    msg["Subject"] = assunto

    corpo_email = "Anexo: Relatório do mês anterior."
    msg.attach(MIMEText(corpo_email, "plain"))

    # Anexando o relatório (exemplo de anexo)
    arquivo_relatorio = "C:\\Users\\Kauan\\OneDrive\\Área de Trabalho\\reste.xlsx"  # type: ignore
    anexo = open(arquivo_relatorio, "rb")

    part = MIMEBase("application", "octet-stream")
    part.set_payload((anexo).read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", "attachment; filename= %s" % arquivo_relatorio
    )

    msg.attach(part)

    # Conexão com o servidor SMTP
    servidor_smtp = smtplib.SMTP("smtp.gmail.com", 587)
    servidor_smtp.starttls()
    servidor_smtp.login(remetente, senha)  # type: ignore

    # Envio do e-mail
    servidor_smtp.sendmail(remetente, destinatario, msg.as_string())  # type: ignore
    servidor_smtp.quit()

    print(f"Relatório enviado para {destinatario}.")


def agendar_relatorio():
    hoje = datetime.datetime.now()
    if hoje.day == 1:  # Verifica se hoje é o primeiro dia do mês
        enviar_relatorio()


# Agendamento para verificar a cada dia às 08:00 se é o primeiro dia do mês e enviar o relatório
schedule.every().day.at("08:00").do(agendar_relatorio)

while True:
    schedule.run_pending()
    time.sleep(1)
