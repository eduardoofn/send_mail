from sqlalchemy import create_engine
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText


# Função para enviar e-mail com anexo
def enviar_email(from_email, senha_email, to_emails, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ', '.join(to_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
        msg.attach(part)

    try:
        server = smtplib.SMTP('mail.servidor_email.com.br', 587)
        server.starttls()
        server.login(from_email, senha_email)
        text = msg.as_string()
        server.sendmail(from_email, to_emails, text)
        server.quit()
        print(f"E-mail enviado com sucesso para {', '.join(to_emails)}!")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {', '.join(to_emails)}: {e}")

# Conexão ao banco de dados SQL Server usando SQLAlchemy
engine = create_engine('mssql+pyodbc://USUARIO:SENHA@SERVIDOR/BANCO?driver=SQL+Server')

# Consulta SQL
query_last_date = "SCRIPT PARA PEGAR A DATA DO RELÁTORIO OU DO DIA OU ULTIMO DIA"
ultima_data = pd.read_sql(query_last_date, engine)['ultima_data'].iloc[0]

# Convertendo a data para o formato datetime no SQL
ultima_data_str = pd.to_datetime(ultima_data).strftime('%Y-%m-%d')

# Consulta na última data disponível
query = f"""
    SELECT DOS DADOS QUE VÃO NO RELATÓRIO
    AND CAST(DATA AS date) = '{ultima_data_str}'
"""
df = pd.read_sql(query, engine)
excel_path = 'RELATORIO.xlsx'
df.to_excel(excel_path, index=False)


# Configurações do e-mail
from_email = 'email do remetente'
senha_email = 'senha'

# Envia e-mail
to_emails = ['email do destinario']
subject = 'Dados de PRODUÇÃO - RELATORIO - CLIENTE'
body = f'Segue em anexo a base do módulo de Produção de Produto Acabado do ERP - CLIENTE na data {ultima_data_str}.'
enviar_email(from_email, senha_email, to_emails, subject, body, excel_path)

