import gspread
from oauth2client.service_account import ServiceAccountCredentials
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Configurações da planilha
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("your credenciais.json here ", scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("email spreadsheet here").worksheet("emails")

# Configurações de autenticação do Gmail
client_id = 'your client_id'
client_secret = 'your client_secret'
refresh_token = 'you refresh_token'
credentials = Credentials.from_authorized_user_info({
    'client_id': client_id,
    'client_secret': client_secret,
    'refresh_token': refresh_token,
})
credentials.refresh(Request())
gmail_service = build('gmail', 'v1', credentials=credentials)

# Lista para rastrear emails já enviados
emails_enviados = set()

def verify_email(email):
    email_list = sheet.col_values(1)  # Obtém todos os emails da coluna
    if email in email_list:
        return True
    return False

def send_quiz_link(email):
    # Verifica se o email já foi enviado
    if email in emails_enviados:
        print("Email já foi enviado anteriormente.")
        return

    # Criação da mensagem de email
    msg = MIMEMultipart()
    msg["From"] = "seu_email@gmail.com"  # Substitua pelo seu endereço de email
    msg["To"] = email
    msg["Subject"] = "Acesso ao Quiz"

    # Corpo do email
    message = "Ao confirmar seu pagamento, você terá acesso ao quiz! Clique no link abaixo:\n\n"
    quiz_link = "private form"
    message += quiz_link
    msg.attach(MIMEText(message, "plain"))

    # Envia o email via Gmail API
    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    gmail_service.users().messages().send(userId="me", body={"raw": raw_message}).execute()

    # Adiciona o email à lista de emails enviados
    emails_enviados.add(email)

    print("Email enviado com sucesso!")

def access_quiz(email):
    if verify_email(email):
        send_quiz_link(email)
    else:
        print("Email não encontrado na planilha.")

# Aqui, você pode iterar por todos os emails na planilha
def process_emails_in_sheet():
    email_list = sheet.col_values(1)[1:]  # Obtém todos os emails da coluna, excluindo o cabeçalho
    for email in email_list:
        access_quiz(email)

# Exemplo de uso
process_emails_in_sheet()
