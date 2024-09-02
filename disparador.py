import requests
from docx import Document
import time
import re

# Função para configurar o cabeçalho de autorização
def get_headers(access_token):
    return {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

# Função para enviar emails usando a API Microsoft Graph
def send_emails(emails, message_template, batch_number, access_token):
    headers = get_headers(access_token)
    success_count = 0
    failure_count = 0
    for email in emails:
        message = {
            "message": {
                "subject": "VENDA PRIVADO",
                "body": {
                    "contentType": "Text",
                    "content": message_template
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": email
                        }
                    }
                ],
                "from": {
                    "emailAddress": {
                        "address": "curriculo@gruposs.net"
                    }
                },
                "sender": {
                    "emailAddress": {
                        "address": "ti@gruposs.net"
                    }
                }
            },
            "saveToSentItems": "true"
        }
        response = requests.post(
            'https://graph.microsoft.com/v1.0/users/curriculo@gruposs.net/sendMail',
            headers=headers,
            json=message
        )
        if response.status_code == 202:  # 202 Accepted indicates the request was accepted for processing
            success_count += 1
        else:
            failure_count += 1
            print(f"Falha ao enviar e-mail para {email} Status: {response.status_code}")
            print(f"Detalhes do erro: {response.text}")  # Imprime a mensagem de erro detalhada da API

    print(f"Lote {batch_number}: {success_count} e-mails enviados com sucesso, {failure_count} falharam.")

# Carregar os endereços de email do documento Word
doc_path = r'/mnt/data/cleaned_email_addresses_no_from_to.docx'
doc = Document(doc_path)

emails = []
email_pattern = re.compile(r"[^@]+@[^@]+\.[^@]+")

for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    if text:
        # Dividir e-mails separados por vírgulas e limpar espaços em branco
        split_emails = [email.strip() for email in text.split(',')]
        # Adicionar apenas e-mails válidos
        for email in split_emails:
            if email_pattern.match(email):
                emails.append(email)
            else:
                print(f"E-mail inválido ignorado: {email}")

# Remover duplicações na lista de e-mails
emails = list(set(emails))  # Isso remove endereços duplicados

# Definir o corpo do email
message_template = """
Estamos contratando COMERCIAL com AMPLA EXPERIÊNCIA EM VENDAS DE PRESTAÇAO DE SERVIÇOS EM GERAL PARA ORGÃO PRIVADO A NÍVEL BRASIL, Horário de trabalho: DAS 08H ÀS 12:00 e das 13:12 às 18:00 segunda a sexta, TRABALHAMOS 100% HOME OFFICE. Com compensação de horas, caso empresa necessite passar do horário determinados dias. 
R$ 2.500,00 + 5% de comissão sobre o faturamento do 1º mês do contrato adquirido. O objetivo é receber por comissão
VA 20,00 por dia
Experiência obrigatória EM VENDAS DE CONTRATOS DE PRESTAÇÃO DE SERVIÇOS EM GERAL PARA ORGÃO PRIVADO A NÍVEL BRASIL.
Perfil comercial.
Serviços: a) contatar empresas privadas a nível brasil oferecendo serviços e elaborando orçamentos.
OBRIGATORIO EXPERIENCIA EM VENDA DE PRESTAÇÃO DE SERVIÇOS, MÃO DE OBRA TERCEIRIZADA. 

Início imediato. 
É obrigatório dedicação exclusiva, serviço é monitorado em tempo real pelo teamviwer.
Ter Computador com memoria SSD, Windows 10 ou 11.
Início imediato
"""

# Defina seu token de acesso aqui
access_token = 'seu_token_de_acesso_aqui'

# Enviar emails em lotes de 200 para evitar sobrecarga
batch_size = 200
batch_number = 1

# Envio em lotes
for i in range(0, len(emails), batch_size):
    batch_emails = emails[i:i+batch_size]
    send_emails(batch_emails, message_template, batch_number, access_token)
    batch_number += 1
    time.sleep(5)  # Pequeno delay para evitar problemas de limite de taxa
