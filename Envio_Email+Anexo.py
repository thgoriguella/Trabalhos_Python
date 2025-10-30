import smtplib
import os
import datetime
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- Configurações de E-mail --- // Aqui fica as configurações do e-mail que vai utilizar
REMETENTE_EMAIL = "thiago@email.com.br"
REMETENTE_SENHA = "ABCDEFC001"
SERVIDOR_SMTP = "smtp.XXXXX.com"
PORTA_SMTP = 587

# --- Configurações dos Destinatários ---
LISTA_DESTINATARIOS = [
"thiago.silva@email.com.br",
"tiago.ti@email.com.br"
]

# --- Configurações da Mensagem e Anexo ---
ASSUNTO_EMAIL = "Report de lojas - " + datetime.date.today().strftime('%d/%m/%Y')

# --- Edite estes caminhos para pegar o arquivo na pasta destino e o link do sharepoint, caso precise incluir no e-mail ---
CAMINHO_PASTA_DESTINO = "C:\\Users\\thiago.origuella\\Vendas"
LINK_SHAREPOINT = "https://sharepoint.com/:f:/r/sites/USER-SASTLSLPO/Documentos%20Compartilhados/SAS%20TL-SL-PO/Vendas?csf=1&web=1&e=TQCBlQ"

# --- msg do e-mail via html ---
CORPO_EMAIL_HTML = f"""
<html>
<head></head>
<body>
  <p>Prezados, bom dia.</p>
  <p>Encaminho, em anexo, o relatório de vendas.</p>
  <p>O arquivo também pode ser acessado no link abaixo.</p>
  <p><strong><a href="{LINK_SHAREPOINT}">[Clique aqui para acessar]</a></strong></p>
  <br>
  <p>Atenciosamente,</p>
</body>
</html>
"""

# -- local que precisa pegar o arquivo
CAMINHO_COMPLETO_ARQUIVO = "C:\\Users\\thiago.origuella\\Vendas\\RELATÓRIO DE VENDAS.xlsx"
NOME_BASE_ARQUIVO = os.path.basename(CAMINHO_COMPLETO_ARQUIVO)


def enviar_email_com_anexo():
    try:
        # --- Cópia do Arquivo ---
        print(f"Copiando arquivo de origem: {CAMINHO_COMPLETO_ARQUIVO}")
        caminho_arquivo_destino = os.path.join(CAMINHO_PASTA_DESTINO, NOME_BASE_ARQUIVO)
        shutil.copy2(CAMINHO_COMPLETO_ARQUIVO, caminho_arquivo_destino)
        print(f"Arquivo copiado com sucesso para: {caminho_arquivo_destino}")
        print("-" * 30)

        # --- Construindo a Mensagem (MIME) ---
        print("Construindo e-mail...")
        msg = MIMEMultipart()
        msg['From'] = REMETENTE_EMAIL
        msg['To'] = ", ".join(LISTA_DESTINATARIOS)
        msg['Subject'] = ASSUNTO_EMAIL

        msg.attach(MIMEText(CORPO_EMAIL_HTML, 'html', 'utf-8'))

        # --- Lendo e Adicionando o Anexo ---
        with open(CAMINHO_COMPLETO_ARQUIVO, "rb") as attachment:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=('utf-8', '', NOME_BASE_ARQUIVO)
        )
        msg.attach(part)

        # --- Conectando e Enviando ---
        print("Conectando ao servidor SMTP...")
        with smtplib.SMTP(SERVIDOR_SMTP, PORTA_SMTP) as server:
            server.starttls()
            server.login(REMETENTE_EMAIL, REMETENTE_SENHA)
            texto_completo = msg.as_string()
            server.sendmail(REMETENTE_EMAIL, LISTA_DESTINATARIOS, texto_completo)
            print(f"E-mail enviado com sucesso para: {', '.join(LISTA_DESTINATARIOS)}")

    except FileNotFoundError:
        print(f"ERRO: Arquivo de origem NÃO ENCONTRADO em: {CAMINHO_COMPLETO_ARQUIVO}")
        print("Ou a PASTA DE DESTINO não foi encontrada. Verifique os caminhos.")
    except PermissionError:
        print(f"ERRO: Sem permissão para copiar o arquivo para: {CAMINHO_PASTA_DESTINO}")
    except smtplib.SMTPException as e:
        print(f"ERRO de SMTP: {e}")
    except Exception as e:
        print(f"ERRO inesperado: {e}")


# --- Executa a função ---
if __name__ == "__main__":
    enviar_email_com_anexo()
