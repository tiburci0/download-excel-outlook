import win32com.client as client
from pathlib import Path

outlook = client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

remetente = str(input('Digite o e-mail do remetente: '))

inbox = outlook.GetDefaultFolder(6)
getEmailsInInbox = inbox.Items

for email in getEmailsInInbox:
    assunto = email.subject
    dataFormtada = email.SentOn.strftime("%d-%m-%y")

    if remetente in email.SenderEmailAddress: 
        print('--------------------------------------------------------')
        print(f"📩  {assunto}" )
        print(f"📅  {dataFormtada}")
        print('🔄  Download em andamento...')
        print('✅  Download concluído!')
        for anexos in email.attachments:
                if str(anexos).__contains__("xlsx"):
                    destino = Path.cwd() / f'arquivos/[{dataFormtada}] {assunto} - {remetente}'
                    destino.mkdir(parents=True, exist_ok=True)

                    anexos.SaveAsFile(destino / str(anexos))
                