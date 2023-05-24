from imap_tools import MailBox, AND
usuario = ""
senha = ""

meu_email = MailBox("impa.gmail.com".login(usuario, senha))

lista_emails = meu_email.fetch(AND(from_=""))
for email in lista_emails:
    
    print(email.suject)#Titulo do E-mail
    print(email.text)#Texto do e-mail

#pegar anexo de um e-mail 

lista_emails = meu_email.fetch(AND(from_=("Andersonlimaz@mail.com")))
for email in lista_emails:
    if len(email.attachments) > 0: #attachments usado para localizar anexo
        for anexo in email.attachments:
            if "RealtorioExcel" in anexo.filename: #filename procura o nome 
                informacoes_anexos = anexo.payload #transforma arquivo em byts, assim podendo recriar o arquivo 
                with open("RelarioExcel.xlsx", "wb") as arquivo_excel:#conversão em um arquivo excel
                    arquivo_excel.write(informacoes_anexos) # escreve informações do arquivo excel no computador
                    
