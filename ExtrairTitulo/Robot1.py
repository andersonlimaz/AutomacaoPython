from imap_tools import MailBox, AND
import openpyxl
login = "estagiario1.naj.prfn3@pgfn.gov.br"
senha = "hzrctlidbdooqvax"


meu_email = MailBox("imap.gmail.com").login(login, senha)

# Pegar email que foram enviados por um remetente especifico 
workbook = openpyxl.Workbook()
sheet = workbook.active

lista_emails = meu_email.fetch(AND(from_="andersonzlima7@gmail.com"))

for email in lista_emails:
    subject = email.subject
    
    sheet.append([subject])

workbook.save("emails.xlsx")