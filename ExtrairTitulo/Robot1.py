from imap_tools import MailBox, AND
import openpyxl
import re 

login = "push.demandas.prfn3regiao@pgfn.gov.br"
senha = "smybwsrarkugrcls"


meu_email = MailBox("imap.gmail.com").login(login, senha)

workbook = openpyxl.Workbook()
sheet = workbook.active

lista_emails = meu_email.fetch(AND(from_="pje@trf3.jus.br"))

for email in lista_emails:
    subject = email.subject
    split_subject = subject.split()
    
    if len(split_subject) >= 5:
        quinto_split = split_subject[4]
    else:
        quinto_split = "Não disponível"
    
    sheet.append([quinto_split])

workbook.save("emails.xlsx")
