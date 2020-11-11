import smtplib
from openpyxl import load_workbook
from email.mime.text import MIMEText                # Текст/HTML
from email.mime.multipart import MIMEMultipart

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
pas=''
server.login('celene.fsr.helper@gmail.com', pas)

l=''
p=''
mails=[]
passwords=[]
logins=[]
wb = load_workbook('mag.xlsx')
sheet = wb.get_sheet_by_name('Лист1')
n=0#количество человек

for i in range(0,n):
    passwords.append(sheet['C'+str(2+i)].value)
    logins.append(sheet['B'+str(2+i)].value)
    mails.append(sheet['D'+str(2+i)].value)
print(passwords)
print(logins)
for i in range(0,n):
    l=str(logins[i])
    p=str(passwords[i])
    msg = MIMEMultipart()  # Создаем сообщение
    msg['From'] = 'celene.fsr.helper@gmail.com' # Адресат
    msg['To'] = mails[i]  # Получатель
    msg['Subject'] = 'Пароль для Селена'  # Тема сообщения
    body='Здравствуй!\nНа факультете запускается учебный портал Selene\nТвой логин: ' + l + '\nТвой пароль: ' + p+'\nСам сайт: https://selene.cosmos.msu.ru/\nЗдесь будут лежать материалы, и все такое. \nСайт пока в разработке, поэтому возможности сменить пароль здесь временно нет, но она обязательно скоро появится!'
    msg.attach(MIMEText(body, 'plain'))
    server.send_message(msg)
server.quit()