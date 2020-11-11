import time
import requests as r
from openpyxl import load_workbook

MY_TOKEN=''

ids=[]
passwords=[]
logins=[]
p=''
l=''

wb = load_workbook('spec2.xlsx')
sheet = wb.get_sheet_by_name('Лист1')


for i in range(0,3):
    passwords.append(sheet['E'+str(2+i)].value)
    logins.append(sheet['C'+str(2+i)].value)
    ids.append(sheet['F'+str(2+i)].value)
print(ids)

print(passwords)
print(logins)
id=''
counter=0

for i in range(0, 3):
    id=str(ids[i]);
    p=str(passwords[i])
    l=str(logins[i])
    message = 'Привет!\nАртем Маркович просил разослать всем логины и пароли от Selene (наш новый учебный портал).\nТвой логин: ' + l + '\nТвой пароль: ' + p+'\nСам сайт: https://selene.cosmos.msu.ru/\nЗдесь будут лежать материалы, и все такое. \nСайт пока в разработке, поэтому возможности сменить пароль здесь временно нет, но если ты вдруг забудешь свой, то просто спроси его у меня.'
    send = 'https://api.vk.com/method/messages.send?peer_id=' + id + '&message=' + message + '&access_token=' + MY_TOKEN + '&v=5.89'
    response=r.get(send)
    if response.status_code==200:
        counter+=1
    time.sleep(0.25)
print('скрипт закончил работу, результат '+str(counter)+'/11')
#print(response.status_code)
