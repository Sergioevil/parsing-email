# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import email
import imaplib
import xlsxwriter
import os, time, re

login = input("Почта: ") 

password = input("Пароль: ")

data_list = []

imap = imaplib.IMAP4_SSL('imap.yandex.ru')
imap.login(login, password)

os.system('cls')
print('Выберите папку')
for ui in imap.list()[1]:
	print(re.search(r'"\|" ([\w\-\d]*)',ui.decode('utf-8')).group(1))

folder_name = str(input())

imap.select(folder_name)
mail_amount = int(re.search(r'(\d*)', imap.select(folder_name)[1][0].decode('utf-8')).group(1))

workbook = xlsxwriter.Workbook('заявки.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Что интересует')
worksheet.write('B1', 'ФИО')
worksheet.write('C1', 'Возраст')
worksheet.write('D1', 'Дата')
worksheet.write('E1', 'Время')
worksheet.write('F1', 'Почта')
worksheet.write('G1', 'Телефон')
worksheet.write('H1', 'Сумма')
worksheet.write('I1', 'Что рефинансировать')
worksheet.write('J1', 'КИ')

for i in range(mail_amount):
	os.system('cls')
	print(f'{(i+1)/mail_amount*100:.2f}', '%')
	try:
		status,data = imap.fetch(str(int(i)+1), '(RFC822)')
		msg = email.message_from_bytes(data[0][1])
		payload = msg.get_payload()
		try:
			ty, rt = payload
			payload = rt.get_payload(decode=True).decode('utf-8')
		except:
			payload= msg.get_payload(decode=True).decode('utf-8')
		soup = BeautifulSoup(payload, 'lxml')
		x = str(payload)
		match_email = re.search(r'Email: (.{3,}[mu])<br>', x) 
		mainemail = match_email[1] if match_email else 'Not found' 
		
		match_what_inter = re.search(r'Что интересует:\s?([\w\s]*)', x)
		what_inter = match_what_inter[1] if match_what_inter else 'Not found' 
		
		match_summ = re.search(r'Сумма:\s?(\d{1,20})', x)
		summ = match_summ[1] if match_summ else 'Not found'
			
		match_age = re.search(r'Возраст:\s?(\d{2})', x)
		age = match_age[1] if match_age else 'Not found'
	
		match_name = re.search(r'Имя:\s?([\w\s]*)', x)
		name = match_name[1] if match_name else 'Not found'
	
		match_what_ref = re.search(r'Что рефинансировать:\s?([\w\s]*)', x)
		what_ref = match_what_ref[1] if match_what_ref else 'Not found'
	
		match_ki = re.search(r'История:\s?([\w\s,]*)', x)
		ki = match_ki[1] if match_ki else 'Not found'
	
		match_number = re.search(r'(\+?(\d(\s|\()?)?\d{3}(\s|\))?[\d\-]{4,9})', x)
		number = match_number[1] if match_number else 'Not found'
	
		match_time_to_parse = re.search(r'(\d{1,2}:\d{1,2})', x)
		time_to_parse = match_time_to_parse[1] if match_time_to_parse else 'Not found'
	
		match_date_to_parse = re.search(r'(\d{1,2}\.\d{1,2}\.\d{1,2})', x)
		date_to_parse = match_date_to_parse[1] if match_date_to_parse else 'Not found'

		worksheet.write(f'A{i+2}', what_inter)
		worksheet.write(f'B{i+2}', name)
		worksheet.write(f'C{i+2}', age)
		worksheet.write(f'D{i+2}', date_to_parse)
		worksheet.write(f'E{i+2}', time_to_parse)
		worksheet.write(f'F{i+2}', mainemail)
		worksheet.write(f'G{i+2}', number)
		worksheet.write(f'H{i+2}', summ)
		worksheet.write(f'I{i+2}', what_ref)
		worksheet.write(f'J{i+2}', ki)
	except:
		pass

workbook.close()
print('Done!')
time.sleep(3)