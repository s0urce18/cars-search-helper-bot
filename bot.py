import telebot
from telebot import types
from datetime import datetime, date, time
import win32com.client
import pythoncom
import uuid
import os

print("Bot is running")

bot = telebot.TeleBot('')

admin="0"
ausername=""

#пасхалки-------------------------------------------------------------------
@bot.message_handler(commands=['c250'])
def c250(message):
	bot.send_message(message.from_user.id, "23:27 06.07.2021 — 250 авто добавлены!")
#admin---------------------------------------------------------------------
@bot.message_handler(commands=['admin'])
def admin_message(message):
    global admin
    sent=bot.send_message(message.from_user.id, "Введи пароль")
    bot.register_next_step_handler(sent, admin_password)

def admin_password(message):
    global admin
    global ausername
    if message.text == "":
        admin=message.from_user.id
        print(admin)
        if bool(message.from_user.username)==True:
            ausername="@"+message.from_user.username
        else:
            ausername='"'+message.from_user.first_name+'"'
        bot.send_message(admin, "Ты теперь админ")
    else:
        bot.send_message(message.from_user.id, "Пароль неверный")
        bot.send_message(message.from_user.id, "Ты не админ")

@bot.message_handler(commands=['adminn'])
def admin_message(message):
    global admin
    global ausername
    if admin=="0":
        bot.send_message(message.from_user.id, "Админ не зарегистрирован")
    else:
        bot.send_message(message.from_user.id, ausername)
#/admin---------------------------------------------------------------------

@bot.message_handler(commands=['help'])
def help(message):
    bot.send_message(message.from_user.id, "*Помощь*")

@bot.message_handler(commands=['start'])
def start(message):
	user = bot.get_me()
	print("____Time____: "+str(datetime.today()))
	print("____Bot____: "+str(user))
	print("____User____: "+str(message))
	bot.send_message(message.from_user.id, "Привет")
	pythoncom.CoInitializeEx(0)
	excelc = win32com.client.Dispatch("Excel.Application")
	wbc = excelc.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\cars.xlsx')
	wsc = wbc.ActiveSheet
	wsc.Range('A1:E'+str(wsc.UsedRange.Rows.Count)).Sort(Key1=wsc.Range('A1'), Order1=1, Orientation=1)
	marks = []
	keys = []
	keyboard = types.InlineKeyboardMarkup()
	keys.append(types.InlineKeyboardButton(text="Любая", callback_data='any'))
	for a in range(1, wsc.UsedRange.Rows.Count+1):
		if wsc.Range('A'+str(a)).value in marks:
			pass
		else:
			marks.append(wsc.Range('A'+str(a)).value)
			keys.append(types.InlineKeyboardButton(text=str(wsc.Range('A'+str(a)).value), callback_data=str(wsc.Range('A'+str(a)).value)))
	for b in range(0, len(keys)):
		keyboard.add(keys[b])
	question = 'Укажи марку'
	bot.send_message(message.from_user.id, text=question, reply_markup=keyboard)
	wsc.Range('A1:E'+str(wsc.UsedRange.Rows.Count)).Sort(Key1=wsc.Range('E1'), Order1=1, Orientation=1)
	wbc.Save()
	wbc.Close()

answ = ""
fname = ""

@bot.callback_query_handler(func=lambda call: True)
def callback(call):
	global fname
	if fname == "":
		fname = uuid.uuid4().hex+".xlsx"
	bot.send_message(call.from_user.id, "Хм...")
	global answ
	answ = ""	
	if call.data == "<5":
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\'+fname)
		ws = wb.ActiveSheet
		a = False
		y = 1
		while a==False:
			if y in range(1, ws.UsedRange.Rows.Count+1):
				if int(ws.Range('E'+str(y)).value) <= 0:
					y+=1
				elif int(ws.Range('E'+str(y)).value) < 5000:
					answ += str(str(ws.Range('A'+str(y)).value)+" "+str(ws.Range('B'+str(y)).value)+" — "+str(ws.Range('D'+str(y)).value)+"₽"+" — "+str(ws.Range('E'+str(y)).value)+"$"+"\n")
					y+=1
				else:
					a = True
			else:
				a = True
		bot.send_message(call.from_user.id, answ)
		wb.Close()
		os.remove(fname)
		fname = ""
	elif call.data == "5-10":
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\'+fname)
		ws = wb.ActiveSheet
		a = False
		y = 1
		while a==False:
			if y in range(1, ws.UsedRange.Rows.Count+1):
				if int(ws.Range('E'+str(y)).value) <= 5000:
					y+=1
				elif int(ws.Range('E'+str(y)).value) > 5001 and int(ws.Range('E'+str(y)).value) < 10000:
					answ += str(str(ws.Range('A'+str(y)).value)+" "+str(ws.Range('B'+str(y)).value)+" — "+str(ws.Range('D'+str(y)).value)+"₽"+" — "+str(ws.Range('E'+str(y)).value)+"$"+"\n")
					y+=1
				else:
					a = True
			else:
				a = True
		bot.send_message(call.from_user.id, answ)
		wb.Close()
		os.remove(fname)
		fname = ""
	elif call.data == "10-20":
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\'+fname)
		ws = wb.ActiveSheet
		a = False
		y = 1
		while a==False:
			if y in range(1, ws.UsedRange.Rows.Count+1):
				if int(ws.Range('E'+str(y)).value) <= 10000:
					y+=1
				elif int(ws.Range('E'+str(y)).value) > 10001 and int(ws.Range('E'+str(y)).value) < 20000:
					answ += str(str(ws.Range('A'+str(y)).value)+" "+str(ws.Range('B'+str(y)).value)+" — "+str(ws.Range('D'+str(y)).value)+"₽"+" — "+str(ws.Range('E'+str(y)).value)+"$"+"\n")
					y+=1
				else:
					a = True
			else:
				a = True
		bot.send_message(call.from_user.id, answ)
		wb.Close()
		os.remove(fname)
		fname = ""
	elif call.data == "20-50":
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\'+fname)
		ws = wb.ActiveSheet
		a = False
		y = 1
		while a==False:
			if y in range(1, ws.UsedRange.Rows.Count+1):
				if int(ws.Range('E'+str(y)).value) <= 20000:
					y+=1
				elif int(ws.Range('E'+str(y)).value) > 20001 and int(ws.Range('E'+str(y)).value) < 50000:
					answ += str(str(ws.Range('A'+str(y)).value)+" "+str(ws.Range('B'+str(y)).value)+" — "+str(ws.Range('D'+str(y)).value)+"₽"+" — "+str(ws.Range('E'+str(y)).value)+"$"+"\n")
					y+=1
				else:
					a = True
			else:
				a = True
		bot.send_message(call.from_user.id, answ)
		wb.Close()
		os.remove(fname)
		fname = ""
	elif call.data == "50-100":
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\'+fname)
		ws = wb.ActiveSheet
		a = False
		y = 1
		while a==False:
			if y in range(1, ws.UsedRange.Rows.Count+1):
				if int(ws.Range('E'+str(y)).value) <= 50000:
					y+=1
				elif int(ws.Range('E'+str(y)).value) > 50001 and int(ws.Range('E'+str(y)).value) < 100000:
					answ += str(str(ws.Range('A'+str(y)).value)+" "+str(ws.Range('B'+str(y)).value)+" — "+str(ws.Range('D'+str(y)).value)+"₽"+" — "+str(ws.Range('E'+str(y)).value)+"$"+"\n")
					y+=1
				else:
					a = True
			else:
				a = True
		bot.send_message(call.from_user.id, answ)
		wb.Close()
		os.remove(fname)
		fname = ""
	elif call.data == ">100":
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\'+fname)
		ws = wb.ActiveSheet
		a = False
		y = 1
		while a==False:
			if y in range(1, ws.UsedRange.Rows.Count+1):
				if int(ws.Range('E'+str(y)).value) <= 100000:
					y+=1
				elif int(ws.Range('E'+str(y)).value) > 100001:
					answ += str(str(ws.Range('A'+str(y)).value)+" "+str(ws.Range('B'+str(y)).value)+" — "+str(ws.Range('D'+str(y)).value)+"₽"+" — "+str(ws.Range('E'+str(y)).value)+"$"+"\n")
					y+=1
				else:
					a = True
			else:
				a = True
		bot.send_message(call.from_user.id, answ)
		wb.Close()
		os.remove(fname)
		fname = ""
	elif call.data == "any":
		excel = win32com.client.Dispatch("Excel.Application")
		wb1 = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\cars.xlsx')
		ws1 = wb1.ActiveSheet
		wb1.SaveAs(u'C:\\Works\\pyBotCarsParse\\'+fname)
		keyboard = types.InlineKeyboardMarkup()
		key_1 = types.InlineKeyboardButton(text='<5000$', callback_data='<5')
		keyboard.add(key_1)
		key_2= types.InlineKeyboardButton(text='5001-10000$', callback_data='5-10')
		keyboard.add(key_2)
		key_3= types.InlineKeyboardButton(text='10001-20000$', callback_data='10-20')
		keyboard.add(key_3)
		key_4= types.InlineKeyboardButton(text='20001-50000$', callback_data='20-50')
		keyboard.add(key_4)
		key_5= types.InlineKeyboardButton(text='50001-100000$', callback_data='50-100')
		keyboard.add(key_5)
		key_6= types.InlineKeyboardButton(text='>100001$', callback_data='>100')
		keyboard.add(key_6)
		question = 'Укажи ценовой сегмент'
		bot.send_message(call.from_user.id, text=question, reply_markup=keyboard)
		wb1.Save()
		wb1.Close()
	else:
		pythoncom.CoInitializeEx(0)
		excel = win32com.client.Dispatch("Excel.Application")
		wb1 = excel.Workbooks.Open(u'C:\\Works\\pyBotCarsParse\\cars.xlsx')
		ws1 = wb1.ActiveSheet
		wb2 = excel.Workbooks.Add()
		ws2 = wb2.Worksheets.Add()
		wb2.SaveAs(u'C:\\Works\\pyBotCarsParse\\'+fname)
		d = 1
		for c in range(1, ws1.UsedRange.Rows.Count+1):
			if ws1.Range('A'+str(c)).value == call.data:
				ws2.Range('A'+str(d)+':E'+str(d)).value = ws1.Range('A'+str(c)+':E'+str(c)).value
				d += 1
		keyboard = types.InlineKeyboardMarkup()
		key_1 = types.InlineKeyboardButton(text='<5000$', callback_data='<5')
		keyboard.add(key_1)
		key_2= types.InlineKeyboardButton(text='5001-10000$', callback_data='5-10')
		keyboard.add(key_2)
		key_3= types.InlineKeyboardButton(text='10001-20000$', callback_data='10-20')
		keyboard.add(key_3)
		key_4= types.InlineKeyboardButton(text='20001-50000$', callback_data='20-50')
		keyboard.add(key_4)
		key_5= types.InlineKeyboardButton(text='50001-100000$', callback_data='50-100')
		keyboard.add(key_5)
		key_6= types.InlineKeyboardButton(text='>100001$', callback_data='>100')
		keyboard.add(key_6)
		question = 'Укажи ценовой сегмент'
		bot.send_message(call.from_user.id, text=question, reply_markup=keyboard)
		wb1.Close()
		wb2.Save()
		wb2.Close()

bot.polling(none_stop=True, interval=0)