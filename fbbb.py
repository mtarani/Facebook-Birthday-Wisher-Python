import pyexcel as p
import datetime
from fbchat import Client
from fbchat.models import *


username='mttarrat@gmail.com'
password='asdfghjklmm'
default_message="Happy Birthday!!!"
default_cake='E:/Projects/Python/Birthday Wisher/cake.jpg'


now= datetime.datetime.now()
date=int(now.strftime("%d"))
month=int(now.strftime("%m"))

records = p.get_records(file_name="E:/Projects/Python/Birthday Wisher/birthdays.xlsx")
birthday_data = {'Name': [], 'Message': [], 'Image': []}
for record in records:
    if(record['Date']==date and record['Month']==month):
        print("Its "+record['Name']+"'s Birthday today")
        birthday_data['Name'].append(record['Name'])
        if record['Message']:
            birthday_data['Message'].append(record['Message'])
        else:
            birthday_data['Message'].append(default_message)
        if record['Image']:
            birthday_data['Image'].append(record['Image'])
        else:
            birthday_data['Image'].append(default_cake)

if len(birthday_data['Name'])>0:
    client = Client(username, password)
    for i in range(0,len(birthday_data['Name'])):
        users = client.searchForUsers(birthday_data['Name'][i])
        user =users[0]
        client.sendMessage(birthday_data['Message'][i], thread_id=user.uid, thread_type=ThreadType.USER)
        print (birthday_data['Name'][i] +" has been wished with text")
        client.sendLocalImage(birthday_data['Image'][i] , thread_id=user.uid, thread_type=ThreadType.USER)
        print(birthday_data['Name'][i] + " has been wished with image")