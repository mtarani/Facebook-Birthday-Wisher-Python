import pyexcel as p
import datetime
from fbchat import Client
from fbchat.models import *


#STORE FACEBOOK USERNAME
username='YOUR FACEBOOK USERNAME'
#STORE FACEBOOK PASSWORD
password='YOUR FACEBOOK PASSWORD'
#DEFAULT TEXT MESSAGE
default_message="Happy Birthday!!!"
#DEFAULT PATH FOR CAKE IMAGE
default_cake='cake.jpg'

#DATE EXTRACTION
now= datetime.datetime.now()
date=int(now.strftime("%d"))
month=int(now.strftime("%m"))

#PATH FOR "birthday.xlsx" BIRTHDAY DATA EXCEL FILE
records = p.get_records(file_name="birthdays.xlsx")
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
        # SENDING TEXT MESSAGE
        client.sendMessage(birthday_data['Message'][i], thread_id=user.uid, thread_type=ThreadType.USER)
        print (birthday_data['Name'][i] +" has been wished with text")
        # SENDING CAKE IMAGE REMOVE THIS LINE IF DO NOT WISH TO SEND IMAGE OF CAKE
        client.sendLocalImage(birthday_data['Image'][i] , thread_id=user.uid, thread_type=ThreadType.USER)
        print(birthday_data['Name'][i] + " has been wished with image")