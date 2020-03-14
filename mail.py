import openpyxl
import smtplib
import ssl
import getpass

loc = "NCC.xlsx"

smtp_server = "smtp.gmail.com"
port = 465
sender = "credenzuser@gmail.com"
passwd = getpass.getpass("Enter your password:- ")

wb = openpyxl.load_workbook(loc)
sheet = wb.active
rows = sheet.max_row
context = ssl.create_default_context()


context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, port) as server:
    server.login(sender, passwd)
    for i in range(2, rows+1):
        r_email = str(sheet.cell(row=i, column=4).value)
        name1 = str(sheet.cell(row=i, column=2).value)
        name2 = str(sheet.cell(row=i, column=3).value)        
        username = sheet.cell(row=i, column=7).value
        password = sheet.cell(row=i, column=8).value
        subject = "Here are your login credentials"
        if name2 == 'None':
        	body1 = f"Good morning, {name1}. Thank you for registering in NCC 2020."
        	body2 = f"\nHere are your login credentials. Username:- {username} and Password:- {password}."
        	body3 = f"\nWe hope you will enjoy our event"
        	message = f'Subject: {subject}\n\n{body1}{body2}{body3}' 
        else:
        	body1 = f"Good morning, {name1} and {name2}. Thank you for registering in NCC 2020."
        	body2 = f"\nHere are your login credentials. Username:- {username} and Password:- {password}."
        	body3 = f"\nWe hope you will enjoy our event"
        	message = f'Subject: {subject}\n\n{body1}{body2}{body3}'
        
        server.sendmail(
            sender, r_email, message
        )
        print("Email sent to "+r_email)
        
        
        
        
        
        
        
