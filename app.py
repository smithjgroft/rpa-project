# IMPORT PACKAGES IN THIS CELL
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows



# INSERT ANSWER FOR QUESTION 1 HERE

directory_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx','Order4.xlsx']


while True:
    count = 0
    directory = select_file('Please select any file in the folder containing invoice documentation.')
    last_file = directory.split('/')[-1]
    directory = directory.split(last_file)[0]
    files = os.listdir(directory)
    for file in directory_files:
        if file in files:
            count += 1
    if count == 6:
        break
    else:
        show_message('error', 'directory missing necessary files')



# INSERT ANSWER FOR QUESTION 2 HERE
product_description = {}
product_prices = {}

descriptions = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Descriptions')
for i, row in descriptions.iterrows():
    product_description[row['Unit number']] = row['Description']

prices = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Prices')
for i, row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']




# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    workbook = openpyxl.load_workbook(directory+filename)
    sheet = workbook['OrderForm']
    name = sheet['C14'].value
    email = sheet['C13'].value
    date = sheet['C9'].value.strftime('%m/%d/%Y')
    
    order_data = {}
    total = 0
    start = 24
    i = 0
    while True:
        unit_cell = 'B' + str(start + i)
        unit = sheet[unit_cell].value
        if unit is not None:
            price = product_prices[unit]
            quantity_cell = 'C' + str(start + i)
            quantity = sheet[quantity_cell].value
            desc = product_description[unit]
            order_data[f'Price{i}'] = price
            order_data[f'Quantity{i}'] = quantity
            order_data[f'Description{i}'] = desc
            amount = float(price) * float(quantity)
            total += amount
            order_data[f'Amount{i}'] = f'${amount:,.2f}'
            i += 1
        else:
            break
    order_data['Total'] = f'${total:,.2f}'

    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = date

    order_data['CustomerAddress1'] = sheet['C16'].value
    order_data['CustomerAddress2'] = sheet['C17'].value + ', ' + sheet['C18'].value + ' ' + str(sheet['C19'].value)
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = sheet['C15'].value
    
    return order_data, name, email, date


# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)



# INSERT ANSWER FOR QUESTION 5 HERE

def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "smithjgroft@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""
    Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.
    
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['CC'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = f'{directory}Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file1, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()
        



# INSERT ANSWER FOR QUESTION 6 HERE

filenames = sorted(os.listdir(directory))

inv_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(inv_num, file, directory)
        create_invoice(inv_num, order_data, directory)
        send_email(inv_num, name, email, date, directory)
        inv_num += 1

