# Insert your python code here
# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *
import datetime as dt

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE

while True:
    directory = single_input("Please input the path to the folder your invoices are saved in.")
    directory = directory.strip('"').strip("'")
    if directory[-1] != '\\':
        directory = directory+'\\'
    # print(directory)
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message("Error", "The folder does not have the correct files. Please try again.")


# INSERT ANSWER FOR QUESTION 2 HERE

product_description = {}
product_prices = {}

descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index,row in descriptions.iterrows():
    Unit_number = row['Unit number']
    Description = row.Description
    product_description[Unit_number] = Description
    # print(f'This row contains {Unit_number} in {Description}.')

    price = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for index,row in price.iterrows():
    Unit_number = row['Unit number']
    prices = row['Price per unit']
    product_prices[Unit_number] = prices
    # print(f'This row contains {Unit_number} in {prices}.')


    # INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num, filename, directory = directory):
    template = openpyxl.load_workbook(directory+filename)
    # pr(template)
    selected_sheet = template['OrderForm']
    Unit_Number0 = selected_sheet['B24'].value
    Unit_Number1 = selected_sheet['B25'].value
    Unit_Number2 = selected_sheet['B26'].value
    Unit_Number3 = selected_sheet['B27'].value
    Unit_Number4 = selected_sheet['B28'].value
    date = str(selected_sheet[('C9')].value)
    date = date.split()[0].split('-')
    date = f'{date[1]}/{date[2]}/{date[0]}'
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value

    order_data = {}
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = date #fix this code
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {str(selected_sheet['C19'].value)}'
    if Unit_Number0 != None:
        order_data['Price0'] = (product_prices[Unit_Number0])
        order_data['Quantity0'] = str(selected_sheet['C24'].value)
        order_data['Amount0_Raw'] = order_data['Price0'] * order_data['Quantity0']
        order_data['Amount0'] =     f'${order_data['Amount0_Raw']:,.2f}'
        order_data['Description0'] = product_description[Unit_Number0]
        order_data['Total'] = (order_data['Amount0_Raw'])
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'

    if Unit_Number1 != None:
        order_data['Price1'] = (product_prices[Unit_Number1])
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        order_data['Amount1_Raw'] = order_data['Price1'] * order_data['Quantity1']
        order_data['Amount1'] =   f'${order_data['Amount1_Raw']:,.2f}'
        order_data['Description1'] = product_description[Unit_Number1]
        order_data['Total'] = (order_data['Amount0_Raw']) + (order_data['Amount1_Raw'])
        order_data['Price1'] = f'${order_data['Price1']:,.2f}' 

    if Unit_Number2 != None:
        order_data['Price2'] = (product_prices[Unit_Number2])  #The Problem is that this is a null value currently
        order_data['Quantity2'] = str(selected_sheet['C26'].value)
        order_data['Amount2_Raw'] = order_data['Price2'] * order_data['Quantity2']
        order_data['Amount2'] =   f'${order_data['Amount2_Raw']:,.2f}'
        order_data['Description2'] = product_description[Unit_Number2]
        order_data['Total'] = (order_data['Amount0_Raw']) + (order_data['Amount1_Raw']) + (order_data['Amount2_Raw'])
        order_data['Price2'] = f'${order_data['Price2']:,.2f}'

    if Unit_Number3 != None:
        order_data['Price3'] = (product_prices[Unit_Number3])
        order_data['Quantity3'] = str(selected_sheet['C27'].value)
        order_data['Amount3_Raw'] = order_data['Price3'] * order_data['Quantity3']
        order_data['Amount3'] =   f'${order_data['Amount3_Raw']:,.2f}'
        order_data['Description3'] = product_description[Unit_Number3]
        order_data['Total'] = (order_data['Amount0_Raw']) + (order_data['Amount1_Raw']) + (order_data['Amount2_Raw']) + (order_data['Amount3_Raw']) 
        order_data['Price3'] = f'${order_data['Price3']:,.2f}'

    if Unit_Number4 != None:
        order_data['Price4'] = (product_prices[Unit_Number4])
        order_data['Quantity4'] = str(selected_sheet['C28'].value)
        order_data['Amount4_Raw'] = order_data['Price4'] * order_data['Quantity4']
        order_data['Amount4'] =  f'${order_data['Amount4_Raw']:,.2f}'
        order_data['Description4'] = product_description[Unit_Number4] 
        order_data['Total'] = (order_data['Amount0_Raw']) + (order_data['Amount1_Raw']) + (order_data['Amount2_Raw']) + (order_data['Amount3_Raw']) + (order_data['Amount4_Raw'])
        order_data['Price4'] = f'${order_data['Price4']:,.2f}'

   
    order_data['Total'] = f'${order_data['Total']:,.2f}'
    return order_data, name, email, date


# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data = order_data, directory = directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)


   # INSERT ANSWER FOR QUESTION 5 HERE
gmail_password_path = select_file('Please select the gmail password file')
Gmail_input = single_input("Please input your gmail username")
    
def send_email(invoice_num, name = name, email = email, date = date):
    gmail_user = Gmail_input
    gmail_password = open(gmail_password_path).read()


    # List of recipients

    recipients = email 

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = name

    body = f"""Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062 Phone: 
(801) 555-2372
E-mail: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


    # INSERT ANSWER FOR QUESTION 6 HERE
invoice_num = 8515

filenames = os.listdir(directory)
for f in filenames:
    if 'Order' in f and 'xlsx' in f:
        order_data = get_order_data(invoice_num, f, directory)[0]
        name = get_order_data(invoice_num, f, directory)[1]
        email = get_order_data(invoice_num, f, directory)[2]
        date = get_order_data(invoice_num, f, directory)[3]
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date)
        invoice_num = invoice_num + 1