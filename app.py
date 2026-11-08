#question 1 
valid_files = ['ProductData.xlsx', 'MasterInvoice.pdf',
                  'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input('Please enter path.')
    directory = directory.strip().strip('"').strip("'")

    if directory[-1:] != '\\':
        directory = directory + '\\'

    try:
        filenames = os.listdir(directory)
        file_count = 0
        for file in valid_files:
            if file in filenames:
                file_count = file_count + 1
            else:
                show_message('Incorrect', f'{file} is not in the directory.')

        if file_count == 6:
            show_message("Correct", "This path contains all of the files.")
            break
        else:
            show_message("Incorrect", "Path is missing files. Please try again.")
    except:
        show_message("Incorrect", "That path does not exist. Please try again.")
#question2
description_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions= {}

for index, row in description_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices= {}

for index, row in prices_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
#question3 
def get_order_data(invoice_num, filename, directory):
    order_form = openpyxl.load_workbook(directory + filename)
    selected_sheet = order_form['OrderForm']
    order_data = {}

    order_data['InvoiceNumber'] = str(invoice_num)
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['Date'] = date

    # Total
    item0 = selected_sheet['B24'].value
    unit_price0 = product_prices.get(item0, 0)
    total = float(selected_sheet['C24'].value * unit_price0)
    if pd.isnull(selected_sheet['B25'].value) == False:
        item1 = selected_sheet['B25'].value
        unit_price1 = product_prices.get(item1, 0)
        total = total + float(selected_sheet['C25'].value * unit_price1)
    if pd.isnull(selected_sheet['B26'].value) == False:
        item2 = selected_sheet['B26'].value
        unit_price2 = product_prices.get(item2, 0)
        total = total + float(selected_sheet['C26'].value * unit_price2)
    if pd.isnull(selected_sheet['B27'].value) == False:
        item3 = selected_sheet['B27'].value
        unit_price3 = product_prices.get(item3, 0)
        total = total + float(selected_sheet['C27'].value * unit_price3)
    if pd.isnull(selected_sheet['B28'].value) == False:
        item4 = selected_sheet['B28'].value
        unit_price4 = product_prices.get(item4, 0)
        total = total + float(selected_sheet['C28'].value * unit_price4)
    order_data['Total'] = f'${total:,.2f}'

    # Customer info
    email = selected_sheet['C13'].value
    name = selected_sheet['C14'].value
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value

    # Address
    order_data['CustomerAddress1'] = str(address)
    order_data['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    # Customer name
    order_data['CustomerName'] = str(name)

    # Company name
    order_data['CustomerCompany'] = str(selected_sheet['C15'].value)

    # Order price, quantity, amount
    item0 = selected_sheet['B24'].value
    price0 = product_prices.get(item0, 0)
    quantity0 = selected_sheet['C24'].value
    order_data['Description0'] = str(product_descriptions.get(item0, ''))
    order_data['Price0'] = f'${price0:,.2f}'
    order_data['Quantity0'] = str(quantity0)
    order_data['Amount0'] = f'${price0 * quantity0:,.2f}'

    if pd.isnull(selected_sheet['B25'].value) == False:
        item1 = selected_sheet['B25'].value
        price1 = product_prices.get(item1, 0)
        quantity1 = selected_sheet['C25'].value
        order_data['Description1'] = str(product_descriptions.get(item1, ''))
        order_data['Price1'] = f'${price1:,.2f}'
        order_data['Quantity1'] = str(quantity1)
        order_data['Amount1'] = f'${price1 * quantity1:,.2f}'

    if pd.isnull(selected_sheet['B26'].value) == False:
        item2 = selected_sheet['B26'].value
        price2 = product_prices.get(item2, 0)
        quantity2 = selected_sheet['C26'].value
        order_data['Description2'] = str(product_descriptions.get(item2, ''))
        order_data['Price2'] = f'${price2:,.2f}'
        order_data['Quantity2'] = str(quantity2)
        order_data['Amount2'] = f'${price2 * quantity2:,.2f}'

    if pd.isnull(selected_sheet['B27'].value) == False:
        item3 = selected_sheet['B27'].value
        price3 = product_prices.get(item3, 0)
        quantity3 = selected_sheet['C27'].value
        order_data['Description3'] = str(product_descriptions.get(item3, ''))
        order_data['Price3'] = f'${price3:,.2f}'
        order_data['Quantity3'] = str(quantity3)
        order_data['Amount3'] = f'${price3 * quantity3:,.2f}'

    if pd.isnull(selected_sheet['B28'].value) == False:
        item4 = selected_sheet['B28'].value
        price4 = product_prices.get(item4, 0)
        quantity4 = selected_sheet['C28'].value
        order_data['Description4'] = str(product_descriptions.get(item4, ''))
        order_data['Price4'] = f'${price4:,.2f}'
        order_data['Quantity4'] = str(quantity4)
        order_data['Amount4'] = f'${price4 * quantity4:,.2f}'

    return order_data, name, email, date


order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name, email, date)
#question4
def create_invoice(invoice_num, order_data, directory): 
    create_pdf(directory + 'MasterInvoice.pdf', f"Invoice {invoice_num}.pdf", order_data)
create_invoice(8515, order_data, directory) 
#question5 
def send_email(invoice_num, name, email, date, directory): 

    gmail_user = "sales@pythontechco.com"
    gmail_password_path = select_file('Please select the gmail password file')

    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"

    # Body of the email
    

    body = f"""Dear {name},
        
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
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment
    path_to_file1 = f'{directory}Invoice {invoice_num}.pdf'
    #class notes:   invoice_template.save(f"{directory}Invoice_{row['Invoice Number']}.xlsx")
    #path_to_file1 = os.getcwd()+'\\Invoice Template.xlsx'
    #path_to_file2 = os.getcwd()+'\\Invoice Data.xlsx'

    message = add_attachment(path_to_file1, message)
    #message = add_attachment(path_to_file2, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'harleyarogers@gmail.com', '11/13/2023', directory)
#question6 
filenames = os.listdir(directory)
filenames.sort()
invoice_num=8515 
for file in filenames: 
    if 'Order' in file and 'xlsx' in file: 
        order_data= get_order_data(invoice_num, file, directory) 

        name=order_data[1]
        email=order_data[2]
        date=order_data[3]
        create_invoice(invoice_num, order_data[0],directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num=invoice_num+1
      
