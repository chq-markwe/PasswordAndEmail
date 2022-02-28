# Python code to send email to a list of
# emails from a spreadsheet

# import the required libraries
import pandas as pd
import win32com.client

def send_email(path):
    #Generate the email
    outlook = win32com.client.Dispatch('outlook.application')

    # choose sender account
    send_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'eccosupport@expeditors.com':
            send_account = account
            break

    # reading the spreadsheet
    email_list = pd.read_excel(path)

    # getting the names and the emails
    first_names = email_list['FIRST NAME']
    usernames = email_list['USERNAME']
    emails = email_list['EMAIL']
    passwords = email_list['PASSWORD']

    # iterate through the records
    for i in range(len(emails)):

        mail = outlook.CreateItem(0)

        #Code to assign the sending account
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

        # for every record get the name and the email addresses
        first_name = first_names[i]
        username = usernames[i]
        email = emails[i]
        password = passwords[i]

        # the message to be emailed
        mail.To = email
        mail.Subject = f'User Account Created for {username}'
        mail.HTMLBody =(f"""Hello {first_name},<br><br>
             Your account has been created.<br><br>
             Please see below for your username and temporary password:<br><br>
             Username: {username}<br>
             Password: {password}<br><br>
             Please note: you will have to change your password upon logging in to ECCO<br><br><br>
             Best Regards, <br><br>
             <strong>Critical Logistics Management</strong>
             <strong><p style='color:red;font-size:.7rem'>Expeditors Carrier Capacity Optimization</p><strong><br><br>
             <strong>Email</strong> eccosupport@expeditors.com <br><br>
             <a href='https://www.expeditors.com'>
             <img src ='https://info.expeditors.com/hubfs/without%20tag%20line.gif', alt = 'Expeditors Logo no tag', width = '150px'> 
             </a>
             """)
        mail.BCC = 'eccosupport@expeditors.com'

        mail.Send()
