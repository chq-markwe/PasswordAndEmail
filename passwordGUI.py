import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import pandas as pd
from email_function import send_email
import random
import string
import openpyxl

# --- classes ---

class MyWindow:
    def __init__(self, parent):

        self.parent = parent
        self.parent.title('Password Generator and Email Sender')
        self.parent.iconbitmap('flag.ico')

        self.filename = None
        self.df = None

        self.text = tk.Text(self.parent)
        self.text.pack()

        # Generate Passwords button
        self.script_button = tk.Button(self.parent, text='Generate Passwords', command=self.generate_passwords)
        self.script_button.pack()

        # Generate Emails button
        self.script_button = tk.Button(self.parent, text='Send Emails', command=self.send_emails)
        self.script_button.pack()

        # Exit button
        self.script_button = tk.Button(self.parent, text="Exit", command=self.parent.destroy)
        self.script_button.pack(pady=20)

    def generate_passwords(self):
        def generate_password():
            length = 10

            lower = string.ascii_lowercase
            upper = string.ascii_uppercase
            num = string.digits
            symbols = string.punctuation

            all = lower + upper + num + symbols

            temp = random.sample(all, length)

            # create the password
            password = "".join(temp)

            # return the password
            return (password)

        # Ask for input file
        name = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])

        if name:
            if name.endswith('.csv'):
                self.df = pd.read_csv(name)
            else:
                self.df = pd.read_excel(name)

            self.filename = name

            # display directly
            self.text.insert('end', str(self.df.head()) + '\n')

        wb = openpyxl.load_workbook(name)
        wb.sheetnames  # get names of all spreadsheet in the file
        ws = wb["Sheet1"]  # get the first spreadsheet by name
        ws.max_row  # get the number of rows in the sheet

        for i in range(2, ws.max_row + 1):
            pw = generate_password()
            ws[f"F{i}"].value = pw

        # Ask for saved file name
        fname = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                        ("All files", "*.*")))
        # note: this will fail unless user ends the fname with ".xlsx"
        # self.df.to_excel(fname)
        wb.save(fname)

        self.filename = fname
        self.df = pd.read_excel(fname)

        self.text.insert('end', self.filename + '\n')
        self.text.insert('end', str(self.df.head()) + '\n')

        self.text.insert('insert', '\n'+ 'Passwords Generated'+ '\n'+ '\n')


    def send_emails(self):

        # Ask for input file
        email_list = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])

        if email_list:
            if email_list.endswith('.csv'):
                self.df = pd.read_csv(email_list)
            else:
                self.df = pd.read_excel(email_list)

            self.filename = email_list

        send_email(email_list)
        self.text.insert('insert', '\n' + 'Emails Sent')

# --- main ---

if __name__ == '__main__':
    root = tk.Tk()
    top = MyWindow(root)
    root.mainloop()