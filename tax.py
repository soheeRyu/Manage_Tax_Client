'''
Created by: Sohee Ryu

This program is for accountants, accounting technicians, or administrators working in accountin firms.
Users can see the list of clients, search clients by their SIN or first name.
Users can add a client, edit a client's tax amount if they need to change tax amount due to any missing documents.
If users want to see a list of clients who owe taxes to remind them, they can also do it with this program. 
Users can send emails to clients regarding tax filing result before they file taxes.
EMAIL_PASSWORD: Enter app password for EMAIL.
'''
import os
from email.message import EmailMessage
import smtplib, ssl
import pandas as pd
from openpyxl import load_workbook

class Client:
    def __init__(self, sin, firstName, lastName, amount, ccb, caip, email):
        self.sin = sin
        self.firstName = firstName
        self.lastName = lastName
        self.amount = amount
        self.ccb = ccb
        self.caip = caip
        self.email = email
        
    def enterNewClientInfo(self):
        '''Ask the user to enter the new client information'''
        global client
        self.sin = int(input("Enter client SIN:\n"))
        if self.sin not in sinList:
            self.firstName = input("Enter the client's first name:\n")
            self.lastName = input("Enter the client's last name:\n")
            self.amount = int(input("Enter the client tax amount(-:Refund | +: Owing):\n"))
            self.ccb = int(input("Enter the client CCB amount per month:\n"))
            self.caip = int(input("Enter the client CAIP amount if eligible, if not, enter 0:\n")       )
            self.email = input("Enter the client email address:\n")
            client = Client(self.sin, self.firstName, self.lastName, self.amount, self.ccb, self.caip, self.email)
            clientList.append(client)
            addClientToFile()
        else:
            print("The SIN already exists on file.")
            
    def formatInfo(self):
        '''Format client info'''
        return (str(self.sin), str(self.firstName), str(self.lastName), str(self.amount), str(self.ccb), str(self.caip), str(self.email))
    
    def emailTemplate(self):
        '''Create email template to be sent to clients regarding their tax filing'''
        global taxResult, ccbResult, caipResult, emailContent
        emailContent = [
        "Hello " + client.firstName + "\n"
        "Your tax filing prep is done.\n\n"
        "Below is the detail. If you have any questions or concerns, please let me know.\n"
        "Otherwise, we will e-file your tax in 5 days.\n\n"
        "Name: " + client.firstName + " " + client.lastName + "\n"
        "Tax Result: " + taxResult + "\n"
        "CCB Amount: " + ccbResult + "\n"
        "CAIP Amount: "+ caipResult + "\n"
        "Thank you very much for your business with us."
        ]
        client.sendEmails()
        
    def sendEmails(self):
        '''Send emails to clients'''
        global emailContent
        sender = 'EMAILADDRESS'
        password = "EMAILPASSWORD"
        receiver = client.email 
        subject = '2022 Tax results'
        body = str(*emailContent).strip()
        context = ssl.create_default_context()
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context)
        server.login(sender, password)
        
        em = EmailMessage()
        em['From'] = sender
        em['To'] = receiver
        em['Subject'] = subject
        em.set_content(body)
        server.send_message(em)
        print("The email sent to " + client.firstName + " " + client.lastName + ".\n\n")
    
    def __str__(self):
        return Client.formatInfo(self)
    
def readClientsFile():
    '''Read from clients.xlsx'''
    
    global wb, sheet, path, dataframe1, clientList, sinList, firstNameList, lastNameList, amountList, ccbList, caipList, emailList, client
    clientList = []
    path = 'clients.xlsx'
    if os.path.exists (path):
        wb = load_workbook(path)
        sheet = wb.active

        dataframe1 = pd.read_excel(path)  
        sinList = list(dataframe1['SIN'])
        firstNameList = list(dataframe1['First name'])
        lastNameList = list(dataframe1['Last name'])
        amountList = list(dataframe1['Amount'])
        ccbList = list(dataframe1['CCB'])
        caipList = list(dataframe1['CAIP'])
        emailList = list(dataframe1['Email'])
        
        for i in range(0,len(sinList)):
            sin = sinList[i]
            firstName = firstNameList[i]
            lastName = lastNameList[i]
            amount = amountList[i]
            ccb = ccbList[i]
            caip = caipList[i]
            email = emailList[i]
            client= Client(sin,firstName,lastName, amount, ccb, caip, email)
            clientList.append(client)

def addClientToFile():
    '''Add a new client to the excel file'''
    clientData = [client.sin, client.firstName, client.lastName, client.amount, client.ccb, client.caip, client.email]
    sheet.append(clientData)   
    wb.save(path)

def displayClientList():
    '''Display the list of clients as a table'''
    print(dataframe1)

def displayClientInfo():
    '''Display client information'''
    global n
    print(f'{clientList[n].sin: <10}{clientList[n].firstName: <10}{clientList[n].lastName: <10}{clientList[n].amount: <10}{clientList[n].ccb: <10}{clientList[n].caip: <10}{clientList[n].email:}')
    
def searchClientBySin():
    '''Search for a client using their sin that user enters'''
    global n
    enterClientsin = int(input("Enter the client sin:\n"))
    if enterClientsin in sinList:
        n = sinList.index(enterClientsin)
        displayClientInfo()
    else:
        print("Not in the client list.")
        
def searchClientByFirstName():
    '''Search for a client using their first name that user enters'''
    global n, client
    enterClientFirstName = input("Enter the client first name:\n")
    enterClientFirstName = enterClientFirstName.capitalize()
    if enterClientFirstName in firstNameList:
        for client in clientList:
            if client.firstName == enterClientFirstName:
                data = client.formatInfo()
                print('{: <10}{: <10}{: <10}{: <10}{: <10}{: <10}{: <10}'.format(*data))
    else:
        print("Not in the client list.")
       
def editClientTaxAmount():
    '''Ask the user to edit a client info'''
    global n, client, dataframe1
    clientSIN = int(input("Enter the client's SIN\n"))
    if clientSIN in sinList:
        clientNewAmount = int(input("Enter the new client's tax amount. (-: Refund | +: Owing)\n"))
        n = sinList.index(clientSIN)
        sheet["D"+str(n+2)] = clientNewAmount
        wb.save(path)
    else:
        print("Invalid SIN")
    
def searchClientOwing():
    '''Search clients who owe taxes'''
    global n, client
    for i in amountList:
        if i > 0:
            n = amountList.index(i)
            displayClientInfo()
      
def taxRefundCalculate():
    '''See if the client will get tax refund or has tax owing'''
    global taxResult
    if client.amount < 0:
        taxResult = "You have tax refund of $" + str(-client.amount) + ".\n" + "The amount will be deposited to your bank account that you have registered on CRA."
    elif client.amount == 0:
        taxResult = 'Your refund for this year is $0.'
    else:
        taxResult = "You owe taxes of $" + str(client.amount) + ".\n" + "Please pay the amount by April 30, 2023."
    
def ccbCalculate():
    '''See the CCB amount if the client is eligible'''
    global ccbResult
    if client.ccb > 0:
        ccbResult = "Your monthly CCB amount is $" + str(client.ccb) +"."
    else:
        ccbResult = "You are not eligible for CCB."
        
def caipCalculate():
    '''See the CAIP amount if the client is eligible'''
    global caipResult
    if client.caip > 0:
        caipResult = "You will received $" + str(client.caip) + " quarterly."
    else:
        caipResult = "You are not eligible for CAIP due to the province where you lived on December 31, 2022."
                 
while True:
    options = input("""Option:\n
1 - Display the client list\n
2 - Search for client by SIN\n
3 - Search for client by first name\n
4 - Add a client\n
5 - Edit client tax return amount\n
6 - Search for clients who owe taxes\n
7 - Create an email template and send an email to client\n
8 - Back to the main options\n
0 - EXIT\n""")
    
    readClientsFile()
    
    if options == '1':
        displayClientList()
    elif options == '2':
        searchClientBySin()
    elif options == '3':
        searchClientByFirstName()
    elif options == '4':
        Client.enterNewClientInfo(self=Client)
    elif options == '5':
        editClientTaxAmount()
    elif options == '6':
        searchClientOwing()
    elif options == '7':
        for client in clientList:
            taxRefundCalculate()
            ccbCalculate()
            caipCalculate()
            Client.emailTemplate(self= Client)
        print("Done")
    elif options == '8':
        pass
    elif options == '0':
        break
    else:
        print("Invalid option")
    print("\nBack to the previous option")
