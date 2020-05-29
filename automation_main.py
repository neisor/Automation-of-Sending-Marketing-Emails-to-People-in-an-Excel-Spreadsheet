#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu May 28 20:20:30 2020

@author: https://github.com/neisor
"""

#Import necessary libraries
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

#Read the excel file with contact information
df = pd.read_excel("PATH_TO_YOUR_EXCEL_FILE_WITH_EMAILS.xlsx")

#Setting the maximum amount of rows if we want to print all of the data from our marketing list
pd.set_option('display.max_rows', None)

#Variable to later display the total amount of for iterations and to determine the line in our Excel spreadsheet
i = 0

#Initialize the variable to count the amount of e-mails sent
emailCounter = 0

#Initialize a variable to count the amount of refused recipients (wrong e-mail addresses)
numberOfRefusedRecipients = 0

#Loop for sending e-mails based on the data from the Excel spreadsheet (from columns with name: e-mail & Company name)
for email in df['e-mail']:
    #If e-mail value in Excel spreadsheet equals to NaN, then continue to next loop and increase the value of loop counter to it's previous value + 1
    if pd.isnull(email) is True:
        i+=1
        continue

    #If e-mail value in Excel spreadsheet equals to null, then continue to next loop and increase the value of loop counter to it's previous value + 1
    if email == 'null':
        i+=1
        continue

    #Increase the value of the email counter variable to it's previous value + 1
    emailCounter+=1

    #sender == my email address
    #receiver == recipient's email address
    sender = "your@email.com"
    receiver = str(email)
    print('Receiver is ' + receiver + '\n')

    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Subject of your email"
    msg['From'] = sender
    msg['To'] = receiver

    #Get the name of the company to which we are sending the e-mail in the given instance of the loop
    nameOfCompany = df['Company name'][i]
    print('Company is ' + nameOfCompany + '\n')

    #Increase the value of loop counter to it's previous value + 1
    i+=1

    # Create the body of the message (a plain-text and an HTML version).
    text = """Dear company """ + str(nameOfCompany) + """,
CHANGE THIS TO YOUR TEXT - IN PLAIN TEXT"""

    html = """\
<html>
  <head></head>
  <body>
    <p>CHANGE THIS TO YOUR TEXT - IN HTML FORMAT</p>
  </body>
</html>
"""

    # Record the MIME types of both parts - text/plain and text/html.
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    msg.attach(part1)
    msg.attach(part2)

    # Connect to the SMTP server and send the message via SMTP server.
    smtpObj = smtplib.SMTP('your.smtp.server.com', 587) #587 is the default SMTP port number, change it if necessary
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login('your@email.com', 'PASSWORDtoYOURe-mail')

    # sendmail function takes 3 arguments: sender's address, recipient's address
    # and message to send - here it is sent as one string.
    try:
        smtpObj.sendmail(sender, receiver, msg.as_string())
    except:
        numberOfRefusedRecipients =+1
        print('This e-mail address seems to be wrong: ' + str(email) + '. Check it once the program ends.')

    #Close the smtplib object for this instance
    smtpObj.quit()

    #Print a line which acts as a separator so that the Company and Receiver in the next loop are more visible
    print('-------------')

print('Done. All e-mails were sent. In total, ' + str(emailCounter) + ' e-mails were sent.')
print('Amount of refused recipients (recipients with potentially wrong e-mail address) is: ' + str(numberOfRefusedRecipients) + '.\n')
