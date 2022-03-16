import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import nse
import bse

def sent_mail():

    body=nse.nsemarket()+bse.bsemarket()
    
    #The mail addresses and password
    sender_address = ''  #Enter sender mail address here
    sender_pass = ''   #Enter password here
    receiver_address = '' #sender receiver address here

    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'DAILY-STOCK-INFO'

    message.attach(MIMEText(body, 'plain'))

    #Attach excel
    excel_name = 'Stock-Book.xlsx'
    binary_excel = open(excel_name,'rb')
    payload = MIMEBase('application','octate-stream',Name=excel_name)
    payload.set_payload((binary_excel).read())

    encoders.encode_base64(payload)

    payload.add_header('Content-Decomposition','attachment',filename=excel_name)
    message.attach(payload)
    text = message.as_string()
    # establishing connection with gmail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    session.sendmail(sender_address, receiver_address,text)
    session.quit()
    print('Mail Sent')


