import csv
import time
import smtplib
import ssl
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Define the subject and write the email
subject = input('Type the subject of your email: ')
body = """

{formule_de_politesse} {prenom} {nom},

Here you can hardcode your email within the program and use brackets to personalize it, 
with the infos given by your csv file.

"""
# Define the sender and other variables for the email campaign
sender_email = input("Type your email address: ")
password = input("Type your password and press enter: ")
path = input(
    "Enter the pathway to your contacts file to who you want to send emails to: ")
path_to_attached_document = ""
filename_attached_document = None
part = ""
invalid_input = True


# Function to ask if the user wants to attach a file to the email campaign
def attached_file():
    global path_to_attached_document
    global filename_attached_document
    global part
    global invalid_input
    attach_file = input(
        "Do you want to attach a file to this email campaign ? (y/n): ")
    if attach_file == "y":
        path_to_attached_document = input(
            "Enter the pathway to the file you want to attach: ")
        filename_attached_document = path_to_attached_document.split('/')[-1]

        print('Your attached document will be send ender name "{}"'.format(
            filename_attached_document))
        # Open PDF file in binary mode
        with open(path_to_attached_document, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            "attachment; filename= {}".format(filename_attached_document),
        )
        invalid_input = False
    elif attach_file == "n":
        invalid_input = False
    else:
        print('You have to respond by "y" (yes) or "n" (no).')


while invalid_input:
    attached_file()

df = pd.read_csv(path)

# Starting the massive email campaign
context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(sender_email, password)
    print('Connected to SMTP server...')
    time.sleep(3)
    print('Starting the massive email campaign...')
    time.sleep(3)
    # for i in range(len(df["Nom"])):
    for i in range(518):
        # Create a multipart message
        message = MIMEMultipart()

        # Add body to email
        message["From"] = sender_email
        message["Subject"] = subject

        commission_capitalized = df['Commission'][i].capitalize()
        prenom = df['Prenom'][i]
        nom = df['Nom'][i]
        if df['Genre'][i] == "M.":
            formule_de_politesse = "EXAMPLE: Monsieur le député"
        elif df['Genre'][i] == "Mme":
            formule_de_politesse = "EXAMPLE: Madame la députée"
        personalized_message = body.format(formule_de_politesse=formule_de_politesse, prenom=prenom,
                                           nom=nom, commission=commission_capitalized)
        message["To"] = df['Mail'][i]
        # Add text to email
        message.attach(MIMEText(personalized_message, "plain"))
        # Add attached file to email
        if filename_attached_document != None:
            message.attach(part)
        text_to_send = message.as_string()
        server.sendmail(
            sender_email,
            df['Mail'][i],
            text_to_send,
        )
        print('Email sent to {} {} with success'.format(
            df['Prenom'][i], df['Nom'][i]))
