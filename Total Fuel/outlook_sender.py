import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Replace these variables with your Outlook email details
sender_email = 'pokuttu@yalelo.ug'
sender_password = 'Aligator@1'
recipient_email = 'eokello@yalelo.ug'
subject = 'Excel Files Fuel Usage'
body = 'Hello, This is an automated File from D&A'

# Create the MIME object
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = recipient_email
message['Subject'] = subject

# Attach the body of the email
message.attach(MIMEText(body, 'plain'))

# Connect to the SMTP server (Outlook)
smtp_server = 'smtp-mail.outlook.com'
smtp_port = 587
with smtplib.SMTP(smtp_server, smtp_port) as server:
    # Start TLS (Transport Layer Security) mode
    server.starttls()
    
    # Login to your Outlook account
    server.login(sender_email, sender_password)
    
    # Send the email
    server.sendmail(sender_email, recipient_email, message.as_string())
