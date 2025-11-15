# log-in to email server
"""
# ######################################################################
# # Email With Attachments Python Script
# # 
# ######################################################################
# """
# 
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
# 
smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp-mail.outlook.com"  # Google SMTP Server
print("Connecting to server...")

def email_function(sender_email,sender_password,receipient_email,subject_line,email_body):
    TIE_server = smtplib.SMTP(smtp_server, smtp_port)
    TIE_server.starttls()
    TIE_server.login(sender_email, sender_password)
    print("Succesfully connected to server")
    #%%
    # Make a MIME object to define parts of the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['Subject'] = subject_line
    msg['To'] = ', '.join(receipient_email)
    # 
    # 
    # Attach the body of the message
    msg.attach(MIMEText(email_body, 'plain'))
# 
    # Define the file to attach
    # filename = attachment_path
# 
    # # Open the file in python as a binary
    # with open(filename, 'rb') as attachment:
    #     # Encode as base 64
    #     attachment_package = MIMEBase('application', 'octet-stream')
    #     attachment_package.set_payload(attachment.read())
    #     encoders.encode_base64(attachment_package)
    #     attachment_package.add_header('Content-Disposition', f"attachment; filename= {filename}")
    #     msg.attach(attachment_package)
    # Send emails to all recipients at once
    TIE_server.sendmail(sender_email,receipient_email, msg.as_string())
    #clear space
    TIE_server.quit()

    # %%
