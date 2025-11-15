import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from email.utils import formataddr


from credentials import *

# Gmail SMTP configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = personal_ac
sender_password = personal_app_pw


print("Connecting to Gmail server...")
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, sender_password)
print("Login successful!")


def email_function(receipient_email,subject_line,email_body,attachment_path):
    # Create the email message
    msg = MIMEMultipart()
    # msg["From"] = sender_email
    msg["From"] = formataddr(("Data Team Uganda", sender_email))
    msg["To"] = receipient_email
    msg["Subject"] = subject_line

    # Attach body text
    msg.attach(MIMEText(email_body, "plain"))

    # Attach the file if it exists
    if os.path.exists(attachment_path):
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file to ASCII characters
        encoders.encode_base64(part)

        # Add header
        filename = os.path.basename(attachment_path)
        part.add_header("Content-Disposition", f"attachment; filename={filename}")

        # Attach the file
        msg.attach(part)
        # print(f"Attached file: {filename}")

    else:
        print(f"⚠️ No valid attachment found — sending email without attachment: {attachment_path}")
    server.send_message(msg)
    time.sleep(5)