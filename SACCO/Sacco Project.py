import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas 
from datetime import datetime#, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from cryptography.fernet import Fernet
from email.mime.text import MIMEText


formatted_date = datetime.today().strftime('%m/%d/%Y')
print(formatted_date)

from creds import *

sacco = pd.read_excel("P:/Pertinent Files/Python/scripts/SACCO/Contributions.xlsx")



def send_email_with_statement(data, sender_email, sender_password):
    # Group data by unique combinations of name and email
    grouped_data = data.groupby(["Name", "Email"])

    # SMTP Configuration
    smtp_server = "smtp-mail.outlook.com"  # Outlook SMTP Server
    smtp_port = 587  # Standard secure SMTP port

    for (name, email), group_data in grouped_data:
        # Extract recipient information
        recipient_name = name.split()[0]  # Get first name
        recipient_email = email

        # Create PDF attachment with statement
        pdf_buffer = BytesIO()
        pdf = canvas.Canvas(pdf_buffer)
        pdf.setFont("Helvetica", 10)
        pdf.drawString(50, 800, f"Date: {formatted_date}")
        pdf.drawString(50, 750,  f"Statement for {name}")
        pdf.drawString(50, 725, "Email: " + recipient_email)
        pdf.drawString(50, 700 ,"Position: " + recipient_email)

        # Add table headers
        pdf.drawString(50, 650, "Date")
        pdf.drawString(200, 650, "Contribution")

        # Write statement data
        y_pos = 630
        subtotal=0
        for _, row in group_data.iterrows():
            pdf.drawString(50, y_pos, str(row["Date"]))
            pdf.drawString(200, y_pos, str(row["Contribution"]))
            subtotal += row["Contribution"]
            y_pos -= 20
        # Add subtotal row
        pdf.drawString(50, y_pos, "Subtotal:")
        pdf.drawString(200, y_pos, str(subtotal))

        pdf.setFillColorRGB(1, 0, 0)
        pdf.drawString(50, 10, "Should you have any issues regarding this statement, please reach out to management")

        pdf.save()
        pdf_data = pdf_buffer.getvalue()
        pdf_buffer.close()

        # Save PDF locally
        filename = f"{name}_statement.pdf"
        with open(f"P:/Pertinent Files/Python/scripts/SACCO/{filename}", "wb") as pdf_file:
            pdf_file.write(pdf_data)

        print(f"{recipient_name}'s PDF file saved successfully as {filename}!")

        
        from pypdf import PdfReader, PdfWriter

        reader = PdfReader(f"P:/Pertinent Files/Python/scripts/SACCO/{filename}")

        writer = PdfWriter()
        writer.append_pages_from_reader(reader)
        writer.encrypt(recipient_email)

        with open(f"P:/Pertinent Files/Python/scripts/SACCO/{filename}", "wb") as out_file:
            writer.write(out_file)

        print(f"{recipient_name}'s PDF file encrypted successfully as {filename}!")

        body = """
        Hello Team,
        
        Please find Accounts Receivables Report attached.

        Regards,
        Audit Team
        """

        # Email configuration
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = f"Statement for {recipient_name}"
        msg.attach(MIMEText(body, 'plain'))

        # Attach PDF
        with open(f"P:/Pertinent Files/Python/scripts/SACCO/{filename}", "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {filename}")
            msg.attach(part)

        # Connect to SMTP server and send email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())

        print(f"Email sent successfully to {recipient_email}!")

# Example usage
send_email_with_statement(sacco, pbi_user, pbi_pass)

