import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import PyPDF2
import threading

# Step 1: Read the Excel file
df = pd.read_excel('contacts.xlsx')  # Replace 'contacts.xlsx' with your file name

# Extract email addresses and names
emails = df['email'].tolist()
names = df['name'].tolist()

# Step 2: Define the email body template
email_body_template = """
Dear {name},

This is a test email with a PDF attachment.

Best regards,
Your Name
"""

# Step 3: Function to send email with PDF attachment
def send_email(email, name):
    from_email = 'your_email'  # Replace with your email address
    to_email = email
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = 'Test Email with PDF Attachment'

    # Customize the email body with the recipient's name
    body = email_body_template.format(name=name)
    msg.attach(MIMEText(body, 'plain'))

    # Attach PDF file
    filename = 'filename.pdf'  # Replace with your PDF file name
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
    msg.attach(part)

    # Connect to the SMTP server
    smtp_server = 'smtp.gmail.com'  # Replace with your SMTP server
    smtp_port = 587  # Replace with your SMTP port
    smtp_username = 'your email'  # Replace with your SMTP username
    smtp_password = 'password'  # Replace with your SMTP password, app password generated using google account 

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, to_email, msg.as_string())
        print(f'Email with PDF attachment sent to {name} ({to_email})')
        server.quit()
    except Exception as e:
        print(f'Failed to send email with PDF attachment to {name} ({to_email}): {str(e)}')

# Step 4: Function to send emails using threads
def send_emails_in_threads(emails, names):
    threads = []
    for email, name in zip(emails, names):
        thread = threading.Thread(target=send_email, args=(email, name))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

# Step 5: Call the function to send emails
send_emails_in_threads(emails, names)
