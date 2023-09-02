import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Email configuration
sender_email = "*********@outlook.com"
sender_password = "**********"
recipient_email = "***********@gmail.com"
subject = "Email with Attachment"
message_body = "Please find the attached file."

# Create an email message
msg = MIMEMultipart()
msg["From"] = sender_email
msg["To"] = recipient_email
msg["Subject"] = subject

# Add the email body (plain text)
msg.attach(MIMEText(message_body, "plain"))

# Attach a file
file_path = r"D:\abc\xyz.txt"
attachment = open(file_path, "rb")

part = MIMEApplication(attachment.read(), Name="xyz.txt")
part['Content-Disposition'] = f'attachment; filename="{file_path}"'
msg.attach(part)

# SMTP server settings
smtp_server = "smtp.outlook.com"
smtp_port = 587  # The default SMTP port is 587 for secure connections (TLS/STARTTLS)

# Create an SMTP connection
try:
    smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
    smtp_connection.starttls()  # Enable TLS encryption for the connection

    # Log in to the SMTP server
    smtp_connection.login(sender_email, sender_password)

    # Send the email with attachment
    smtp_connection.sendmail(sender_email, recipient_email, msg.as_string())

    # Close the SMTP connection
    smtp_connection.quit()
    print("Email with attachment sent successfully!")

except Exception as e:
    print(f"Error: {str(e)}")
