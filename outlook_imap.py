import imaplib
import email
from email.header import decode_header
import os

# Your Outlook email credentials
email_address = "********@outlook.com"
password = "**********"

# Connect to the Outlook server
imap_server = "imap.outlook.com"  # Use the appropriate IMAP server for your Outlook account
mailbox = imaplib.IMAP4_SSL(imap_server)

# Log in to your account
mailbox.login(email_address, password)

# Select the mailbox (inbox in this case)
mailbox.select("inbox")

# Search for emails
status, email_ids = mailbox.search(None, "ALL")

# Get the list of email IDs
email_id_list = email_ids[0].split()

# Loop through email IDs and fetch the emails
for email_id in email_id_list:
    # Fetch the email
    status, email_data = mailbox.fetch(email_id, "(RFC822)")
    raw_email = email_data[0][1]

    # Parse the raw email data
    msg = email.message_from_bytes(raw_email)

    # Extract email details
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8")

    from_name, encoding = decode_header(msg.get("From"))[0]
    if isinstance(from_name, bytes):
        from_name = from_name.decode(encoding or "utf-8")

    print(f"Subject: {subject}")
    print(f"From: {from_name}")

    # Check for attachments
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if "attachment" in content_disposition:
                filename = part.get_filename()

                if filename:
                    # Download the attachment
                    save_path = os.path.join(os.getcwd(), filename)
                    with open(save_path, "wb") as attachment_file:
                        attachment_file.write(part.get_payload(decode=True))
                    print(f"Attachment '{filename}' saved to {save_path}")

    # Print the email content
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if "text/plain" in content_type:
                body = part.get_payload(decode=True).decode("utf-8")
                print(f"Email Content:\n{body}")

# Close the mailbox
mailbox.logout()
