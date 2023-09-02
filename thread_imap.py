import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta

# Your Outlook email credentials
email_address = "*******@outlook.com"
password = "**********"

# Connect to the Outlook server
imap_server = "imap.outlook.com"  # Use the appropriate IMAP server for your Outlook account
mailbox = imaplib.IMAP4_SSL(imap_server)

# Log in to your account
mailbox.login(email_address, password)

# Select the mailbox (inbox in this case)
mailbox.select("inbox")

# Calculate the date 2 days ago
two_days_ago = (datetime.now() - timedelta(days=2)).strftime("%d-%b-%Y")

# Search for emails from the last 2 days
status, email_ids = mailbox.search(None, f'SINCE "{two_days_ago}"')

# Get the list of email IDs
email_id_list = email_ids[0].split()

# Dictionary to store threads
threads = {}

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

    # Extract threading information from headers
    in_reply_to = msg.get("In-Reply-To", "")
    references = msg.get("References", "")

    # Use the "In-Reply-To" and "References" headers to determine the thread
    if in_reply_to:
        thread_id = in_reply_to
    elif references:
        thread_id = references.split()[0]
    else:
        thread_id = email_id

    # Create a thread if it doesn't exist
    if thread_id not in threads:
        threads[thread_id] = {
            "subject": subject,
            "from": from_name,
            "messages": []
        }

    # Add the email message to the thread
    threads[thread_id]["messages"].append(msg)

print(f"threads are {threads}")

# Now, `threads` is a dictionary where each key is a thread ID (usually an email message ID),
# and the value is another dictionary containing the thread's subject, sender, and a list of
# email messages in that thread.

# Close the mailbox
mailbox.logout()
