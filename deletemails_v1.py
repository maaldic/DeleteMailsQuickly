import imaplib
import email
from email.header import decode_header

# It also works for other non-business domains like gmail.
server = 'outlook.office365.com'

# IMAP is not appliable for business mail domains.
user = "YOUR_MAIL_ADDRESS"
password = "YOUR_PASSWORD"

imap = imaplib.IMAP4_SSL(server)

# authenticate
imap.login(user, password)

# Lists for available mailboxes we want to delete:
# imap.list()
imap.select("Inbox")

# to get mails after a specific date
#status, messages = imap.search(None, 'SINCE "01-JUL-2023"')

# to get mails before a specific date
status, messages = imap.search(None, 'BEFORE "31-JAN-2021"')

# messages is returned as a list of a single byte string of mail IDs separated by a space,
# let's convert it to a list of integers:
messages = messages[0].split(b' ')

counter = 0
"""
# This block helps you to see the subject part of the mails will be deleted. 
# Uncomment if you do not have time issues, because it will slow down the deleting process. 

for mail in messages:
    _, msg = imap.fetch(mail, "(RFC822)")
    counter = counter + 1
    # you can delete the for loop for performance if you have a long list of emails
    # because it is only for printing the SUBJECT of target email to delete
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                # if it's a bytes type, decode to str
                subject = subject.decode('latin-1')
            print("Deleting", subject)
    # mark the mail as deleted
"""
for mail in messages:
    counter = counter+1
    imap.store(mail, "+FLAGS", "\\Deleted")

# Shows how many mails erased.
print(counter)

# permanently remove mails that are marked as deleted from the selected mailbox (in this case, INBOX).
#imap.expunge()

# close the mailbox
imap.close()
# logout from the account
imap.logout()

