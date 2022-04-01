import email
from imapclient import IMAPClient
import os
import mimetypes
from datetime import datetime

# Input your Username and password here
username = 'Lorem'
password = 'Ipsum'

# Input your search text here. Currently searches for FROM field. Ref 'Part 2/2' comment
search_text='xxx@yy.ac.in'

client = IMAPClient("outlook.office365.com", ssl=True)
client.login(username, password)

# Uncomment to see the folder list
# result = client.list_folders()
# print(result)

# Change if different folder needs to be read
folder_selected = client.select_folder('INBOX', readonly=True)
print('%d messages in INBOX' % folder_selected[b'EXISTS'])

# Change the criterion as necessary. Check the Py3 documentation. FROM, TO, SINCE, BEFORE supported. Part 2/2
messages = client.search(['FROM', search_text])

# Uncomment to do a blind search
# messages = client.search()

print('%d messages found from search' % len(messages))

for msgid, data in client.fetch(messages, 'RFC822').items():

    email_message = email.message_from_bytes(data[b"RFC822"])

    counter = 1
    for part in email_message.walk():
        filename = part.get_filename()
        content_type = part.get_content_type()

        if part.get_content_maintype() == 'multipart':
            continue
        if not filename:
            ext = mimetypes.guess_extension(content_type)
            if not ext:
                ext = ".bin"
            elif 'html' in content_type:
                ext = ".html"
            filename = 'msg-%03d%s' % (counter, ext)
        counter += 1

        save_location = os.path.join(os.getcwd(), "Inbox", datetime.strptime(email_message.get("Date"),
                                                                             '%a, %d %b %Y %H:%M:%S +0530').strftime(
            "%d-%m-%Y"), email_message.get("Subject").replace(":", "_").replace("\n", "_").replace("\r", "_").replace(
            "\t", "_").replace("!", "_").replace("'", "").replace('"', "")).strip()

        try:
            if not os.path.exists(save_location):
                os.makedirs(save_location)
            with open(os.path.join(save_location, filename.strip()), 'wb') as fp:
                fp.write(part.get_payload(decode=True))
        except OSError:
            print("Could not save files from "+search_text+" of "+email_message.get("Date"))

    try:
        if not os.path.exists(save_location):
            os.makedirs(save_location)
        with open(os.path.join(save_location, "metadata.txt"), 'w') as fp:
            text = 'ID #%d \n Subject: "%s" \n Date: %s \n From: %s \n To: %s \n CC: %s \n' % (
                msgid, email_message.get("Subject"), email_message.get("Date"), email_message.get("From"),
                email_message.get("To"), email_message.get("CC"))
            fp.writelines(text)
    except OSError:
        continue


client.logout()
