import email
import mimetypes
import os
import subprocess
import sys
import traceback
from datetime import datetime

from imapclient import IMAPClient

# implement pip as a subprocess:
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'imapclient'])

# process output with an API in the subprocess module:
reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze'])
installed_packages = [r.decode().split('==')[0] for r in reqs.split()]

# print(installed_packages)

config_props = {}
with open("env.properties") as props_file:
    for line in props_file:
        if line.startswith("#"):
            continue
        elif line.find("=") > 0:
            key, value = line.partition("=")[::2]
            config_props[key.strip()] = value.strip()

# print(config_props)

# Input your Username and password here
username = config_props.get('username')
while username == '':
    print()
    username = input(
        "Please enter the username associated with the account. (You can also set it in the env.properties file): ")
password = config_props.get('password')
while password == '':
    password = input(
        "Please enter the password associated with the account. (You can also set it in the env.properties file) : ")

# Currently searches for FROM field. Ref 'Part 2/2' comment
search_text = config_props.get("search_from_text")

client = IMAPClient(config_props.get("IMAPClient_host"), ssl=True)
client.login(username, password)

# Uncomment to see the folder list
# result = client.list_folders()
# print(result)

# Change if different folder needs to be read
folder_selected = client.select_folder(config_props.get("selected_folder"), readonly=True)
print('%d messages in INBOX' % folder_selected[b'EXISTS'])

# Change the criterion as necessary. Check the Py3 documentation. FROM, TO, SINCE, BEFORE supported. Part 2/2

if search_text == '':
    messages = client.search()
else:
    messages = client.search(['FROM', search_text])

print('%d messages found from search' % len(messages))

current_email_num = 0

for msgid, data in client.fetch(messages, 'RFC822').items():

    email_message = email.message_from_bytes(data[b"RFC822"])

    counter = 1
    # print("Processing Email with ID %d " % msgid)
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

        subject = email_message.get("Subject").replace(":", "_").replace("\n", "_").replace("\r", "_").replace("\\n",
                                                                                                               "_").replace(
            "\\r", "_").replace("\r\n", "").replace("\\r\\n", "").replace(
            "\\t", "_").replace("!", "_").replace("'", "").replace('"', "").replace("<", "_").replace(">", "_").replace(
            "\\", "_").replace("/", "_").replace("|", "_").replace("?", "_").replace("*", "_").replace(",",
                                                                                                       "_").replace(".",
                                                                                                                    "_")
        subject = (subject[:200]) if len(subject) > 200 else subject
        # subject = (subject[:-1]) if subject[-1] == '.' else subject

        save_location = os.path.join(os.getcwd(), "EmailBackups", "Inbox", datetime.strptime(email_message.get("Date"),
                                                                                             '%a, %d %b %Y %H:%M:%S +0530').strftime(
            "%d-%m-%Y"), subject).strip()

        try:
            if not os.path.exists(save_location):
                os.makedirs(save_location)
            with open(os.path.join(save_location, filename.strip()), 'wb') as fp:
                fp.write(part.get_payload(decode=True))
        except OSError:
            print("Could not save files from " + search_text + " of " + email_message.get("Date"))
            print(traceback.format_exc())

    try:
        if not os.path.exists(save_location):
            os.makedirs(save_location)
        with open(os.path.join(save_location, "metadata.txt"), 'w') as fp:
            text = 'ID #%d \n Subject: "%s" \n Date: %s \n From: %s \n To: %s \n CC: %s \n' % (
                msgid, email_message.get("Subject"), email_message.get("Date"), email_message.get("From"),
                email_message.get("To"), email_message.get("CC"))
            fp.writelines(text)
        current_email_num += 1
    except OSError:
        continue
        print(traceback.format_exc())

    print('%d messages processed from %d messages' % (current_email_num, len(messages)))

client.logout()
print("Email Backup Successful")
