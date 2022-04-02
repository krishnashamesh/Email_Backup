Do you want to create a local copy of your email account with attachments and messages?

Do you want to fetch only a couple of emails and cannot afford to import the whole .pst file?

Then this script is for you!!!

# Creates a local Backup of the Email Account

Uses IMAP (Internet Message Access Protocol)

Prerequisites : Python 3

Steps and Guidelines:
1. Input Username, Password and Search Text in the env.properties file
2. Folders will be created as per timestamp and subject
3. Attachments will be downloaded into respective folders
4. Email messages will be saved as HTML to preserve formatting
5. Metadata is preserved inside the last sub-folder to enable search. [This will require 'Windows Search finding file content']

Execution Steps
1. Run "python .\backup.py" without the quotes.
