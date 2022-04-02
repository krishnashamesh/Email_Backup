# Creates a local Backup of the Outlook Account

Uses IMAP (Internet Message Access Protocol)

Prerequisites
1. Python 3
2. IMAPClient [pip install imapclient]

Steps and Guidelines:
1. Input Username, Password and Search Text in the env.properties file
2. Folders will be created as per timestamp and subject
3. Attachments will be downloaded into respective folders
4. Email messages will be saved as HTML to preserve formatting
5. Metadata is preserved inside the last sub-folder to enable search. [This will require 'Windows Search finding file content']

Execution Steps
1. Run "python .\backup.py" without the quotes.
