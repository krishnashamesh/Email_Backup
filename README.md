Do you want to create a local copy of your email account with attachments and messages?

Do you want to fetch only a couple of emails and cannot afford to import the whole .pst file?

Then this script is for you!!!

# Creates a local Backup of the Email Account

# For users without Python Installation

1. Download "Windows Executable.zip" from the repository (https://github.com/krishnashamesh/Email_Backup/raw/main/Windows%20Executable.zip)
2. Unzip and Double Click on exe_backup.exe

# For users with Python Installation

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

Please refer to the current issues tab in case of issues.

Known issues:
1. Windows executable does not run on Win7 machines
2. Windows executable throws error when Python version < 3.0 is installed in the host sytem.
3. Login fails when MFA is activated in Outlook accounts.