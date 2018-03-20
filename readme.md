# VBS_Utilities

> A list of small, stand-alone and straight-forward, VBS utilities

## Table of Contents

- [Scripts](#scripts)
    - [Files](#files)
        - [Get list of files](#get-list-of-files)
    - [Folders](#folders)
        - [Get folder size](#get-folder-size)
    - [Outlook](#outlook)
        - [Get mail's subject](#get-mail-s-subject)
        - [Send email](#send-email)
- [Author](#author)
- [Licence](#licence)

## Scripts

### Files

#### Get list of files

Take advantage of Windows Desktop Search and very fast, get the list
of MS Access applications present on the computer (local drives so
also return files present, f.i. on a D: drive if you run the script
from the C: drive; don't scan network drives).

Just adjust the searched extension for searching for any other type 
of files like .docx, .png, .xlsx, ...

This script is really, really fast but only works for local drives.

[go to files/list_of_files](https://github.com/cavo789/vbs_utilities/tree/master/src/files/list_of_files)

### Folders 

#### Get folder size

Scan a folder recursively and display the size of each folders (first level)

[go to folders/get_folder_size](https://github.com/cavo789/vbs_utilities/tree/master/src/folders/get_folder_size)

### Outlook 

#### Get mail's subject

Simple pattern VBS script for retrieving the list of MS Outlook
objects like emails, contacts, ...

The current demo script will just echoed the email's subject in a DOS prompt

Adjust the constant for GetDefaultFolder to retrieve emails,
contacts, ...

[go to mails/get_subject](https://github.com/cavo789/vbs_utilities/tree/master/src/mails/get_subject)

#### Send email

Simple pattern VBS script for demonstrating how to retrieve the default's mail signature, create a new email, add a file to it and send (or display) the mail

[go to mails/get_subject](https://github.com/cavo789/vbs_utilities/tree/master/src/mails/send_mail)

## Author

See copyrights when mentionned or comments at the top of the script.
If no one are mentionned, the author is Christophe Avonture.

## Licence

[MIT](LICENSE)