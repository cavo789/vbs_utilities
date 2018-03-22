# VBS_Utilities

> A list of small, stand-alone and straight-forward, VBS utilities

I've written some of these scripts, modified somes others or just get a copy of an exising script (original author always mentionned and linked).

## Table of Contents

- [Scripts](#scripts)
	- [Files](#files)
		- [Find](#find)
	- [Folders](#folders)
		- [Get folder size](#get-folder-size)
		- [Get list of files](#get-list-of-files)
		- [Select folder](#select-folder)
	- [GitHub](#github)
		- [auto_update](#auto_update)
	- [Outlook](#outlook)
		- [Get mail's subject](#get-mails-subject)
		- [Send email](#send-email)
- [Author](#author)
- [Licence](#licence)

## Scripts

### Files

#### Find

Take advantage of Windows Desktop Search and very fast, get the list
of MS Access applications present on the computer (local drives so
also return files present, f.i. on a D: drive if you run the script
from the C: drive; don't scan network drives).

Just adjust the searched extension for searching for any other type 
of files like .docx, .png, .xlsx, ...

This script is really, really fast but only works for local drives.

[go to files/find](https://github.com/cavo789/vbs_utilities/tree/master/src/files/find)

### Folders 

#### Get folder size

Scan a folder recursively and display the size of each folders (first level)

[go to folders/get_folder_size](https://github.com/cavo789/vbs_utilities/tree/master/src/folders/get_folder_size)

#### Get list of files

Get the list of files of the current folder + subfolders and generate a .csv file with files informations like path, size, extensions, author, ... making then easy to work with that list in Excel

[go to folders/get_list_of_files](https://github.com/cavo789/vbs_utilities/tree/master/src/folders/get_list_of_files)

#### Select folder

Display a select folder dialog box and then return then selected foldername

[go to folders/select_folder](https://github.com/cavo789/vbs_utilities/tree/master/src/folders/select_folder)

### GitHub 

#### Auto_update 

Sample script to demonstrate how it's possible to add an auto-update feature in a VBS script.

The script will check for newer version on GitHub and if there is one, the script will overwrite himself with that newer version.

[go to github/auto_update](https://github.com/cavo789/vbs_utilities/tree/master/src/github/auto_update)

### Outlook 

#### Get mail's subject

Simple pattern VBS script for retrieving the list of MS Outlook
objects like emails, contacts, ...

The current demo script will just echoed the email's subject in a DOS prompt

Adjust the constant for GetDefaultFolder to retrieve emails,
contacts, ...

[go to mails/get_subject](https://github.com/cavo789/vbs_utilities/tree/master/src/mails/get_subject)

#### Retrieve old messages

Retrieve messages over 14 days old from Outlook's `Sent Items` folder

[go to mails/retrieve_old](https://github.com/cavo789/vbs_utilities/tree/master/src/mails/retrieve_old)

#### Send email

Simple pattern VBS script for demonstrating how to retrieve the default's mail signature, create a new email, add a file to it and send (or display) the mail

[go to mails/send_mail](https://github.com/cavo789/vbs_utilities/tree/master/src/mails/send_mail)

## Author

See copyrights when mentionned or comments at the top of the script.
If no one are mentionned, the author is Christophe Avonture.

## Licence

[MIT](LICENSE)