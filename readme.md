# VBS_Utilities

![Banner](./banner.svg)

> A list of small, stand-alone and straight-forward, VBS utilities

I've written some of these scripts, modified some others or just get a copy of an existing script (original author always mentioned and linked).

## Table of Contents

* [Database](#database)
	* [Test connection](#test-connection)
* [Files](#files)
	* [Find](#find)
* [Folders](#folders)
	* [Get folder size](#get-folder-size)
	* [Get list of files](#get-list-of-files)
	* [Select folder](#select-folder)
* [GitHub](#github)
	* [Auto_update](#autoupdate)
* [Outlook](#outlook)
	* [Get mail's subject](#get-mails-subject)
	* [Retrieve old messages](#retrieve-old-messages)
	* [Send email](#send-email)
* [Author](#author)
* [Licence](#licence)

## Scripts

### Database

#### Test connection

> Try to establish a connection to a SQL database

This script will allow you to quickly check if access to your SQL Server database is possible.

The objective is to establish a connection and check if it works before starting, e. g., to investigate your program code or the permissions required for the user to use your tables, views, stored procedures,...

This script will only do this, i.e. try to connect to the database, which will eliminate the possibility of a login problem.

[go to database/test_connection](https://github.com/cavo789/vbs_utilities/tree/master/src/database/test_connection)

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
