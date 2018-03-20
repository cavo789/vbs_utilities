' -----------------------------------------------------------------
'
' Simple pattern VBS script for retrieving the list of MS Outlook
' objects like emails, contacts, ...
'
' Output : Currently, just echoed the email's subject in a DOS prompt
'
' Note : Adjust the constant for GetDefaultFolder to retrieve emails,
'       contacts, ...
'
' -----------------------------------------------------------------

Set olApp=CreateObject("Outlook.Application")
Set olns=olApp.GetNameSpace("MAPI")

' Get the list of constants here : 
' https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/oldefaultfolders-enumeration-outlook
Set objFolder=olns.GetDefaultFolder(6)

wScript.echo "eMails found in your Inbox folder : "
wScript.echo "=================================== "
wScript.echo ""

For each item1 in objFolder.Items
    wscript.echo item1.subject
Next