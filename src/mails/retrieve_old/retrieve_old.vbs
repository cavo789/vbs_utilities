' -----------------------------------------------------------------
'
' Retrieve messages over 14 days old from Outlook's
' "Sent Items" folder
'
' Note: You may want to insert a 60 seconds delay and then
' add shortcuts to this script and to Outlook in your Startup
' folder
'
' @Author : Rob van der Woude
' http://www.robvanderwoude.com/vbstech_automation_outlook.php#CleanupSent
'
' -----------------------------------------------------------------

Option Explicit

Dim intMax, intOld
Dim objFolder, objItem, objNamespace, objOutlook

Const SENT =  5 ' Sent Items folder

intMax = 14 ' Messages older than this will be deleted (#days)
intOld =  0 ' Counter for the number of deleted messages

Set objOutlook = CreateObject( "Outlook.Application" )
Set objNamespace = objOutlook.GetNamespace( "MAPI" )

' Open default account (will fail if Outlook is closed)
' and delete Sent messages over 2 weeks old
objNamespace.Logon "Default Outlook Profile", , False, False

Set objFolder = objNamespace.GetDefaultFolder( SENT )

For Each objItem In objFolder.Items
	' Check the age of the message against the maximum allowed age
	If DateDiff( "d", objItem.CreationTime, Now ) > intMax Then
		intOld = intOld + 1

		'objItem.Delete  '<-- Uncomment to really delete mail

		' Or just echo the mail's creation datetime and subject
		wScript.echo objItem.CreationTime & "---" & objItem.Subject

	End If
Next

Set objFolder = Nothing
Set objNamespace = Nothing
Set objOutlook = Nothing