' -----------------------------------------------------------------
'
' Simple pattern VBS script for demonstrating how to retrieve
' the default's mail signature, create a new email, add a file 
' to it and send (or display) the mail
'
' -----------------------------------------------------------------

Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)

'Creating HTML outlook body for correct signature format
MsgHTML = "<HTML><style>p {font-size: 1.6em;}</style><pTest email</p></HTML>"

' Display a mail window so we can retrieve the signature
With objMail
    .Display
    Signature = .HTMLbody
End With

With objMail
    .To = "john@doe.com"
    .subject = "Un test d'envoi"
    .HTMLbody = MsgHTML & Signature
    '.Attachments.Add sFileName
    .Send ' or .Display
End With

Set objMail = Nothing
Set objOutlook = Nothing