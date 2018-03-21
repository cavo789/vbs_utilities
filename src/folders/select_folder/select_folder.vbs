' -----------------------------------------------------------------
'
' Display a select folder dialog box and then return then selected
' foldername
'
' @Author : Rob van der Woude
' Slightly modified by Christophe Avonture
'
' @Link http://www.robvanderwoude.com/vbstech_ui_selectfolder.php
'
' -----------------------------------------------------------------

Option Explicit

Dim strPath

strPath = SelectFolder("")

If strPath = vbNull Then
	wScript.echo "Cancelled"
Else
	wScript.echo "Selected Folder: """ & strPath & """"
End If

' ----------------------------------------------------------------
' This function opens a "Select Folder" dialog and will
' return the fully qualified path of the selected folder
'
' Argument:
'	 sStartFolder	The root folder where you can start browsing;
'		  if an empty string is used, browsing starts
'		  on the local computer
' ----------------------------------------------------------------
Function SelectFolder(sStartFolder)

	Dim objFolder, objItem, objShell

	' Custom error handling
	On Error Resume Next

	SelectFolder = vbNull

	' Create a dialog object
	Set objShell  = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0, "Select Folder", 0, sStartFolder)

	' Return the path of the selected folder
	If IsObject(objfolder) Then
		SelectFolder = objFolder.Self.Path
	End If

	Set objFolder = Nothing
	Set objshell  = Nothing

	On Error Goto 0

End Function