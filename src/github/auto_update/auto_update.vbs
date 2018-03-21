' -----------------------------------------------------------------
'
' Connect to GitHub, get the raw version of a file and download it
'
' The script will check for newer version on GitHub and if
' there is one, the script will overwrite himself with
' that newer version.
'
' Based on a script of
' @author Michi Lehenauer §https://github.com/michiil)
' then modified by Christophe Avonture
'
' @Link https://github.com/michiil/vbs_scrips/blob/master/WZV-Excel.vbs
'
' -----------------------------------------------------------------

' The content to download
url = "https://raw.githubusercontent.com/cavo789/vbs_utilities/master/src/folders/get_folder_size/get_folder_size.vbs"

Set objReq = CreateObject("Msxml2.ServerXMLHttp.6.0")
objReq.setTimeouts 500, 500, 500, 500
' If you're behind a firewall, uncomment the following line
' and mention the proxy address and port
'objReq.setProxy 2, "your.proxy.net:8080", ""
objReq.open "GET", url, False
objReq.send

If Err.Number = 0 Then
	If objReq.Status = 200 Then
		' Get the content, just downloaded
		sDownloadedContent = objReq.responseText

		' Get the original content, this script

	 	sScriptName = wScript.ScriptFullName
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objTextFile = objFSO.OpenTextFile(sScriptName, 1)
		sOriginalContent = objTextFile.ReadAll
		objTextFile.Close

		' Compare if the two contents are differents
		If (sOriginalContent <> sDownloadedContent) Then
			' If yes, for instance, rewrite this script by
			' the new content ==> auto-update
			Set objTextFile = objFSO.OpenTextFile(sScriptName, 2)

			objTextFile.Write (sDownloadedContent)
			objTextFile.Close

			'wScript.echo sScriptName & " has been updated"

		End If

		Set objTextFile = Nothing
		Set objFSO = Nothing

	End If
Else
	' Ok, in case of error, don't panic, just do nothing
	Err.Clear

End If