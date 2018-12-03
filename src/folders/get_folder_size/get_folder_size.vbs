' -----------------------------------------------------------------
'
' Scan a folder recursively and display the size of each folders 
' (first level)
'
' AUTH: TY WSH
'
' @Link https://github.com/hadrins/VBscript/blob/master/dirscan.vbs'
'
' Output will be something like 
'
' Root Directory is: C:\Christophe
'
'           5.459.941 C:\Christophe\archives
'         509.572.286 C:\Christophe\Dropbox
'         509.572.286 C:\Christophe\Repository
'         509.572.286 C:\Christophe\Sites
'       6.846.646.798 C:\Christophe\tools
'       2.305.454.237 C:\Christophe\zip
' -----------------------------------------------------------------

On Error Resume Next

Dim RootDir, FileSystem, RootFolder, SubFolders, Folder
Dim FolderSize, Tmp

If Wscript.Arguments.count <> 1 Then 

	wScript.Echo "Usage: get_folder_size.vbs [root directory]"
	wScript.Echo ""
	wScript.Echo "Given a root directory,dirscan will scan"
	wScript.Echo "all directories and output the size of"
	wScript.Echo "each subdirectory."
	
	' And quit
	wScript.Quit 0
	
Else

	RootDir = Wscript.Arguments(0)
	wScript.Echo "Root Directory is: " & RootDir
	wScript.Echo " "

End If 

Set FileSystem = CreateObject("Scripting.FileSystemObject")
Set RootFolder = FileSystem.GetFolder(RootDir)

If Err.Number <> 0 Then

	wScript.Echo "(" & Err.Number & ") " & Err.Description
	wScript.Echo ""
	wScript.Echo "The path you entered is invalid, please " & _
		"select a different path."
		
	Wscript.Quit Err.Number
	
End If

Set SubFolders = RootFolder.SubFolders

For Each Folder In SubFolders

	FolderSize = Folder.Size
	Tmp = FormatNumber (FolderSize, 0, 0, 0, -1)
	Tmp = Right (Space(20) & Tmp, 20)
	wScript.Echo Tmp & " " & Folder.Path

Next

wScript.Echo "Press enter to exit"

Input = wscript.stdin.Read(1)
