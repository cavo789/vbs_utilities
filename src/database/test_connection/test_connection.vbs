' -----------------------------------------------------------------
'
' This script will allow you to quickly check if access to 
' your SQL Server database is possible.
'
' The objective is to establish a connection and check if it 
' works before starting, e. g., to investigate your program 
' code or the permissions required for the user to use your 
' tables, views, stored procedures,...
'
' This script will only do this, i.e. try to connect to the database,
' which will eliminate the possibility of a login problem.
'
' @src https://github.com/cavo789/vbs_utilities
'
' -----------------------------------------------------------------

Option Explicit

Class clsHelper

	Sub ForceCScriptExecution()

		Dim sArguments, Arg, sCommand

		If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then

			' Get command lines paramters'
			sArguments = ""
			For Each Arg In WScript.Arguments
				sArguments=sArguments & Chr(34) & Arg & Chr(34) & Space(1)
			Next

			sCommand = "cmd.exe cscript.exe //nologo " & Chr(34) & _
			WScript.ScriptFullName & Chr(34) & Space(1) & Chr(34) & sArguments & Chr(34)

			' 1 to activate the window
			' true to let the window opened
			Call CreateObject("Wscript.Shell").Run(sCommand, 1, true)

			' This version of the script (started with WScript) can be terminated
			wScript.quit

		End If

	End Sub

End Class

Dim cHelper, objConnection, objRecordSet
Dim serverName, dbName, login, password, sDSN, sSQL
Dim StartTime, EndTime

Set cHelper = New clsHelper
Call cHelper.ForceCScriptExecution()
Set cHelper = Nothing

If Wscript.Arguments.count < 3 Then 

	wScript.Echo "Usage: test_connection.vbs [server_name] [database_name] [login] [password]"
	wScript.Echo ""
	wScript.Echo "Try to connect to a SQL Server database by using a specific"
	wScript.Echo "SQL user login and password."
	
	' And quit
	wScript.Quit 0
	
Else

	serverName = trim(Wscript.Arguments(0))
	dbName     = trim(Wscript.Arguments(1))
	login      = trim(Wscript.Arguments(2))
	password   = trim(Wscript.Arguments(3))
	
End if

If ((serverName = "") or (dbName = "") or (login = "")) Then 
    wScript.Echo "Error, [server_name], [database_name] or [login] can't be empty"
    wScript.Quit 
End If

wScript.Echo "Try to connect " & dbName & " on " & serverName & "... "
wScript.Echo ""

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

sDSN = "Provider=SQLOLEDB;Data Source=" & serverName & ";" & _
	"Trusted_Connection=False;Initial Catalog=" & dbName & ";" & _
	"User ID=" & login & ";Password=" & password & ";"

wScript.Echo "Connection string: " & sDSN
wScript.Echo ""

objConnection.Open sDSN

' Just retrieve the active database name
' (should be equal to the specified dbName but it isn't important)
sSQL = "SELECT db_name()"
wScript.Echo sSQL

StartTime = Timer()

' Run the query
objRecordSet.Open sSQL, objConnection

EndTime = Timer()

' If the code comes here, no error was encountered, connection is thus OK
wScript.Echo "Active database name= " & objRecordSet.Fields(0).Value
wScript.Echo ""
wScript.Echo "Test successful, database connection has been successfully established"

' Display the elapsed time in seconds; how many seconds needed to get 
' the answer from the server
wScript.Echo ""
WScript.Echo "Time taken: " & FormatNumber(EndTime - StartTime, 2)

objRecordSet.close

Set objRecordSet=nothing
Set objConnection=nothing
