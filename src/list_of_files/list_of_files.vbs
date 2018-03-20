' -----------------------------------------------------------------
'
' Take advantage of Windows Desktop Search and very fast, 
' get the list of MS Access applications present on the computer 
' (local drives so also return files present, f.i. on a D: drive if 
' you run the script from the C: drive; don't scan network drives).
'
' Source : https://blogs.technet.microsoft.com/heyscriptingguy/2006/12/15/how-can-i-list-all-the-access-database-files-on-a-computer/
'
' Output : This script will display the list (absolute filenames) 
' 		of MS Access files like f.i. 
'
'		C:\Temp\Inventory.accdb
'		D:\Dev\SubFolder\BigApp.mdb
'
' Note : just change searched extensions like .xlsx, .png, ... 
'		for searching other type of files
'
' -----------------------------------------------------------------

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

' Just change extensions for searching for other files like .xlsx f.i.
objRecordSet.Open _
	"SELECT System.ItemPathDisplay " & _
	"FROM SYSTEMINDEX " & _
	"WHERE (System.FileExtension='.accdb') OR " & _ 
	"(System.FileExtension='.mdb')", objConnection

objRecordSet.MoveFirst

wScript.echo "List of files retrieved :"
wScript.echo "========================="
wScript.echo ""

Do Until objRecordset.EOF
	wScript.Echo objRecordset.Fields.Item("System.ItemPathDisplay")
	objRecordset.MoveNext
Loop

objRecordSet.Close
Set objRecordSet = Nothing

objConnection.Close
Set objConnection = Nothing