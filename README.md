# vbalibrary
## Description
This file simplify the use of ADODB for MSSQL connection
## How to include this connection class?
1. Include the Reference **Microsoft ActiveX Data Objects 6.1 Library
2. In the Microsoft Visual Basic for Applications Editor, right click on the Class Modules and use the import file.
## Usage
### Retrieve Record
```vba
Public Sub TestRetrieve()
    Dim dbconnection As New dbconnection
    
    ' https://support.microsoft.com/en-us/help/181734/how-to-invoke-a-parameterized-ado-query-using-vba-c-java
    Dim parameters As New Collection
    With parameters
        .Add dbconnection.createParameter(adInteger, 1)
        .Add dbconnection.createParameter(adInteger, 2)
    End With
    
    Call dbconnection.retrieveRecord("select * from assessmenttask where AT_id = ? or AT_id = ?", parameters)
End Sub
```
### Insert/ Update Record
```vba
Public Sub TestSave()
	Dim dbconnection As New dbconnection
	Call dbconnection.beginTransaction
	Dim rs as ADODB.Recordset
	Set rs = dbconnection.save("Insert into....")
	'OR
	'Set rs = dbconnection.save("Update into....")
	Call dbconnection.endTransaction
End Sub
```
