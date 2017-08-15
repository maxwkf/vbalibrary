# vbalibrary
## Description
This file simplify the use of ADODB for MSSQL connection
## How to include this connection class?
1. Include the Reference **Microsoft ActiveX Data Objects 6.1 Library**
2. In the Microsoft Visual Basic for Applications Editor, right click on the Class Modules and use the import file.
## Usage
### Retrieve Record
```vba
Public Sub TestRetrieve()
    Call dbconnection.beginTransaction
    With dbconnection.cmd
        .parameters.Append .createParameter(, DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, , 10)
        .parameters.Append .createParameter(, DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, , 11)
        .CommandText = "select * from assessmenttask where AT_id = ? or AT_id = ?"
    End With

    Dim a As Collection
    Set a = dbconnection.runsql
    Call dbconnection.endTransaction
End Sub
```
### Insert/ Update Record
```vba
Public Sub TestSave()
    Call dbconnection.beginTransaction
    With dbconnection.cmd
        .parameters.Append .createParameter(, DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 100, "PPP")
        .CommandText = "update assessmenttask set task = ? where AT_id = 1"
    End With

    Call dbconnection.runsql
    Call dbconnection.endTransaction
End Sub
```
