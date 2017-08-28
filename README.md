# DBConnection.cls
## Description
This file simplify the use of ADODB for MSSQL connection
## How to include this connection class?
1. Include the Reference **Microsoft ActiveX Data Objects 6.1 Library**
2. In the Microsoft Visual Basic for Applications Editor, right click on the Class Modules and use the import file.
## Usage
### Retrieve and Loop Records
```vba
Public Sub TestRetrieve()
    Dim dbresult As Collection
    Dim row As Scripting.Dictionary
    
    
    With dbconnection.cmd
        .parameters.Append .createParameter(, DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput, , Yr)
        .parameters.Append .createParameter(, DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 100, course)
        .commandText = "Select * from CohortCourse where Yr = ? and Course = ?"
    End With

    Set dbresult = dbconnection.runsql
    
    If (dbresult.Count > 1) Then
        Err.Raise 9999, "getCohortCourseByYrCourse", "More than 1 result in which is abnormal. Please contact system administrator."
    ElseIf (dbresult.Count = 1) Then
        For Each row In dbresult
            Debug.Print row.Item("CC_id")
            Debug.Print row.Item("Course")
        Next row
    Else
    End If
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
