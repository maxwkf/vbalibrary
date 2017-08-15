VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' use the import function to import this vbs
Public connection As ADODB.connection
Public Recordset As ADODB.Recordset
Private connectionString As String

' Event - triggered when class created
Private Sub Class_Initialize()
	' no need to replace {SQL Server}
    connectionString = "Driver={SQL Server};Server={{SERVER}};Database=accreditation; UID={{UID}}; 
End Sub

'https://www.youtube.com/watch?v=5Zv-iNyzMzI
'https://www.youtube.com/watch?v=MshpY1VTZIY
Public Sub beginTransaction()

On Error GoTo Cleanup
        
    Set connection = New ADODB.connection
        
    connection.Open (connectionString)
    connection.BeginTrans
        
Cleanup:
    If (Err.Number <> 0) Then
        Debug.Print Err.Description
        If (Not (connection Is Nothing) And connection.State = ObjectStateEnum.adStateOpen) Then
            connection.Close
        End If
    End If
    
   
End Sub

Public Function save(sql As String) As ADODB.Recordset
On Error GoTo Cleanup
    Set save = connection.execute(sql)
Cleanup:
    If (Err.Number <> 0) Then
        Debug.Print Err.Description
        If (Not (connection Is Nothing) And connection.State = ObjectStateEnum.adStateOpen) Then
            connection.Close
        End If
    End If
End Function

Public Sub endTransaction()
On Error GoTo Cleanup

    connection.CommitTrans
    connection.Close

Cleanup:
    If (Err.Number <> 0) Then
        Debug.Print Err.Description
        If (Not (connection Is Nothing) And connection.State = ObjectStateEnum.adStateOpen) Then
            connection.Close
        End If
    End If

End Sub

Public Function retrieveRecord(sql As String, cParameters As Collection) As Collection

    Dim rs As New ADODB.Recordset
    Dim result As New Collection
    Dim oParamDictionary As scripting.Dictionary
    Dim cmd As New ADODB.Command
    Dim param As ADODB.parameter
    
    ' Setup connection
    Set connection = New ADODB.connection
    connection.Open (connectionString)
    cmd.ActiveConnection = connection
    
    ' Build SQL and plug parameters
    cmd.CommandText = sql
    For Each oParamDictionary In cParameters
        Set param = cmd.createParameter(, oParamDictionary.Item("type"), adParamInput)
        param.value = oParamDictionary.Item("value")
        
        cmd.parameters.Append param
    Next oParamDictionary
    
    ' Retrieve result
    Set rs = cmd.execute
       
    ' Construct result as collection
    Do While Not rs.EOF
        result.Add rs
        rs.MoveNext
    Loop
    rs.Close
    
    Set retrieveRecord = result
   
End Function

Public Function createParameter(iType As Integer, vValue As Variant) As scripting.Dictionary

    Dim parameter As New scripting.Dictionary
    With parameter
        .Add Key:="type", Item:=iType
        .Add Key:="value", Item:=vValue
    End With
    
    Set createParameter = parameter
End Function









