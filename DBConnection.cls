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

Public connection As ADODB.connection
Private connectionString As String
Private level As Integer
Public cmd As ADODB.command

' Event - triggered when class created
Private Sub Class_Initialize()
    connectionString = "Driver={SQL Server};Server=wt5.cs.cityu.edu.hk;Database=accreditation; UID=milo; PWD=milo1234"
End Sub

'https://www.youtube.com/watch?v=5Zv-iNyzMzI
'https://www.youtube.com/watch?v=MshpY1VTZIY
Public Sub beginTransaction()

On Error GoTo Cleanup
        
    Set connection = New ADODB.connection
        
    connection.Open (connectionString)
    level = connection.BeginTrans
    
    Set cmd = New ADODB.command
    cmd.ActiveConnection = connection
Exit Sub
Cleanup:
    If (Not (connection Is Nothing) And connection.State = ObjectStateEnum.adStateOpen) Then
        connection.Close
    End If
End Sub

Public Sub endTransaction()
On Error GoTo Cleanup

    connection.CommitTrans
    connection.Close
Exit Sub
Cleanup:
    If (Not (connection Is Nothing) And connection.State = ObjectStateEnum.adStateOpen) Then
        connection.Close
    End If
End Sub

' Return collection of recordset as dictionary
Public Function runsql() As Collection
On Error GoTo Cleanup
    Dim rs As New ADODB.recordset
    Dim result As New Collection
    Dim row As Scripting.Dictionary
    Dim fields As fields
    Dim field As field
    
    Set rs = cmd.execute
    
    ' Construct result as collection
    If (rs.State > 0) Then
        Do While Not rs.EOF
            Set row = New Scripting.Dictionary
            Set fields = rs.fields
            
            For Each field In fields
                row.Add Key:=field.Name, Item:=Trim(field.Value)
            Next field
            result.Add row
            rs.MoveNext
        Loop
        
        Set runsql = result
    Else
        Set runsql = Nothing
    End If
    
    
Exit Function
Cleanup:
    If connection.State = adStateOpen Then
        If (level > 1) Then
            connection.RollbackTrans
            connection.Close
        End If
    End If
    Debug.Print Err.Description
End Function

Public Function getRecordsetCount(rs As ADODB.recordset) As Integer

   If rs.EOF Then
      getRecordsetCount = 0
   Else
      rs.MoveLast
      getRecordsetCount = rs.RecordCount
   End If
   rs.MoveFirst

End Function
