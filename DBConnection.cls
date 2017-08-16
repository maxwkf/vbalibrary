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


Public Function runsql() As Collection
On Error GoTo Cleanup
    Dim rs As New ADODB.recordset
    Dim result As New Collection
    
    Set rs = cmd.execute
       
    ' Construct result as collection
    If (rs.State > 0) Then
        Do While Not rs.EOF
            result.Add rs
            rs.MoveNext
        Loop
    End If
    
    Set runsql = result
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
