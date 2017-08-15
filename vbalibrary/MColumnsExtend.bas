Attribute VB_Name = "MColumnsExtend"
Option Explicit

Public Function getColumnAddressByNumber(col As Long) As String
    Dim sColumn As String
    On Error Resume Next
    sColumn = Split(Columns(col).Address(, False), ":")(1)
    On Error GoTo 0
    getColumnAddressByNumber = sColumn
End Function
