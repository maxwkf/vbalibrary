Attribute VB_Name = "MRangeExtend"
Option Explicit

' @param Variant target - This can be Cells(1,1) or Range("B4") or Range("B4:C9") format
'   cannot be Range("B") or Range(1)
' @return Collection - collection.Item("Column") => "B" / "B:C"
'   collection.Item("Row") => 4 / "4:9"
Public Function getColumnRowAddress(target As Variant) As scripting.Dictionary
    Dim result As New scripting.Dictionary

    Dim r As Range
    Set r = target

    Dim normalizedAddress As String
    normalizedAddress = target.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Dim WrdArray() As String
    WrdArray() = Split(normalizedAddress, ":")
    
    Dim col1 As String
    Dim col2 As String
    Dim row1 As Integer
    Dim row2 As Integer
    
    Dim columnRowSet1 As scripting.Dictionary
    Dim columnRowSet2 As scripting.Dictionary
    
    If MCommon.ArraySize(WrdArray) > 0 Then
        Set columnRowSet1 = splitColumnAndRowAddress(WrdArray(0))
    End If

    If MCommon.ArraySize(WrdArray) > 1 Then
        Set columnRowSet2 = splitColumnAndRowAddress(WrdArray(1))
    End If

    If Not columnRowSet1 Is Nothing Then
        If Not columnRowSet2 Is Nothing Then
            result.Add columnRowSet1.Item("Column") & ":" & columnRowSet2.Item("Column"), "Column"
            result.Add columnRowSet1.Item("Row") & ":" & columnRowSet2.Item("Row"), "Row"
        Else
            Set result = columnRowSet1
        End If
    End If
    
    Set getColumnRowAddress = result
    
End Function

' Only accept Range("AB12") style but not Range("AB12:CC13")
'   only if for single column or row use Range("B:B")
Public Function splitColumnAndRowAddress(rangeStr As String) As scripting.Dictionary
    Dim lnRow As String
    Dim strCol As String
    
    If InStr(1, rangeStr, ":") > 0 Then
        Dim WrdArray() As String
        WrdArray() = Split(rangeStr, ":")
        
        strCol = WrdArray(0)
    Else
        lnRow = Range(rangeStr).row
        strCol = Left(rangeStr, Len(rangeStr) - Len(CStr(lnRow)))
    End If
    
    Dim dict As scripting.Dictionary
    Set dict = New scripting.Dictionary
    
    dict.Add Key:="Column", Item:=strCol
    dict.Add Key:="Row", Item:=lnRow
    
    Set splitColumnAndRowAddress = dict
End Function

Public Sub simpleRangeSort(rangeStr As String, Optional Sht As String = vbNullString)
    Dim currentSheet As Worksheet
    If Sht = vbNullString Then
        Set currentSheet = Application.ActiveSheet
    Else
        Set currentSheet = Worksheets(Sht)
    End If
    currentSheet.Select
    Range(rangeStr).Sort key1:=Range(rangeStr), order1:=xlAscending, Header:=xlNo
End Sub
