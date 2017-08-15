Attribute VB_Name = "MCommon"
Option Explicit

'Public Function selectToRowEnd(sheet As Worksheet, sRangeStart As String) As Range
'
'    Dim col As String
'    col = Left(sRangeStart, 1)
'
'    Let srange = sRangeStart & ":" & col & GetSaTableStudentCount(sheet)
'    'Debug.Print ("This is the range -> " & srange);
'    Set selectToRowEnd = sheet.Range(srange)
'
'End Function

' Put "coll" Collection
' to Range start at "rangeStart"
' in sheet with name "sht"
Public Sub writeCollectionToRangeVertically(Sht As String, rangeStart As String, coll As Collection)
    Dim rng As Range
    
    Set rng = Sheets(Sht).Range(rangeStart).Resize(coll.Count, 1)


    Dim i As Integer
    For i = 1 To coll.Count
        rng.Cells(i) = coll(i)
    Next i
End Sub

Public Sub writeDictionaryToRangeVertically(Sht As String, rangeStart As String, dict As scripting.Dictionary)
    Dim rng As Range
    
    Set rng = Sheets(Sht).Range(rangeStart).Resize(dict.Count, 1)


    Dim i As Long
    For i = 0 To dict.Count - 1
       rng.Cells(i + 1) = dict.Items(i)
    Next i

End Sub


Public Sub writeStringArrayToRangeVertically(Sht As String, rangeStart As String, strs As Variant)
    Dim rng As Range
    
    Set rng = Sheets(Sht).Range(rangeStart).Resize(MArrayExtend.size(strs), 1)
    
    
    Dim i As Long
    For i = 0 To MArrayExtend.size(strs) - 1
       rng.Cells(i + 1) = strs(i)
    Next i
End Sub
' Put "coll" Collection
' to Range start at "rangeStart"
' in sheet with name "sht"
Public Sub writeCollectionToRangeHorizontally(Sht As String, rangeStart As String, coll As Collection)
    Dim rng As Range
    
    Set rng = Sheets(Sht).Range(rangeStart).Resize(1, coll.Count)


    Dim i As Integer
    For i = 1 To coll.Count
        rng.Cells(i) = coll(i)
    Next i
End Sub

Public Sub writeDictionaryToRangeHorizontally(Sht As String, rangeStart As String, dict As scripting.Dictionary)
    Dim rng As Range
    
    Set rng = Sheets(Sht).Range(rangeStart).Resize(1, dict.Count)


    Dim i As Long
    For i = 0 To dict.Count - 1
       rng.Cells(i + 1) = dict.Items(i)
    Next i

End Sub


Public Sub writeStringArrayToRangeHorizontally(Sht As String, rangeStart As String, strs As Variant)
    Dim rng As Range
    
    Set rng = Sheets(Sht).Range(rangeStart).Resize(1, MArrayExtend.size(strs))
    
    
    Dim i As Long
    For i = 0 To MArrayExtend.size(strs) - 1
       rng.Cells(i + 1) = strs(i)
    Next i
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This routine uses the "heap sort" algorithm to sort a VB collection.
' It returns the sorted collection.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function sortCollection(ByVal c As Collection) As Collection
   Dim N As Long: N = c.Count
   If N = 0 Then Set sortCollection = New Collection: Exit Function
   ReDim Index(0 To N - 1) As Long                    ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To N - 1: Index(i) = i + 1: Next         ' fill index array
   For i = N \ 2 - 1 To 0 Step -1                     ' generate ordered heap
      Heapify c, Index, i, N
      Next
   For m = N To 2 Step -1                             ' sort the index array
      Exchange Index, 0, m - 1                        ' move highest element to top
      Heapify c, Index, 0, m - 1
      Next
   Dim c2 As New Collection
   For i = 0 To N - 1: c2.Add c.Item(Index(i)): Next  ' fill output collection
   Set sortCollection = c2
   End Function

Private Sub Heapify(ByVal c As Collection, Index() As Long, ByVal i1 As Long, ByVal N As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   Dim nDiv2 As Long: nDiv2 = N \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim K As Long: K = 2 * i + 1
      If K + 1 < N Then
         If c.Item(Index(K)) < c.Item(Index(K + 1)) Then K = K + 1
         End If
      If c.Item(Index(i)) >= c.Item(Index(K)) Then Exit Do
      Exchange Index, i, K
      i = K
      Loop
   End Sub

Private Sub Exchange(Index() As Long, ByVal i As Long, ByVal j As Long)
   Dim Temp As Long: Temp = Index(i)
   Index(i) = Index(j)
   Index(j) = Temp
   End Sub

Public Function selectToLastRowFromCell(ws As Worksheet, sRangeStart As String) As Range

    Dim col As String
    col = MRangeExtend.getColumnRowAddress(Range(sRangeStart)).Item("Column")

    'Find the last used row in a Column: column A in this example
    Dim LastRow As Long
    With ws
        LastRow = .Cells(.Rows.Count, col).End(xlUp).row
    End With
        
    'Select until last row
    Dim srange As String
    Let srange = sRangeStart & ":" & col & LastRow
    Set selectToLastRowFromCell = ws.Range(srange)
    
End Function

Public Function selectToLastColumnFromCell(ws As Worksheet, sRangeStart As String) As Range

    Dim row As String
    row = MRangeExtend.getColumnRowAddress(Range(sRangeStart)).Item("Row")

    'Find the last used col in a Row: row 1 in this example
    Dim LastCol As Long
    With ws
        LastCol = .Cells(row, .Columns.Count).End(xlToLeft).Column
    End With
    
    Dim sLastCol As String
    sLastCol = MColumnsExtend.getColumnAddressByNumber(LastCol)
        
    'Select until last row
    Dim srange As String
    Let srange = sRangeStart & ":" & sLastCol & row
    Set selectToLastColumnFromCell = ws.Range(srange)
    
End Function


Private Sub LastRowInOneColumn()
'Find the last used row in a Column: column A in this example
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    MsgBox LastRow
End Sub

Private Sub LastColumnInOneRow()
'Find the last used column in a Row: row 1 in this example
    Dim LastCol As Integer
    With ActiveSheet
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    MsgBox LastCol
End Sub

Public Function ArraySize(Arr As Variant) As Integer
    ArraySize = Application.CountA(Arr)
End Function

' @param String datasheetName - "SA - 201415 - CS2115"
Public Function getDatasheetCourseCode(datasheetName As String) As String
    Dim nameArray() As String
    nameArray() = Split(datasheetName, " - ")
    
    getDatasheetCourseCode = nameArray(2)
End Function

Public Function IsWorksheetExists(ByVal WorksheetName As String) As Boolean
    Dim Sht As Worksheet

    For Each Sht In ThisWorkbook.Worksheets
        If Application.Proper(Sht.Name) = Application.Proper(WorksheetName) Then
            IsWorksheetExists = True
            Exit Function
        End If
    Next Sht
    IsWorksheetExists = False
End Function
