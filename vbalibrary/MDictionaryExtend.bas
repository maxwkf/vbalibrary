Attribute VB_Name = "MDictionaryExtend"
Option Explicit

Public Function rangeToDictionary(rng As Range) As scripting.Dictionary
    Dim newDict As New scripting.Dictionary
    Dim cell As Variant
    Dim arbitraryKey As String
    Dim keystore As New scripting.Dictionary
    
    For Each cell In rng
        arbitraryKey = CStr(cell.value)
        If newDict.Exists(arbitraryKey) Then
            keystore.Item(cell.value) = keystore.Item(cell.value) + 1
            newDict.Add Key:=CStr(cell.value) & "_" & keystore.Item(cell.value), Item:=cell.value
        Else
            newDict.Add Key:=CStr(cell.value), Item:=cell.value
            
            keystore.Add Key:=cell.value, Item:=1
            
        End If
    Next cell
    
    Set rangeToDictionary = newDict
End Function
