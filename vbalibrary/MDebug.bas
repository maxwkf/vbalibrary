Attribute VB_Name = "MDebug"
Option Explicit

Public Sub printDictionary(dict As scripting.Dictionary)
    Dim K As Variant
    Debug.Print ""
    Debug.Print "Dictionary:"
    Debug.Print "------------------------------"
    For Each K In dict.Keys
        Debug.Print "[""" & K & """] => """ & dict.Item(K) & """"
    Next
    Debug.Print ""
End Sub

Public Sub printCollection(coll As Collection)
    Dim cell As Variant
    For Each cell In coll
        'Debug.Print "Key: " & cell(1), "Value: " & cell(0)
        Debug.Print "Value: " & cell
    Next
End Sub


