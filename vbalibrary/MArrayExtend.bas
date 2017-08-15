Attribute VB_Name = "MArrayExtend"
Option Explicit


Public Function size(data As Variant) As Integer
    'http://www.excel-easy.com/vba/examples/size-of-an-array.html
    size = UBound(data, 1) - LBound(data, 1) + 1
End Function
