Attribute VB_Name = "MNeatSplit"
Option Explicit
Option Base 1

' +-------------------------------------------------------------------
'
' +-------------------------------------------------------------------
Public Function NeatSplit(ByVal Expression As String, _
Optional ByVal Delimiter As String = " ", _
Optional ByVal Limit As Long = -1, _
Optional Compare As VbCompareMethod = vbBinaryCompare) _
As Variant

Dim varItems As Variant, i As Long

varItems = Split(Expression, Delimiter, Limit, Compare)

For i = LBound(varItems) To UBound(varItems)

    If Len(varItems(i)) = 0 Then varItems(i) = Delimiter

Next i

NeatSplit = Filter(varItems, Delimiter, False)
    
End Function


