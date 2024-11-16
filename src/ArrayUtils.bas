Attribute VB_Name = "ArrayUtils"
'@Folder("Utils")

Option Explicit

Public Function IsInArray(ByVal arr As Variant, ByVal val As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = val Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Public Function InsertAtEnd(ByRef arr As Variant, ByVal val As Variant) As Long
    Dim position As Long
    For position = LBound(arr) To UBound(arr)
        If arr(position) = vbNullString Then
            arr(position) = val
            InsertAtEnd = position
            Exit Function
        End If
    Next position
    InsertAtEnd = -1
End Function

