Attribute VB_Name = "StringUtils"
'@Folder("Utils")
Option Explicit

Public Function Replace1Char(ByVal str As String, ByVal replacement As String, ByVal charPos As Long) As String
    Replace1Char = Mid(str, 1, charPos - 1) & replacement & Mid(str, charPos + 1, Len(str))
End Function

Public Function InsertAtEndStr(ByRef str As String, ByVal val As String) As String
    InsertAtEndStr = str & val
End Function
