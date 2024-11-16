VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelDisplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Displayer")
Option Explicit

Implements IDisplayer

Private Sub IDisplayer_DisplayCell(ByVal sheetName As String, ByVal value As Variant, ByVal row As Long, ByVal column As Long)
    ActiveWorkbook.Worksheets(sheetName).Cells(row, column).value = value
End Sub

Private Sub IDisplayer_ChangeRangeColorIndex(ByVal sheetName As String, ByVal value As Integer, ByVal rng As String)
    ActiveWorkbook.Worksheets(sheetName).Range(rng).Interior.ColorIndex = value
End Sub

Private Sub IDisplayer_ShowMsgBox(ByVal message As String)
    MsgBox message
End Sub
