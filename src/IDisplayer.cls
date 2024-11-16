VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDisplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Displayer")
Option Explicit

' Sets a value to a given cell with it's coordinates
'
' Parameters :
'   sheetName : Name of the sheet containing the cell.
'   value : Value to give to the cell
'   row : Row of the cell
'   column : Column of the cell
Public Sub DisplayCell(ByVal sheetName As String, ByVal value As Variant, ByVal row As Long, ByVal column As Long)

End Sub

' Sets a  color index to a given Range
'
' Parameters :
'   sheetName : Name of the sheet containing the cell.
'   value : Value of ColorIndex to give to the cell to the cell
'   rng : Range given as a String. Ex : "A1:A10" or "A1" to represent single cell
Public Sub ChangeRangeColorIndex(ByVal sheetName As String, ByVal value As Integer, ByVal rng As String)
    
End Sub

' Shows a messsage to the user.
'
' Parameters
'   message : Message to show.
Public Sub ShowMsgBox(ByVal message As String)

End Sub
