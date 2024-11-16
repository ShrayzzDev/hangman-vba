VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model")

Option Explicit

Private m_displayer As IDisplayer
Private m_state As GameState

' Initialize the object
' Method should only be used by the factory
Public Sub Init(ByVal displayer As IDisplayer)
    Set m_displayer = displayer
    Set m_state = CreateGameState()
End Sub

' Starts a game
' Initialize the GameState and shows UI
Public Sub StartGame(ByVal word As String)
    m_state.Init word
    UpdateLife
    UpdateWord
End Sub

' Update the display of life in the UI.
Private Sub UpdateLife()
    m_displayer.ChangeRangeColorIndex "game", 0, "C3:C13"
    m_displayer.ChangeRangeColorIndex "game", 3, "C3:C" & CStr(2 + m_state.Life)
End Sub

' Update the guessed word in the UI.
Private Sub UpdateWord()
    m_displayer.DisplayCell "game", m_state.GuessedWord, 2, 3
End Sub

'
Public Sub MakeAGuess(ByVal guess As String)
    If Not m_state.IsStarted Then
        Exit Sub
    End If
    Dim isEnded As Boolean
    isEnded = m_state.MakeAGuess(guess)
    UpdateLife
    UpdateWord
    If isEnded Then
        If m_state.Life = 0 Then
            m_displayer.ShowMsgBox "You have lost."
        Else
            m_displayer.ShowMsgBox "You have won !"
        End If
    End If
End Sub
