Attribute VB_Name = "Factory"
'@Folder("Model")
Option Explicit

' Factories are pretty much necessary since
' VBA does not support constructors.

' Creates an instance of GameLogic
Public Function CreateGameLogic(ByVal displayer As IDisplayer) As GameLogic
    Dim value As GameLogic
    Set value = New GameLogic
    value.Init displayer
    Set CreateGameLogic = value
End Function

' Creates an instance of GameState
' Here we do not init because it is initialized when
' the game is started in GameLogic. Initializing here
' is useless because it would be overriden anyway
Public Function CreateGameState() As GameState
    Dim value As GameState
    Set value = New GameState
    Set CreateGameState = value
End Function
