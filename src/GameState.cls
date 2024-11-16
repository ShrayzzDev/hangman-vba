VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model")

Option Explicit

Private m_wordToFind As String
Private m_life As Integer
Private m_guessedWord As String
Private m_guessedLetters As String
Private m_isStarted As Boolean

' Initialize the object
' Should only be called in the GameLogic's own init.
Public Sub Init(ByVal word As String)
    m_wordToFind = word
    Dim n As Long
    
    For n = 1 To Len(word)
        m_guessedWord = Replace1Char(m_guessedWord, "_", n)
    Next n
    
    m_life = 10
    m_isStarted = True
End Sub

' Contains the logic of when a user makes a guess
Public Function MakeAGuess(ByVal guess As String) As Boolean
    If Len(guess) > 1 Or Not InStr(m_guessedLetters, guess) = 0 Then
        Exit Function
    End If
    m_guessedLetters = InsertAtEndStr(m_guessedLetters, guess)
    Dim n As Long
    Dim found As Boolean
    found = False
    For n = 1 To Len(m_wordToFind)
        If Mid(m_wordToFind, n, 1) = guess Then
            m_guessedWord = Replace1Char(m_guessedWord, guess, n)
            found = True
        End If
    Next n
    If Not found Then
        m_life = m_life - 1
    End If
    If m_life = 0 Or m_wordToFind = m_guessedWord Then
        Reset
        MakeAGuess = True
    End If
End Function

Private Sub Reset()
    m_wordToFind = vbNullString
    m_life = 10
    m_guessedWord = vbNullString
    m_guessedLetters = vbNullString
    m_isStarted = False
End Sub

Public Property Get Life() As Integer
    Life = m_life
End Property

Public Property Get GuessedWord() As String
    GuessedWord = m_guessedWord
End Property

Public Property Get IsStarted() As Boolean
    IsStarted = m_isStarted
End Property
