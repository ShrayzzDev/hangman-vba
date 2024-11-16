Attribute VB_Name = "IsInArrayTest"
'@TestModule
'@Folder("Tests.ArrayUtils")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestMethod("Contains")
Private Sub TestContainsWith1Element()
    Dim arr(2) As String
    Dim result As Long
    arr(0) = "yes"
    
    Assert.IsTrue IsInArray(arr, "yes")
    
End Sub

'@TestMethod("Contains")
Private Sub TestContainsWithMultipleElements()
    Dim arr(3) As String
    Dim result As Long
    arr(0) = "no"
    arr(1) = "yes"
    arr(2) = "other"
    
    Assert.IsTrue IsInArray(arr, "yes")
    
End Sub

'@TestMethod("Contains")
Private Sub TestNotContainsWith1Element()
    Dim arr(2) As String
    Dim result As Long
    arr(0) = "yes"
    
    Assert.IsFalse IsInArray(arr, "no")
    
End Sub

'@TestMethod("Contains")
Private Sub TestNotContainsWithMultipleElements()
    Dim arr(3) As String
    Dim result As Long
    arr(0) = "no"
    arr(1) = "yes"
    arr(2) = "other"
    
    Assert.IsFalse IsInArray(arr, "still no")
    
End Sub

'@TestMethod("Contains")
Private Sub TestNotContainsWith0Element()
    Dim arr(2) As String
    Dim result As Long
    
    Assert.IsFalse IsInArray(arr, "no")
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

