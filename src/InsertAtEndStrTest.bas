Attribute VB_Name = "InsertAtEndStrTest"
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

'@TestMethod("Insert")
Private Sub TestWithNoElement()
    Dim arr(3) As String
    Dim position As Long
    
    position = InsertAtEnd(arr, "oui")
    
    Assert.AreEqual CLng(0), position
    Assert.AreEqual "oui", arr(0)
       
End Sub

'@TestMethod("Insert")
Private Sub TestWithOneElement()
    Dim arr(3) As String
    Dim position As Long
    arr(0) = "yes"
    
    position = InsertAtEnd(arr, "oui")
    
    Assert.AreEqual CLng(1), position
    Assert.AreEqual "oui", arr(1)
    
End Sub

'@TestMethod("Insert")
Private Sub TestWithTwoElement()
    Dim arr(3) As String
    Dim position As Long
    arr(0) = "yes"
    arr(1) = "other"
    
    position = InsertAtEnd(arr, "oui")
    
    Assert.AreEqual CLng(2), position
    Assert.AreEqual "oui", arr(2)
    
End Sub

'@TestMethod("Insert")
Private Sub TestWithFullList()
    Dim arr(3) As String
    arr(0) = "yes"
    arr(1) = "other"
    arr(2) = "Hatsune Miku"
    
    InsertAtEnd arr, "oui"
    
    Assert.AreEqual CLng(-1), InsertAtEnd(arr, "oui")
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
