Attribute VB_Name = "TestModule2"
' TestCase
Sub testAdd()
    assert add(1, 2), 3
End Sub

' Method
Function add(a As Integer, b As Integer) As Integer
    add = a + b
End Function

' Same Method Name
Sub testFailedMessageTrue()
    assertTrue False
End Sub

Sub testExcludedTest()
    assertTrue False
End Sub

