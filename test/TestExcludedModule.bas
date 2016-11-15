Attribute VB_Name = "TestExcludedModule"
' TestCase
Sub testAdd()
    assert add(1, 2), 3
End Sub

' Method
Function add(a, b)
    add = a + b
End Function

