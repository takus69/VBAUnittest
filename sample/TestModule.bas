Attribute VB_Name = "TestModule"
' Test Case
Sub testAdd()
    assert 3, add(1, 2)
    assert 5, add(2, 3)
End Sub

Sub testSubtraction()
    assert 3, subtraction(4, 1)
    assert -1, subtraction(1, 2)
End Sub

