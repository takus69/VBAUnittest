Attribute VB_Name = "TestModule"
Dim log As String

' TestCase
Sub setUp()
    log = "setUp "
End Sub

Sub tearDown()
    log = log & "tearDown "
End Sub

Sub assertTest()
    assertTrue assertFalse(False)
    assertTrue assert("A", "A")
End Sub

Sub testFailedResultFormatting()
    assertTrue False
    assertTrue False
End Sub

Sub testFailedMessageTrue()
    assertTrue False
End Sub

Sub testFailedMessageFalse()
    assertFalse False
    assertFalse True
End Sub

Sub testFailedMessageEqual()
    assertTrue True
    assertTrue True
    assert 1, 2
End Sub

Sub testTweAssertionFailed()
    assertTrue False
    assert 1, 2
End Sub

Sub testExcludedTest()
    assertTrue False
End Sub

' TestRunner
Sub eachTestRun()
    testStatus = True
    
    testStatus = resultCheck(assertTrue(True), True)
    
    testRun "TestModule.assertTest"
    testStatus = resultCheck(testSummary, "1 run, 0 failed")
    
    testRun "TestModule.testMethod"
    testStatus = resultCheck(log, "setUp testMethod tearDown ")
    testStatus = resultCheck(testSummary, "1 run, 0 failed")
    
    testRun "TestModule.testFailedResultFormatting"
    testStatus = resultCheck(testSummary, "1 run, 1 failed")
    
    testRun "TestModule.testFailedMessageTrue"
    testStatus = resultCheck(failedMessage, "TestModule.testFailedMessageTrue, Assertion1, Expected:True, Actual:False")

    testRun "TestModule.testFailedMessageFalse"
    testStatus = resultCheck(failedMessage, "TestModule.testFailedMessageFalse, Assertion2, Expected:False, Actual:True")

    testRun "TestModule.testFailedMessageEqual"
    testStatus = resultCheck(failedMessage, "TestModule.testFailedMessageEqual, Assertion3, Expected:1, Actual:2")
    
    testRun "TestModule.testTweAssertionFailed"
    testStatus = resultCheck(failedMessage, "TestModule.testTweAssertionFailed, Assertion2, Expected:1, Actual:2")
    
    If testStatus Then
        Debug.Print "All test green!"
    Else
        Debug.Print "Some tests have red."
    End If
End Sub

Function resultCheck(a, b) As Boolean
    If a = b Then
        resultCheck = True
    Else
        resultCheck = False
    End If
End Function

Sub SuiteTest()
    testStatus = True
    
    testInit
    addTest "TestModule.assertTest"
    addTest "TestModule.testFailedResultFormatting"
    addTest "TestModule.assertTest"
    addTest "TestModule.testFailedMessageTrue"
    addTest "TestModule.testFailedMessageFalse"
    addTest "TestModule.testFailedMessageEqual"
    addTest "TestModule.testTweAssertionFailed"
    suiteRun
    testStatus = resultCheck(testSummary, "7 run, 5 failed")
    
    If testStatus Then
        Debug.Print "SuiteTest is green!"
    Else
        Debug.Print "Some tests have red."
    End If
End Sub

Sub ModuleRunTest()
    testStatus = True
    
    testModuleRun "TestModule"
    testStatus = resultCheck(testSummary, "6 run, 5 failed")
    
    If testStatus Then
        Debug.Print "ModuleRunTest is green!"
    Else
        Debug.Print "Some tests have red."
    End If
End Sub

Sub AllTestRunTest()
    testStatus = True
    
    allTestRun
    testStatus = resultCheck(testSummary, "8 run, 6 failed")
    
    If testStatus Then
        Debug.Print "AllTestRunTest is green!"
    Else
        Debug.Print "Some tests have red."
    End If
End Sub

' Method for test
Sub testMethod()
    log = log & "testMethod "
End Sub

