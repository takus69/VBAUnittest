Attribute VB_Name = "VBAUnittest"
''
' VBAUnittest v1.0.1
' Copyright(c) 2016 takus - https://github.com/takus69/VBAUnittest
'
' @author takus4649@gmail.com
' @license MIT (https://opensource.org/licenses/MIT)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Dim tests As Object
Dim runCount As Integer
Dim failedCount As Integer
Dim assertFailedCount As Integer
Dim assertCount As Integer
Dim assertMessage As String
Dim runningTest As String
Dim excludedTests As Object
Dim excludedModules As Object


' Public procedure
' Setting excluding tests or modules
Sub setExcludedTests()
'    addExcludedTest "TestModule.testExcludedTest"
'    addExcludedTest "TestModule2.testExcludedTest"
End Sub

Sub setExcludedModules()
'    addExcludedModule "TestExcludedModule"
End Sub

' Test runner
Sub testRun(test As String)
    testInit
    setUp
    oneTestRun test
    tearDown
    
    showResult
End Sub

Sub testInit()
    runCount = 0
    failedCount = 0
    Set tests = CreateObject("Scripting.Dictionary")
    Set excludedTests = CreateObject("Scripting.Dictionary")
    Set excludedModules = CreateObject("Scripting.Dictionary")
    setExcludedTests
    setExcludedModules
End Sub

Sub addTest(test As String)
    tests.add tests.Count, test
End Sub

Sub suiteRun()
    Dim i As Integer
    
    setUp
    For i = 0 To tests.Count - 1
        oneTestRun tests.Item(i)
    Next i
    tearDown
    
    showResult
End Sub

Sub testModuleRun(TestModule As String)
    testInit
    addTestsInTestModule TestModule
    suiteRun
End Sub

Sub allTestRun()
    testInit
    addAllTest
    suiteRun
End Sub

' Assertion
Function assertTrue(status As Boolean) As Boolean
    assertMessage = setAssert(True, status)
    assertCount = assertCount + 1
    If Not status Then
        assertFailedCount = assertFailedCount + 1
    End If
    assertTrue = status
End Function

Function assertFalse(status As Boolean) As Boolean
    assertFalse = assertTrue(Not status)
    assertMessage = setAssert(False, status)
End Function

Function assert(a, b) As Boolean
    assert = assertTrue(a = b)
    assertMessage = setAssert(a, b)
End Function

' Messages
Function testSummary() As String
    testSummary = runCount & " run, " & failedCount & " failed"
End Function

Function failedMessage() As String
    failedMessage = runningTest & ", Assertion" & assertCount & ", " & assertMessage
End Function

' Private procedure
Private Sub oneTestRun(test As String)
    Dim runStatus As Boolean
    assertFailedCount = 0
    assertCount = 0
    runningTest = test
    
    runCount = runCount + 1
    Application.Run test
    
    If assertFailedCount > 0 Then
        failedCount = failedCount + 1
        showFailed
    End If
End Sub

Private Sub showResult()
    Dim result As String
    
    If failedCount = 0 Then
        result = "green"
    Else
        result = "red"
    End If
    Debug.Print result & " : " & testSummary
End Sub

Private Sub showFailed()
    Debug.Print failedMessage
End Sub

Private Function setAssert(expected, actual) As String
    setAssert = "Expected:" & expected & ", " & "Actual:" & actual
End Function

Private Function fetchProcs(TestModule As String) As String()
    Dim buf As String, testName As String, procNames() As String, i As Long, cnt As Integer
    cnt = -1
    With ThisWorkbook.VBProject.VBComponents(TestModule).CodeModule
        For i = 1 To .CountOfLines
            testName = TestModule & "." & .ProcOfLine(i, 0)
            If buf <> testName And .ProcOfLine(i, 0) Like "test*" Then
                buf = testName
                If Not isTestExcluded(testName) Then
                    cnt = cnt + 1
                    ReDim Preserve procNames(cnt)

                    procNames(cnt) = testName
                End If
            End If
        Next i
    End With
    
    fetchProcs = procNames
End Function

Private Function isTestExcluded(testName As String) As Boolean
    isTestExcluded = excludedTests.exists(testName)
End Function

Private Function isModuleExcluded(moduleName As String) As Boolean
    isModuleExcluded = excludedModules.exists(moduleName)
End Function

Private Sub addTestsInTestModule(TestModule As String)
    Dim procNames() As String
    Dim i As Integer
    
    procNames = fetchProcs(TestModule)
    For i = 0 To UBound(procNames)
        addTest procNames(i)
    Next i
End Sub

Private Sub addAllTest()
    Dim comp As Object, procNames() As String
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Name Like "Test*" And Not isModuleExcluded(comp.Name) Then
            procNames = fetchProcs(comp.Name)
            addTestsInTestModule comp.Name
        End If
    Next comp
End Sub

Private Sub addExcludedTest(excludedTest As String)
    excludedTests.add excludedTest, True
End Sub

Private Sub addExcludedModule(excludedModule As String)
    excludedModules.add excludedModule, True
End Sub
