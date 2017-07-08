Attribute VB_Name = "VBAUnittest"
''
' VBAUnittest v1.0.2
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
Const setUpMethodName As String = "setUp"
Const tearDownMethodName As String = "tearDown"


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
    oneTestRun test
    
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
    
    For i = 0 To tests.Count - 1
        oneTestRun tests.Item(i)
    Next i
    
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
    assertCount = assertCount + 1
    If Not status Then
        assertFailedCount = assertFailedCount + 1
        assertMessage = setAssert(True, status)
    End If
    assertTrue = status
End Function

Function assertFalse(status As Boolean) As Boolean
    Dim ret As Boolean
    ret = assertTrue(Not status)
    If Not ret Then
        assertMessage = setAssert(False, status)
    End If
    assertFalse = ret
End Function

Function assert(expected, actual) As Boolean
    Dim ret As Boolean
    ret = assertTrue(expected = actual)
    If Not ret Then
        assertMessage = setAssert(expected, actual)
    End If
    assert = ret
End Function

' Messages
Function testSummary() As String
    testSummary = runCount & " run, " & failedCount & " failed"
End Function

Function failedMessage() As String
    failedMessage = runningTest & ", Count of assertion is " & assertCount & ", " & assertMessage
End Function

Function isSetUp(TestModule As String) As Boolean
    Dim methodName As String, i As Long
    
    With ThisWorkbook.VBProject.VBComponents(TestModule).CodeModule
        For i = 1 To .CountOfLines
            methodName = .ProcOfLine(i, 0)
            If methodName = setUpMethodName Then
                isSetUp = True
                Exit Function
            End If
        Next i
    End With
    
    isSetUp = False
End Function

Function isTearDown(TestModule As String) As Boolean
    Dim methodName As String, i As Long
    
    With ThisWorkbook.VBProject.VBComponents(TestModule).CodeModule
        For i = 1 To .CountOfLines
            methodName = .ProcOfLine(i, 0)
            If methodName = tearDownMethodName Then
                isTearDown = True
                Exit Function
            End If
        Next i
    End With
    
    isTearDown = False
End Function

' Private procedure
Private Sub oneTestRun(test As String)
    Dim runStatus As Boolean, arr() As String, runningModule As String
    assertFailedCount = 0
    assertCount = 0
    runningTest = test
    runningModule = fetchModule(test)

    If isSetUp(runningModule) Then
        Application.Run runningModule & "." & setUpMethodName
    End If
    
    runCount = runCount + 1
    Application.Run test
    
    If isTearDown(runningModule) Then
        Application.Run runningModule & "." & tearDownMethodName
    End If
    
    If assertFailedCount > 0 Then
        failedCount = failedCount + 1
        showFailed
    End If
End Sub

Private Function fetchModule(testMethod)
    Dim runningModule As String, arr() As String
    
    runningModule = ""
    arr = Split(testMethod, ".")
    If UBound(arr) = 1 Then
        runningModule = arr(0)
    End If
    
    fetchModule = runningModule
End Function

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
    If (Not procNames) <> -1 Then ' Check no array data
        For i = 0 To UBound(procNames)
            addTest procNames(i)
        Next i
    End If
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
