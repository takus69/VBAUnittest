# VBAUnittest
VBAUnittest is testing framework for excel VBA.

## Installation
- Import VBAUnitest.bas to Visual Basic Editor.
- Set up below, if you use "allTestRun" or "testModuleRun".

![Setting](https://github.com/takus69/VBAUnittest/blob/master/setting.png)

## Sample
1. Make product and test codes.

``` MainModule.bas
Attribute VB_Name = "MainModule"
Function add(a, b)
    add = a + b
End Function

Function subtraction(a, b)
    subtraction = a - b
End Function
```

``` TestModule.bas
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
```
2. Run tests.
Execute "allTestRun" macro.
Result is below.

```
green : 2 run, 0 failed
```


## Usage
### Naming conventions
- Name of test module start with "Test". (e.g. TestModule)
- Name of test procedure start with "test". (e.g. testMethod)

### Assertion
The following assertions are defined.
- assertTrue
- assertFalse
- assert(a, b)

### allTestRun
- "allTestRun" executes all test procedure started with "test" in module started with "Test".
- Each test module or each test procedure can be excluded. (Example is below.)
- Add below in "setExcludedTests".
  - e.g. addExcludedTest "TestModule.testExcludedTest"
- Add below in "setExcludedModules".
  - e.g. addExcludedModule "TestExcludedModule"

### testModuleRun
- "testModuleRun" executes all test procedure started with "test" in module of argument.
- e.g. testModuleRun("TestModule")
- Each test procedure can be excluded. (Same with "allTestRun")

### suiteRun
- "suiteRun" executes tests that you need.
- e.g.

```
Sub suiteTestRun()
  testInit
  addTest "TestModule.testMethod"
  addTest "TestModule.testAssert"
  suiteRun
End Sub
```

### testRun
- "testRun" executes a test.
- e.g.

```
Sub eachTestRun()
  testRun "TestModule.testMethod"
End Sub
```

### Test Case
- Example of test case is below.

```
Sub testMethod()
  assert(1+2, 3)
  assertTrue(isTrue)
End Sub
```
### setUp and tearDown
- "setUp" executes before each test procedure.
- "tearDown" executes after each test procedure.


## Contributing
Bug reports and pull requests are welcome on GitHub at (https://github.com/takus69/VBAUnittest.git).

## License
This source codes is available as opne source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Release Notes
### v1.0.2
- Bug fix
  - Error occures if there are no "setUp" or "tearDown".
  - Error occures if there is no procedure.
  - The assertion error message shows last assertion even if it is success.
  - Error occures if there are two test modules with "setUp".

### v1.0.1
- Add data type of return value for function.
- Add license.

### v1.0.0
- First release
