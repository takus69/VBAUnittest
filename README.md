# VBAUnittest
VBAUnittest is testing framework for excel VBA.

## Installation
- ExcelのVisual Basic Editorにimportしてください。
- Excelの拡張子は、xlsmにする必要があります。
- 全てのテストを実行するallTestRun or テストモジュールの全てのテストを実行するtestModuleRunを使用する場合は以下の設定が必要です。
- 「開発」タブの「マクロのセキュリティ」から「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを付けて下さい。

![Setting](https://github.com/takus69/VBAUnittest/blob/master/setting.png)

## Usage
### Naming conventions
- テストモジュール名は、Testから開始して下さい。(e.g. TestModule)
- テストプロシージャ名は、testから開始して下さい。(e.g. testMethod)

### Assertion
以下のAssertionが定義してあります。
- assertTrue
- assertFalse
- assert(a, b)

### allTestRun
- allTestRunを実行すると、Testから開始するモジュール内にある、testから開始するプロシージャを全て実行します。
- テスト対象から除外したい場合は、モジュール単位に、プロシージャ単位で除外が可能です。(以下に設定例を示します。)
- setExcludedTestsに以下を追加します。(プロシージャの例)
  - addExcludedTest "TestModule.testExcludedTest"
- setExcludedModulesに以下を追加します。(モジュールの例)
  - addExcludedModule "TestExcludedModule"

### testModuleRun
- testModuleRunで設定したモジュール内にある、testから開始するプロシージャを全て実行します。
- e.g. testModuleRun("TestModule")
- テスト対象から除外したい場合は、上記allTestRunと同様です。

### suiteRun
- 実行したいテストプロシージャを設定して、まとめてテストできます。
- e.g.

```
Sub testSuite()
  testInit
  addTest "TestModule.testMethod"
  addTest "TestModule.testAssert"
  suiteRun
End Sub
```

### testRun
- テストを一つずつ実行できます。
- e.g.
  - testRun "TestModule.testMethod"

### Test Case
- テストケースは、以下のように設定します。
- e.g.

```
Sub testMethod()
  assert(1+2, 3)
  assertTrue(isTrue)
End Sub
```

## Contributing
Bug reports and pull requests are welcome on GitHub at (https://github.com/takus69/VBAUnittest.git).

## License
This source codes is available as opne source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
