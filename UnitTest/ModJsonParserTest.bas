Attribute VB_Name = "ModJsonParserTest"
'==============================================================================
' ModJsonParserTest - clsJsonParser テストモジュール
'
' 使い方:
'   1. clsJsonParser.cls をVBAプロジェクトにインポート
'   2. このモジュールをインポート
'   3. test_sample.json を C:\temp\ に配置
'   4. RunAllTests を実行 (イミディエイトウィンドウから)
'   5. 結果はイミディエイトウィンドウ (Ctrl+G) に出力されます
'
' ※参照設定は不要です (遅延バインディング)
'==============================================================================
Option Explicit

Private m_passCount As Long
Private m_failCount As Long

'==============================================================================
' メインエントリーポイント
'==============================================================================
Public Sub RunAllTests()
    m_passCount = 0
    m_failCount = 0
    
    Debug.Print "=============================================="
    Debug.Print " clsJsonParser テスト開始"
    Debug.Print "=============================================="
    Debug.Print ""
    
    ' --- loads テスト ---
    TestBasicObject
    TestNestedObject
    TestArrayValues
    TestNestedArray
    TestMixedArray
    TestNumberTypes
    TestBooleanAndNull
    TestEscapeSequences
    TestUnicodeEscape
    TestJapaneseString
    TestEmptyObjectAndArray
    TestRootArray
    TestCaseSensitiveKeys
    TestDeepNest
    TestScientificNotation
    
    ' --- dumps テスト ---
    TestDumpsBasic
    TestDumpsCompact
    TestDumpsNested
    TestDumpsArray
    TestDumpsJapanese
    TestDumpsEscape
    TestDumpsEmptyObjectAndArray
    
    ' --- ラウンドトリップ テスト ---
    TestRoundTrip
    
    ' --- エラー系テスト ---
    TestErrorEmptyString
    TestErrorInvalidJson
    TestErrorNothingDict
    
    ' --- ファイル読み込みテスト (手動実行) ---
    ' イミディエイトウィンドウから:
    '   TestLoadFromFile "C:\temp\test_sample.json"
    
    ' --- 結果サマリー ---
    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print " テスト結果: " & m_passCount & " PASS / " & m_failCount & " FAIL"
    Debug.Print "=============================================="
End Sub

'==============================================================================
' loads テスト
'==============================================================================

Private Sub TestBasicObject()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""name"":""test"",""value"":42}")
    
    AssertEqual "BasicObject - key count", 2, d.Count
    AssertEqual "BasicObject - name", "test", d("name")
    AssertEqual "BasicObject - value", 42, d("value")
    AssertEqual "BasicObject - value type", "Long", TypeName(d("value"))
End Sub

Private Sub TestNestedObject()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""outer"":{""inner"":{""deep"":""found""}}}")
    
    Dim outer As Object
    Set outer = d("outer")
    AssertEqual "NestedObject - outer type", "Dictionary", TypeName(outer)
    
    Dim inner As Object
    Set inner = outer("inner")
    AssertEqual "NestedObject - deep value", "found", inner("deep")
End Sub

Private Sub TestArrayValues()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""items"":[10,20,30]}")
    
    Dim arr As Variant
    arr = d("items")
    AssertTrue "ArrayValues - is array", IsArray(arr)
    AssertEqual "ArrayValues - length", 3, UBound(arr) - LBound(arr) + 1
    AssertEqual "ArrayValues - (0)", 10, arr(0)
    AssertEqual "ArrayValues - (1)", 20, arr(1)
    AssertEqual "ArrayValues - (2)", 30, arr(2)
End Sub

Private Sub TestNestedArray()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""matrix"":[[1,2],[3,4]]}")
    
    Dim matrix As Variant
    matrix = d("matrix")
    AssertTrue "NestedArray - is array", IsArray(matrix)
    
    Dim row0 As Variant
    row0 = matrix(0)
    AssertEqual "NestedArray - (0)(0)", 1, row0(0)
    AssertEqual "NestedArray - (0)(1)", 2, row0(1)
    
    Dim row1 As Variant
    row1 = matrix(1)
    AssertEqual "NestedArray - (1)(0)", 3, row1(0)
    AssertEqual "NestedArray - (1)(1)", 4, row1(1)
End Sub

Private Sub TestMixedArray()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""mix"":[1,""two"",true,null,{""a"":1},[10]]}")
    
    Dim arr As Variant
    arr = d("mix")
    AssertEqual "MixedArray - (0) number", 1, arr(0)
    AssertEqual "MixedArray - (1) string", "two", arr(1)
    AssertEqual "MixedArray - (2) bool", True, arr(2)
    AssertTrue "MixedArray - (3) null", IsNull(arr(3))
    AssertEqual "MixedArray - (4) obj type", "Dictionary", TypeName(arr(4))
    AssertTrue "MixedArray - (5) nested array", IsArray(arr(5))
End Sub

Private Sub TestNumberTypes()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""int"":42,""neg"":-7,""float"":3.14,""zero"":0}")
    
    AssertEqual "NumberTypes - int", 42, d("int")
    AssertEqual "NumberTypes - int type", "Long", TypeName(d("int"))
    AssertEqual "NumberTypes - neg", -7, d("neg")
    AssertEqual "NumberTypes - float", 3.14, d("float")
    AssertEqual "NumberTypes - float type", "Double", TypeName(d("float"))
    AssertEqual "NumberTypes - zero", 0, d("zero")
End Sub

Private Sub TestBooleanAndNull()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""t"":true,""f"":false,""n"":null}")
    
    AssertEqual "BoolNull - true", True, d("t")
    AssertEqual "BoolNull - false", False, d("f")
    AssertTrue "BoolNull - null", IsNull(d("n"))
End Sub

Private Sub TestEscapeSequences()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    ' VBAでは \ はそのままリテラル。JSONエスケープには \ 1つでOK
    Set d = jp.Loads("{""esc"":""line1\nline2\ttab\""quote\\ backslash""}")
    
    Dim expected As String
    expected = "line1" & vbLf & "line2" & vbTab & "tab""quote\ backslash"
    AssertEqual "EscapeSequences", expected, d("esc")
End Sub

Private Sub TestUnicodeEscape()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    ' \u3053\u3093\u306b\u3061\u306f = こんにちは
    Set d = jp.Loads("{""greet"":""\u3053\u3093\u306b\u3061\u306f""}")
    
    Dim expected As String
    expected = ChrW$(&H3053) & ChrW$(&H3093) & ChrW$(&H306B) & ChrW$(&H3061) & ChrW$(&H306F)
    AssertEqual "UnicodeEscape", expected, d("greet")
End Sub

Private Sub TestJapaneseString()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""msg"":""\u65E5\u672C\u8A9E\u30C6\u30B9\u30C8""}")
    
    AssertEqual "JapaneseString", ChrW$(&H65E5) & ChrW$(&H672C) & ChrW$(&H8A9E) & ChrW$(&H30C6) & ChrW$(&H30B9) & ChrW$(&H30C8), d("msg")
End Sub

Private Sub TestEmptyObjectAndArray()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""obj"":{},""arr"":[]}")
    
    Dim emptyDict As Object
    Set emptyDict = d("obj")
    AssertEqual "EmptyObject - count", 0, emptyDict.Count
    
    Dim emptyArr As Variant
    emptyArr = d("arr")
    AssertTrue "EmptyArray - is array", IsArray(emptyArr)
    
    Dim arrLen As Long
    On Error Resume Next
    arrLen = UBound(emptyArr) - LBound(emptyArr) + 1
    If Err.Number <> 0 Then arrLen = 0
    On Error GoTo 0
    AssertEqual "EmptyArray - length", 0, arrLen
End Sub

Private Sub TestRootArray()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    ' ルート配列 -> Dict("0", "1", ...)
    Set d = jp.Loads("[{""a"":1},{""b"":2},{""c"":3}]")
    
    AssertEqual "RootArray - key count", 3, d.Count
    AssertTrue "RootArray - has key 0", d.Exists("0")
    AssertTrue "RootArray - has key 1", d.Exists("1")
    AssertTrue "RootArray - has key 2", d.Exists("2")
    
    Dim item0 As Object
    Set item0 = d("0")
    AssertEqual "RootArray - item0.a", 1, item0("a")
End Sub

Private Sub TestCaseSensitiveKeys()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""Key"":""upper"",""key"":""lower""}")
    
    AssertEqual "CaseSensitive - count", 2, d.Count
    AssertEqual "CaseSensitive - Key", "upper", d("Key")
    AssertEqual "CaseSensitive - key", "lower", d("key")
End Sub

Private Sub TestDeepNest()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""a"":{""b"":{""c"":{""d"":""deep""}}}}")
    
    Dim val As String
    val = d("a")("b")("c")("d")
    AssertEqual "DeepNest - value", "deep", val
End Sub

Private Sub TestScientificNotation()
    Dim jp As New clsJsonParser
    Dim d As Object
    
    Set d = jp.Loads("{""big"":2.1e5,""small"":1.5e-3,""neg"":-3.0E+2}")
    
    AssertEqual "Scientific - big", 210000#, d("big")
    AssertEqual "Scientific - big type", "Double", TypeName(d("big"))
    AssertEqual "Scientific - small", 0.0015, d("small")
    AssertEqual "Scientific - neg", -300#, d("neg")
End Sub

'==============================================================================
' dumps テスト
'==============================================================================

Private Sub TestDumpsBasic()
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    d("name") = "test"
    d("value") = 42
    
    Dim result As String
    result = jp.Dumps(d)
    
    ' 再パースして検証
    Dim d2 As Object
    Set d2 = jp.Loads(result)
    AssertEqual "DumpsBasic - name", "test", d2("name")
    AssertEqual "DumpsBasic - value", 42, d2("value")
End Sub

Private Sub TestDumpsCompact()
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    d("a") = 1
    d("b") = 2
    
    Dim result As String
    result = jp.Dumps(d, 0)
    
    ' コンパクト出力には改行が含まれない
    AssertTrue "DumpsCompact - no newline", InStr(result, vbCrLf) = 0
    AssertTrue "DumpsCompact - has content", Len(result) > 0
End Sub

Private Sub TestDumpsNested()
    Dim jp As New clsJsonParser
    
    Dim inner As Object
    Set inner = CreateObject("Scripting.Dictionary")
    inner.CompareMode = vbBinaryCompare
    inner("deep") = "value"
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    Set d("nested") = inner
    
    Dim result As String
    result = jp.Dumps(d)
    
    ' 再パースして検証
    Dim d2 As Object
    Set d2 = jp.Loads(result)
    AssertEqual "DumpsNested", "value", d2("nested")("deep")
End Sub

Private Sub TestDumpsArray()
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    d("items") = Array(1, "two", True)
    
    Dim result As String
    result = jp.Dumps(d)
    
    ' 再パースして検証
    Dim d2 As Object
    Set d2 = jp.Loads(result)
    Dim arr As Variant
    arr = d2("items")
    AssertEqual "DumpsArray - (0)", 1, arr(0)
    AssertEqual "DumpsArray - (1)", "two", arr(1)
    AssertEqual "DumpsArray - (2)", True, arr(2)
End Sub

Private Sub TestDumpsJapanese()
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    d("msg") = ChrW$(&H65E5) & ChrW$(&H672C) & ChrW$(&H8A9E)  ' 日本語
    
    Dim result As String
    result = jp.Dumps(d)
    
    ' 日本語がそのまま出力されているか
    AssertTrue "DumpsJapanese - contains", InStr(result, ChrW$(&H65E5)) > 0
    
    ' 再パースして検証
    Dim d2 As Object
    Set d2 = jp.Loads(result)
    AssertEqual "DumpsJapanese - roundtrip", d("msg"), d2("msg")
End Sub

Private Sub TestDumpsEscape()
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    d("esc") = "tab" & vbTab & "newline" & vbLf & "quote"""
    
    Dim result As String
    result = jp.Dumps(d)
    
    ' エスケープが含まれているか
    AssertTrue "DumpsEscape - has \t", InStr(result, "\t") > 0
    AssertTrue "DumpsEscape - has \n", InStr(result, "\n") > 0
    AssertTrue "DumpsEscape - has \""", InStr(result, "\""") > 0
    
    ' 再パースして検証
    Dim d2 As Object
    Set d2 = jp.Loads(result)
    AssertEqual "DumpsEscape - roundtrip", d("esc"), d2("esc")
End Sub

Private Sub TestDumpsEmptyObjectAndArray()
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    Dim emptyDict As Object
    Set emptyDict = CreateObject("Scripting.Dictionary")
    emptyDict.CompareMode = vbBinaryCompare
    Set d("obj") = emptyDict
    d("arr") = Array()
    
    Dim result As String
    result = jp.Dumps(d)
    
    AssertTrue "DumpsEmpty - has {}", InStr(result, "{}") > 0
    AssertTrue "DumpsEmpty - has []", InStr(result, "[]") > 0
End Sub

'==============================================================================
' ラウンドトリップ テスト
'==============================================================================

Private Sub TestRoundTrip()
    Dim jp As New clsJsonParser
    
    ' 複雑なDictを構築
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    
    d("string") = ChrW$(&H30C6) & ChrW$(&H30B9) & ChrW$(&H30C8)  ' テスト
    d("integer") = 42
    d("float") = 3.14
    d("boolTrue") = True
    d("boolFalse") = False
    d("nothing") = Null
    d("array") = Array(1, "a", True)
    
    Dim child As Object
    Set child = CreateObject("Scripting.Dictionary")
    child.CompareMode = vbBinaryCompare
    child("key") = "value"
    Set d("nested") = child
    
    ' dumps -> loads -> dumps で一致するか
    Dim json1 As String
    json1 = jp.Dumps(d)
    
    Dim d2 As Object
    Set d2 = jp.Loads(json1)
    
    Dim json2 As String
    json2 = jp.Dumps(d2)
    
    AssertEqual "RoundTrip - json match", json1, json2
    
    ' 各値の検証
    AssertEqual "RoundTrip - string", d("string"), d2("string")
    AssertEqual "RoundTrip - integer", 42, d2("integer")
    AssertEqual "RoundTrip - float", 3.14, d2("float")
    AssertEqual "RoundTrip - boolTrue", True, d2("boolTrue")
    AssertEqual "RoundTrip - boolFalse", False, d2("boolFalse")
    AssertTrue "RoundTrip - null", IsNull(d2("nothing"))
    AssertEqual "RoundTrip - nested", "value", d2("nested")("key")
End Sub

'==============================================================================
' エラー系テスト
'==============================================================================

Private Sub TestErrorEmptyString()
    Dim jp As New clsJsonParser
    
    On Error Resume Next
    Dim d As Object
    Set d = jp.Loads("")
    
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    
    AssertEqual "ErrorEmpty - err number", 10001, errNum
End Sub

Private Sub TestErrorInvalidJson()
    Dim jp As New clsJsonParser
    
    On Error Resume Next
    Dim d As Object
    Set d = jp.Loads("{""key"":}")
    
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    
    AssertTrue "ErrorInvalid - raised error", errNum <> 0
End Sub

Private Sub TestErrorNothingDict()
    Dim jp As New clsJsonParser
    
    On Error Resume Next
    Dim result As String
    result = jp.Dumps(Nothing)
    
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    
    AssertEqual "ErrorNothing - err number", 10002, errNum
End Sub

'==============================================================================
' ファイル読み込みテスト
'==============================================================================

Public Sub TestLoadFromFile(ByVal filePath As String)
    
    Debug.Print ""
    Debug.Print "--- ファイル読み込みテスト: " & filePath & " ---"
    
    ' ファイル存在チェック
    If Dir(filePath) = "" Then
        Debug.Print "  [SKIP] ファイルが見つかりません: " & filePath
        Exit Sub
    End If
    
    ' UTF-8 読み込み (ADODB.Stream)
    Dim jsonStr As String
    jsonStr = ReadUtf8File(filePath)
    
    If Len(jsonStr) = 0 Then
        Debug.Print "  [SKIP] ファイルが空または読み込み失敗"
        Exit Sub
    End If
    
    Dim jp As New clsJsonParser
    Dim d As Object
    Set d = jp.Loads(jsonStr)
    
    ' 基本値
    ' ※ 日本語の正確性はエンコーディング環境に依存するため、
    '    プレフィックスと文字数で検証する
    AssertTrue "File - project starts with CAE", Left$(d("project"), 3) = "CAE"
    AssertTrue "File - project has content", Len(d("project")) > 3
    AssertEqual "File - version type", "Double", TypeName(d("version"))
    AssertEqual "File - enabled", True, d("enabled")
    AssertTrue "File - author null", IsNull(d("author"))
    
    ' タブ文字のエスケープ
    AssertTrue "File - description has tab", InStr(d("description"), vbTab) > 0
    
    ' ネストオブジェクト
    Dim settings As Object
    Set settings = d("settings")
    AssertEqual "File - meshSize", 3.5, settings("meshSize")
    AssertEqual "File - autoSave", False, settings("autoSave")
    AssertTrue "File - outputPath backslash", InStr(settings("outputPath"), "\") > 0
    
    ' 配列
    Dim materials As Variant
    materials = d("materials")
    AssertEqual "File - materials count", 2, UBound(materials) - LBound(materials) + 1
    AssertEqual "File - material0 name", "Steel", materials(0)("name")
    
    ' 指数表記
    AssertEqual "File - youngModulus type", "Double", TypeName(materials(0)("youngModulus"))
    
    ' 文字列配列
    Dim loadCases As Variant
    loadCases = d("loadCases")
    AssertEqual "File - LC count", 3, UBound(loadCases) - LBound(loadCases) + 1
    AssertEqual "File - LC1", "LC1", loadCases(0)
    
    ' ネスト配列 (matrix)
    Dim matrix As Variant
    matrix = d("matrix")
    Dim row0 As Variant
    row0 = matrix(0)
    AssertEqual "File - matrix(0)(0)", 1, row0(0)
    
    ' 空オブジェクト
    AssertEqual "File - emptyObj count", 0, d("emptyObject").Count
    
    ' 大文字小文字区別
    AssertEqual "File - CaseSensitive", "UPPER", d("CaseSensitive")
    AssertEqual "File - caseSensitive", "lower", d("caseSensitive")
    
    ' Unicode エスケープ
    Dim expectedUnicode As String
    expectedUnicode = ChrW$(&H3053) & ChrW$(&H3093) & ChrW$(&H306B) & ChrW$(&H3061) & ChrW$(&H306F)
    AssertEqual "File - unicode", expectedUnicode, d("unicode")
    
    ' 深いネスト
    Dim deepVal As String
    deepVal = d("nested")("level1")("level2")("level3")
    AssertTrue "File - deep nest has value", Len(deepVal) > 0
    
    ' 混合配列
    Dim mix As Variant
    mix = d("mixedArray")
    AssertEqual "File - mix(0)", 1, mix(0)
    AssertEqual "File - mix(1)", "two", mix(1)
    AssertEqual "File - mix(2)", True, mix(2)
    AssertTrue "File - mix(3) null", IsNull(mix(3))
    AssertEqual "File - mix(4) type", "Dictionary", TypeName(mix(4))
    AssertTrue "File - mix(5) array", IsArray(mix(5))
    
    ' ラウンドトリップ
    Dim json2 As String
    json2 = jp.Dumps(d)
    Dim d2 As Object
    Set d2 = jp.Loads(json2)
    AssertEqual "File - roundtrip project", d("project"), d2("project")
    
    Debug.Print "--- ファイル読み込みテスト完了 ---"
End Sub

'==============================================================================
' UTF-8 ファイル読み込みヘルパー
'==============================================================================
Private Function ReadUtf8File(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8File = stream.ReadText(-1)  ' adReadAll
    stream.Close
    Set stream = Nothing
    
    ' BOM除去
    If Len(ReadUtf8File) > 0 Then
        If AscW(Left$(ReadUtf8File, 1)) = &HFEFF Then
            ReadUtf8File = Mid$(ReadUtf8File, 2)
        End If
    End If
    Exit Function
    
ErrHandler:
    ReadUtf8File = ""
    If Not stream Is Nothing Then stream.Close
End Function

'==============================================================================
' アサーションヘルパー
'==============================================================================
Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        m_passCount = m_passCount + 1
        Debug.Print "  [PASS] " & testName
    Else
        m_failCount = m_failCount + 1
        Debug.Print "  [FAIL] " & testName & " | Expected: " & CStr(expected) & " | Actual: " & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    If condition Then
        m_passCount = m_passCount + 1
        Debug.Print "  [PASS] " & testName
    Else
        m_failCount = m_failCount + 1
        Debug.Print "  [FAIL] " & testName & " | Expected: True | Actual: False"
    End If
End Sub

