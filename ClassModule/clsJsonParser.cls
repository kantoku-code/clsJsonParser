VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsJsonParser - JSON文字列 <-> Dictionary 相互変換クラス
'
' 概要:
' JSON文字列とScripting.Dictionaryの相互変換を行う軽量ユーティリティ。
' CATIA VBA / Excel VBA 両対応。ステートレス設計。
' 参照設定不要 (遅延バインディング)。
'
' データマッピング:
' JSON Object   -> Scripting.Dictionary (vbBinaryCompare: 大文字小文字区別)
' JSON Array    -> Variant配列 (混合型・ネスト対応)
' JSON String   -> String
' JSON Number   -> Long (整数) / Double (小数・指数)
' JSON Boolean  -> Boolean
' JSON null     -> Null
' ルート配列    -> Dictionary ("0", "1", ... をキーに変換)
'
' パーサー:
' VBScript.RegExp によるトークン分割 + 構造解析のハイブリッド方式
'
' 公開メソッド:
' Loads(sJsonStr As String) As Object  ' Dictionary
' Dumps(dictObj As Object, [iIndent As Long = 4]) As String
'
' 使用例:
' Dim oJP As New clsJsonParser
' Dim dictData As Object
'
' ' パース
' Set dictData = oJP.Loads("{""name"":""テスト"",""values"":[1,2,3]}")
' Debug.Print dictData("name")           ' -> テスト
' Debug.Print dictData("values")(0)      ' -> 1
'
' ' シリアライズ
' Debug.Print oJP.Dumps(dictData)          ' -> 整形JSON (indent=4)
' Debug.Print oJP.Dumps(dictData, 0)       ' -> コンパクトJSON
'==============================================================================
Option Explicit


'--- エラー番号 ---
Private Const ERR_PARSE As Long = 10001
Private Const ERR_SERIALIZE As Long = 10002
Private Const ERR_SOURCE As String = "clsJsonParser"


'--- トークナイザー内部状態 ---
Private m_arrTokens() As String
Private m_iPos As Long
Private m_iCount As Long

'==============================================================================
' 公開メソッド
'==============================================================================

'*********************************************************
'* JSON文字列をパースし、Dictionaryを返す
'*  引数：sJsonStr - JSON文字列
'*  戻り値：パース結果のDictionary (Object型)
'*********************************************************
Public Function Loads( _
    ByVal sJsonStr As String) As Object

  If Len(Trim$(sJsonStr)) = 0 Then
    Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Loads", "空のJSON文字列です。")
  End If

  'トークン分割
  Call Tokenize(sJsonStr)
  m_iPos = 0

  'パース
  Dim vVal As Variant
  Call ParseValue(vVal)

  '残余トークンチェック
  If m_iPos < m_iCount Then
    Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Loads", "JSONの末尾に不正なデータがあります。")
  End If

  '結果の変換
  If IsArray(vVal) Then
    Set Loads = ArrayToDict(vVal)
  ElseIf IsObject(vVal) Then
    If TypeName(vVal) = "Dictionary" Then
      Set Loads = vVal
    Else
      Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Loads", "ルートはオブジェクトまたは配列である必要があります。")
    End If
  Else
    Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Loads", "ルートはオブジェクトまたは配列である必要があります。")
  End If

End Function


'*********************************************************
'* DictionaryをJSON文字列に変換する
'*  引数：dictData - 変換対象のDictionary (Object型)
'*  引数：iIndent - インデント幅 (既定=4: 整形出力、0: コンパクト出力)
'*  戻り値：JSON文字列
'*********************************************************
Public Function Dumps( _
    ByVal dictData As Object, _
    Optional ByVal iIndent As Long = 4) As String

  If dictData Is Nothing Then
    Call Err.Raise(ERR_SERIALIZE, ERR_SOURCE & ".Dumps", "NothingのDictionaryはシリアライズできません。")
  End If

  Dumps = SerializeValue(dictData, 0, iIndent)

End Function

'==============================================================================
' トークナイザー (正規表現ベース)
'==============================================================================

'---------------------------------------------------------
'- JSON文字列をトークン列に分割する
'-  引数：sJsonStr - JSON文字列
'---------------------------------------------------------
Private Sub Tokenize( _
    ByVal sJsonStr As String)

  Dim oRegExp As Object
  Set oRegExp = CreateObject("VBScript.RegExp")
  oRegExp.Global = True
  oRegExp.Pattern = _
    """(?:[^""\\]|\\.)*""" & "|" & _
    "-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?" & "|" & _
    "true|false|null" & "|" & _
    "[\[\]{}:,]"

  Dim oMatches As Object
  Set oMatches = oRegExp.Execute(sJsonStr)

  m_iCount = oMatches.Count
  If m_iCount = 0 Then
    Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Tokenize", "有効なJSONトークンがありません。")
  End If

  ReDim m_arrTokens(0 To m_iCount - 1)
  Dim i As Long
  For i = 0 To m_iCount - 1
    m_arrTokens(i) = oMatches(i).Value
  Next i

  Set oMatches = Nothing
  Set oRegExp = Nothing

End Sub


'---------------------------------------------------------
'- 現在のトークンを参照 (位置は進めない)
'-  戻り値：現在のトークン文字列
'---------------------------------------------------------
Private Function Peek() As String

  If m_iPos >= m_iCount Then
    Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Peek", "予期しないJSONの終端です。")
  End If

  Peek = m_arrTokens(m_iPos)

End Function


'---------------------------------------------------------
'- 現在のトークンを取得し、位置を1つ進める
'-  戻り値：取得したトークン文字列
'---------------------------------------------------------
Private Function Consume() As String

  Consume = Peek()
  m_iPos = m_iPos + 1

End Function


'---------------------------------------------------------
'- 期待するトークンを消費する (不一致時はエラー)
'-  引数：sExpected - 期待するトークン文字列
'---------------------------------------------------------
Private Sub Expect( _
    ByVal sExpected As String)

  Dim sToken As String
  sToken = Consume()

  If sToken <> sExpected Then
    Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".Expect", "'" & sExpected & "' が必要ですが '" & sToken & "' が見つかりました。")
  End If

End Sub

'==============================================================================
' パーサー
'==============================================================================

'---------------------------------------------------------
'- 値のパース (型に応じて分岐)
'-  引数：vOutVal - パース結果を格納する変数 (Variant)
'---------------------------------------------------------
Private Sub ParseValue( _
    ByRef vOutVal As Variant)

  Dim sToken As String
  sToken = Peek()

  Select Case sToken
    Case "{"
      Dim dictObj As Object
      Set dictObj = ParseObject()
      Set vOutVal = dictObj

    Case "["
      vOutVal = ParseArray()

    Case "true"
      vOutVal = True
      m_iPos = m_iPos + 1

    Case "false"
      vOutVal = False
      m_iPos = m_iPos + 1

    Case "null"
      vOutVal = Null
      m_iPos = m_iPos + 1

    Case Else
      If Left$(sToken, 1) = """" Then
        vOutVal = ParseStr()
      ElseIf IsNumericToken(sToken) Then
        vOutVal = ParseNumber()
      Else
        Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".ParseValue", "予期しないトークン: '" & sToken & "'")
      End If
  End Select

End Sub


'---------------------------------------------------------
'- オブジェクトのパース -> Dictionary
'-  戻り値：パースされた Dictionary オブジェクト
'---------------------------------------------------------
Private Function ParseObject() As Object

  Dim dictData As Object
  Set dictData = CreateObject("Scripting.Dictionary")
  dictData.CompareMode = vbBinaryCompare  '大文字小文字を区別

  Call Expect("{")

  '空オブジェクト
  If Peek() = "}" Then
    m_iPos = m_iPos + 1
    Set ParseObject = dictData
    Exit Function
  End If

  Do
    'キー (文字列)
    If Left$(Peek(), 1) <> """" Then
      Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".ParseObject", "オブジェクトキーは文字列である必要があります。")
    End If

    Dim sKey As String
    sKey = ParseStr()

    'コロン
    Call Expect(":")

    '値
    Dim vVal As Variant
    Call ParseValue(vVal)

    If IsObject(vVal) Then
      Set dictData(sKey) = vVal
    Else
      dictData(sKey) = vVal
    End If

    'カンマ or 閉じ括弧
    Dim sToken As String
    sToken = Peek()

    If sToken = "}" Then
      m_iPos = m_iPos + 1
      Exit Do
    ElseIf sToken = "," Then
      m_iPos = m_iPos + 1
    Else
      Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".ParseObject", "',' または '}' が必要ですが '" & sToken & "' が見つかりました。")
    End If
  Loop

  Set ParseObject = dictData

End Function

'---------------------------------------------------------
'- 配列のパース -> Variant配列
'-  戻り値：パースされた Variant 配列
'---------------------------------------------------------
Private Function ParseArray() As Variant

  Dim cItems As New Collection

  Call Expect("[")

  '空配列
  If Peek() = "]" Then
    m_iPos = m_iPos + 1
    ParseArray = Array()
    Exit Function
  End If

  Do
    '値
    Dim vVal As Variant
    Call ParseValue(vVal)
    Call cItems.Add(vVal)

    'カンマ or 閉じ括弧
    Dim sToken As String
    sToken = Peek()

    If sToken = "]" Then
      m_iPos = m_iPos + 1
      Exit Do
    ElseIf sToken = "," Then
      m_iPos = m_iPos + 1
    Else
      Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".ParseArray", "',' または ']' が必要ですが '" & sToken & "' が見つかりました。")
    End If
  Loop

  'Collection -> Variant配列に変換
  ParseArray = ColToArray(cItems)

End Function


'---------------------------------------------------------
'- 文字列トークンのパース (エスケープ展開)
'-  戻り値：デコードされた文字列
'---------------------------------------------------------
Private Function ParseStr() As String

  Dim sToken As String
  sToken = Consume()

  '両端のダブルクォートを除去
  Dim sRaw As String
  sRaw = Mid$(sToken, 2, Len(sToken) - 2)

  'エスケープなしならそのまま返す
  If InStr(1, sRaw, "\", vbBinaryCompare) = 0 Then
    ParseStr = sRaw
    Exit Function
  End If

  'エスケープシーケンスを展開
  ParseStr = UnescapeString(sRaw)

End Function


'---------------------------------------------------------
'- 数値トークンのパース -> Long または Double
'-  戻り値：数値 (Long または Double)
'---------------------------------------------------------
Private Function ParseNumber() As Variant

  Dim sToken As String
  sToken = Consume()

  '小数点・指数があれば Double
  If InStr(sToken, ".") > 0 Or InStr(sToken, "e") > 0 Or InStr(sToken, "E") > 0 Then
    ParseNumber = val(sToken)
  Else
    '整数: Long範囲内ならLong、それ以外はDouble
    Dim dVal As Double
    dVal = val(sToken)

    If dVal >= -2147483648# And dVal <= 2147483647# Then
      ParseNumber = CLng(dVal)
    Else
      ParseNumber = dVal
    End If
  End If

End Function

'==============================================================================
' シリアライザー
'==============================================================================

'---------------------------------------------------------
'- 値のシリアライズ (型に応じて分岐)
'-  引数：vVal - シリアライズ対象の値 (Variant)
'-  引数：iDepth - 現在のネスト深さ
'-  引数：iIndent - インデント幅
'-  戻り値：JSON文字列の一部
'---------------------------------------------------------
Private Function SerializeValue( _
    ByRef vVal As Variant, _
    ByVal iDepth As Long, _
    ByVal iIndent As Long) As String

  If IsNull(vVal) Or IsEmpty(vVal) Then
    SerializeValue = "null"

  ElseIf IsArray(vVal) Then
    SerializeValue = SerializeArray(vVal, iDepth, iIndent)

  ElseIf IsObject(vVal) Then
    If TypeName(vVal) = "Dictionary" Then
      SerializeValue = SerializeDict(vVal, iDepth, iIndent)
    Else
      Call Err.Raise(ERR_SERIALIZE, ERR_SOURCE & ".SerializeValue", "シリアライズ不可能なオブジェクト型: " & TypeName(vVal))
    End If

  ElseIf VarType(vVal) = vbString Then
    SerializeValue = """" & EscapeString(CStr(vVal)) & """"

  ElseIf VarType(vVal) = vbBoolean Then
    If vVal Then
      SerializeValue = "true"
    Else
      SerializeValue = "false"
    End If

  ElseIf IsNumeric(vVal) Then
    SerializeValue = NumToStr(vVal)

  Else
    Call Err.Raise(ERR_SERIALIZE, ERR_SOURCE & ".SerializeValue", "シリアライズ不可能な型: " & TypeName(vVal))
  End If

End Function

'---------------------------------------------------------
'- Dictionaryのシリアライズ
'-  引数：dictData - シリアライズ対象の Dictionary
'-  引数：iDepth - 現在のネスト深さ
'-  引数：iIndent - インデント幅
'-  戻り値：JSONオブジェクト文字列
'---------------------------------------------------------
Private Function SerializeDict( _
    ByVal dictData As Object, _
    ByVal iDepth As Long, _
    ByVal iIndent As Long) As String

  If dictData.Count = 0 Then
    SerializeDict = "{}"
    Exit Function
  End If

  Dim flgCompact As Boolean
  flgCompact = (iIndent = 0)

  Dim sNL As String
  Dim sInnerPad As String
  Dim sOuterPad As String

  If Not flgCompact Then
    sNL = vbCrLf
    sInnerPad = String$((iDepth + 1) * iIndent, " ")
    sOuterPad = String$(iDepth * iIndent, " ")
  End If

  Dim arrParts() As String
  ReDim arrParts(0 To dictData.Count - 1)

  Dim vKeys As Variant
  vKeys = dictData.keys

  Dim i As Long
  For i = 0 To dictData.Count - 1
    Dim sKey As String
    sKey = """" & EscapeString(CStr(vKeys(i))) & """"

    Dim sValue As String
    sValue = SerializeValue(dictData.Item(vKeys(i)), iDepth + 1, iIndent)

    If flgCompact Then
      arrParts(i) = sKey & ": " & sValue
    Else
      arrParts(i) = sInnerPad & sKey & ": " & sValue
    End If
  Next i

  If flgCompact Then
    SerializeDict = "{" & Join(arrParts, ", ") & "}"
  Else
    SerializeDict = "{" & sNL & Join(arrParts, "," & sNL) & sNL & sOuterPad & "}"
  End If

End Function

'---------------------------------------------------------
'- Variant配列のシリアライズ
'-  引数：arrData - シリアライズ対象の配列
'-  引数：iDepth - 現在のネスト深さ
'-  引数：iIndent - インデント幅
'-  戻り値：JSON配列文字列
'---------------------------------------------------------
Private Function SerializeArray( _
    ByRef arrData As Variant, _
    ByVal iDepth As Long, _
    ByVal iIndent As Long) As String

  '空配列チェック
  Dim lLB As Long
  Dim lUB As Long

  On Error Resume Next
  lLB = LBound(arrData)
  lUB = UBound(arrData)

  If Err.Number <> 0 Then
    On Error GoTo 0
    SerializeArray = "[]"
    Exit Function
  End If
  On Error GoTo 0

  If lUB < lLB Then
    SerializeArray = "[]"
    Exit Function
  End If

  Dim flgCompact As Boolean
  flgCompact = (iIndent = 0)

  Dim sNL As String
  Dim sInnerPad As String
  Dim sOuterPad As String

  If Not flgCompact Then
    sNL = vbCrLf
    sInnerPad = String$((iDepth + 1) * iIndent, " ")
    sOuterPad = String$(iDepth * iIndent, " ")
  End If

  Dim arrParts() As String
  ReDim arrParts(lLB To lUB)

  Dim i As Long
  For i = lLB To lUB
    Dim sValue As String
    sValue = SerializeValue(arrData(i), iDepth + 1, iIndent)

    If flgCompact Then
      arrParts(i) = sValue
    Else
      arrParts(i) = sInnerPad & sValue
    End If
  Next i

  If flgCompact Then
    SerializeArray = "[" & Join(arrParts, ", ") & "]"
  Else
    SerializeArray = "[" & sNL & Join(arrParts, "," & sNL) & sNL & sOuterPad & "]"
  End If

End Function

'==============================================================================
' ヘルパー関数
'==============================================================================

'---------------------------------------------------------
'- エスケープシーケンスの展開 (\n, \t, \uXXXX 等)
'-  引数：sStr - 対象文字列
'-  戻り値：展開後の文字列
'---------------------------------------------------------
Private Function UnescapeString( _
    ByVal sStr As String) As String

  Dim arrResult() As String
  Dim lBufIdx As Long
  Dim lBufSize As Long

  lBufSize = 64
  ReDim arrResult(1 To lBufSize)
  lBufIdx = 0

  Dim i As Long
  Dim lLen As Long
  lLen = Len(sStr)
  i = 1

  Do While i <= lLen
    Dim sChar As String
    sChar = Mid$(sStr, i, 1)

    If sChar = "\" And i < lLen Then
      'エスケープシーケンス
      i = i + 1
      Dim sEsc As String
      sEsc = Mid$(sStr, i, 1)

      Dim sUnescaped As String
      Select Case sEsc
        Case """": sUnescaped = """"
        Case "\": sUnescaped = "\"
        Case "/": sUnescaped = "/"
        Case "b": sUnescaped = Chr$(8)
        Case "f": sUnescaped = Chr$(12)
        Case "n": sUnescaped = vbLf
        Case "r": sUnescaped = vbCr
        Case "t": sUnescaped = vbTab
        Case "u"
          '\uXXXX Unicode エスケープ
          If i + 4 > lLen Then
            Call Err.Raise(ERR_PARSE, ERR_SOURCE & ".UnescapeString", "不正なUnicodeエスケープです。")
          End If

          Dim sHex4 As String
          sHex4 = Mid$(sStr, i + 1, 4)
          i = i + 4

          Dim lCP As Long
          lCP = CLng("&H" & sHex4)

          'サロゲートペア判定 (High: D800-DBFF)
          If lCP >= &HD800& And lCP <= &HDBFF& Then
            'Low surrogate (\uDC00-\uDFFF) が続くか確認
            If i + 2 <= lLen Then
              If Mid$(sStr, i + 1, 2) = "\u" Then
                Dim sHex4Low As String
                sHex4Low = Mid$(sStr, i + 3, 4)
                Dim lCPLow As Long
                lCPLow = CLng("&H" & sHex4Low)
                i = i + 6
                sUnescaped = ChrW$(lCP) & ChrW$(lCPLow)
              Else
                sUnescaped = ChrW$(lCP)
              End If
            Else
              sUnescaped = ChrW$(lCP)
            End If
          Else
            sUnescaped = ChrW$(lCP)
          End If
        Case Else
          sUnescaped = sEsc
      End Select

      lBufIdx = lBufIdx + 1
      If lBufIdx > lBufSize Then
        lBufSize = lBufSize * 2
        ReDim Preserve arrResult(1 To lBufSize)
      End If
      arrResult(lBufIdx) = sUnescaped
    Else
      '通常文字
      lBufIdx = lBufIdx + 1
      If lBufIdx > lBufSize Then
        lBufSize = lBufSize * 2
        ReDim Preserve arrResult(1 To lBufSize)
      End If
      arrResult(lBufIdx) = sChar
    End If

    i = i + 1
  Loop

  If lBufIdx = 0 Then
    UnescapeString = ""
  Else
    ReDim Preserve arrResult(1 To lBufIdx)
    UnescapeString = Join(arrResult, "")
  End If

End Function

'---------------------------------------------------------
'- 文字列のエスケープ (シリアライズ用)
'-  引数：sStr - 対象文字列
'-  戻り値：エスケープ後の文字列
'---------------------------------------------------------
Private Function EscapeString( _
    ByVal sStr As String) As String

  If Len(sStr) = 0 Then
    EscapeString = ""
    Exit Function
  End If

  '高速パス: エスケープ対象文字がなければそのまま返す
  Dim oRegExp As Object
  Set oRegExp = CreateObject("VBScript.RegExp")
  oRegExp.Pattern = "[""\\\/\x00-\x1F]"

  If Not oRegExp.test(sStr) Then
    EscapeString = sStr
    Set oRegExp = Nothing
    Exit Function
  End If
  Set oRegExp = Nothing

  'エスケープ処理
  Dim arrResult() As String
  ReDim arrResult(1 To Len(sStr) * 2)
  Dim lBufIdx As Long
  lBufIdx = 0

  Dim i As Long
  For i = 1 To Len(sStr)
    Dim sChar As String
    sChar = Mid$(sStr, i, 1)
    Dim lCode As Long
    lCode = AscW(sChar)

    '負値補正 (AscWが負を返す場合)
    If lCode < 0 Then lCode = lCode + 65536

    Dim sEscaped As String
    Select Case lCode
      Case 34: sEscaped = "\"""     '"
      Case 92: sEscaped = "\\"      '\
      Case 47: sEscaped = "\/"      '/
      Case 8: sEscaped = "\b"      'BS
      Case 12: sEscaped = "\f"      'FF
      Case 10: sEscaped = "\n"      'LF
      Case 13: sEscaped = "\r"      'CR
      Case 9: sEscaped = "\t"      'Tab
      Case 0 To 31
        sEscaped = "\u" & Right$("000" & Hex$(lCode), 4)
      Case Else
        sEscaped = sChar
    End Select

    lBufIdx = lBufIdx + 1
    arrResult(lBufIdx) = sEscaped
  Next i

  ReDim Preserve arrResult(1 To lBufIdx)
  EscapeString = Join(arrResult, "")

End Function

'---------------------------------------------------------
'- 数値を文字列に変換 (ロケール非依存)
'-  引数：vVal - 対象数値
'-  戻り値：文字列化した数値
'---------------------------------------------------------
Private Function NumToStr( _
    ByVal vVal As Variant) As String

  NumToStr = CStr(vVal)
  NumToStr = Replace(NumToStr, ",", ".")

End Function

'---------------------------------------------------------
'- トークンが数値かどうか判定
'-  引数：sToken - 判定対象トークン
'-  戻り値：数値なら True
'---------------------------------------------------------
Private Function IsNumericToken( _
    ByVal sToken As String) As Boolean

  Dim sChar As String
  sChar = Left$(sToken, 1)

  IsNumericToken = (sChar >= "0" And sChar <= "9") Or sChar = "-"

End Function

'---------------------------------------------------------
'- Variant配列をインデックスキーのDictionaryに変換
'-  引数：arrData - 変換対象配列
'-  戻り値：インデックスをキーとする Dictionary
'---------------------------------------------------------
Private Function ArrayToDict( _
    ByRef arrData As Variant) As Object

  Dim dictData As Object
  Set dictData = CreateObject("Scripting.Dictionary")
  dictData.CompareMode = vbBinaryCompare

  Dim lLB As Long
  Dim lUB As Long

  On Error Resume Next
  lLB = LBound(arrData)
  lUB = UBound(arrData)

  If Err.Number <> 0 Then
    On Error GoTo 0
    Set ArrayToDict = dictData
    Exit Function
  End If
  On Error GoTo 0

  Dim i As Long
  For i = lLB To lUB
    If IsObject(arrData(i)) Then
      Set dictData(CStr(i)) = arrData(i)
    Else
      dictData(CStr(i)) = arrData(i)
    End If
  Next i

  Set ArrayToDict = dictData

End Function

'---------------------------------------------------------
'- CollectionをVariant配列に変換
'-  引数：cCol - 変換対象コレクション
'-  戻り値：格納された Variant 配列
'---------------------------------------------------------
Private Function ColToArray( _
    ByVal cCol As Collection) As Variant

  If cCol.Count = 0 Then
    ColToArray = Array()
    Exit Function
  End If

  Dim arrData() As Variant
  ReDim arrData(0 To cCol.Count - 1)

  Dim i As Long
  For i = 1 To cCol.Count
    If IsObject(cCol(i)) Then
      Set arrData(i - 1) = cCol(i)
    Else
      arrData(i - 1) = cCol(i)
    End If
  Next i

  ColToArray = arrData

End Function

