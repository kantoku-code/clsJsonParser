# clsJsonParser(VBA JSON Parser Class)
A lightweight, dependency-free JSON parser class for VBA.

This repository provides:

- `clsJsonParser.cls` — A pure-VBA JSON parser  
- `ModJsonParserTest.bas` — A comprehensive unit test module  
- `test_sample.json` — A structured test dataset used by the unit tests  

The parser supports JSON objects, arrays, numbers, strings, booleans, null, Unicode escapes, nested structures, and scientific notation.  
It is designed for classic VBA environments such as Excel, Access, Word, PowerPoint, Outlook, and VB6.

---

## Features

This JSON parser focuses on clarity, portability, and predictable behavior in classic VBA environments.  
Key characteristics include:

- **Pure VBA implementation**  
  No external libraries, references, or dependencies are required.

- **Single-file class design**  
  The entire parser is contained in one `.cls` file, making it easy to import into any VBA project.

- **Python-like API**  
  Provides `Loads` and `Dumps` methods for intuitive JSON handling.

- **Consistent handling of objects and arrays**  
  - JSON objects → `Scripting.Dictionary`  
  - JSON arrays → `Variant` arrays  

- **Case-sensitive key behavior**  
  Keys are treated exactly as written, following JSON specification rules.

- **Unicode escape support**  
  Fully supports sequences like `\u3053\u3093\u306b\u3061\u306f`.

- **Scientific notation support**  
  Numbers such as `1.2e5` or `-3.0E+2` are parsed correctly.

- **Pretty-print and compact output**  
  `Dumps` can format JSON with indentation or produce compact one-line output.

- **UTF-8 helpers included**  
  Includes BOM-safe UTF-8 read/write utilities using `ADODB.Stream`.

- **Comprehensive unit tests included**  
  The repository contains a full test suite covering objects, arrays, escapes, Unicode, errors, and round-trip validation.

---

## Files
| File | Description |
|------|-------------|
| `clsJsonParser.cls` | JSON parser class |
| `ModJsonParserTest.bas` | Unit test module |
| `test_sample.json` | Test data used by the unit tests |

---

## Usage Example

### Load JSON
```vb
Dim jp As New clsJsonParser
Dim dict As Object
Set dict = jp.Loads(jsonString)
```

### Dump JSON (pretty)
```vb
Debug.Print jp.Dumps(dict)
```

### Dump JSON (compact)
```vb
Debug.Print jp.Dumps(dict, 0)
```

---

## Unit Tests

A full test suite is included in `ModJsonParserTest.bas`.

Run all tests:

```vb
RunAllTests
```

The tests cover:

- Basic objects  
- Nested objects  
- Arrays and mixed arrays  
- Unicode escape sequences  
- Japanese strings  
- Empty objects/arrays  
- Root-level arrays  
- Case-sensitive keys  
- Deep nesting  
- Scientific notation  
- Dumps/Loads round-trip  
- Error handling  
- File loading using UTF-8  

---

## License
**MIT**

---
---

# 以下は、日本語です

# VBA用 JSON パーサークラス  
VBA で JSON を扱うための軽量・シンプルなパーサークラスです。

このリポジトリには以下が含まれます：

- `clsJsonParser.cls` — JSON パーサークラス  
- `ModJsonParserTest.bas` — ユニットテストモジュール  
- `test_sample.json` — テスト用 JSON データ  

---

## 特徴

この JSON パーサーは、VBA 環境での扱いやすさと移植性を重視して設計されています。  
主な特徴は次のとおりです：

- **Pure VBA（純粋な VBA 実装）**  
  外部ライブラリや参照設定は不要です。

- **単一ファイルのクラス構成**  
  パーサーは 1 つの `.cls` ファイルに収まっており、どの VBA プロジェクトにも簡単に追加できます。

- **Python ライクな API**  
  `Loads` / `Dumps` による直感的な JSON 操作が可能です。

- **オブジェクトと配列の明確な扱い**  
  - JSON オブジェクト → `Scripting.Dictionary`  
  - JSON 配列 → `Variant` 配列  

- **キーの大文字小文字を区別**  
  JSON 仕様に従い、キーは厳密に区別されます。

- **Unicode エスケープ対応**  
  `\uXXXX` 形式の文字を正しくデコードします。

- **指数表記に対応**  
  `1.2e5` や `-3.0E+2` のような数値も正しく処理します。

- **整形出力とコンパクト出力**  
  `Dumps` でインデント付き／1 行の JSON を生成できます。

- **UTF-8 読み書きヘルパー付き**  
  BOM を考慮した UTF-8 読み書きが可能です。

- **包括的なユニットテストを同梱**  
  オブジェクト、配列、Unicode、エスケープ、エラー処理、ラウンドトリップなどを網羅したテストを提供しています。

---

## ファイル構成
| ファイル名 | 説明 |
|-----------|------|
| `clsJsonParser.cls` | JSON パーサークラス |
| `ModJsonParserTest.bas` | ユニットテスト |
| `test_sample.json` | テストデータ |

---

## 使用例

### JSON を読み込む
```vb
Dim jp As New clsJsonParser
Dim dict As Object
Set dict = jp.Loads(jsonString)
```

### 整形して出力
```vb
Debug.Print jp.Dumps(dict)
```

### コンパクトに出力
```vb
Debug.Print jp.Dumps(dict, 0)
```

---

## ユニットテスト

`RunAllTests` を実行すると、  
イミディエイトウィンドウに結果が表示されます。

テスト内容：

- 基本オブジェクト  
- ネストオブジェクト  
- 配列・混合配列  
- Unicode エスケープ  
- 日本語文字列  
- 空オブジェクト/配列  
- ルート配列  
- 大文字小文字区別  
- 深いネスト  
- 指数表記  
- Dumps/Loads のラウンドトリップ  
- エラー処理  
- UTF-8 ファイル読み込み  

---

## ライセンス
**MIT**