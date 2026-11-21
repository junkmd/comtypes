引数付き（名前付き）プロパティを実装した新しいテストダブルを `src/source/CppTestSrv` に追加し、`bound_named_property` / `named_property` の挙動を検証するテストを新規作成しました。

## 実施したこと

### 1. テストダブル（C++ COM サーバー）の追加
- [`SERVER.IDL`](src/source/CppTestSrv/SERVER.IDL:236) に引数付きプロパティを持つ dual インタフェース `IDualNamedPropertyTest` と、その dispinterface エイリアス `IDispNamedPropertyTest`、および coclass `CoNamedPropertyTest` を追加しました。プロパティは以下です。
  - `Item(index)` : 読み書き（1引数）
  - `Matrix(row, col)` : 読み書き（2引数、タプルインデックス用）
  - `ReadOnlyItem(index)` : 読み取り専用
  - `WriteOnlyItem(index)` : 書き込み専用
  - `OptionalItem([in, optional] VARIANT index)` : 引数が省略可能（`_all_slice` 経路用）
- [`CoNamedPropertyTest.h`](src/source/CppTestSrv/CoNamedPropertyTest.h:1) / [`CoNamedPropertyTest.cpp`](src/source/CppTestSrv/CoNamedPropertyTest.cpp:1) を新規作成しました。値は `std::map` に保持し、未設定時はインデックスから決まる決定的な値を返します（`Item` は `index*10`、`Matrix` は `row*10+col`、`OptionalItem` は引数省略時 `-1`）。
- [`SERVER.CPP`](src/source/CppTestSrv/SERVER.CPP:93) のファクトリ配列に登録、[`MAKEFILE`](src/source/CppTestSrv/MAKEFILE:78) にビルドルールとリンク対象を追加しました。

### 2. テストの追加
- [`test_named_property.py`](src/comtypes/test/test_named_property.py:1) を新規作成しました。dual 経路（`ComPropertyGenerator`）と disp 経路（`DispPropertyGenerator`）の両方について、`__getitem__`（単一/タプル/`_all_slice`）、`__setitem__`、`__call__`、読み取り専用・書き込み専用での `TypeError`、`__iter__` の `TypeError`、`__repr__` を、1テスト1挙動で明示的に検証します（計30テスト）。

## 実行したテストと結果
- ビルド: `cd tools` → `build_test_server.bat` … 成功（`CoNamedPropertyTest.obj` を含めてリンク成功）
- サーバー登録: `server.exe /RegServer` … 成功
- 新規テスト: `uv run python -m unittest comtypes.test.test_named_property -v` … **30 tests OK**
- 全テストスイート: `cd tools` → `run_tests.bat x64` … **Ran 758 tests, OK (skipped=33)**
- 整形: `uvx ruff==0.6.9 format src/comtypes/.`（154 files left unchanged）→ `uvx ruff==0.6.9 check --fix --extend-select I src/comtypes/.`（All checks passed!）

## 範囲外で気づいた問題（修正はしていません）
- `bound_named_property` は `__len__` を実装していないため `len(obj.Item)` は `TypeError` になります。今回のテスト対象（プロパティの読み書き・呼び出し）から外れるためテストは追加していません。
- `_all_slice` 経路（`obj.Prop[:]`）は「引数がすべて省略可能なプロパティ」でのみ到達可能です。この経路を実際に通すには `OptionalItem` のような省略可能引数が必要で、通常の必須引数プロパティでは `obj.Item[:]` は COM 呼び出しエラーになります。

## 却下されたコマンド・判断に迷った点
- 観察用に一時作成した `src/comtypes/test/test_named_property_observe.py` を削除しようと `del src\comtypes\test\test_named_property_observe.py` を実行しましたが、許可コマンド外のため却下されました。AGENTS.md §3 に従い別コマンドでの再試行はせず、ご指示どおり**空ファイル**に置き換えました（`git status` 上は untracked のまま残ります）。
- `uv run` の副作用で `src/uv.lock` が untracked として生成されています（削除コマンドが許可されていないため未削除）。コミット対象から除外してください。
- テストコードは AGENTS.md §5 の WET 方針に従い、dual/disp で共通化せず明示的に2回書いています。

追加したテストダブルの仕様を詳しく説明します。

## 1. 目的

`comtypes` は「引数を取る COM プロパティ」（parameterized property / named property）を [`named_property`](src/comtypes/_memberspec.py:739) ディスクリプタとして公開します。その束縛後の実体 [`bound_named_property`](src/comtypes/_memberspec.py:698) は `obj.Prop[index]` / `obj.Prop[index] = value` / `obj.Prop(index)` / タプルインデックス `obj.Prop[a, b]` をサポートしますが、これを実装した COM サーバーが Excel 以外になく検証しづらい、というのが出発点です。そこで **引数付きプロパティだけを持つ最小のテストダブル** を `CppTestSrv` に追加しました。

## 2. コンポーネント構成

| 項目 | 値 |
|---|---|
| coclass | `CoNamedPropertyTest` / `{9D3F6A18-2E5B-4C70-8A1F-6B4D9E2C7F53}` |
| ProgID | `Comtypes.NamedPropertyTest`（バージョン付き: `.1`） |
| dual インタフェース | `IDualNamedPropertyTest` / `{2C7E5B41-9A3D-4F82-B6E1-7D4C8A0F5E93}` |
| disp エイリアス | `IDispNamedPropertyTest` / `{6F1A8D24-3B7C-4E59-A0D2-8C5E9B1F4A76}` |
| 型ライブラリ | 既存の `ComtypesCppTestSrvLib` / `{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}` |
| 実行形態 | Out-of-process（Local Server, `Server.exe`） |

`IDualNamedPropertyTest` は `IDispatch` 派生の `[dual]` インタフェースで、`IDispNamedPropertyTest` はその dispinterface エイリアス（`interface IDualNamedPropertyTest;`）です。同じプロパティ群を **vtable 経由**と **`IDispatch::Invoke` 経由**の両方で叩けるため、`comtypes` 側の 2 つの生成経路（`ComPropertyGenerator` と `DispPropertyGenerator`）を両方検証できます。

## 3. インタフェース仕様（プロパティ一覧）

[`SERVER.IDL`](src/source/CppTestSrv/SERVER.IDL:250) で定義した 5 つのプロパティです。dispid は 1〜5。

| プロパティ | dispid | 種別 | シグネチャ | 検証したいこと |
|---|---|---|---|---|
| `Item` | 1 | propget + propput | `get(index) -> long` / `put(index, value)` | 単一インデックスの読み書き・呼び出し |
| `Matrix` | 2 | propget + propput | `get(row, col) -> long` / `put(row, col, value)` | タプルインデックス（複数引数） |
| `ReadOnlyItem` | 3 | propget のみ | `get(index) -> long` | setter が無いときの `__setitem__` の `TypeError` |
| `WriteOnlyItem` | 4 | propput のみ | `put(index, value)` | getter が無いときの `__getitem__`/`__call__` の `TypeError` |
| `OptionalItem` | 5 | propget のみ | `get([in, optional] VARIANT index) -> long` | 引数省略（`obj.OptionalItem[:]`）で getter を無引数呼び出し |

`OptionalItem` だけ引数が `[in, optional] VARIANT` なのは、`comtypes` の [`_all_slice`](src/comtypes/__init__.py:43) 経路（`obj.Prop[:]` → `fget(instance)` を無引数で呼ぶ）に到達するには「引数がすべて省略可能」である必要があるためです。省略時は `comtypes` が `VARIANT.missing`（`vt == VT_ERROR`）を渡します。

## 4. C++ 実装の挙動

[`CoNamedPropertyTest.cpp`](src/source/CppTestSrv/CoNamedPropertyTest.cpp:201) の各メソッドは、README の方針どおり「入力をそのまま返す／インデックスから決まる決定的な値を返す」だけの単純な実装です。状態は **インスタンスごと** に `std::map` で保持します（`CreateObject` ごとに新しいインスタンス）。

| メソッド | 挙動 |
|---|---|
| `get_Item(index)` | 保存済みならその値、未保存なら `index * 10` |
| `put_Item(index, value)` | `m_itemValues[index] = value` |
| `get_Matrix(row, col)` | 保存済みならその値、未保存なら `row * 10 + col` |
| `put_Matrix(row, col, value)` | `m_matrixValues[{row,col}] = value` |
| `get_ReadOnlyItem(index)` | 常に `index * 10`（保存しない） |
| `put_WriteOnlyItem(index, value)` | `m_writeOnlyValues[index] = value`（getter は無い） |
| `get_OptionalItem(index)` | `index.vt == VT_ERROR` なら `-1`、それ以外は `index.lVal` |

`value == NULL` のときは `E_POINTER` を返します。`IDispatch` の 4 メソッド（`GetTypeInfoCount`/`GetTypeInfo`/`GetIDsOfNames`/`Invoke`）は既存の [`CoRecordSafearrayTest`](src/source/CppTestSrv/CoRecordSafearrayTest.cpp:212) と同じパターンで、型情報をロードして `ITypeInfo::Invoke` に委譲します。`NondelegatingQueryInterface` は `IID_IDualNamedPropertyTest` / `DIID_IDispNamedPropertyTest` / `IID_IDispatch` に応答します。

## 5. comtypes 側での見え方

生成モジュールでは、dual 側は [`ComPropertyGenerator`](src/comtypes/_memberspec.py:527)、disp 側は [`DispPropertyGenerator`](src/comtypes/_memberspec.py:552) により、いずれも `nargs > 0` のため `named_property` になります（生成コードの抜粋）:

```python
Item = hints.named_property('Item', _get_Item, _set_Item)
Matrix = hints.named_property('Matrix', _get_Matrix, _set_Matrix)
ReadOnlyItem = hints.named_property('ReadOnlyItem', _get_ReadOnlyItem)
WriteOnlyItem = hints.named_property('WriteOnlyItem', fset=_set_WriteOnlyItem)
OptionalItem = hints.named_property('OptionalItem', _get_OptionalItem)
```

`Item` は `_patch_dunder` により `__call__` / `__getitem__` / `__setitem__` も型ヒント上で付与されます。

## 6. テストとの対応

[`test_named_property.py`](src/comtypes/test/test_named_property.py:1) は dual 用 `TestDualNamedProperty` と disp 用 `TestDispNamedProperty` の 2 クラスで、上記 5 プロパティを使い次の挙動を 1 テスト 1 挙動で固定しています。

- `__getitem__` がインデックスを getter に渡す（`Item[3] == 30`）
- `__setitem__` → `__getitem__` のラウンドトリップ（`Item[3] = 42` → `42`）
- `__call__` が引数を渡す（`Item(3)`、`Matrix(1, 2)`）
- タプルインデックス（`Matrix[1, 2]`）
- 読み取り専用への代入で `TypeError`
- 書き込み専用への読み出し／呼び出しで `TypeError`
- `_all_slice`（`OptionalItem[:] == -1`）と省略可能引数の明示指定（`OptionalItem[5] == 5`）
- `__iter__` の `TypeError`
- `__repr__` がインタフェース名とプロパティ名を含む

このテストダブルは「引数付きプロパティのマーシャリング」と「`bound_named_property` の Python 側プロトコル」の両方を、実プロセスの COM 越しに検証できるようにするものです。
