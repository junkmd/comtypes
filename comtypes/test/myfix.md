## 実施したこと

`find_memleak` のプロセスメモリ系ターゲット（`working_set` / `private_usage`）の測定を、複数のサブフェーズに分割し、その**最小差分**を返すように変更しました。これにより、CI で時折発生していた「一度きりの割り当て（新しいヒープセグメント、遅延生成されたスレッドスタック、RPC バッファ等）をリークと誤判定する」問題を除外できます。実リークはどのサブフェーズでも増加するため最小値でも検出でき、一度きりの変化は1つのサブフェーズにしか現れないため除外されます。サブフェーズは合計呼び出し回数を変えないため、テストは遅くなりません。`mallocspy` は COM タスク割り当てを正確に数えるため変更していません。

変更ファイル:
- [`src/comtypes/test/find_memleak.py`](src/comtypes/test/find_memleak.py:343) — `_PROCESS_MEMORY_SUBPHASES = 3` を追加し、[`find_memleak()`](src/comtypes/test/find_memleak.py:359) をサブフェーズ分割＋最小値採用に変更。
- [`src/comtypes/test/test_find_memleak.py`](src/comtypes/test/test_find_memleak.py:209) — 実リークが検出されること、一度きりの割り当てが除外されることを固定するテストを2件追加。

## 実行したテストと結果

- `uv run python -m unittest comtypes.test.test_find_memleak -v` → 13 tests, OK
- `uv run python -m unittest comtypes.test.test_comserver -v` → 41 tests, OK (skipped=17)
- `run_tests.bat x64`（フルスイート、カバレッジ付き）→ **4回すべて成功**（752 tests, OK (skipped=25)、約143〜145秒。修正前は約140秒で、実行時間の増加はほぼなし）
- `uvx ruff==0.6.9 format src/comtypes/.` → 153 files left unchanged
- `uvx ruff==0.6.9 check --fix --extend-select I src/comtypes/.` → All checks passed!

## 範囲外で気づいた問題（修正していない）

- この環境（Windows Server 2025 / Python 3.15）では、報告された `TestLocalServer.test_get_name` の 311296 バイトのリークを再現できませんでした。`private_usage` の生の差分を20回計測してもすべて 0 でした。CI ランナー固有のタイミング・負荷に依存するノイズと推測しています。そのため、修正が実際のフレークを解消するかは直接検証できていません（ロジック上は一度きりの割り当てを除外します）。
- `mallocspy` はローカルサーバーのクライアント側 BSTR 割り当てを追跡しないことを実験で確認しました（意図的に BSTR を保持しても net 0）。そのためローカルサーバーで `mallocspy` を使う代替案は成立しません。
- 診断用に作成した空ファイル [`src/comtypes/test/test_diag_memleak.py`](src/comtypes/test/test_diag_memleak.py) が未追跡で残っています。許可リストに削除コマンド（`del`/`rm`）が無いため削除できませんでした。また `uv run` が生成した `src/uv.lock` も未追跡です。どちらもコミット対象外です。

## 却下されたコマンド・止まった点

- `git commit -am "..."` を実行しようとしましたが、許可リスト（`git log`/`git diff`/`git status`/`git show` のみ）に無いため却下されました。AGENTS.md の「1タスク = 1コミット」に従えていません。コミットは未実施です。迂回はしていません。
- 同様に `del`（診断ファイル削除）も許可リスト外で却下されました。

## コミットメッセージ案（未実施）

```
test: Filter one-off allocations in `find_memleak` process-memory targets.

`find_memleak` with the `private_usage` target (used by the Local Server
tests) measures the whole process, so a single measurement can report an
unrelated one-off allocation (a new heap segment, a lazily created thread
stack, an RPC buffer, ...) as a leak. This made the Local Server tests fail
occasionally in CI (see #971).

Split the measurement into sub-phases and return the smallest difference.
A real leak grows in every sub-phase, so the smallest difference still
reports it, while a one-off change appears in only one sub-phase and is
filtered out. The sub-phases share the same total number of calls as a
single measurement, so the test does not get slower.

`mallocspy` counts COM task allocations exactly, so it is left unchanged.

Tests: `uv run python -m unittest comtypes.test.test_find_memleak` and
`run_tests.bat x64` (752 tests, OK).
```