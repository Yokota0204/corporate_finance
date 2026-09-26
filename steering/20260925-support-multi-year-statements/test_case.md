# テストケース: 3年分以上の財務データを一括で取得・保存できるようにする

- テストフレームワーク: Jest + ts-jest
- テストファイル:
  - 単体テスト: `test/statements.test.ts`
  - 結合テスト: `test/doPost.test.ts`
- テスト名（`test()` / `describe()` の名前）は日本語で書く。
- テストデータは `makeStatement(month, base)`（`test/fixtures.ts`）で作る。以下の「S(2024/03)」は month が `"2024/03"` の決算を表す。

## 単体テスト

### `normalizeStatements()`

| ID | 前提・入力 | 期待結果 | 対応する受け入れ条件 |
|---|---|---|---|
| U-01 | `{ statements: [S(2023/03), S(2024/03), S(2025/03)] }`（古い順） | 同じ3件を同じ順で返す | 3件以上を一括で追記 |
| U-02 | `{ statements: [S(2025/03), S(2024/03), S(2023/03)] }`（新しい順） | 2023/03, 2024/03, 2025/03 の順で返す | 古い順で追記 |
| U-03 | `{ statements: [S(2024/03), S(2025/03), S(2022/03), S(2023/03)] }`（順不同） | 2022/03, 2023/03, 2024/03, 2025/03 の順で返す | 古い順で追記 |
| U-04 | `{ statements: [S(2025/03), S(2024/03)] }` | 戻り値を並べ替えても、入力の配列の順番は変わらない（2025/03, 2024/03 のまま） | ― （副作用がないこと） |
| U-05 | `{ statements: [S(2024/03)] }`（1件） | 1件の配列を返す | 1件のとき1行追記 |
| U-06 | `{ statements: [S(2024/03, base=1), S(2023/03), S(2024/03, base=100)] }`（同じ month が2件） | 2023/03, 2024/03(base=1), 2024/03(base=100) の順で返す。同じ month 同士の順番は入力のまま | 重複してもスキップ・上書きしない |
| U-07 | `{ before: S(2024/03), after: S(2025/03) }` | [before, after] の順で返す | 旧形式 before/after |
| U-08 | `{ before: S(2025/03), after: S(2024/03) }`（旧形式で month が逆） | 並べ替えず [before, after] の順で返す | 旧形式は今までと同じ |
| U-09 | `{ after: S(2025/03) }` | [after] を返す | 旧形式 after のみ |
| U-10 | `{ statements: [] }` | 「statementsに決算が含まれていません。」でエラー | 空配列はエラー |
| U-11 | `{}` | 「ボディに今期の決算が含まれていません。」でエラー | statements も after もないときはエラー |
| U-12 | `{ before: S(2024/03) }`（after なし） | 「ボディに今期の決算が含まれていません。」でエラー | 同上（現状と同じ） |
| U-13 | `{ statements: [S(2025/03)], after: S(2025/03) }` | 「statementsとbefore/afterは同時に指定できません。」でエラー | 同時指定はエラー |
| U-14 | `{ statements: [S(2025/03)], before: S(2024/03) }` | 同上 | 同時指定はエラー |
| U-15 | `{ statements: "2025/03" }`（配列でない） | 「statementsに決算が含まれていません。」でエラー | 不正な入力はエラー |

### `toPlRow()` / `toBsRow()` / `toCfRow()`

| ID | 前提・入力 | 期待結果 | 対応する受け入れ条件 |
|---|---|---|---|
| U-16 | `toPlRow(S(2025/03))` | `[month, sales, cost, sellingExpenses, nonOperatingIncome, nonOperatingExpenses, specialIncome, specialLosses, net]` の9要素 | 列の並びが現状と同じ |
| U-17 | `toBsRow(S(2025/03))` | `[month, cash, notesAndAccountsReceivableTrade, otherReceivables, depositsPaid, shortTermLoans, allowance, currentAssets, nonCurrentAssets, currentLiabilities, nonCurrentLiabilities, retainedEarnings]` の12要素 | 同上 |
| U-18 | `toCfRow(S(2025/03))` | `[month, operating, investing, financing, proceedsFromShortTermBorrowings, proceedsFromLongTermBorrowings, proceedsFromIssuanceOfBonds, repaymentsOfShortTermBorrowings, repaymentsOfLongTermBorrowings, repaymentsOfBonds, purchaseOfTreasuryStock, retirementOfTreasuryStock, dividendsPaid]` の13要素 | 同上 |
| U-19 | 各 `toXxRow` の要素数 | 定数から計算した列数と一致する（PL: `NET_COL_IN_PL_SHEET - MONTH_COL_IN_PL_SHEET + 1`、BS: `RETAINED_EARNINGS_COL_IN_BS_SHEET - MONTH_COL_IN_BS_SHEET + 1`、CF: `LAST_INPUT_COL_IN_CF_SHEET - MONTH_COL_IN_CF_SHEET + 1`） | 範囲と行の列数がずれない |

## 結合テスト（`doPost`）

前提：3シート（損益計算・資産・キャッシュフロー）のモックを用意し、A列の最終行を 損益計算=10、資産=20、キャッシュフロー=30 とする。`PropertiesService` の `SS_ID` は `"dummy"` を返す。

| ID | シナリオ | 入力（POSTのボディ） | 期待結果 |
|---|---|---|---|
| I-01 | 新形式で3件を書き込む | `statements: [S(2023/03), S(2024/03), S(2025/03)]` | 損益計算は11行目から3行×9列、資産は21行目から3行×12列、キャッシュフローは31行目から3行×13列の範囲に `setValues` される。各行の中身は `toXxRow` の結果と一致する |
| I-02 | 新形式・順不同の5件 | `statements: [S(2025/03), S(2021/03), S(2023/03), S(2024/03), S(2022/03)]` | 3シートとも5行が 2021/03 → 2025/03 の順で書き込まれる |
| I-03 | 新形式で1件 | `statements: [S(2025/03)]` | 3シートとも1行だけ書き込まれる |
| I-04 | 旧形式 before/after | `before: S(2024/03), after: S(2025/03)` | 3シートとも2行が before → after の順で書き込まれる（変更前と同じ結果） |
| I-05 | 旧形式 after のみ | `after: S(2025/03)` | 3シートとも1行だけ書き込まれる（変更前と同じ結果） |
| I-06 | 空配列 | `statements: []` | エラーになり、どのシートの `setValues` も呼ばれない |
| I-07 | 同時指定 | `statements: [S(2025/03)], after: S(2025/03)` | エラーになり、どのシートの `setValues` も呼ばれない |
| I-08 | レスポンス | I-01 と同じ | `{ status: "ok", received: <送ったボディ> }` の JSON が返り、MimeType が JSON |
