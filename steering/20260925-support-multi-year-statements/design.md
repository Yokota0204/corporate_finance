# 設計: 3年分以上の財務データを一括で取得・保存できるようにする

## 実装アプローチ（※）

- 「リクエストボディ → 書き込む決算の配列（古い順）」の変換を、新しいファイル `src/statements.ts` の `normalizeStatements()` にまとめる。入力のチェックもここで行う。
- 「決算1件 → 各シートの1行」の変換を `toPlRow()` / `toBsRow()` / `toCfRow()` に切り出す。`doPost` にあった before/after の分岐と、同じ内容を繰り返し書いている部分をなくす。
- `doPost` は、`normalizeStatements()` の結果の件数分だけ範囲を取り、`map(toXxRow)` で作った行を `setValues` する形にする。
- テストでは `src/*.ts` を TypeScript で JS に変換し、Node の `vm` で**1つのグローバル空間**に読み込む。GAS と同じ実行のされ方になるので、本番コードに `module.exports` などを足さずにすむ。`SpreadsheetApp` などの GAS のグローバルはモックを入れる。

### `src/models.ts`

```diff
 type RequestBody = {
+  statements?: FinancialStatement[];
   before?: FinancialStatement;
   after?: FinancialStatement;
 };
```

### `src/statements.ts`（新規）

```ts
/**
 * リクエストボディからシートに書き込む決算の配列を作る
 * - statements 形式: month の昇順に並べ替えて返す
 * - before/after 形式（旧形式）: [before, after] または [after] を返す
 */
function normalizeStatements(body: RequestBody): FinancialStatement[] {
  const hasLegacy: boolean = body.before !== undefined || body.after !== undefined;

  if (body.statements !== undefined) {
    if (hasLegacy) {
      throw "statementsとbefore/afterは同時に指定できません。";
    }
    if (!Array.isArray(body.statements) || body.statements.length === 0) {
      throw "statementsに決算が含まれていません。";
    }
    // 元の配列を壊さないようにコピーしてから並べ替える（sort は安定ソート）
    return [...body.statements].sort((a, b) =>
      a.month < b.month ? -1 : a.month > b.month ? 1 : 0
    );
  }

  // 今期の決算が含まれない場合はエラー
  if (!body.after) {
    throw "ボディに今期の決算が含まれていません。";
  }
  return body.before ? [body.before, body.after] : [body.after];
}

// 損益計算シートの1行（月〜純利益）
function toPlRow(fs: FinancialStatement): (string | number)[] {
  return [
    fs.month,
    fs.pl.sales,
    fs.pl.cost,
    fs.pl.sellingExpenses,
    fs.pl.nonOperatingIncome,
    fs.pl.nonOperatingExpenses,
    fs.pl.specialIncome,
    fs.pl.specialLosses,
    fs.pl.net,
  ];
}

// 資産シートの1行（月〜利益剰余金）
function toBsRow(fs: FinancialStatement): (string | number)[] {
  return [
    fs.month,
    fs.bs.cash,
    fs.bs.notesAndAccountsReceivableTrade,
    fs.bs.otherReceivables,
    fs.bs.depositsPaid,
    fs.bs.shortTermLoans,
    fs.bs.allowance,
    fs.bs.currentAssets,
    fs.bs.nonCurrentAssets,
    fs.bs.currentLiabilities,
    fs.bs.nonCurrentLiabilities,
    fs.bs.retainedEarnings,
  ];
}

// キャッシュフローシートの1行（月〜配当金の支払額）
function toCfRow(fs: FinancialStatement): (string | number)[] {
  return [
    fs.month,
    fs.cf.operating,
    fs.cf.investing,
    fs.cf.financing,
    fs.cf.proceedsFromShortTermBorrowings,
    fs.cf.proceedsFromLongTermBorrowings,
    fs.cf.proceedsFromIssuanceOfBonds,
    fs.cf.repaymentsOfShortTermBorrowings,
    fs.cf.repaymentsOfLongTermBorrowings,
    fs.cf.repaymentsOfBonds,
    fs.cf.purchaseOfTreasuryStock,
    fs.cf.retirementOfTreasuryStock,
    fs.cf.dividendsPaid,
  ];
}
```

### `src/main.ts`

入力チェックをシートの取得より前に行うので、エラー時はシートに何も書き込まれない（現状と同じ）。

```diff
   const data: RequestBody = JSON.parse(rawBody);
   log("info", JSON.stringify(data));

-  // 今期の決算が含まれない場合はエラー
-  if (!data.after) {
-    throw "ボディに今期の決算が含まれていません。";
-  }
+  // 書き込む決算を古い順に取得（入力が不正な場合はここでエラー）
+  const statements: FinancialStatement[] = normalizeStatements(data);
+  log("info", `書き込む決算の件数: ${statements.length}`);
   ...
   // セットする範囲の行数
-  const setRowCount: number = data.before ? 2 : 1;
+  const setRowCount: number = statements.length;
   ...（各シートの取得・最終行の取得・範囲の取得は変更なし）
-  const beforeFs: FinancialStatement | undefined = data.before;
-  const afterFs: FinancialStatement = data.after;
-
-  // レスポンスの数値をシートに出力
-  if (beforeFs) {
-    plRange.setValues([ ...before..., ...after... ]);
-    bsRange.setValues([ ... ]);
-    cfRange.setValues([ ... ]);
-  } else {
-    plRange.setValues([ ...after... ]);
-    bsRange.setValues([ ... ]);
-    cfRange.setValues([ ... ]);
-  }
+  // 決算の数値をシートに出力
+  plRange.setValues(statements.map(toPlRow));
+  bsRange.setValues(statements.map(toBsRow));
+  cfRange.setValues(statements.map(toCfRow));
```

### テスト基盤

#### `package.json`

```diff
   "scripts": {
     "build": "...",
     "push": "clasp push",
-    "deploy": "npm run build && npm run push"
+    "deploy": "npm run test && npm run build && npm run push",
+    "test": "jest"
   },
   "devDependencies": {
     "@types/google-apps-script": "^1.0.97",
+    "@types/jest": "^30.0.0",
+    "jest": "^30.0.0",
+    "ts-jest": "^29.4.0",
     "typescript": "^5.8.3"
   }
```

※ ts-jest の実際に入るバージョンに合わせて jest のメジャーバージョンを揃える（インストール時に確認）。
※ `deploy` の前にテストを走らせる変更は任意。不要なら外す。

#### `jest.config.cjs`（新規。package.json が `"type": "module"` のため .cjs）

```js
module.exports = {
  testEnvironment: "node",
  roots: ["<rootDir>/test"],
  transform: {
    "^.+\\.ts$": ["ts-jest", { tsconfig: "tsconfig.test.json" }],
  },
};
```

#### `tsconfig.test.json`（新規）

本番ビルド用の `tsconfig.json` は `rootDir: src` なので、テスト用に分ける。

```json
{
  "extends": "./tsconfig.json",
  "compilerOptions": {
    "rootDir": ".",
    "noEmit": true
  },
  "include": ["src", "test"]
}
```

#### `test/helpers/loadGas.ts`（新規）

```ts
import * as fs from "fs";
import * as path from "path";
import * as vm from "vm";
import * as ts from "typescript";

const SRC_DIR = path.resolve(__dirname, "../../src");

// src 配下の .ts をすべて取得（consts/ を含む）
function listSources(dir: string): string[] { /* 再帰的に .ts を列挙し、パス順に並べる */ }

/**
 * src/*.ts を GAS と同じく1つのグローバル空間に読み込み、
 * テストで使う関数を返す
 */
export function loadGas(globals: Record<string, unknown> = {}) {
  const code = listSources(SRC_DIR)
    .map((file) => ts.transpileModule(fs.readFileSync(file, "utf8"), {
      compilerOptions: { target: ts.ScriptTarget.ES2020 },
    }).outputText)
    .join("\n");
  const context = vm.createContext({ console, ...globals });
  return vm.runInContext(
    `${code}\n;({ doPost, normalizeStatements, toPlRow, toBsRow, toCfRow })`,
    context,
  );
}
```

#### `test/helpers/gasMocks.ts`（新規）

- `createSheetMock(lastRow)`: `getMaxRows` / `getRange(row, col, numRows?, numCols?)` / `getNextDataCell` / `getValue` / `getRow` / `setValues` を持つシートのモック。`setValues` に渡された値と、`getRange` の引数を記録する。
- `createGasGlobals(sheets)`: `SpreadsheetApp`（`openById` → `getSheetByName`、`Direction.UP`）、`PropertiesService`（`SS_ID` を返す）、`ContentService`（`createTextOutput` → `setMimeType`、`MimeType.JSON`）のモックを返す。

#### `test/fixtures.ts`（新規）

- `makeStatement(month, base)`: month と、各項目に `base` から連番の数値を入れた `FinancialStatement` を返す。どの行がどの決算か、テストで見分けやすくするため。

## 変更するコンポーネント（※）

| ファイル / コンポーネント | 変更種別 | 概要 |
|---|---|---|
| `src/models.ts` | 修正 | `RequestBody` に `statements?` を追加 |
| `src/statements.ts` | 新規 | `normalizeStatements` / `toPlRow` / `toBsRow` / `toCfRow` |
| `src/main.ts` | 修正 | `doPost` を配列で書き込む形に変更。before/after の分岐を削除 |
| `package.json` / `package-lock.json` | 修正 | jest, ts-jest, @types/jest の追加、`test` スクリプトの追加 |
| `jest.config.cjs` | 新規 | Jest の設定 |
| `tsconfig.test.json` | 新規 | テスト用の TypeScript 設定 |
| `test/helpers/loadGas.ts` | 新規 | src を1つのグローバル空間に読み込むヘルパー |
| `test/helpers/gasMocks.ts` | 新規 | GAS のグローバルのモック |
| `test/fixtures.ts` | 新規 | テスト用の決算データ |
| `test/statements.test.ts` | 新規 | 単体テスト |
| `test/doPost.test.ts` | 新規 | 結合テスト |

## 影響範囲の分析（※）

- **`doPost` の呼び出し元**: 外部から Web アプリの URL に POST される（送信側の仕組み）。旧形式 `{ before, after }` / `{ after }` はこれまでどおり受け付けるため、送信側を改修しなくても今の挙動は変わらない。
- **エラー時の挙動**: `after` がないときのエラー文言「ボディに今期の決算が含まれていません。」は変えない。新しく増えるエラーは `statements` を使ったときだけ。
- **シートへの書き込み**: 書き込む列の範囲・並びは変えない（`toXxRow` は今の `setValues` の中身をそのまま移したもの）。行数が件数分になるだけ。
- **件数が多いとき**: 1シートにつき `setValues` 1回でまとめて書き込むので、件数が増えても API の呼び出し回数は増えない。
- **`onOpen` / `setSpreadSheetName` / `resetSpreadSheetName` / `api.ts`**: 変更なし。影響なし。
- **ビルド（`npm run build`）**: `tsconfig.json` の `include` は `src` のままなので、`test/` は `dist` に出力されない。`src/statements.ts` は `dist/statements.js` として出力され、`clasp push` で他のファイルと同じグローバル空間に読み込まれる。`@types/jest` の型がビルド時にも見えるようになるが、出力には影響しない。
- **既存テスト**: なし。

## データ構造の変更

### リクエストボディ

変更前：

```json
{ "before": { "month": "2024/03", "pl": {...}, "bs": {...}, "cf": {...} },
  "after":  { "month": "2025/03", "pl": {...}, "bs": {...}, "cf": {...} } }
```

変更後（新形式。旧形式もそのまま使える）：

```json
{ "statements": [
    { "month": "2025/03", "pl": {...}, "bs": {...}, "cf": {...} },
    { "month": "2023/03", "pl": {...}, "bs": {...}, "cf": {...} },
    { "month": "2024/03", "pl": {...}, "bs": {...}, "cf": {...} }
] }
```

→ シートには 2023/03, 2024/03, 2025/03 の順で追記される。

### レスポンス

変更なし（`{ status: "ok", received: <受け取ったボディ> }`）。
