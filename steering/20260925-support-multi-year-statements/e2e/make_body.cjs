/**
 * 動作確認用の POST ボディ（JSON）を標準出力に出す
 *
 * 使い方:
 *   node make_body.cjs statements 2025/03:500 2021/03:100
 *     → { "statements": [ 2025/03 の決算, 2021/03 の決算 ] }（指定した順のまま）
 *   node make_body.cjs legacy 2024/03:400 2025/03:500
 *     → { "before": 2024/03 の決算, "after": 2025/03 の決算 }
 *   node make_body.cjs legacy 2025/03:500
 *     → { "after": 2025/03 の決算 }
 *   node make_body.cjs empty
 *     → { "statements": [] }
 *   node make_body.cjs both 2025/03:500
 *     → { "statements": [2025/03 の決算], "after": 2025/03 の決算 }
 *
 * 「月:基準値」の基準値から、各項目に連番を入れる。
 *   損益計算: 売上高 = 基準値, 売上原価 = 基準値+1, ..., 純利益 = 基準値+7
 *   資産:     現金預金 = 基準値+8, ..., 利益剰余金 = 基準値+18
 *   CF:       営業CF = 基準値+19, ..., 配当金の支払額 = 基準値+30
 */
const PL = ["sales", "cost", "sellingExpenses", "nonOperatingIncome", "nonOperatingExpenses",
  "specialIncome", "specialLosses", "net"];
const BS = ["cash", "notesAndAccountsReceivableTrade", "otherReceivables", "depositsPaid",
  "shortTermLoans", "allowance", "currentAssets", "nonCurrentAssets", "currentLiabilities",
  "nonCurrentLiabilities", "retainedEarnings"];
const CF = ["operating", "investing", "financing", "proceedsFromShortTermBorrowings",
  "proceedsFromLongTermBorrowings", "proceedsFromIssuanceOfBonds",
  "repaymentsOfShortTermBorrowings", "repaymentsOfLongTermBorrowings", "repaymentsOfBonds",
  "purchaseOfTreasuryStock", "retirementOfTreasuryStock", "dividendsPaid"];

function makeStatement(arg) {
  const [month, baseStr] = arg.split(":");
  let n = Number(baseStr);
  const fill = (keys) => Object.fromEntries(keys.map((k) => [k, n++]));
  return { month, pl: fill(PL), bs: fill(BS), cf: fill(CF) };
}

const [mode, ...args] = process.argv.slice(2);
const statements = args.map(makeStatement);
const bodies = {
  statements: () => ({ statements }),
  legacy: () =>
    statements.length === 2
      ? { before: statements[0], after: statements[1] }
      : { after: statements[0] },
  empty: () => ({ statements: [] }),
  both: () => ({ statements, after: statements[0] }),
};
if (!bodies[mode]) {
  console.error("mode は statements / legacy / empty / both のいずれか");
  process.exit(1);
}
process.stdout.write(JSON.stringify(bodies[mode]()));
