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
