/**
 * month と、各項目に base からの連番を入れた決算を作る
 * （どの行がどの決算か、テストで見分けやすくするため）
 */
export function makeStatement(month: string, base = 1): FinancialStatement {
  let n = base;
  const next = () => n++;
  return {
    month,
    pl: {
      sales: next(),
      cost: next(),
      sellingExpenses: next(),
      nonOperatingIncome: next(),
      nonOperatingExpenses: next(),
      specialIncome: next(),
      specialLosses: next(),
      net: next(),
    },
    bs: {
      cash: next(),
      notesAndAccountsReceivableTrade: next(),
      otherReceivables: next(),
      depositsPaid: next(),
      shortTermLoans: next(),
      allowance: next(),
      currentAssets: next(),
      nonCurrentAssets: next(),
      currentLiabilities: next(),
      nonCurrentLiabilities: next(),
      retainedEarnings: next(),
    },
    cf: {
      operating: next(),
      investing: next(),
      financing: next(),
      proceedsFromShortTermBorrowings: next(),
      proceedsFromLongTermBorrowings: next(),
      proceedsFromIssuanceOfBonds: next(),
      repaymentsOfShortTermBorrowings: next(),
      repaymentsOfLongTermBorrowings: next(),
      repaymentsOfBonds: next(),
      purchaseOfTreasuryStock: next(),
      retirementOfTreasuryStock: next(),
      dividendsPaid: next(),
    },
  };
}
