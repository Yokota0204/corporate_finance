type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
type Range = GoogleAppsScript.Spreadsheet.Range;

type ProfitAndLoss = {
  sales: number;
  cost: number;
  sellingExpenses: number;
  nonOperatingIncome: number;
  nonOperatingExpenses: number;
  specialIncome: number;
  specialLosses: number;
  net: number;
};
type BalanceSheet = {
  cash: number;
  notesAndAccountsReceivableTrade: number;
  otherReceivables: number;
  depositsPaid: number;
  shortTermLoans: number;
  allowance: number;
  currentAssets: number;
  nonCurrentAssets: number;
  currentLiabilities: number;
  nonCurrentLiabilities: number;
  retainedEarnings: number;
};
type FinancialStatement = {
  month: string;
  pl: ProfitAndLoss;
  bs: BalanceSheet;
};
type RequestBody = {
  before?: FinancialStatement;
  after?: FinancialStatement;
};
