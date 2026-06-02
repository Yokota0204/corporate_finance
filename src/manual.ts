/**
 * スプレッドシートのファイル名を変更
 */
const setSpreadSheetName = (): void => {
  // 設定シートのB1から銘柄コードを取得
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("設定");
  const code = settingSheet?.getRange("B1").getValue() as string;
  if (!code) {
    throw new Error("銘柄コードが設定されていません");
  }
  log("log", `code: ${code}`);

  // 銘柄名を取得
  const companyName: string = getCompanyName(code);

  // スプレッドシートのファイル名を変更
  SpreadsheetApp.getActiveSpreadsheet()
    .rename(`株｜決算短信推移｜${companyName}(${code}) v3.5.4｜月日`);
};

/**
 * ファイル名をもとに戻す
 */
const resetSpreadSheetName = (): void => {
  // 設定シートのB2からデフォルトのファイル名を取得
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("設定");
  const defaultName = settingSheet?.getRange("B2").getValue() as string;

  // スプレッドシートのファイル名をもとに戻す
  SpreadsheetApp
    .getActiveSpreadsheet()
    .rename(defaultName);
};
