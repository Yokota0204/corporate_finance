function getLastRowInColumn(
  sheet: Sheet,
  column: number,
): number {
  const maxRow: number = sheet.getMaxRows();
  const bottomCell: Range = sheet.getRange(maxRow, column);
  const lastDataCell: Range = bottomCell.getNextDataCell(SpreadsheetApp.Direction.UP);

  // 完全に空なら row 1 のセルが返ってくることがあるので、空チェック
  if (lastDataCell.getValue() === "") {
    throw "損益計算シートの最終行の行数の取得に失敗しました。";
  }
  return lastDataCell.getRow();
}


/**
 * WebアプリとしてデプロイしたURLに対して
 * POSTリクエストが飛んできたときに呼ばれる関数
 */
function doPost(e: GoogleAppsScript.Events.DoPost) {
  // 生のリクエストボディ文字列
  const rawBody: string = e.postData.contents;

  // JSONならパースしてオブジェクト化
  const data: RequestBody = JSON.parse(rawBody);
  log("info", JSON.stringify(data));

  // 今期の決算が含まれない場合はエラー
  if (!data.after) {
    throw "ボディに今期の決算が含まれていません。";
  }

  // スプレッドシートのキーを取得
  const props: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();
  const ssId: string | null = props.getProperty("SS_ID");

  if (!ssId) {
    throw "スプレッドシートIDを取得できませんでした。";
  }

  // スプレッドシートを取得
  const ss: Spreadsheet = SpreadsheetApp.openById(ssId);

  // セットする範囲の行数
  const setRowCount: number = data.before ? 2 : 1;

  // 損益計算のシートを取得
  const plSh: Sheet | null = ss.getSheetByName("損益計算");
  if (!plSh) {
    throw "損益計算シートを取得できませんでした。";
  }

  // 貸借対照表のシートを取得
  const bsSh: Sheet | null = ss.getSheetByName("資産");

  // 資産シートが存在しない場合はエラー
  if (!bsSh) {
    throw "資産シートを取得できませんでした。";
  }

  // 損益計算書のシートを取得
  const cfSh: Sheet | null = ss.getSheetByName("キャッシュフロー");

  // キャッシュフローシートが存在しない場合はエラー
  if (!cfSh) {
    throw "キャッシュフローシートを取得できませんでした。";
  }

  // 損益計算シートのA列の最終行を取得
  const plLastRow: number = getLastRowInColumn(plSh, MONTH_COL_IN_PL_SHEET);
  log("info", `損益計算シートのA列の最終行: ${plLastRow}`);

  // 日付から純利益までのセルの範囲を取得
  const plRange: Range = plSh.getRange(
    plLastRow + 1,
    MONTH_COL_IN_PL_SHEET,
    setRowCount,
    NET_COL_IN_PL_SHEET - MONTH_COL_IN_PL_SHEET + 1,
  );

  // 資産シートのA列の最終行を取得
  const bsLastRow: number = getLastRowInColumn(bsSh, MONTH_COL_IN_BS_SHEET);
  log("info", `資産シートのA列の最終行: ${bsLastRow}`);

  // 日付から利益剰余金までのセルの範囲を取得
  const bsRange: Range = bsSh.getRange(
    bsLastRow + 1,
    MONTH_COL_IN_BS_SHEET,
    setRowCount,
    RETAINED_EARNINGS_COL_IN_BS_SHEET - MONTH_COL_IN_BS_SHEET + 1,
  );

  // キャッシュフローのA列の最終行を取得
  const cfLastRow: number = getLastRowInColumn(cfSh, MONTH_COL_IN_CF_SHEET);
  log("info", `キャッシュフローシートのA列の最終行: ${cfLastRow}`);

  // 日付から財務キャッシュフローまでのセルの範囲を取得
  const cfRange: Range = cfSh.getRange(
    cfLastRow + 1,
    MONTH_COL_IN_CF_SHEET,
    setRowCount,
    FINANCING_COL_IN_CF_SHEET - MONTH_COL_IN_CF_SHEET + 1,
  );

  const beforeFs: FinancialStatement | undefined = data.before;
  const afterFs: FinancialStatement = data.after;

  // レスポンスの数値をシートに出力
  if (beforeFs) {
    plRange.setValues([
      [
        beforeFs.month,
        beforeFs.pl.sales,
        beforeFs.pl.cost,
        beforeFs.pl.sellingExpenses,
        beforeFs.pl.nonOperatingIncome,
        beforeFs.pl.nonOperatingExpenses,
        beforeFs.pl.specialIncome,
        beforeFs.pl.specialLosses,
        beforeFs.pl.net,
      ],
      [
        afterFs.month,
        afterFs.pl.sales,
        afterFs.pl.cost,
        afterFs.pl.sellingExpenses,
        afterFs.pl.nonOperatingIncome,
        afterFs.pl.nonOperatingExpenses,
        afterFs.pl.specialIncome,
        afterFs.pl.specialLosses,
        afterFs.pl.net,
      ],
    ]);
    bsRange.setValues([
      [
        beforeFs.month,
        beforeFs.bs.cash,
        beforeFs.bs.notesAndAccountsReceivableTrade,
        beforeFs.bs.otherReceivables,
        beforeFs.bs.depositsPaid,
        beforeFs.bs.shortTermLoans,
        beforeFs.bs.allowance,
        beforeFs.bs.currentAssets,
        beforeFs.bs.nonCurrentAssets,
        beforeFs.bs.currentLiabilities,
        beforeFs.bs.nonCurrentLiabilities,
        beforeFs.bs.retainedEarnings,
      ],
      [
        afterFs.month,
        afterFs.bs.cash,
        afterFs.bs.notesAndAccountsReceivableTrade,
        afterFs.bs.otherReceivables,
        afterFs.bs.depositsPaid,
        afterFs.bs.shortTermLoans,
        afterFs.bs.allowance,
        afterFs.bs.currentAssets,
        afterFs.bs.nonCurrentAssets,
        afterFs.bs.currentLiabilities,
        afterFs.bs.nonCurrentLiabilities,
        afterFs.bs.retainedEarnings,
      ],
    ]);
    cfRange.setValues([
      [
        beforeFs.month,
        beforeFs.cf.operating,
        beforeFs.cf.investing,
        beforeFs.cf.financing,
      ],
      [
        afterFs.month,
        afterFs.cf.operating,
        afterFs.cf.investing,
        afterFs.cf.financing,
      ],
    ]);
  } else {
    plRange.setValues([
      [
        afterFs.month,
        afterFs.pl.sales,
        afterFs.pl.cost,
        afterFs.pl.sellingExpenses,
        afterFs.pl.nonOperatingIncome,
        afterFs.pl.nonOperatingExpenses,
        afterFs.pl.specialIncome,
        afterFs.pl.specialLosses,
        afterFs.pl.net,
      ],
    ]);
    bsRange.setValues([
      [
        afterFs.month,
        afterFs.bs.cash,
        afterFs.bs.notesAndAccountsReceivableTrade,
        afterFs.bs.otherReceivables,
        afterFs.bs.depositsPaid,
        afterFs.bs.shortTermLoans,
        afterFs.bs.allowance,
        afterFs.bs.currentAssets,
        afterFs.bs.nonCurrentAssets,
        afterFs.bs.currentLiabilities,
        afterFs.bs.nonCurrentLiabilities,
        afterFs.bs.retainedEarnings,
      ],
    ]);
    cfRange.setValues([
      [
        afterFs.month,
        afterFs.cf.operating,
        afterFs.cf.investing,
        afterFs.cf.financing,
      ],
    ]);
  }

  // 何か処理をしたあと、レスポンスを返す例
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      received: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
