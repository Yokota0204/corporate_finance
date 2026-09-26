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
 * シートの月の列の最終行の次の行から、rows を書き込む範囲を取得
 */
function getAppendRange(
  sheet: Sheet,
  monthCol: number,
  rows: (string | number)[][],
): Range {
  const lastRow: number = getLastRowInColumn(sheet, monthCol);
  log("info", `${sheet.getName()}シートのA列の最終行: ${lastRow}`);
  return sheet.getRange(lastRow + 1, monthCol, rows.length, rows[0].length);
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

  // 書き込む決算を古い順に取得（入力が不正な場合はここでエラー）
  const statements: FinancialStatement[] = normalizeStatements(data);
  log("info", `書き込む決算の件数: ${statements.length}`);

  // スプレッドシートのキーを取得
  const props: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();
  const ssId: string | null = props.getProperty("SS_ID");

  if (!ssId) {
    throw "スプレッドシートIDを取得できませんでした。";
  }

  // スプレッドシートを取得
  const ss: Spreadsheet = SpreadsheetApp.openById(ssId);

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

  // 各シートに書き込む行
  const plRows: (string | number)[][] = statements.map(toPlRow);
  const bsRows: (string | number)[][] = statements.map(toBsRow);
  const cfRows: (string | number)[][] = statements.map(toCfRow);

  // 一部のシートだけに書き込まれないよう、先にすべての範囲を取得してから書き込む
  const plRange: Range = getAppendRange(plSh, MONTH_COL_IN_PL_SHEET, plRows);
  const bsRange: Range = getAppendRange(bsSh, MONTH_COL_IN_BS_SHEET, bsRows);
  const cfRange: Range = getAppendRange(cfSh, MONTH_COL_IN_CF_SHEET, cfRows);

  // 決算の数値をシートに出力
  plRange.setValues(plRows);
  bsRange.setValues(bsRows);
  cfRange.setValues(cfRows);

  // 何か処理をしたあと、レスポンスを返す例
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      received: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// メニューにGASを追加
function onOpen(): void {
  const ui: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();
  ui.createMenu('GAS')
    .addItem('ファイル名を変更', 'setSpreadSheetName')
    .addItem('ファイル名をもとに戻す', 'resetSpreadSheetName')
    .addToUi();
}
