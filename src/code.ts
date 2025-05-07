/**
 * WebアプリとしてデプロイしたURLに対して
 * POSTリクエストが飛んできたときに呼ばれる関数
 */
function doPost(e: GoogleAppsScript.Events.DoPost) {
  // 生のリクエストボディ文字列
  const rawBody: string = e.postData.contents;

  // JSONならパースしてオブジェクト化
  const data = JSON.parse(rawBody);
  log("info", data);

  // スプレッドシートのキーを取得
  const props: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();
  const ssId: string | null = props.getProperty("SS_ID");

  if (!ssId) {
    throw "スプレッドシートIDを取得できませんでした。";
  }

  // スプレッドシートを取得
  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ssId);

  // 損益計算書のシートを取得
  const plSh: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName("損益計算");

  // 貸借対照表のシートを取得
  const bsSh: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName("資産負債表");

  // 何か処理をしたあと、レスポンスを返す例
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      received: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
