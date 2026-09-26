export type SheetMock = {
  name: string;
  lastRow: number;
  // getRange(row, column, numRows, numColumns) の引数と setValues に渡された値の記録
  written: { row: number; column: number; numRows: number; numColumns: number; values: unknown[][] }[];
  getName: () => string;
  getMaxRows: () => number;
  getRange: (row: number, column: number, numRows?: number, numColumns?: number) => unknown;
};

const MAX_ROWS = 1000;

/**
 * A列の最終行が lastRow のシートのモック
 */
export function createSheetMock(name: string, lastRow: number): SheetMock {
  const sheet: SheetMock = {
    name,
    lastRow,
    written: [],
    getName: () => name,
    getMaxRows: () => MAX_ROWS,
    getRange: (row, column, numRows = 1, numColumns = 1) => ({
      // 最下行から上方向に探すと、最終行のセルが返る
      getNextDataCell: () => ({
        getValue: () => (sheet.lastRow > 0 ? "データ" : ""),
        getRow: () => sheet.lastRow,
      }),
      setValues: (values: unknown[][]) => {
        sheet.written.push({ row, column, numRows, numColumns, values });
      },
    }),
  };
  return sheet;
}

/**
 * doPost が使う GAS のグローバル（SpreadsheetApp など）のモック
 */
export function createGasGlobals(sheets: SheetMock[]): Record<string, unknown> {
  return {
    SpreadsheetApp: {
      Direction: { UP: "UP" },
      openById: () => ({
        getSheetByName: (name: string) => sheets.find((s) => s.name === name) ?? null,
      }),
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key: string) => (key === "SS_ID" ? "dummy" : null),
      }),
    },
    ContentService: {
      MimeType: { JSON: "JSON" },
      createTextOutput: (content: string) => {
        const output = {
          content,
          mimeType: undefined as string | undefined,
          setMimeType: (mimeType: string) => {
            output.mimeType = mimeType;
            return output;
          },
        };
        return output;
      },
    },
  };
}

/**
 * doPost に渡すイベントオブジェクト
 */
export function createPostEvent(body: unknown): GoogleAppsScript.Events.DoPost {
  return { postData: { contents: JSON.stringify(body) } } as unknown as GoogleAppsScript.Events.DoPost;
}
