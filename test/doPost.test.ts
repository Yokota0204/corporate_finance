import { loadGas, Gas } from "./helpers/loadGas";
import { createGasGlobals, createPostEvent, createSheetMock, SheetMock } from "./helpers/gasMocks";
import { makeStatement } from "./fixtures";

describe("doPost", () => {
  let pl: SheetMock;
  let bs: SheetMock;
  let cf: SheetMock;
  let gas: Gas;

  beforeEach(() => {
    pl = createSheetMock("損益計算", 10);
    bs = createSheetMock("資産", 20);
    cf = createSheetMock("キャッシュフロー", 30);
    gas = loadGas(createGasGlobals([pl, bs, cf]));
  });

  const writtenMonths = (sheet: SheetMock) =>
    sheet.written.flatMap((w) => w.values.map((row) => row[0]));

  test("I-01 statementsの3件を各シートの最終行の次から書き込む", () => {
    const statements = [
      makeStatement("2023/03", 1),
      makeStatement("2024/03", 100),
      makeStatement("2025/03", 200),
    ];
    gas.doPost(createPostEvent({ statements }));

    expect(pl.written).toEqual([
      { row: 11, column: 1, numRows: 3, numColumns: 9, values: statements.map(gas.toPlRow) },
    ]);
    expect(bs.written).toEqual([
      { row: 21, column: 1, numRows: 3, numColumns: 12, values: statements.map(gas.toBsRow) },
    ]);
    expect(cf.written).toEqual([
      { row: 31, column: 1, numRows: 3, numColumns: 13, values: statements.map(gas.toCfRow) },
    ]);
  });

  test("I-02 順不同のstatements5件を古い順に書き込む", () => {
    const statements = ["2025/03", "2021/03", "2023/03", "2024/03", "2022/03"].map((m) =>
      makeStatement(m),
    );
    gas.doPost(createPostEvent({ statements }));

    const expected = ["2021/03", "2022/03", "2023/03", "2024/03", "2025/03"];
    expect(writtenMonths(pl)).toEqual(expected);
    expect(writtenMonths(bs)).toEqual(expected);
    expect(writtenMonths(cf)).toEqual(expected);
  });

  test("I-03 statementsが1件なら各シートに1行だけ書き込む", () => {
    gas.doPost(createPostEvent({ statements: [makeStatement("2025/03")] }));

    for (const sheet of [pl, bs, cf]) {
      expect(sheet.written).toHaveLength(1);
      expect(sheet.written[0].numRows).toBe(1);
      expect(writtenMonths(sheet)).toEqual(["2025/03"]);
    }
  });

  test("I-04 旧形式のbeforeとafterはbefore・afterの順で2行書き込む", () => {
    const before = makeStatement("2024/03", 1);
    const after = makeStatement("2025/03", 100);
    gas.doPost(createPostEvent({ before, after }));

    expect(pl.written[0].values).toEqual([gas.toPlRow(before), gas.toPlRow(after)]);
    expect(bs.written[0].values).toEqual([gas.toBsRow(before), gas.toBsRow(after)]);
    expect(cf.written[0].values).toEqual([gas.toCfRow(before), gas.toCfRow(after)]);
  });

  test("I-05 旧形式のafterのみなら1行だけ書き込む", () => {
    const after = makeStatement("2025/03");
    gas.doPost(createPostEvent({ after }));

    expect(pl.written[0].values).toEqual([gas.toPlRow(after)]);
    expect(bs.written[0].values).toEqual([gas.toBsRow(after)]);
    expect(cf.written[0].values).toEqual([gas.toCfRow(after)]);
  });

  test("I-06 statementsが空配列ならエラーでどのシートにも書き込まない", () => {
    expect(() => gas.doPost(createPostEvent({ statements: [] }))).toThrow(
      "statementsに決算が含まれていません。",
    );
    for (const sheet of [pl, bs, cf]) {
      expect(sheet.written).toHaveLength(0);
    }
  });

  test("I-07 statementsとafterを同時に指定するとエラーでどのシートにも書き込まない", () => {
    const body = { statements: [makeStatement("2025/03")], after: makeStatement("2025/03") };
    expect(() => gas.doPost(createPostEvent(body))).toThrow(
      "statementsとbefore/afterは同時に指定できません。",
    );
    for (const sheet of [pl, bs, cf]) {
      expect(sheet.written).toHaveLength(0);
    }
  });

  test("I-08 受け取ったボディをstatus okとともにJSONで返す", () => {
    const body = {
      statements: [makeStatement("2023/03"), makeStatement("2024/03"), makeStatement("2025/03")],
    };
    // ContentService はモックなので、モックが持つプロパティで確認する
    const output = gas.doPost(createPostEvent(body)) as unknown as { content: string; mimeType: string };

    expect(JSON.parse(output.content)).toEqual({ status: "ok", received: body });
    expect(output.mimeType).toBe("JSON");
  });
});
