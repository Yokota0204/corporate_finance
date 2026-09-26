import { loadGas } from "./helpers/loadGas";
import { makeStatement } from "./fixtures";

const gas = loadGas();

const months = (statements: FinancialStatement[]) => statements.map((s) => s.month);

describe("normalizeStatements", () => {
  test("U-01 古い順のstatementsはそのままの順で返す", () => {
    const body = {
      statements: [makeStatement("2023/03"), makeStatement("2024/03"), makeStatement("2025/03")],
    };
    expect(months(gas.normalizeStatements(body))).toEqual(["2023/03", "2024/03", "2025/03"]);
  });

  test("U-02 新しい順のstatementsは古い順に並べ替えて返す", () => {
    const body = {
      statements: [makeStatement("2025/03"), makeStatement("2024/03"), makeStatement("2023/03")],
    };
    expect(months(gas.normalizeStatements(body))).toEqual(["2023/03", "2024/03", "2025/03"]);
  });

  test("U-03 順不同のstatementsは古い順に並べ替えて返す", () => {
    const body = {
      statements: [
        makeStatement("2024/03"),
        makeStatement("2025/03"),
        makeStatement("2022/03"),
        makeStatement("2023/03"),
      ],
    };
    expect(months(gas.normalizeStatements(body))).toEqual([
      "2022/03",
      "2023/03",
      "2024/03",
      "2025/03",
    ]);
  });

  test("U-04 並べ替えても入力の配列の順番は変わらない", () => {
    const statements = [makeStatement("2025/03"), makeStatement("2024/03")];
    gas.normalizeStatements({ statements });
    expect(months(statements)).toEqual(["2025/03", "2024/03"]);
  });

  test("U-05 statementsが1件なら1件の配列を返す", () => {
    const body = { statements: [makeStatement("2024/03")] };
    expect(months(gas.normalizeStatements(body))).toEqual(["2024/03"]);
  });

  test("U-06 同じmonthが複数あっても除外せず入力の順を保って返す", () => {
    const first = makeStatement("2024/03", 1);
    const second = makeStatement("2024/03", 100);
    const body = { statements: [first, makeStatement("2023/03"), second] };
    const result = gas.normalizeStatements(body);
    expect(months(result)).toEqual(["2023/03", "2024/03", "2024/03"]);
    expect(result[1].pl.sales).toBe(1);
    expect(result[2].pl.sales).toBe(100);
  });

  test("U-07 旧形式のbeforeとafterはbefore・afterの順で返す", () => {
    const body = { before: makeStatement("2024/03"), after: makeStatement("2025/03") };
    expect(months(gas.normalizeStatements(body))).toEqual(["2024/03", "2025/03"]);
  });

  test("U-08 旧形式はmonthが逆でも並べ替えない", () => {
    const body = { before: makeStatement("2025/03"), after: makeStatement("2024/03") };
    expect(months(gas.normalizeStatements(body))).toEqual(["2025/03", "2024/03"]);
  });

  test("U-09 旧形式のafterのみならafterだけを返す", () => {
    const body = { after: makeStatement("2025/03") };
    expect(months(gas.normalizeStatements(body))).toEqual(["2025/03"]);
  });

  test("U-10 statementsが空配列ならエラー", () => {
    expect(() => gas.normalizeStatements({ statements: [] })).toThrow(
      "statementsに決算が含まれていません。",
    );
  });

  test("U-11 statementsもafterもなければエラー", () => {
    expect(() => gas.normalizeStatements({})).toThrow("ボディに今期の決算が含まれていません。");
  });

  test("U-12 beforeだけでafterがなければエラー", () => {
    expect(() => gas.normalizeStatements({ before: makeStatement("2024/03") })).toThrow(
      "ボディに今期の決算が含まれていません。",
    );
  });

  test("U-13 statementsとafterを同時に指定するとエラー", () => {
    const body = { statements: [makeStatement("2025/03")], after: makeStatement("2025/03") };
    expect(() => gas.normalizeStatements(body)).toThrow(
      "statementsとbefore/afterは同時に指定できません。",
    );
  });

  test("U-14 statementsとbeforeを同時に指定するとエラー", () => {
    const body = { statements: [makeStatement("2025/03")], before: makeStatement("2024/03") };
    expect(() => gas.normalizeStatements(body)).toThrow(
      "statementsとbefore/afterは同時に指定できません。",
    );
  });

  test("U-15 statementsが配列でなければエラー", () => {
    expect(() => gas.normalizeStatements({ statements: "2025/03" } as unknown as RequestBody)).toThrow(
      "statementsに決算が含まれていません。",
    );
  });
});

describe("toPlRow / toBsRow / toCfRow", () => {
  // makeStatement は pl → bs → cf の順に base から連番を振る
  const fs = makeStatement("2025/03", 1);

  test("U-16 toPlRowは月から純利益までを損益計算シートの列順で返す", () => {
    expect(gas.toPlRow(fs)).toEqual(["2025/03", 1, 2, 3, 4, 5, 6, 7, 8]);
  });

  test("U-17 toBsRowは月から利益剰余金までを資産シートの列順で返す", () => {
    expect(gas.toBsRow(fs)).toEqual(["2025/03", 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]);
  });

  test("U-18 toCfRowは月から配当金の支払額までをキャッシュフローシートの列順で返す", () => {
    expect(gas.toCfRow(fs)).toEqual([
      "2025/03", 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31,
    ]);
  });

  test("U-19 各行の要素数はシートに書き込む範囲の列数と一致する", () => {
    expect(gas.toPlRow(fs)).toHaveLength(gas.NET_COL_IN_PL_SHEET - gas.MONTH_COL_IN_PL_SHEET + 1);
    expect(gas.toBsRow(fs)).toHaveLength(
      gas.RETAINED_EARNINGS_COL_IN_BS_SHEET - gas.MONTH_COL_IN_BS_SHEET + 1,
    );
    expect(gas.toCfRow(fs)).toHaveLength(
      gas.LAST_INPUT_COL_IN_CF_SHEET - gas.MONTH_COL_IN_CF_SHEET + 1,
    );
  });
});
