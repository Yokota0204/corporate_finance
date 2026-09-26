import { loadGas } from "./helpers/loadGas";

type FetchCall = { url: string; params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions };

/**
 * J-Quants API の呼び出しに使う GAS のグローバル（PropertiesService・UrlFetchApp）のモック
 */
function createApiGlobals(properties: Record<string, string>) {
  const calls: FetchCall[] = [];
  const globals = {
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key: string) => properties[key] ?? null,
      }),
    },
    UrlFetchApp: {
      fetch: (url: string, params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions) => {
        calls.push({ url, params });
        return { getContentText: () => JSON.stringify({ data: [] }) };
      },
    },
  };
  return { globals, calls };
}

describe("J-Quants APIのトークン", () => {
  test("A-01 getCompanyByCodeはスクリプトプロパティJ_QUANTS_TOKENの値をx-api-keyに設定する", () => {
    const { globals, calls } = createApiGlobals({ J_QUANTS_TOKEN: "token-from-props" });
    const gas = loadGas(globals);

    gas.getCompanyByCode("7203");

    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe("https://api.jquants.com/v2/equities/master?code=7203");
    expect(calls[0].params.headers).toMatchObject({ "x-api-key": "token-from-props" });
  });

  test("A-02 getFinanceStatementはスクリプトプロパティJ_QUANTS_TOKENの値をx-api-keyに設定する", () => {
    const { globals, calls } = createApiGlobals({ J_QUANTS_TOKEN: "token-from-props" });
    const gas = loadGas(globals);

    gas.getFinanceStatement("7203");

    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe("https://api.jquants.com/v2/fins/summary?code=7203");
    expect(calls[0].params.headers).toMatchObject({ "x-api-key": "token-from-props" });
  });

  test("A-03 J_QUANTS_TOKENが未設定ならエラーにしてAPIを呼ばない", () => {
    const { globals, calls } = createApiGlobals({});
    const gas = loadGas(globals);

    expect(() => gas.getCompanyByCode("7203")).toThrow("J-Quantsのトークンを取得できませんでした。");
    expect(() => gas.getFinanceStatement("7203")).toThrow("J-Quantsのトークンを取得できませんでした。");
    expect(calls).toHaveLength(0);
  });
});
