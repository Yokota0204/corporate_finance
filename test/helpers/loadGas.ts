import * as fs from "fs";
import * as path from "path";
import * as vm from "vm";
import * as ts from "typescript";

const SRC_DIR = path.resolve(__dirname, "../../src");

// テストから参照する src のグローバルな関数・定数
export type Gas = {
  doPost: typeof doPost;
  normalizeStatements: typeof normalizeStatements;
  toPlRow: typeof toPlRow;
  toBsRow: typeof toBsRow;
  toCfRow: typeof toCfRow;
  MONTH_COL_IN_PL_SHEET: number;
  NET_COL_IN_PL_SHEET: number;
  MONTH_COL_IN_BS_SHEET: number;
  RETAINED_EARNINGS_COL_IN_BS_SHEET: number;
  MONTH_COL_IN_CF_SHEET: number;
  LAST_INPUT_COL_IN_CF_SHEET: number;
};

const EXPOSED_NAMES: (keyof Gas)[] = [
  "doPost",
  "normalizeStatements",
  "toPlRow",
  "toBsRow",
  "toCfRow",
  "MONTH_COL_IN_PL_SHEET",
  "NET_COL_IN_PL_SHEET",
  "MONTH_COL_IN_BS_SHEET",
  "RETAINED_EARNINGS_COL_IN_BS_SHEET",
  "MONTH_COL_IN_CF_SHEET",
  "LAST_INPUT_COL_IN_CF_SHEET",
];

// src 配下の .ts をすべて取得（consts/ を含む）
function listSources(dir: string): string[] {
  return fs
    .readdirSync(dir, { withFileTypes: true })
    .flatMap((entry) => {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        return listSources(fullPath);
      }
      return entry.name.endsWith(".ts") ? [fullPath] : [];
    })
    .sort();
}

/**
 * src/*.ts を GAS と同じく1つのグローバル空間に読み込み、
 * テストで使う関数・定数を返す
 */
export function loadGas(globals: Record<string, unknown> = {}): Gas {
  const code = listSources(SRC_DIR)
    .map(
      (file) =>
        ts.transpileModule(fs.readFileSync(file, "utf8"), {
          compilerOptions: { target: ts.ScriptTarget.ES2020 },
        }).outputText,
    )
    .join("\n");

  // 未実装の名前があっても読み込み自体は失敗しないように typeof で取り出す
  const exposed = EXPOSED_NAMES.map(
    (name) => `${name}: typeof ${name} === "undefined" ? undefined : ${name}`,
  ).join(",\n");

  const context = vm.createContext({ console, ...globals });
  return vm.runInContext(`${code}\n;({ ${exposed} })`, context);
}
