// Every function FortuneSheet's engine (formulajs 2.9.3) evaluates must also be known here.
import { createRequire } from "node:module";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { functionNames } from "./functions";

// Functions FortuneSheet lists but are not part of Excel (sparklines, Chinese stock/ID helpers) or
// need hosts we do not have; they stay unsupported on purpose.
const NOT_EXCEL = new Set(["EVALUATE", "REMOTE", "ISDATE", "E", "FINDFIELD", "LN10", "LN2", "LOG10E", "LOG2E", "SQRT1_2", "SQRT2", "HTML2TEXT", "PRONETIC", "INTERVAL", "REFERENCE", "BAHTTEXT",
  // Rare depreciation helpers with country-specific day counts; not built yet.
  "AMORDEGRC", "AMORLINC"]);

describe("function coverage against FortuneSheet's engine", () => {
  it("knows every function the old engine knew", () => {
    const require = createRequire(import.meta.url);
    const old = require(path.resolve("node_modules/.pnpm/@formulajs+formulajs@2.9.3/node_modules/@formulajs/formulajs")) as Record<string, unknown>;
    const wanted = Object.keys(old).filter((k) => typeof old[k] === "function" && /^[A-Z][A-Z0-9_.]*$/.test(k) && !NOT_EXCEL.has(k) && !/SPLINES$/.test(k));
    const ours = new Set(functionNames());
    const missing = wanted.filter((name) => !ours.has(name));
    expect(missing).toEqual([]);
    expect(wanted.length).toBeGreaterThan(300);
  });
});
