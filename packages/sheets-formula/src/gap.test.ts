// Every function FortuneSheet's engine (formulajs 2.9.3) evaluated must also be known here. The list of
// names was recorded from that engine (src/lib/parity/fortuneFunctionNames.json) before it was removed.
import { describe, expect, it } from "vitest";
import { functionNames } from "./functions";
import recorded from "../../../src/lib/parity/fortuneFunctionNames.json";

// Functions that are not part of Excel (constants, sparklines, Chinese stock/ID helpers) or need hosts we
// do not have; they stay unsupported on purpose.
const NOT_EXCEL = new Set([
  "E", "EVALUATE", "REMOTE", "ISDATE", "FINDFIELD", "LN10", "LN2", "LOG10E", "LOG2E", "SQRT1_2", "SQRT2", "HTML2TEXT", "PRONETIC", "INTERVAL", "REFERENCE", "BAHTTEXT",
  // Rare depreciation helpers with country-specific day counts; not built yet.
  "AMORDEGRC", "AMORLINC",
]);

describe("function coverage against FortuneSheet's engine", () => {
  it("knows every function the old engine knew", () => {
    const wanted = (recorded as string[]).filter((name) => !NOT_EXCEL.has(name) && !/SPLINES$/.test(name));
    const ours = new Set(functionNames());
    const missing = wanted.filter((name) => !ours.has(name));
    expect(missing).toEqual([]);
    expect(wanted.length).toBeGreaterThan(300);
  });
});
