import { describe, expect, it } from "vitest";
import { formatFor, matchesCondition, withoutArea } from "./bigConditional";

const rule = (type: string, values: string[] = [], r = { r1: 0, c1: 0, r2: 9, c2: 3 }) => ({
  ranges: [r],
  condition: { type, values },
  format: { cellColor: "#ffcccc" },
});

describe("conditional formatting", () => {
  it("compares numbers, text and empties", () => {
    const cell = (value: string | number) => ({ value });
    expect(matchesCondition({ type: "gt", values: ["5"] }, cell(6))).toBe(true);
    expect(matchesCondition({ type: "gt", values: ["5"] }, cell(5))).toBe(false);
    expect(matchesCondition({ type: "gte", values: ["5"] }, cell(5))).toBe(true);
    expect(matchesCondition({ type: "lt", values: ["5"] }, cell("4"))).toBe(true); // text that reads as a number
    expect(matchesCondition({ type: "gt", values: ["5"] }, cell("abc"))).toBe(false);
    expect(matchesCondition({ type: "between", values: ["10", "1"] }, cell(5))).toBe(true);
    expect(matchesCondition({ type: "between", values: ["1", "3"] }, cell(5))).toBe(false);
    expect(matchesCondition({ type: "eq", values: ["Done"] }, cell("done"))).toBe(true);
    expect(matchesCondition({ type: "neq", values: ["Done"] }, cell("open"))).toBe(true);
    expect(matchesCondition({ type: "contains", values: ["ERR"] }, cell("an error"))).toBe(true);
    expect(matchesCondition({ type: "empty", values: [] }, undefined)).toBe(true);
    expect(matchesCondition({ type: "empty", values: [] }, cell(""))).toBe(true);
    expect(matchesCondition({ type: "notEmpty", values: [] }, cell(0))).toBe(true);
    expect(matchesCondition({ type: "gt", values: ["1"] }, null)).toBe(false);
  });

  it("uses the first matching rule that covers the cell", () => {
    const rules = [rule("gt", ["100"], { r1: 0, c1: 0, r2: 4, c2: 0 }), { ...rule("gt", ["0"]), format: { textColor: "#0000ff" } }];
    expect(formatFor(rules, 1, 0, { value: 500 })).toEqual({ cellColor: "#ffcccc" });
    expect(formatFor(rules, 1, 0, { value: 50 })).toEqual({ textColor: "#0000ff" });
    expect(formatFor(rules, 7, 0, { value: 500 })).toEqual({ textColor: "#0000ff" }); // outside the first rule's range
    expect(formatFor(rules, 20, 0, { value: 500 })).toBeUndefined();
    expect(formatFor(undefined, 0, 0, { value: 1 })).toBeUndefined();
  });

  it("takes an area out of a rule's ranges, keeping the rest", () => {
    const rules = [rule("gt", ["1"], { r1: 0, c1: 0, r2: 9, c2: 3 })];
    const left = withoutArea(rules, { r1: 3, c1: 1, r2: 5, c2: 2 });
    expect(left).toHaveLength(1);
    const covered = (r: number, c: number) => left[0].ranges.some((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2);
    expect(covered(4, 1)).toBe(false);
    expect(covered(4, 0)).toBe(true);
    expect(covered(4, 3)).toBe(true);
    expect(covered(2, 1)).toBe(true);
    expect(covered(6, 2)).toBe(true);
    expect(withoutArea(rules, { r1: 0, c1: 0, r2: 20, c2: 20 })).toEqual([]);
  });
});
