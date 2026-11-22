import { describe, expect, it } from "vitest";
import { buildRule, validate } from "./bigValidation";

describe("data validation", () => {
  it("checks a dropdown list and lets empty cells through", () => {
    const rule = { type: "list", options: ["Yes", "No"] };
    expect(validate(rule, "Yes")).toBeNull();
    expect(validate(rule, "Maybe")).toMatch(/Yes, No/);
    expect(validate(rule, "  ")).toBeNull();
  });

  it("checks numbers, whole numbers and text length with every operator", () => {
    const num = (op: string, a: number, b = 0) => ({ type: "number", params: { op, a: String(a), b: String(b) } });
    expect(validate(num("between", 1, 10), "5")).toBeNull();
    expect(validate(num("between", 10, 1), "5")).toBeNull(); // limits in either order
    expect(validate(num("between", 1, 10), "11")).toMatch(/between 1 and 10/);
    expect(validate(num("notBetween", 1, 10), "5")).not.toBeNull();
    expect(validate(num("gt", 3), "3")).not.toBeNull();
    expect(validate(num("gte", 3), "3")).toBeNull();
    expect(validate(num("lt", 3), "2.5")).toBeNull();
    expect(validate(num("eq", 3), "3")).toBeNull();
    expect(validate(num("neq", 3), "3")).not.toBeNull();
    expect(validate(num("gt", 3), "1,000")).toBeNull(); // thousands separators are fine
    expect(validate(num("gt", 3), "abc")).toBe("Enter a number.");
    expect(validate({ type: "integer", params: { op: "gte", a: "0" } }, "2.5")).toBe("Enter a whole number.");
    expect(validate({ type: "length", params: { op: "lte", a: "3" } }, "abcd")).toMatch(/at most 3/);
    expect(validate({ type: "length", params: { op: "lte", a: "3" } }, "abc")).toBeNull();
  });

  it("builds rules from the dialog and reports what is missing", () => {
    expect(buildRule("list", { list: "a, b\nb,c", op: "between", a: "", b: "" })).toEqual({ type: "list", options: ["a", "b", "c"] });
    expect(buildRule("list", { list: " , ", op: "between", a: "", b: "" })).toMatch(/at least one/);
    expect(buildRule("number", { list: "", op: "gt", a: "5", b: "" })).toEqual({ type: "number", params: { op: "gt", a: "5", b: "0" } });
    expect(buildRule("number", { list: "", op: "between", a: "1", b: "" })).toBe("Enter both limits.");
    expect(buildRule("number", { list: "", op: "gt", a: "x", b: "" })).toBe("Enter a number for the limit.");
  });
});
