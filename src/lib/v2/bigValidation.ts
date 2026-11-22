// Data validation rules for cells of big sheets: what a cell accepts, checked when someone types into it.
import type { WorkbookValidation } from "../workbook";

export type Operator = "between" | "notBetween" | "eq" | "neq" | "gt" | "gte" | "lt" | "lte";

export const RULE_TYPES = [
  { value: "list", label: "Dropdown list" },
  { value: "number", label: "Number" },
  { value: "integer", label: "Whole number" },
  { value: "length", label: "Text length" },
] as const;

export const OPERATORS: { value: Operator; label: string }[] = [
  { value: "between", label: "is between" },
  { value: "notBetween", label: "is not between" },
  { value: "eq", label: "is equal to" },
  { value: "neq", label: "is not equal to" },
  { value: "gt", label: "is greater than" },
  { value: "gte", label: "is at least" },
  { value: "lt", label: "is less than" },
  { value: "lte", label: "is at most" },
];

function compare(value: number, op: Operator, a: number, b: number): boolean {
  switch (op) {
    case "between":
      return value >= Math.min(a, b) && value <= Math.max(a, b);
    case "notBetween":
      return value < Math.min(a, b) || value > Math.max(a, b);
    case "eq":
      return value === a;
    case "neq":
      return value !== a;
    case "gt":
      return value > a;
    case "gte":
      return value >= a;
    case "lt":
      return value < a;
    case "lte":
      return value <= a;
  }
}

/** Says what is wrong with `text` for this rule, or returns null when it is fine (an empty cell is always fine). */
export function validate(rule: WorkbookValidation, text: string): string | null {
  const value = text.trim();
  if (value === "") return null;
  if (rule.type === "list") {
    const options = rule.options ?? [];
    return options.includes(value) ? null : `Pick one of: ${options.join(", ")}.`;
  }
  const op = (rule.params?.op as Operator | undefined) ?? "between";
  const a = Number(rule.params?.a ?? 0);
  const b = Number(rule.params?.b ?? 0);
  const phrase = OPERATORS.find((item) => item.value === op)?.label ?? op;
  const range = op === "between" || op === "notBetween" ? `${a} and ${b}` : String(a);
  if (rule.type === "length") {
    return compare(text.length, op, a, b) ? null : `The text length must be ${phrase.replace(/^is /, "")} ${range}.`;
  }
  const n = Number(value.replace(/,/g, ""));
  if (!Number.isFinite(n)) return "Enter a number.";
  if (rule.type === "integer" && !Number.isInteger(n)) return "Enter a whole number.";
  return compare(n, op, a, b) ? null : `The value must be ${phrase.replace(/^is /, "")} ${range}.`;
}

/** Turns what the dialog collected into a rule (or an error message when something is missing). */
export function buildRule(
  type: string,
  input: { list: string; op: Operator; a: string; b: string },
): WorkbookValidation | string {
  if (type === "list") {
    const options = input.list
      .split(/[\n,]/)
      .map((item) => item.trim())
      .filter(Boolean);
    if (options.length === 0) return "Add at least one option, separated by commas.";
    return { type, options: [...new Set(options)] };
  }
  const a = Number(input.a);
  const b = Number(input.b);
  if (input.a.trim() === "" || !Number.isFinite(a)) return "Enter a number for the limit.";
  if ((input.op === "between" || input.op === "notBetween") && (input.b.trim() === "" || !Number.isFinite(b))) {
    return "Enter both limits.";
  }
  return { type, params: { op: input.op, a: String(a), b: String(Number.isFinite(b) ? b : 0) } };
}
