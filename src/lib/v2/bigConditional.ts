// Conditional formatting for big sheets: rules that colour a cell when its value matches.
// A rule is stored with the tab's look (see TabLayout.conditionalFormats) and checked while drawing,
// so it works on any size of sheet without touching the cells.
import type { WorkbookCell, WorkbookConditionalRule } from "../workbook";
import { cellDisplay } from "../workbook/format";

export const CONDITIONS = [
  { value: "gt", label: "Greater than", values: 1 },
  { value: "gte", label: "Greater than or equal to", values: 1 },
  { value: "lt", label: "Less than", values: 1 },
  { value: "lte", label: "Less than or equal to", values: 1 },
  { value: "eq", label: "Equal to", values: 1 },
  { value: "neq", label: "Not equal to", values: 1 },
  { value: "between", label: "Between", values: 2 },
  { value: "contains", label: "Text contains", values: 1 },
  { value: "empty", label: "Is empty", values: 0 },
  { value: "notEmpty", label: "Is not empty", values: 0 },
] as const;

/** Does the cell satisfy the rule's condition? */
export function matchesCondition(condition: WorkbookConditionalRule["condition"], cell: WorkbookCell | null | undefined): boolean {
  const empty = !cell || cell.value === undefined || cell.value === "";
  if (condition.type === "empty") return empty;
  if (condition.type === "notEmpty") return !empty;
  if (empty) return false;
  const text = cellDisplay(cell!);
  const [first = "", second = ""] = condition.values;
  if (condition.type === "contains") return text.toLowerCase().includes(first.toLowerCase());
  const number = typeof cell!.value === "number" ? cell!.value : Number(String(cell!.value).replace(/,/g, ""));
  const a = Number(first);
  const b = Number(second);
  const numeric = Number.isFinite(number) && Number.isFinite(a) && (condition.type !== "between" || Number.isFinite(b));
  switch (condition.type) {
    case "eq":
      return numeric ? number === a : text.toLowerCase() === first.toLowerCase();
    case "neq":
      return numeric ? number !== a : text.toLowerCase() !== first.toLowerCase();
    case "gt":
      return numeric && number > a;
    case "gte":
      return numeric && number >= a;
    case "lt":
      return numeric && number < a;
    case "lte":
      return numeric && number <= a;
    case "between":
      return numeric && number >= Math.min(a, b) && number <= Math.max(a, b);
    default:
      return false;
  }
}

/** The colours the first matching rule gives a cell, or undefined. */
export function formatFor(
  rules: readonly WorkbookConditionalRule[] | undefined,
  row: number,
  col: number,
  cell: WorkbookCell | null | undefined,
): WorkbookConditionalRule["format"] | undefined {
  if (!rules) return undefined;
  for (const rule of rules) {
    if (!rule.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2)) continue;
    if (matchesCondition(rule.condition, cell)) return rule.format;
  }
  return undefined;
}

/** Rules with the part of `area` taken out of their ranges (used when a rule is replaced or removed). */
export function withoutArea(
  rules: readonly WorkbookConditionalRule[],
  area: { r1: number; c1: number; r2: number; c2: number },
): WorkbookConditionalRule[] {
  const out: WorkbookConditionalRule[] = [];
  for (const rule of rules) {
    const ranges: WorkbookConditionalRule["ranges"] = [];
    for (const r of rule.ranges) {
      const overlaps = r.r1 <= area.r2 && r.r2 >= area.r1 && r.c1 <= area.c2 && r.c2 >= area.c1;
      if (!overlaps) {
        ranges.push(r);
        continue;
      }
      // Keep what is left: the strips above, below, left and right of the removed area.
      if (r.r1 < area.r1) ranges.push({ ...r, r2: area.r1 - 1 });
      if (r.r2 > area.r2) ranges.push({ ...r, r1: area.r2 + 1 });
      const r1 = Math.max(r.r1, area.r1);
      const r2 = Math.min(r.r2, area.r2);
      if (r.c1 < area.c1) ranges.push({ r1, r2, c1: r.c1, c2: area.c1 - 1 });
      if (r.c2 > area.c2) ranges.push({ r1, r2, c1: area.c2 + 1, c2: r.c2 });
    }
    if (ranges.length > 0) out.push({ ...rule, ranges });
  }
  return out;
}
