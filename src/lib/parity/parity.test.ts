// Differential parity harness: the same scenario runs on FortuneSheet (headless) and on the
// candidate grid, and what each produced must match. One scenario per row of the parity matrix in
// the plan. DIVERGENCES documents places where FortuneSheet itself is wrong or odd and we choose not
// to copy it.
import { describe, expect, it } from "vitest";
import { normalizeWorkbook, WORKBOOK_SCHEMA, type WorkbookDocument } from "../workbook";
import { runOnFortuneSheet, type Step } from "./fortuneHeadless";
import { runOnStub } from "./stubGrid";

function sample(): WorkbookDocument {
  return normalizeWorkbook({
    schema: WORKBOOK_SCHEMA,
    version: 1,
    title: "Parity",
    activeSheetId: "s1",
    createdAt: 1,
    updatedAt: 1,
    sheets: [
      {
        id: "s1",
        name: "Sheet 1",
        order: 0,
        rowCount: 200,
        colCount: 26,
        cells: {
          "0:0": { value: "Item", type: "text" },
          "0:1": { value: "Cost", type: "text" },
          "1:0": { value: "Rent", type: "text" },
          "1:1": { value: 1200, type: "number" },
          "2:0": { value: "Food", type: "text" },
          "2:1": { value: 300, type: "number" },
          "3:0": { value: "Gas", type: "text" },
          "3:1": { value: 80, type: "number" },
        },
      },
    ],
  });
}

/** What parity compares: cell values and the style fields the steps touch, merges, sheet size. */
function view(doc: WorkbookDocument) {
  const sheet = doc.sheets[0];
  const cells = Object.fromEntries(
    Object.entries(sheet.cells)
      .map(([cellKey, cell]) => [cellKey, { value: cell.value ?? null, style: cell.style ?? {} }] as const)
      .filter(([, cell]) => cell.value !== null || Object.keys(cell.style).length > 0)
      .sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true })),
  );
  const merges = [...(sheet.merges ?? [])].sort((a, b) => a.r - b.r || a.c - b.c);
  return { cells, merges, rowCount: sheet.rowCount, colCount: sheet.colCount };
}

const SCENARIOS: Record<string, Step[]> = {
  "bold a cell": [{ op: "format", row: 1, col: 0, attr: "bl", value: 1 }],
  "bold then unbold": [
    { op: "format", row: 1, col: 0, attr: "bl", value: 1 },
    { op: "format", row: 1, col: 0, attr: "bl", value: 0 },
  ],
  "fill and font colour": [
    { op: "format", row: 0, col: 1, attr: "bg", value: "#ffeeaa" },
    { op: "format", row: 0, col: 1, attr: "fc", value: "#112233" },
  ],
  "merge a range": [{ op: "merge", r1: 5, c1: 2, r2: 6, c2: 4 }],
  "merge then cancel": [
    { op: "merge", r1: 5, c1: 2, r2: 6, c2: 4 },
    { op: "merge", r1: 5, c1: 2, r2: 6, c2: 4, type: "merge-cancel" },
  ],
  "set a value": [{ op: "set", row: 4, col: 1, value: 42 }],
  "insert a row above": [{ op: "insertRow", index: 2, count: 1, direction: "above" }],
  "insert two rows below": [{ op: "insertRow", index: 1, count: 2, direction: "below" }],
};

describe("parity: FortuneSheet vs candidate grid", () => {
  for (const [name, steps] of Object.entries(SCENARIOS)) {
    it(name, () => {
      const reference = view(runOnFortuneSheet(sample(), steps));
      const candidate = view(runOnStub(sample(), steps));
      expect(candidate).toEqual(reference);
    });
  }
});
