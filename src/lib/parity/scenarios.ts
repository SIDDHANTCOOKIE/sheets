// The scenarios the parity harness runs, and the view of a workbook that is compared.
import { normalizeWorkbook, WORKBOOK_SCHEMA, type WorkbookDocument } from "../workbook";
import type { Step } from "./step";

export function sample(): WorkbookDocument {
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

/** What parity compares: cell values and styles, merges, sheet size. */
export function view(doc: WorkbookDocument) {
  const sheet = doc.sheets[0];
  const cells = Object.fromEntries(
    Object.entries(sheet.cells)
      .map(([cellKey, cell]) => [cellKey, { value: typeof cell.value === "number" ? Number(cell.value.toPrecision(12)) : (cell.value ?? null), formula: (cell.formula ?? "").replace(/^=/, "") || null, style: cell.style ?? {} }] as const)
      .filter(([, cell]) => cell.value !== null || cell.formula !== null || Object.keys(cell.style).length > 0)
      .sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true })),
  );
  const merges = [...(sheet.merges ?? [])].sort((a, b) => a.r - b.r || a.c - b.c);
  return {
    cells,
    merges,
    rowCount: sheet.rowCount,
    colCount: sheet.colCount,
    rowHeights: sheet.rowHeights ?? {},
    columnWidths: sheet.columnWidths ?? {},
    hiddenRows: [...(sheet.hiddenRows ?? [])].sort((a, b) => a - b),
    hiddenColumns: [...(sheet.hiddenColumns ?? [])].sort((a, b) => a - b),
    frozen: sheet.frozen ?? null,
  };
}

export const SCENARIOS: Record<string, Step[]> = {
  "bold a cell": [{ op: "format", row: 1, col: 0, attr: "bl", value: 1 }],
  "bold then unbold": [
    { op: "format", row: 1, col: 0, attr: "bl", value: 1 },
    { op: "format", row: 1, col: 0, attr: "bl", value: 0 },
  ],
  "italic, underline and strike": [
    { op: "format", row: 2, col: 1, attr: "it", value: 1 },
    { op: "format", row: 2, col: 1, attr: "un", value: 1 },
    { op: "format", row: 2, col: 1, attr: "cl", value: 1 },
  ],
  "fill and font colour": [
    { op: "format", row: 0, col: 1, attr: "bg", value: "#ffeeaa" },
    { op: "format", row: 0, col: 1, attr: "fc", value: "#112233" },
  ],
  "font size": [{ op: "format", row: 3, col: 1, attr: "fs", value: 18 }],
  "align left, right, centre": [
    { op: "format", row: 1, col: 0, attr: "ht", value: 1 },
    { op: "format", row: 2, col: 0, attr: "ht", value: 2 },
    { op: "format", row: 3, col: 0, attr: "ht", value: 0 },
  ],
  "vertical align top and bottom": [
    { op: "format", row: 1, col: 1, attr: "vt", value: 1 },
    { op: "format", row: 2, col: 1, attr: "vt", value: 2 },
  ],
  "merge a range": [{ op: "merge", r1: 5, c1: 2, r2: 6, c2: 4 }],
  "merge then cancel": [
    { op: "merge", r1: 5, c1: 2, r2: 6, c2: 4 },
    { op: "merge", r1: 5, c1: 2, r2: 6, c2: 4, type: "merge-cancel" },
  ],
  "set a value": [{ op: "set", row: 4, col: 1, value: 42 }],
  "overwrite a value": [{ op: "set", row: 1, col: 1, value: 999 }],
  "insert a row above": [{ op: "insertRow", index: 2, count: 1, direction: "above" }],
  "insert two rows below": [{ op: "insertRow", index: 1, count: 2, direction: "below" }],
  "insert a column left": [{ op: "insertColumn", index: 1, count: 1, direction: "left" }],
  "insert two columns right": [{ op: "insertColumn", index: 0, count: 2, direction: "right" }],
  "delete a row": [{ op: "deleteRow", start: 2, end: 2 }],
  "delete two rows": [{ op: "deleteRow", start: 1, end: 2 }],
  "delete a column": [{ op: "deleteColumn", start: 0, end: 0 }],
  "insert then delete a row": [
    { op: "insertRow", index: 1, count: 1, direction: "above" },
    { op: "deleteRow", start: 1, end: 1 },
  ],
  "style survives an inserted row": [
    { op: "format", row: 3, col: 0, attr: "bl", value: 1 },
    { op: "insertRow", index: 1, count: 2, direction: "above" },
  ],
  "merge moves with an inserted row": [
    { op: "merge", r1: 5, c1: 2, r2: 6, c2: 4 },
    { op: "insertRow", index: 2, count: 1, direction: "above" },
  ],
  "a formula is calculated": [{ op: "set", row: 5, col: 1, value: "=B2*2" }],
  "a formula keeps its target when a row is inserted above it": [
    { op: "set", row: 5, col: 1, value: "=B2*2" },
    { op: "insertRow", index: 1, count: 1, direction: "above" },
  ],
  "a formula keeps its target when a row is inserted below the target": [
    { op: "set", row: 5, col: 1, value: "=B2+B3" },
    { op: "insertRow", index: 2, count: 2, direction: "above" },
  ],
  "a range grows when rows are inserted inside it": [
    { op: "set", row: 8, col: 1, value: "=SUM(B2:B4)" },
    { op: "insertRow", index: 2, count: 1, direction: "above" },
  ],
  "a formula follows an inserted column": [
    { op: "set", row: 5, col: 3, value: "=B2*2" },
    { op: "insertColumn", index: 0, count: 1, direction: "left" },
  ],
  "a formula whose target row is deleted": [
    { op: "set", row: 5, col: 1, value: "=B3*2" },
    { op: "deleteRow", start: 2, end: 2 },
  ],
  "a range shrinks when a row inside it is deleted": [
    { op: "set", row: 8, col: 1, value: "=SUM(B2:B4)" },
    { op: "deleteRow", start: 2, end: 2 },
  ],
  "a formula moves up with a deleted row above it": [
    { op: "set", row: 8, col: 1, value: "=B2*2" },
    { op: "deleteRow", start: 5, end: 6 },
  ],
  "clear a value, keep its style": [
    { op: "format", row: 1, col: 0, attr: "bl", value: 1 },
    { op: "clear", row: 1, col: 0 },
  ],
  "clear a plain cell": [{ op: "clear", row: 1, col: 1 }],
  "set a row height": [{ op: "rowHeight", row: 3, height: 60 }],
  "set a column width": [{ op: "columnWidth", col: 1, width: 200 }],
  "hide and show a row": [
    { op: "hide", axis: "row", index: 2 },
    { op: "hide", axis: "row", index: 4 },
    { op: "show", axis: "row", index: 2 },
  ],
  "hide a column": [{ op: "hide", axis: "column", index: 3 }],
  "freeze the first row": [{ op: "freeze", type: "row", row: 0, col: 0 }],
  "freeze two rows and a column": [{ op: "freeze", type: "both", row: 1, col: 0 }],
  "a row height moves with an inserted row": [
    { op: "rowHeight", row: 3, height: 60 },
    { op: "insertRow", index: 1, count: 2, direction: "above" },
  ],
  "a column width moves with an inserted column": [
    { op: "columnWidth", col: 2, width: 150 },
    { op: "insertColumn", index: 0, count: 1, direction: "left" },
  ],
  "a hidden row moves with an inserted row": [
    { op: "hide", axis: "row", index: 4 },
    { op: "insertRow", index: 1, count: 1, direction: "above" },
  ],
  "a hidden row disappears when deleted": [
    { op: "hide", axis: "row", index: 4 },
    { op: "deleteRow", start: 4, end: 4 },
  ],
  "a number format": [{ op: "numberFormat", row: 1, col: 1, format: "#,##0.00" }],
  "a percent format": [{ op: "numberFormat", row: 3, col: 1, format: "0%" }],
  "text rotation": [{ op: "rotation", row: 1, col: 0, value: 2 }],
  "fill a number series down": [
    { op: "set", row: 5, col: 3, value: 1 },
    { op: "set", row: 6, col: 3, value: 2 },
    { op: "fill", r1: 5, c1: 3, r2: 6, c2: 3, direction: "down", count: 3 },
  ],
  "fill a formula down": [
    { op: "set", row: 5, col: 3, value: "=B2*2" },
    { op: "fill", r1: 5, c1: 3, r2: 5, c2: 3, direction: "down", count: 2 },
  ],
  "fill text right": [{ op: "fill", r1: 1, c1: 0, r2: 1, c2: 0, direction: "right", count: 2 }],
  "merge moves with an inserted column": [
    { op: "merge", r1: 5, c1: 2, r2: 6, c2: 4 },
    { op: "insertColumn", index: 0, count: 1, direction: "left" },
  ],
};

// Many common formulas, one scenario each: both engines must give the same results.
const FORMULAS = [
  "=SUM(B2:B4)", "=AVERAGE(B2:B4)", "=MAX(B2:B4)", "=MIN(B2:B4)", "=COUNT(B1:B4)", "=COUNTA(A1:A4)", "=IF(B2>1000,\"big\",\"small\")",
  "=ROUND(B2/7,2)", "=ABS(-B3)", "=B2*B3+B4", "=B2^2", "=SQRT(B3+B4)", "=MOD(B2,7)", "=INT(B2/7)", "=LEN(A2)", "=UPPER(A2)", "=LOWER(A3)",
  "=LEFT(A2,2)", "=RIGHT(A3,2)", "=CONCATENATE(A2,\"-\",A3)", "=A2&\" \"&A3", "=COUNTIF(B2:B4,\">100\")", "=SUMIF(B2:B4,\">100\")",
  "=AND(B2>1,B3>1)", "=OR(B2>5000,B3>5000)", "=NOT(B2>1)", "=MEDIAN(B2:B4)", "=PRODUCT(B3:B4)", "=SUM(B2:B4)/COUNT(B2:B4)",
  "=VLOOKUP(\"Food\",A2:B4,2,FALSE)", "=INDEX(B2:B4,2)", "=MATCH(\"Gas\",A2:A4,0)", "=ISNUMBER(B2)", "=ISTEXT(A2)", "=IFERROR(B2/0,\"oops\")",
  "=TEXT(B2,\"0.00\")", "=TRIM(\"  a  b \")", "=SUBSTITUTE(A2,\"e\",\"E\")", "=ROUNDUP(B2/7,1)", "=ROUNDDOWN(B2/7,1)", "=CEILING(B2,50)", "=FLOOR(B2,50)",
  "=POWER(2,10)", "=LARGE(B2:B4,2)", "=SMALL(B2:B4,2)", "=COUNTBLANK(B1:B6)", "=ROW(B3)", "=COLUMN(C1)", "=PI()", "=EXP(1)", "=LN(10)", "=LOG10(1000)",
];
for (const formula of FORMULAS) SCENARIOS[`formula ${formula}`] = [{ op: "set", row: 8, col: 3, value: formula }];

export type View = ReturnType<typeof view>;
