// The candidate side of the parity harness: the same steps applied straight to a WorkbookDocument.
// This is a STUB standing in for the new grid's operation layer (packages/sheets-model once it
// exists). Each step here is the behaviour the new grid must reproduce; when the real implementation
// lands, the harness runs it in place of this file and the same scenarios must stay green.
import type { WorkbookCell, WorkbookCellStyle, WorkbookDocument, WorkbookMerge, WorkbookSheet } from "../workbook";
import type { Step } from "./fortuneHeadless";

const key = (row: number, col: number) => `${row}:${col}`;

function parse(cellKey: string): [number, number] {
  const [row, col] = cellKey.split(":").map(Number);
  return [row, col];
}

const STYLE_OF: Record<string, (value: number | string) => Partial<WorkbookCellStyle> | undefined> = {
  bl: (value) => ({ bold: value === 1 }),
  it: (value) => ({ italic: value === 1 }),
  un: (value) => ({ underline: value === 1 }),
  cl: (value) => ({ strike: value === 1 }),
  bg: (value) => ({ background: String(value) }),
  fc: (value) => ({ color: String(value) }),
  fs: (value) => ({ fontSize: Number(value) }),
  ht: (value) => ({ align: value === 1 ? "left" : value === 2 ? "right" : "center" }),
  vt: (value) => ({ verticalAlign: value === 1 ? "top" : value === 2 ? "bottom" : "middle" }),
  tb: (value) => ({ wrap: value === "2" || value === 2 }),
};

function shiftRows(sheet: WorkbookSheet, from: number, delta: number): void {
  const cells: Record<string, WorkbookCell> = {};
  for (const [cellKey, cell] of Object.entries(sheet.cells)) {
    const [row, col] = parse(cellKey);
    const moved = row >= from ? row + delta : row;
    if (moved >= 0 && !(delta < 0 && row >= from + delta && row < from)) cells[key(moved, col)] = cell;
  }
  sheet.cells = cells;
  if (sheet.merges) {
    sheet.merges = sheet.merges.map((merge): WorkbookMerge => (merge.r >= from ? { ...merge, r: merge.r + delta } : merge));
  }
  if (sheet.rowHeights) {
    sheet.rowHeights = Object.fromEntries(
      Object.entries(sheet.rowHeights).map(([row, height]) => [Number(row) >= from ? Number(row) + delta : Number(row), height]),
    );
  }
  sheet.rowCount += delta;
}

export function runOnStub(document: WorkbookDocument, steps: Step[]): WorkbookDocument {
  const doc = structuredClone(document);
  const sheet = doc.sheets.find((item) => item.id === doc.activeSheetId) ?? doc.sheets[0];
  for (const step of steps) {
    switch (step.op) {
      case "format": {
        const cell = sheet.cells[key(step.row, step.col)] ?? {};
        const style: Record<string, unknown> = { ...cell.style, ...STYLE_OF[step.attr](step.value) };
        // The canonical model stores a flag only while it is on (bold/italic/underline/strike/wrap).
        for (const flag of ["bold", "italic", "underline", "strike", "wrap"]) if (style[flag] === false) delete style[flag];
        sheet.cells[key(step.row, step.col)] = { ...cell, style: style as WorkbookCellStyle };
        break;
      }
      case "merge": {
        if (step.type === "merge-cancel") {
          sheet.merges = (sheet.merges ?? []).filter((merge) => !(merge.r === step.r1 && merge.c === step.c1));
        } else {
          sheet.merges = [...(sheet.merges ?? []), { r: step.r1, c: step.c1, rs: step.r2 - step.r1 + 1, cs: step.c2 - step.c1 + 1 }];
        }
        break;
      }
      case "set": {
        const cell = sheet.cells[key(step.row, step.col)] ?? {};
        sheet.cells[key(step.row, step.col)] = { ...cell, value: step.value, type: typeof step.value === "number" ? "number" : "text" };
        break;
      }
      case "insertRow":
        shiftRows(sheet, step.direction === "above" ? step.index : step.index + 1, step.count);
        break;
      case "deleteRow":
        shiftRows(sheet, step.end + 1, -(step.end - step.start + 1));
        break;
      default:
        throw new Error(`stub grid does not implement ${step.op} yet`);
    }
  }
  return doc;
}
