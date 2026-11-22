// FortuneSheet's own engine, run without a browser: the REFERENCE for the parity harness.
// A scenario is a list of steps applied to a workbook; the same steps are applied to the new grid
// and the two resulting WorkbookDocuments are compared (see parity.test.ts).
import { api, defaultContext, initSheetIndex, ensureSheetIndex, type Context, type Sheet } from "@fortune-sheet/core";
import { sheetsForUpdate } from "../sheetData";
import { fortuneSheetsToWorkbook, workbookToFortuneSheets, type WorkbookDocument } from "../workbook";

/** One user-level action, described independently of either engine. */
export type Step =
  | { op: "format"; row: number; col: number; attr: "bl" | "it" | "un" | "cl" | "bg" | "fc" | "fs" | "ht" | "vt" | "tb"; value: number | string }
  | { op: "merge"; r1: number; c1: number; r2: number; c2: number; type?: "merge-all" | "merge-cancel" }
  | { op: "set"; row: number; col: number; value: string | number }
  | { op: "insertRow"; index: number; count: number; direction: "above" | "below" }
  | { op: "insertColumn"; index: number; count: number; direction: "left" | "right" }
  | { op: "deleteRow"; start: number; end: number }
  | { op: "deleteColumn"; start: number; end: number }
  | { op: "clear"; row: number; col: number }
  | { op: "rowHeight"; row: number; height: number }
  | { op: "columnWidth"; col: number; width: number }
  | { op: "hide"; axis: "row" | "column"; index: number }
  | { op: "show"; axis: "row" | "column"; index: number }
  | { op: "freeze"; type: "row" | "column" | "both"; row: number; col: number };

function newContext(sheets: Sheet[]): Context {
  const refs = {
    globalCache: {},
    cellInput: { current: null },
    fxInput: { current: null },
    canvas: { current: null },
    scrollbarX: { current: null },
    scrollbarY: { current: null },
    cellArea: { current: null },
    workbookContainer: { current: null },
  } as unknown as Parameters<typeof defaultContext>[0];
  const ctx = defaultContext(refs);
  ensureSheetIndex(sheets, () => Math.random().toString(36).slice(2));
  ctx.luckysheetfile = sheets;
  ctx.currentSheetId = sheets.find((sheet) => sheet.status === 1)?.id ?? sheets[0].id ?? "";
  initSheetIndex(ctx);
  return ctx;
}

/** Applies the steps with FortuneSheet's engine and returns the workbook it ends up with. */
export function runOnFortuneSheet(document: WorkbookDocument, steps: Step[]): WorkbookDocument {
  // The editor hands FortuneSheet a data matrix (not celldata), exactly as the app does on load.
  const sheets = structuredClone(sheetsForUpdate(workbookToFortuneSheets(document)));
  const ctx = newContext(sheets);
  for (const step of steps) {
    switch (step.op) {
      case "format":
        api.setCellFormat(ctx, step.row, step.col, step.attr, step.value);
        break;
      case "merge":
        api.mergeCells(
          ctx,
          [{ row: [step.r1, step.r2], column: [step.c1, step.c2] }],
          step.type ?? "merge-all",
        );
        break;
      case "set":
        api.setCellValue(ctx, step.row, step.col, step.value, null);
        break;
      case "insertRow":
        api.insertRowOrColumn(ctx, "row", step.index, step.count, step.direction === "above" ? "lefttop" : "rightbottom");
        break;
      case "insertColumn":
        api.insertRowOrColumn(ctx, "column", step.index, step.count, step.direction === "left" ? "lefttop" : "rightbottom");
        break;
      case "deleteRow":
        api.deleteRowOrColumn(ctx, "row", step.start, step.end);
        break;
      case "deleteColumn":
        api.deleteRowOrColumn(ctx, "column", step.start, step.end);
        break;
      case "clear":
        api.clearCell(ctx, step.row, step.col);
        break;
      case "rowHeight":
        api.setRowHeight(ctx, { [String(step.row)]: step.height });
        break;
      case "columnWidth":
        api.setColumnWidth(ctx, { [String(step.col)]: step.width });
        break;
      case "hide":
        api.hideRowOrColumn(ctx, [String(step.index)], step.axis);
        break;
      case "show":
        api.showRowOrColumn(ctx, [String(step.index)], step.axis);
        break;
      case "freeze":
        api.freeze(ctx, step.type, { row: step.row, column: step.col });
        break;
    }
  }
  return fortuneSheetsToWorkbook({
    title: document.title,
    sheets: api.getAllSheets(ctx),
    createdAt: document.createdAt,
    updatedAt: document.updatedAt,
  });
}
