// The candidate side of the parity harness, for real: the same steps applied to a BigSheetStore (the
// store the big-sheet grid edits), with the formula-aware structure editor, and the result read back
// as a WorkbookDocument so it can be compared with what FortuneSheet's engine produced.
import ExcelJS from "exceljs";
import type { WorkbookCell, WorkbookCellStyle, WorkbookDocument, WorkbookMerge } from "../workbook";
import { BigSheetStore } from "../v2/bigSheet";
import { calculateTyped, FormulaIndex, refreshDependents } from "../v2/bigFormulas";
import { editStructure } from "../v2/bigStructure";
import { dragFill } from "../v2/bigClipboard";
import { importXlsxToBlossom } from "../v2/streamImport";
import { fakeBlossom, localSigner } from "../v2/testUtils";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";
import type { Step } from "./fortuneHeadless";

function inProcessWorker(): WorkerLike {
  const worker: WorkerLike = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: MainToWorker) {
      queueMicrotask(() => handle(message));
    },
    terminate() {
      handle({ t: "cancel" });
    },
  };
  const handle = createWorkerHandler((message) => setTimeout(() => worker.onmessage?.({ data: message }), 0));
  return worker;
}

/** Builds a big-sheet store holding the first sheet of a document (through a real .xlsx and the streaming import). */
export async function documentToStore(document: WorkbookDocument): Promise<BigSheetStore> {
  const sheet = document.sheets[0];
  const book = new ExcelJS.Workbook();
  const worksheet = book.addWorksheet(sheet.name);
  for (const [cellKey, cell] of Object.entries(sheet.cells)) {
    if (cell.value === undefined) continue;
    const [row, col] = cellKey.split(":").map(Number);
    worksheet.getCell(row + 1, col + 1).value = cell.value as ExcelJS.CellValue;
  }
  const bytes = await book.xlsx.writeBuffer();
  const servers = ["https://parity.test"];
  const fake = fakeBlossom();
  const signer = localSigner();
  const env = { fetch: fake.fetch, sleep: async () => {} };
  const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
    title: document.title,
    servers,
    signer,
    env,
    createWorker: inProcessWorker,
    start: { maxChunkBytes: 64 * 1024, headRawBytes: 100_000, rawCap: 1_000_000 },
  });
  const store = new BigSheetStore(root, servers, env);
  await store.open();
  await store.load(0, 0);
  return store;
}

const STYLE_OF: Record<string, (value: number | string) => Partial<WorkbookCellStyle>> = {
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

/** Applies the steps with the big-sheet store and returns the workbook it ends up with. */
export async function runOnBigStore(document: WorkbookDocument, steps: Step[]): Promise<WorkbookDocument> {
  const store = await documentToStore(document);
  const index = new FormulaIndex();
  const rowsAtStart = store.tabs[0].rows;
  const colsAtStart = store.tabs[0].cols;
  for (const step of steps) {
    switch (step.op) {
      case "format": {
        await store.load(0, store.chunkFor(0, step.row));
        await store.edit(0, step.row, step.col, step.row, step.col, (cell) => {
          const style: Record<string, unknown> = { ...cell?.style, ...STYLE_OF[step.attr](step.value) };
          for (const flag of ["bold", "italic", "underline", "strike", "wrap"]) if (style[flag] === false) delete style[flag];
          return { ...(cell ?? {}), style: style as WorkbookCellStyle };
        });
        break;
      }
      case "merge": {
        const merges = store.tabs[0].layout.merges ?? [];
        if (step.type === "merge-cancel") {
          await store.setLayout(0, { merges: merges.filter((m) => !(m.r === step.r1 && m.c === step.c1)) });
        } else {
          await store.setLayout(0, { merges: [...merges, { r: step.r1, c: step.c1, rs: step.r2 - step.r1 + 1, cs: step.c2 - step.c1 + 1 }] });
        }
        break;
      }
      case "set": {
        await store.load(0, store.chunkFor(0, step.row));
        if (typeof step.value === "string" && step.value.startsWith("=")) {
          // Typed like in the grid: calculate, store formula and result, tell the index.
          const formula = step.value.slice(1);
          const result = await calculateTyped(store, 0, step.row, step.col, formula);
          await store.edit(0, step.row, step.col, step.row, step.col, () => ({ formula, value: result.value, ...(result.error ? { type: "error" as const } : {}) }));
          await index.set(store, 0, step.row, step.col, formula);
        } else {
          await store.setCell(0, step.row, step.col, String(step.value));
        }
        break;
      }
      case "numberFormat":
        await store.load(0, store.chunkFor(0, step.row));
        await store.edit(0, step.row, step.col, step.row, step.col, (cell) => ({ ...(cell ?? {}), style: { ...cell?.style, numberFormat: step.format } }));
        break;
      case "rotation":
        await store.load(0, store.chunkFor(0, step.row));
        await store.edit(0, step.row, step.col, step.row, step.col, (cell) => ({ ...(cell ?? {}), style: { ...cell?.style, rotation: step.value } }));
        break;
      case "fill": {
        const source = { r1: step.r1, c1: step.c1, r2: step.r2, c2: step.c2 };
        const extended = step.direction === "down" ? { ...source, r2: step.r2 + step.count } : { ...source, c2: step.c2 + step.count };
        const done = await dragFill(store, 0, source, extended);
        // Like the grid: the filled formulas are calculated afterwards.
        if (done) await refreshDependents(store, index, [{ tab: 0, ...done }], undefined, true);
        break;
      }
      case "clear":
        await store.load(0, store.chunkFor(0, step.row));
        await store.edit(0, step.row, step.col, step.row, step.col, (cell) => (cell?.style ? { style: cell.style } : undefined));
        break;
      case "rowHeight":
        await store.setLayout(0, { rowHeights: { ...store.tabs[0].layout.rowHeights, [step.row]: step.height } });
        break;
      case "columnWidth":
        await store.setLayout(0, { columnWidths: { ...store.tabs[0].layout.columnWidths, [step.col]: step.width } });
        break;
      case "hide": {
        const key = step.axis === "row" ? "hiddenRows" : "hiddenColumns";
        await store.setLayout(0, { [key]: [...new Set([...(store.tabs[0].layout[key] ?? []), step.index])].sort((a, b) => a - b) });
        break;
      }
      case "show": {
        const key = step.axis === "row" ? "hiddenRows" : "hiddenColumns";
        const left = (store.tabs[0].layout[key] ?? []).filter((i) => i !== step.index);
        await store.setLayout(0, { [key]: left.length > 0 ? left : undefined });
        break;
      }
      case "freeze": {
        const rows = step.type === "column" ? 0 : step.row + 1;
        const columns = step.type === "row" ? 0 : step.col + 1;
        await store.setLayout(0, { frozen: { rows, columns } });
        break;
      }
      case "insertRow":
        await editStructure(store, index, 0, { kind: "insert", axis: "row", index: step.direction === "above" ? step.index : step.index + 1, count: step.count });
        break;
      case "insertColumn":
        await editStructure(store, index, 0, { kind: "insert", axis: "col", index: step.direction === "left" ? step.index : step.index + 1, count: step.count });
        break;
      case "deleteRow":
        await editStructure(store, index, 0, { kind: "delete", axis: "row", index: step.start, count: step.end - step.start + 1 });
        break;
      case "deleteColumn":
        await editStructure(store, index, 0, { kind: "delete", axis: "col", index: step.start, count: step.end - step.start + 1 });
        break;
    }
  }

  const sheet = document.sheets[0];
  const cells: Record<string, WorkbookCell> = {};
  for (let c = 0; c < store.tabs[0].chunks.length; c++) {
    await store.load(0, c);
    for (const [row, col, cell] of store.cellsOfChunk(0, c)) cells[`${row}:${col}`] = cell;
  }
  const merges: WorkbookMerge[] = (store.tabs[0].layout.merges ?? []).map((m) => ({ ...m }));
  return {
    ...document,
    sheets: [
      {
        ...sheet,
        cells,
        merges,
        ...(store.tabs[0].layout.rowHeights ? { rowHeights: store.tabs[0].layout.rowHeights } : {}),
        ...(store.tabs[0].layout.columnWidths ? { columnWidths: store.tabs[0].layout.columnWidths } : {}),
        ...(store.tabs[0].layout.hiddenRows ? { hiddenRows: store.tabs[0].layout.hiddenRows } : {}),
        ...(store.tabs[0].layout.hiddenColumns ? { hiddenColumns: store.tabs[0].layout.hiddenColumns } : {}),
        ...(store.tabs[0].layout.frozen ? { frozen: store.tabs[0].layout.frozen } : {}),
        // The two engines start from different default sizes: what counts is how much each step changed them.
        rowCount: sheet.rowCount + (store.tabs[0].rows - rowsAtStart),
        colCount: sheet.colCount + (store.tabs[0].cols - colsAtStart),
      },
    ],
  };
}
