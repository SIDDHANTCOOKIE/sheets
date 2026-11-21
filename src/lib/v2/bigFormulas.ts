// Formulas typed into a big sheet.
//
// The sheet is never loaded whole. To calculate one formula we look at what it names (`=SUM(A:A)`
// names column A), download only the chunks that hold those rows, copy the needed values out
// (the chunk cache may drop the chunks right after) and run the formula engine on that copy.
//
// Not done here: recalculating other formulas when a cell they read changes. That needs an index
// of which cells hold formulas, which is the next step.
import type { WorkbookCell } from "../workbook";
import type { BigSheetStore } from "./bigSheet";

export interface FormulaResult {
  value: string | number | boolean | null;
  /** Set when the formula evaluates to an error (value then holds the code, e.g. "#DIV/0!"). */
  error: boolean;
}

/** Most cells one formula may read (matches the engine's own limit). */
const MAX_READ_CELLS = 5_000_000;
const KEY = 16384;

interface Rect {
  tab: number;
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/** Calculates `formula` (without the leading "=") as if typed into the cell. */
export async function calculateTyped(
  store: BigSheetStore,
  tab: number,
  row: number,
  col: number,
  formula: string,
  onProgress?: (done: number, total: number) => void,
): Promise<FormulaResult> {
  const engine = await import("@formstr/sheets-formula");
  const { FormulaError } = engine;
  const home = store.tabs[tab];
  const tabIndex = (name: string | undefined) => (name === undefined ? tab : store.tabs.findIndex((t) => t.name === name));

  let node;
  try {
    node = engine.parse(formula);
  } catch {
    return { value: "#NAME?", error: true };
  }

  // What the formula reads, as rectangles per tab.
  const rects: Rect[] = [];
  for (const ref of engine.references(node)) {
    const at = tabIndex(ref.sheet);
    if (at < 0) return { value: "#REF!", error: true };
    const info = store.tabs[at];
    if (ref.kind === "cell" && ref.cell) {
      rects.push({ tab: at, r1: ref.cell.row, r2: ref.cell.row, c1: ref.cell.col, c2: ref.cell.col });
    } else if (ref.range) {
      const { r1, c1, r2, c2 } = ref.range;
      rects.push({ tab: at, r1, c1, r2: r2 ?? info.rows - 1, c2: c2 ?? info.cols - 1 });
    }
  }
  let cells = 0;
  for (const rect of rects) cells += Math.max(0, rect.r2 - rect.r1 + 1) * Math.max(0, rect.c2 - rect.c1 + 1);
  if (cells > MAX_READ_CELLS) return { value: "#VALUE!", error: true };

  // Copy the values out, one chunk at a time.
  const values = new Map<number, Map<number, WorkbookCell>>();
  const grab = (tabNo: number) => {
    let map = values.get(tabNo);
    if (!map) values.set(tabNo, (map = new Map()));
    return map;
  };
  const jobs: { rect: Rect; chunk: number }[] = [];
  for (const rect of rects) {
    const info = store.tabs[rect.tab];
    if (info.chunks.length === 0 || rect.r2 < rect.r1) continue;
    const from = store.chunkFor(rect.tab, Math.max(0, rect.r1));
    const to = store.chunkFor(rect.tab, Math.max(0, rect.r2));
    for (let chunk = from; chunk <= to; chunk++) jobs.push({ rect, chunk });
  }
  let done = 0;
  for (const { rect, chunk } of jobs) {
    await store.load(rect.tab, chunk);
    const target = grab(rect.tab);
    for (const [r, c, cell] of store.cellsOfChunk(rect.tab, chunk)) {
      if (r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2) target.set(r * KEY + c, cell);
    }
    onProgress?.(++done, jobs.length);
  }
  // Cells edited since (not yet saved) are in the chunks, so they are read here too.

  const result = engine.calculate(formula, {
    sheet: home.name,
    row,
    col,
    getCell(sheetName, r, c) {
      const at = tabIndex(sheetName);
      const cell = values.get(at)?.get(r * KEY + c);
      if (!cell || cell.value === undefined || cell.value === null) return undefined;
      if (cell.type === "error" && typeof cell.value === "string") return new FormulaError(cell.value as never);
      return cell.value;
    },
    extent: (sheetName) => {
      const info = store.tabs[tabIndex(sheetName)] ?? home;
      return { rows: info.rows, cols: info.cols };
    },
    sheetNames: () => store.tabs.map((t) => t.name),
  });
  if (result instanceof FormulaError) return { value: result.code, error: true };
  // Excel keeps 15 significant digits, so 0.1+0.2 shows as 0.3 and a long sum has no float noise.
  return { value: typeof result === "number" && Number.isFinite(result) ? Number(result.toPrecision(15)) : result, error: false };
}
