// Filtering a big sheet by the values of one column. Rows that do not match are hidden (the same
// way hidden rows work), and the filter is remembered with the tab so it can be changed or cleared.
import { cellDisplay } from "../workbook/format";
import type { BigSheetStore } from "./bigSheet";

export interface ColumnFilter {
  col: number;
  /** The first row is a header and is never hidden. */
  header: boolean;
  /** Rows hidden by this filter (kept apart from rows hidden by hand). */
  hidden: number[];
  /** What it shows, for display ("is one of 3 values", "contains abc"). */
  summary: string;
}

export const BLANK = "(Blank)";
/** Filtering reads every row of the column; stay well inside what fits in memory. */
export const MAX_FILTER_ROWS = 2_000_000;
export const MAX_DISTINCT = 5_000;

export interface ColumnScan {
  /** Shown text of every non-empty cell, by row. */
  texts: Map<number, string>;
  /** Distinct texts (blank rows count as BLANK), most common first. */
  values: { value: string; count: number }[];
  /** More than MAX_DISTINCT different texts: only the first ones are listed. */
  truncated: boolean;
}

/** Reads one column of the whole tab, a chunk at a time. */
export async function scanColumn(
  store: BigSheetStore,
  tab: number,
  col: number,
  onProgress?: (done: number, total: number) => void,
  stop?: () => boolean,
): Promise<ColumnScan | null> {
  const info = store.tabs[tab];
  if (info.rows > MAX_FILTER_ROWS) throw new Error(`Filtering works up to ${MAX_FILTER_ROWS.toLocaleString()} rows.`);
  const texts = new Map<number, string>();
  for (let chunk = 0; chunk < info.chunks.length; chunk++) {
    if (stop?.()) return null;
    await store.load(tab, chunk);
    for (let attempt = 0; attempt < 3 && !store.hasChunk(tab, chunk); attempt++) await store.load(tab, chunk);
    for (const [row, c, cell] of store.cellsOfChunk(tab, chunk)) {
      if (c !== col) continue;
      const text = cellDisplay(cell);
      if (text !== "") texts.set(row, text);
    }
    onProgress?.(chunk + 1, info.chunks.length);
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  const counts = new Map<string, number>();
  for (const text of texts.values()) counts.set(text, (counts.get(text) ?? 0) + 1);
  const blanks = info.rows - texts.size;
  if (blanks > 0) counts.set(BLANK, blanks);
  const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], undefined, { numeric: true }));
  return {
    texts,
    values: sorted.slice(0, MAX_DISTINCT).map(([value, count]) => ({ value, count })),
    truncated: sorted.length > MAX_DISTINCT,
  };
}

/** The rows (from the first data row on) that `keep` rejects. `keep` gets the cell's text, or null for an empty cell. */
export function rowsToHide(scan: ColumnScan, rows: number, header: boolean, keep: (text: string | null) => boolean): number[] {
  const hidden: number[] = [];
  for (let row = header ? 1 : 0; row < rows; row++) if (!keep(scan.texts.get(row) ?? null)) hidden.push(row);
  return hidden;
}
