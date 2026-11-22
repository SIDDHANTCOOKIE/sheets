// Formula references by identity (spec 4.1).
//
// A formula typed as `=SUM(B2:B4)+Sheet2!C1` names positions, and positions move when rows or columns are
// inserted, deleted or moved by someone else. Stored by id, the same formula names "the cell in row r7,
// column cB of tab t1", which means the same thing whatever order edits arrive in. The A1 text is only
// what the user types and sees: `formulaToIds` converts it when a formula is saved and `formulaFromIds`
// converts it back for display, editing and calculation, so the calculation engine keeps working in A1.
//
// Stored form: every reference becomes `~(tabId:rowId:colId:flags)`; a range is two of them joined by ":".
// `flags` has bit 1 = column is $absolute, bit 2 = row is $absolute. A whole-column reference has "-" for the
// row, a whole-row reference "-" for the column. A row or column that has no id yet (past the end of the
// sheet) is written "@<position>" and stays positional. Everything else in the formula text is untouched.
//
// Deleting rows or columns leaves a tombstone (the neighbours the id had), so a range whose corner was deleted
// can still resolve: its start moves to the next surviving row or column and its end to the previous one, as
// in Excel. A single reference to a deleted row or column, or a range with nothing left in it, is #REF!.
import { columnIndex, columnName, quoteSheet } from "@formstr/sheets-formula";
import { indexOfRow, rowIdAt, type RowRun } from "../workbook/ids";

/** What is left behind by a deleted id: the surviving ids on each side of it when it was deleted. */
export type Tombstone = [prev: string | null, next: string | null];
export type Tombstones = Record<string, Tombstone>;

/** The rows or the columns of one tab, in order, with the ids that were deleted from them. */
export interface Axis {
  readonly length: number;
  idAt(index: number): string | undefined;
  /** Position of a live id, or -1. */
  indexOf(id: string): number;
  tombstone(id: string): Tombstone | undefined;
}

/** An axis held as a plain list of ids (columns, and rows of small sheets). */
export class ArrayAxis implements Axis {
  private positions = new Map<string, number>();
  constructor(
    public ids: string[],
    public tombstones: Tombstones = {},
  ) {
    this.reindex();
  }
  private reindex(): void {
    this.positions = new Map(this.ids.map((id, i) => [id, i]));
  }
  get length(): number {
    return this.ids.length;
  }
  idAt(index: number): string | undefined {
    return this.ids[index];
  }
  indexOf(id: string): number {
    return this.positions.get(id) ?? -1;
  }
  tombstone(id: string): Tombstone | undefined {
    return this.tombstones[id];
  }
  insert(index: number, ids: string[]): void {
    this.ids.splice(index, 0, ...ids);
    this.reindex();
  }
  remove(index: number, count: number): void {
    const removed = this.ids.splice(index, count);
    const prev = this.ids[index - 1] ?? null;
    const next = this.ids[index] ?? null;
    for (const id of removed) this.tombstones[id] = [prev, next];
    this.reindex();
  }
}

/** An axis held as runs of ids (rows of a big sheet), so a million rows are a handful of entries. */
export class RunAxis implements Axis {
  constructor(
    public runs: readonly RowRun[],
    public tombstones: Tombstones = {},
  ) {}
  get length(): number {
    let total = 0;
    for (const run of this.runs) total += run[2];
    return total;
  }
  idAt(index: number): string | undefined {
    return rowIdAt(this.runs, index);
  }
  indexOf(id: string): number {
    return indexOfRow(this.runs, id);
  }
  tombstone(id: string): Tombstone | undefined {
    return this.tombstones[id];
  }
}

/** The tombstones for `count` ids deleted from position `index` of `axis` (call before the delete). */
export function tombstonesFor(axis: Axis, index: number, count: number): Tombstones {
  const prev = index > 0 ? (axis.idAt(index - 1) ?? null) : null;
  const next = axis.idAt(index + count) ?? null;
  const out: Tombstones = {};
  for (let i = index; i < index + count; i++) {
    const id = axis.idAt(i);
    if (id !== undefined) out[id] = [prev, next];
  }
  return out;
}

export interface TabAxes {
  id: string;
  name: string;
  rows: Axis;
  cols: Axis;
}

export interface RefContext {
  tabs: readonly TabAxes[];
  /** Id of the tab the formula lives on (a reference without a sheet name means this tab). */
  current: string;
}

// ---- the text <-> id conversion -----------------------------------------------------------------

const STRING = /("(?:[^"]|"")*")/;
const A1 =
  /(?<![A-Za-z0-9_.])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![A-Za-z0-9_(])/g;
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
const COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$/;
const ROWS = /^(\$?)(\d+):(\$?)(\d+)$/;
const ID = /~\(([^:()]+):([^:()]+):([^:()]+):([0-3])\)(?::~\(([^:()]+):([^:()]+):([^:()]+):([0-3])\))?/g;

const tabNamed = (ctx: RefContext, name: string) => ctx.tabs.find((tab) => tab.name.toLowerCase() === name.toLowerCase());
const tabById = (ctx: RefContext, id: string) => ctx.tabs.find((tab) => tab.id === id);

/** `~(tab:row:col:flags)` for a position; a position without an id yet stays positional ("@5"). */
function token(tab: TabAxes, row: number | null, col: number | null, colAbs: boolean, rowAbs: boolean): string {
  const part = (axis: Axis, index: number | null) => (index === null ? "-" : (axis.idAt(index) ?? `@${index}`));
  return `~(${tab.id}:${part(tab.rows, row)}:${part(tab.cols, col)}:${(colAbs ? 1 : 0) | (rowAbs ? 2 : 0)})`;
}

/** One A1 reference (without its sheet prefix) as id tokens, or null when it is not a reference we understand. */
function refToIds(tab: TabAxes, text: string): string | null {
  const unsheet = (part: string) => part;
  const cell = (s: string) => {
    const m = CELL.exec(unsheet(s));
    return m ? { colAbs: m[1] === "$", col: columnIndex(m[2]), rowAbs: m[3] === "$", row: Number(m[4]) - 1 } : null;
  };
  const colon = text.indexOf(":");
  if (colon < 0) {
    const c = cell(text);
    return c && c.row >= 0 ? token(tab, c.row, c.col, c.colAbs, c.rowAbs) : null;
  }
  const colsMatch = COLS.exec(text);
  if (colsMatch) {
    const a = columnIndex(colsMatch[2]);
    const b = columnIndex(colsMatch[4]);
    return `${token(tab, null, a, colsMatch[1] === "$", false)}:${token(tab, null, b, colsMatch[3] === "$", false)}`;
  }
  const rowsMatch = ROWS.exec(text);
  if (rowsMatch) {
    const a = Number(rowsMatch[2]) - 1;
    const b = Number(rowsMatch[4]) - 1;
    if (a < 0 || b < 0) return null;
    return `${token(tab, a, null, false, rowsMatch[1] === "$")}:${token(tab, b, null, false, rowsMatch[3] === "$")}`;
  }
  const a = cell(text.slice(0, colon));
  const b = cell(text.slice(colon + 1));
  if (!a || !b || a.row < 0 || b.row < 0) return null;
  return `${token(tab, a.row, a.col, a.colAbs, a.rowAbs)}:${token(tab, b.row, b.col, b.colAbs, b.rowAbs)}`;
}

const unquote = (name: string) => (name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name);

/** The formula (without "=") with every A1 reference replaced by id tokens. A sheet that does not exist is left as text. */
export function formulaToIds(formula: string, ctx: RefContext): string {
  const here = tabById(ctx, ctx.current);
  if (!here) return formula;
  return formula
    .split(STRING)
    .map((part, i) => {
      if (i % 2 === 1) return part; // inside a string
      return part.replace(A1, (whole: string, prefix: string | undefined, ref: string) => {
        const tab = prefix ? tabNamed(ctx, unquote(prefix.slice(0, -1))) : here;
        if (!tab) return whole;
        return refToIds(tab, ref) ?? whole;
      });
    })
    .join("");
}

/** Does this formula text hold id references? */
export const hasIdRefs = (formula: string): boolean => formula.includes("~(");

type Resolved = { kind: "live"; index: number } | { kind: "gone" };

/** Position of an id; a deleted id resolves along its tombstone toward `direction` (1 = next survivor, -1 = previous). */
function resolve(axis: Axis, id: string, direction: 1 | -1 | 0): Resolved {
  if (id.startsWith("@")) return { kind: "live", index: Number(id.slice(1)) };
  const seen = new Set<string>();
  let at: string | null = id;
  while (at !== null && !seen.has(at)) {
    seen.add(at);
    const index = axis.indexOf(at);
    if (index >= 0) return { kind: "live", index };
    if (direction === 0) return { kind: "gone" };
    const tomb = axis.tombstone(at);
    if (!tomb) return { kind: "gone" };
    at = direction === 1 ? tomb[1] : tomb[0];
  }
  return { kind: "gone" };
}

const REF_ERROR = "#REF!";

/**
 * The A1 text for stored id references (the formula without "="). `current` is the tab the formula lives on:
 * references to it are written without a sheet name.
 */
export function formulaFromIds(stored: string, ctx: RefContext): string {
  if (!hasIdRefs(stored)) return stored;
  return stored
    .split(STRING)
    .map((part, i) => {
      if (i % 2 === 1) return part;
      return part.replace(ID, (_whole, tabId, row1, col1, flags1, _tab2, row2, col2, flags2) => {
        const tab = tabById(ctx, tabId);
        if (!tab) return REF_ERROR;
        const prefix = tab.id === ctx.current ? "" : `${quoteSheet(tab.name)}!`;
        const abs = (flags: string) => ({ col: (Number(flags) & 1) !== 0, row: (Number(flags) & 2) !== 0 });
        const cellText = (index: number | null, col: number | null, f: { col: boolean; row: boolean }) =>
          `${col === null ? "" : `${f.col ? "$" : ""}${columnName(col)}`}${index === null ? "" : `${f.row ? "$" : ""}${index + 1}`}`;

        if (row2 === undefined) {
          // one cell
          const r = row1 === "-" ? ({ kind: "live", index: -1 } as Resolved) : resolve(tab.rows, row1, 0);
          const c = col1 === "-" ? ({ kind: "live", index: -1 } as Resolved) : resolve(tab.cols, col1, 0);
          if (r.kind === "gone" || c.kind === "gone") return REF_ERROR;
          return prefix + cellText(row1 === "-" ? null : r.index, col1 === "-" ? null : c.index, abs(flags1));
        }
        // a range: the start moves forward and the end backward when a corner row or column was deleted
        const rowRange = row1 !== "-";
        const colRange = col1 !== "-";
        const r1 = rowRange ? resolve(tab.rows, row1, 1) : null;
        const r2 = rowRange ? resolve(tab.rows, row2, -1) : null;
        const c1 = colRange ? resolve(tab.cols, col1, 1) : null;
        const c2 = colRange ? resolve(tab.cols, col2, -1) : null;
        if ([r1, r2, c1, c2].some((x) => x?.kind === "gone")) return REF_ERROR;
        // a start corner that ran off the end has nothing left after it, an end corner nothing before
        const idx = (x: Resolved | null) => (x && x.kind === "live" ? x.index : null);
        const [ri1, ri2, ci1, ci2] = [idx(r1), idx(r2), idx(c1), idx(c2)];
        if ((ri1 !== null && ri2 !== null && ri2 < ri1) || (ci1 !== null && ci2 !== null && ci2 < ci1)) return REF_ERROR;
        const f1 = abs(flags1);
        const f2 = abs(flags2);
        return `${prefix}${cellText(ri1, ci1, f1)}:${cellText(ri2, ci2, f2)}`;
      });
    })
    .join("");
}
