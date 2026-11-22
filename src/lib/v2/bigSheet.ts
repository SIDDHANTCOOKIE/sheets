// A big sheet, read (and optionally edited) a chunk at a time, never whole.
//
// The root lists every tab's chunks with the rows they hold (rowStart/rowCount), so finding the
// chunk for a row is a binary search and showing rows 5,000,000-5,000,040 downloads one chunk.
// Decoded chunks sit in a small LRU cache; edited ("dirty") chunks are pinned until saved.
//
// Saving re-encrypts only the dirty chunks (plus a tab's index blob when its chunk list lives in
// one) and uploads them under ONE signed authorization, then returns the updated root. Formulas
// are not recalculated: an edit stores the typed value.
import type { WorkbookCell } from "../workbook";
import { encryptFrame, uploadStaged, type BlossomEnv, type StagedBlob, type UploadOptions } from "./blossom";
import { concatRuns, deleteRuns, insertRuns, legacyRows, mintRows, newDeviceTag, padCols, padRuns, rowCountOf, type RowRun } from "../workbook/ids";
import { readChunk, writeChunk, type ChunkBody as StoredChunkBody } from "./chunks";
import { frame } from "./codec";
import { downloadTabBody } from "./blossom";
import { editLayout, type StructEdit } from "./layoutEdit";
import { FRAME_TYPE_TAB_INDEX, resolveChunks, type ChunkEntry, type SheetRootV2 } from "./root";

/** A decoded chunk: cells are positional ("<row - chunk start>:<column index>"), rows are its row ids. */
interface ChunkBody {
  rows?: RowRun[];
  cells?: Record<string, WorkbookCell>;
  rowCount?: number;
  colCount?: number;
  [layout: string]: unknown;
}

/** Decoded chunks kept in memory (each is ~6 MB compressed, a few tens of MB as objects). */
const CACHE_CHUNKS = 6;
/** Largest selection one toolbar action may change (it may have to download several chunks). */
export const MAX_EDIT_CELLS = 200_000;
const MAX_COLUMNS = 16_384;
const MAX_INSERT_ROWS = 100_000;
const UNDO_STEPS = 100;

export interface CellAddress {
  tab: number;
  row: number;
  col: number;
}

interface Change {
  tab: number;
  row: number;
  col: number;
  before?: WorkbookCell;
  after?: WorkbookCell;
}

/** What a tab remembers about its look besides cell contents (stored with the tab's first chunk). */
export interface TabLayout {
  columnWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  frozen?: { rows: number; columns: number };
  merges?: { r: number; c: number; rs: number; cs: number }[];
  borders?: { r1: number; c1: number; r2: number; c2: number; type: "all" | "outside" | "inside" | "top" | "bottom" | "left" | "right" | "none"; color?: string; style?: string }[];
}

const LAYOUT_KEYS = ["columnWidths", "rowHeights", "hiddenRows", "hiddenColumns", "frozen", "merges", "borders"] as const;

export interface BigTab {
  layout: TabLayout;
  name: string;
  chunks: ChunkEntry[];
  /** Chunk list stored in an index blob (must be re-uploaded when a chunk changes). */
  indexed: boolean;
  rows: number;
  cols: number;
  /** Column ids in order (legacy ids for columns that predate ids). Written back on save. */
  colIds: string[];
}

export class BigSheetStore {
  readonly root: SheetRootV2;
  tabs: BigTab[] = [];
  private readonly servers: string[];
  private readonly env: BlossomEnv;
  private cache = new Map<string, ChunkBody>();
  private loading = new Map<string, Promise<ChunkBody>>();
  private dirty = new Set<string>();
  private undoStack: Change[][] = [];
  private redoStack: Change[][] = [];
  private listeners = new Set<() => void>();
  /** Called each time a chunk finishes loading (the formula index learns its formulas here). */
  onChunkLoaded?: (tab: number, chunk: number) => void;
  /** Ids for rows and columns this device adds. */
  private readonly idTag = newDeviceTag();
  private idCounter = 0;

  constructor(root: SheetRootV2, servers: string[], env: BlossomEnv = {}) {
    this.root = root;
    this.servers = [...root.servers, ...servers];
    this.env = env;
  }

  private fetchBody = (ref: ChunkEntry["ref"], servers: string[], type?: number) =>
    downloadTabBody(ref, [...servers, ...this.servers], this.env, type);

  /** Reads every tab's chunk list and each tab's first chunk (it carries the size and layout). */
  async open(): Promise<void> {
    this.tabs = [];
    for (const tab of this.root.tabs) {
      const chunks = (await resolveChunks(tab, this.root.servers, this.fetchBody)) ?? [];
      const last = chunks[chunks.length - 1];
      this.tabs.push({
        name: tab.name,
        chunks,
        indexed: !!tab.chunks && !Array.isArray(tab.chunks),
        rows: last ? last.rowStart + last.rowCount : 0,
        cols: 26,
        colIds: [],
        layout: {},
      });
    }
    await Promise.all(
      this.tabs.map(async (tab, index) => {
        if (tab.chunks.length === 0) return;
        const head = await this.load(index, 0);
        tab.rows = Math.max(tab.rows, head.rowCount ?? 0, 100);
        tab.cols = Math.max(tab.cols, head.colCount ?? 0, this.root.tabs[index].cols?.length ?? 0);
        tab.colIds = padCols(this.root.tabs[index].cols, tab.cols);
        for (const key of LAYOUT_KEYS) if (head[key] !== undefined) (tab.layout as Record<string, unknown>)[key] = head[key];
      }),
    );
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of this.listeners) listener();
  }

  /** Index of the chunk holding `row` (the last chunk takes rows past the end). */
  chunkFor(tab: number, row: number): number {
    const chunks = this.tabs[tab].chunks;
    let lo = 0;
    let hi = chunks.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (chunks[mid].rowStart <= row) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  private key = (tab: number, chunk: number) => `${tab}:${chunk}`;

  /** Downloads and decodes a chunk (once), evicting the least recently used clean chunks. */
  load(tab: number, chunk: number): Promise<ChunkBody> {
    const key = this.key(tab, chunk);
    const hit = this.cache.get(key);
    if (hit) {
      this.cache.delete(key);
      this.cache.set(key, hit);
      return Promise.resolve(hit);
    }
    let pending = this.loading.get(key);
    if (!pending) {
      const entry = this.tabs[tab].chunks[chunk];
      pending = (this.fetchBody(entry.ref, this.root.servers) as Promise<StoredChunkBody>).then((stored) => {
        // The current column ids (they change when columns are inserted or deleted).
        const colIds = this.tabs[tab].colIds.length > 0 ? this.tabs[tab].colIds : this.root.tabs[tab].cols;
        const read = readChunk(stored, entry.rowStart, entry.rowCount, colIds);
        const body: ChunkBody = { ...read.layout, rows: read.rows, cells: read.cells };
        this.cache.set(key, body);
        this.loading.delete(key);
        for (const old of this.cache.keys()) {
          if (this.cache.size <= CACHE_CHUNKS) break;
          if (!this.dirty.has(old) && old !== key) this.cache.delete(old);
        }
        this.onChunkLoaded?.(tab, chunk);
        this.emit();
        return body;
      });
      pending.catch(() => this.loading.delete(key));
      this.loading.set(key, pending);
    }
    return pending;
  }

  /** The cell if its chunk is in memory; `undefined` when not loaded yet, `null` when empty. */
  cell(tab: number, row: number, col: number): WorkbookCell | null | undefined {
    const chunk = this.chunkFor(tab, row);
    const body = this.cache.get(this.key(tab, chunk));
    if (!body) return undefined;
    const start = this.tabs[tab].chunks[chunk].rowStart;
    return body.cells?.[`${row - start}:${col}`] ?? null;
  }

  /** Is the chunk in memory right now? */
  hasChunk(tab: number, chunk: number): boolean {
    return this.cache.has(this.key(tab, chunk));
  }

  /** Every cell of a loaded chunk as [row, column, cell] (empty when the chunk is not in memory). */
  *cellsOfChunk(tab: number, chunk: number): Generator<[number, number, WorkbookCell]> {
    const body = this.cache.get(this.key(tab, chunk));
    if (!body?.cells) return;
    const start = this.tabs[tab].chunks[chunk].rowStart;
    for (const [key, cell] of Object.entries(body.cells)) {
      const at = key.indexOf(":");
      yield [start + Number(key.slice(0, at)), Number(key.slice(at + 1)), cell];
    }
  }

  isLoaded(tab: number, row: number): boolean {
    return this.cache.has(this.key(tab, this.chunkFor(tab, row)));
  }

  get dirtyCount(): number {
    return this.dirty.size;
  }

  /** Parses what the user typed: a number when it reads as one, else text; empty clears the value. */
  static parseInput(text: string, previous?: WorkbookCell): WorkbookCell | undefined {
    const trimmed = text.trim();
    const style = previous?.style;
    if (trimmed === "") return style ? { style } : undefined;
    const isNumber = /^[-+]?[\d,]*\.?\d+(e[-+]?\d+)?$/i.test(trimmed);
    const value = isNumber ? Number(trimmed.replace(/,/g, "")) : text;
    return { value, ...(style ? { style } : {}) };
  }

  async setCell(tab: number, row: number, col: number, text: string): Promise<void> {
    await this.edit(tab, row, col, row, col, (cell) => BigSheetStore.parseInput(text, cell));
  }

  /**
   * Changes every cell in rows r1..r2, columns c1..c2 (inclusive) with `change`, which gets a copy
   * of the cell (or undefined) and returns the new cell (undefined clears it). One undo step.
   */
  async edit(
    tab: number,
    r1: number,
    c1: number,
    r2: number,
    c2: number,
    change: (cell: WorkbookCell | undefined, row: number, col: number) => WorkbookCell | undefined,
  ): Promise<void> {
    const count = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (count > MAX_EDIT_CELLS) {
      throw new Error(`Select at most ${MAX_EDIT_CELLS.toLocaleString()} cells at once in large sheet mode.`);
    }
    const from = this.chunkFor(tab, r1);
    const to = this.chunkFor(tab, r2);
    for (let c = from; c <= to; c++) await this.load(tab, c);
    const changes: Change[] = [];
    for (let row = r1; row <= r2; row++) {
      for (let col = c1; col <= c2; col++) {
        const before = this.rawCell(tab, row, col);
        const after = change(before ? structuredClone(before) : undefined, row, col);
        if (JSON.stringify(before) === JSON.stringify(after)) continue;
        changes.push({ tab, row, col, before, after });
      }
    }
    if (changes.length === 0) return;
    this.apply(changes, "after");
    this.undoStack.push(changes);
    if (this.undoStack.length > UNDO_STEPS) this.undoStack.shift();
    this.redoStack = [];
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Undoes the last edit; returns the cells it changed (so formulas that read them can be refreshed). */
  undo(): CellAddress[] {
    const step = this.undoStack.pop();
    if (!step) return [];
    this.apply(step, "before");
    this.redoStack.push(step);
    return step.map(({ tab, row, col }) => ({ tab, row, col }));
  }

  redo(): CellAddress[] {
    const step = this.redoStack.pop();
    if (!step) return [];
    this.apply(step, "after");
    this.undoStack.push(step);
    return step.map(({ tab, row, col }) => ({ tab, row, col }));
  }

  /**
   * Stores a calculated value (a formula's new result). Not an undo step: it is recomputed from the
   * cells it reads, so undoing the edit that caused it recalculates it again.
   */
  async setDerived(tab: number, row: number, col: number, cell: WorkbookCell): Promise<void> {
    await this.load(tab, this.chunkFor(tab, row));
    const before = this.rawCell(tab, row, col);
    if (JSON.stringify(before) === JSON.stringify(cell)) return;
    this.apply([{ tab, row, col, before, after: cell }], "after");
  }

  /**
   * Changes what a tab remembers about its look (sizes, hidden rows, frozen panes). It is stored in
   * the tab's first chunk, which is then saved with the next save. Not an undo step.
   */
  async setLayout(tab: number, patch: Partial<TabLayout>): Promise<void> {
    const body = await this.load(tab, 0);
    const info = this.tabs[tab];
    for (const key of LAYOUT_KEYS) {
      if (!(key in patch)) continue;
      const value = patch[key];
      if (value === undefined) {
        delete body[key];
        delete info.layout[key];
      } else {
        body[key] = value;
        (info.layout as Record<string, unknown>)[key] = value;
      }
    }
    this.dirty.add(this.key(tab, 0));
    this.emit();
  }

  // --- inserting and deleting rows and columns ------------------------------------------------

  /**
   * Inserts or deletes rows or columns: moves the cells, the ids and the tab's saved look (sizes,
   * merges, borders). Formulas are NOT rewritten here (see bigStructure.ts). Not undoable: the undo
   * history is cleared because it names cells by position.
   */
  async applyEdit(tab: number, edit: StructEdit): Promise<void> {
    if (edit.count <= 0) return;
    const info = this.tabs[tab];
    const total = edit.axis === "row" ? info.rows : info.cols;
    if (edit.kind === "delete" && edit.index + edit.count > total) edit = { ...edit, count: total - edit.index };
    if (edit.count <= 0) return;
    if (edit.kind === "delete" && edit.count >= total) {
      throw new Error(`A sheet needs at least one ${edit.axis === "row" ? "row" : "column"}.`);
    }
    if (edit.kind === "insert" && edit.axis === "col" && info.cols + edit.count > MAX_COLUMNS) {
      throw new Error(`A sheet has at most ${MAX_COLUMNS.toLocaleString()} columns.`);
    }
    if (edit.kind === "insert" && edit.axis === "row" && edit.count > MAX_INSERT_ROWS) {
      throw new Error(`Insert at most ${MAX_INSERT_ROWS.toLocaleString()} rows at once.`);
    }
    // The first chunk carries the size and look, so it is always rewritten.
    await this.load(tab, 0);
    if (edit.axis === "row") await (edit.kind === "insert" ? this.insertRows(tab, edit) : this.deleteRows(tab, edit));
    else if (edit.kind === "insert") this.insertCols(tab, edit);
    else this.deleteCols(tab, edit);

    const head = this.cache.get(this.key(tab, 0)) as ChunkBody;
    info.layout = editLayout(info.layout, edit);
    for (const key of LAYOUT_KEYS) {
      const value = info.layout[key];
      if (value === undefined) delete head[key];
      else head[key] = value;
    }
    this.dirty.add(this.key(tab, 0));
    this.undoStack = [];
    this.redoStack = [];
    this.emit();
  }

  private async insertRows(tab: number, edit: StructEdit): Promise<void> {
    const info = this.tabs[tab];
    const at = Math.min(edit.index, info.rows);
    const last = info.chunks[info.chunks.length - 1];
    const coverage = last.rowStart + last.rowCount;
    if (at < coverage) {
      const k = this.chunkFor(tab, at);
      const body = await this.load(tab, k);
      const entry = info.chunks[k];
      const local = at - entry.rowStart;
      body.cells = moveCells(body.cells, (row, col) => [row >= local ? row + edit.count : row, col]);
      body.rows = insertRuns(padRuns(body.rows ?? [], entry.rowCount), local, mintRows(this.idTag, this.idCounter, edit.count));
      this.idCounter += edit.count;
      info.chunks[k] = { ...entry, rowCount: entry.rowCount + edit.count };
      for (let j = k + 1; j < info.chunks.length; j++) info.chunks[j] = { ...info.chunks[j], rowStart: info.chunks[j].rowStart + edit.count };
      this.dirty.add(this.key(tab, k));
    }
    info.rows += edit.count;
  }

  private async deleteRows(tab: number, edit: StructEdit): Promise<void> {
    const info = this.tabs[tab];
    const last = info.chunks[info.chunks.length - 1];
    const coverage = last.rowStart + last.rowCount;
    const from = edit.index;
    const to = Math.min(edit.index + edit.count, coverage); // rows [from, to) live in chunks
    let removed = 0;
    if (from < to) {
      const first = this.chunkFor(tab, from);
      const lastChunk = this.chunkFor(tab, to - 1);
      // Load what must be edited before changing anything, so a failed download leaves the sheet alone.
      const partial: number[] = [];
      for (let k = first; k <= lastChunk; k++) {
        const entry = info.chunks[k];
        const covered = from <= entry.rowStart && to >= entry.rowStart + entry.rowCount;
        if (!covered || k === 0) partial.push(k);
      }
      for (const k of partial) await this.load(tab, k);
      const chunks = [...info.chunks];
      for (let k = 0; k < chunks.length; k++) {
        const entry = chunks[k];
        const start = entry.rowStart;
        const end = start + entry.rowCount;
        const lo = Math.max(from, start);
        const hi = Math.min(to, end);
        if (hi > lo && entry.rowCount > 0) {
          const gone = hi - lo;
          if (partial.includes(k)) {
            const body = this.cache.get(this.key(tab, k)) as ChunkBody;
            const a = lo - start;
            const b = hi - start;
            body.cells = moveCells(body.cells, (row, col) => (row >= a && row < b ? null : [row >= b ? row - gone : row, col]));
            body.rows = deleteRuns(padRuns(body.rows ?? [], entry.rowCount), a, gone);
            this.dirty.add(this.key(tab, k));
          } else {
            // Every row of the chunk is deleted: it is simply emptied (nothing to download).
            this.cache.delete(this.key(tab, k));
            this.dirty.delete(this.key(tab, k));
          }
          chunks[k] = { ...entry, rowStart: start - removed, rowCount: entry.rowCount - gone, cells: partial.includes(k) ? entry.cells : 0 };
          removed += gone;
        } else {
          chunks[k] = { ...entry, rowStart: start - removed };
        }
      }
      info.chunks = chunks;
    }
    info.rows -= edit.count;
  }

  private insertCols(tab: number, edit: StructEdit): void {
    const info = this.tabs[tab];
    const at = Math.min(edit.index, info.cols);
    const ids = padCols(info.colIds, info.cols);
    const fresh = Array.from({ length: edit.count }, () => `x${this.idTag}${this.idCounter++}`);
    info.colIds = [...ids.slice(0, at), ...fresh, ...ids.slice(at)];
    info.cols += edit.count;
    this.mapLoadedCells(tab, (row, col) => [row, col >= at ? col + edit.count : col]);
  }

  private deleteCols(tab: number, edit: StructEdit): void {
    const info = this.tabs[tab];
    const ids = padCols(info.colIds, info.cols);
    const to = edit.index + edit.count;
    info.colIds = [...ids.slice(0, edit.index), ...ids.slice(to)];
    info.cols -= edit.count;
    this.mapLoadedCells(tab, (row, col) => (col >= edit.index && col < to ? null : [row, col >= to ? col - edit.count : col]));
    // Chunks written before may still hold the deleted columns; they are ignored when read (their ids are gone).
  }

  /** Moves the cells of every loaded chunk of a tab (positions are chunk-relative rows and columns). */
  private mapLoadedCells(tab: number, move: (row: number, col: number) => [number, number] | null): void {
    for (const [key, body] of this.cache) {
      if (!key.startsWith(`${tab}:`)) continue;
      body.cells = moveCells(body.cells, move);
    }
  }

  /** Chunks currently in memory, as [tab, chunk]. */
  loadedChunks(): [number, number][] {
    return [...this.cache.keys()].map((key) => key.split(":").map(Number) as [number, number]);
  }

  private rawCell(tab: number, row: number, col: number): WorkbookCell | undefined {
    const chunk = this.chunkFor(tab, row);
    const body = this.cache.get(this.key(tab, chunk));
    return body?.cells?.[`${row - this.tabs[tab].chunks[chunk].rowStart}:${col}`];
  }

  /** Writes changes into their (pinned) chunks and marks them dirty. */
  private apply(changes: Change[], side: "before" | "after"): void {
    for (const item of changes) {
      const chunk = this.chunkFor(item.tab, item.row);
      const key = this.key(item.tab, chunk);
      const body = this.cache.get(key);
      if (!body) continue; // edited chunks are pinned, so this does not happen
      const cells = (body.cells ??= {});
      const cellKey = `${item.row - this.tabs[item.tab].chunks[chunk].rowStart}:${item.col}`;
      const value = item[side];
      if (value) cells[cellKey] = structuredClone(value);
      else delete cells[cellKey];
      const tabInfo = this.tabs[item.tab];
      tabInfo.rows = Math.max(tabInfo.rows, item.row + 1);
      tabInfo.cols = Math.max(tabInfo.cols, item.col + 1);
      if (tabInfo.colIds.length < tabInfo.cols) tabInfo.colIds = padCols(tabInfo.colIds, tabInfo.cols);
      this.dirty.add(key);
    }
    this.emit();
  }

  /**
   * Uploads the edited chunks under one signature and returns the root that points at them.
   * The caller signs and publishes the root event.
   */
  async save(options: UploadOptions): Promise<SheetRootV2> {
    if (this.dirty.size === 0) return this.root;
    const keys = [...this.dirty];
    const staged: StagedBlob[] = [];
    const targets: { tab: number; chunk: number }[] = [];
    for (const key of keys) {
      const [tab, chunk] = key.split(":").map(Number);
      const body = this.cache.get(key) as ChunkBody;
      const entry = this.tabs[tab].chunks[chunk];
      const next = this.tabs[tab].chunks[chunk + 1];
      if (!next) entry.rowCount = Math.max(entry.rowCount, this.tabs[tab].rows - entry.rowStart);
      if (chunk === 0) {
        body.rowCount = this.tabs[tab].rows;
        body.colCount = this.tabs[tab].cols;
      }
      const { rows, cells, ...layout } = body;
      const have = rowCountOf(rows ?? []);
      const ids = concatRuns(rows ?? [], legacyRows(entry.rowStart + have, entry.rowCount - have));
      const stored = writeChunk({ rows: ids, cells: cells ?? {}, layout, hasIds: true }, this.tabs[tab].colIds);
      staged.push(await encryptFrame(await frame(stored)));
      targets.push({ tab, chunk });
    }
    // New refs are known before upload, so a tab's index blob joins the same batch (same signature).
    const before = this.tabs.map((tab) => [...tab.chunks]);
    targets.forEach(({ tab, chunk }, i) => {
      const entry = this.tabs[tab].chunks[chunk];
      const body = this.cache.get(this.key(tab, chunk)) as ChunkBody;
      this.tabs[tab].chunks[chunk] = { ...entry, ref: staged[i].ref, cells: Object.keys(body.cells ?? {}).length };
    });
    const touched = [...new Set(targets.map((t) => t.tab))];
    const indexOf = new Map<number, number>();
    for (const tab of touched) {
      if (!this.tabs[tab].indexed) continue;
      indexOf.set(tab, staged.length);
      staged.push(await encryptFrame(await frame({ chunks: this.tabs[tab].chunks }, FRAME_TYPE_TAB_INDEX)));
    }
    let uploaded;
    try {
      uploaded = await uploadStaged(staged, options);
    } catch (error) {
      before.forEach((chunks, tab) => (this.tabs[tab].chunks = chunks)); // edits stay dirty; try again
      throw error;
    }
    const holders = new Set(this.root.servers);
    uploaded.forEach((blob) => blob.servers.forEach((server) => holders.add(server)));
    const root: SheetRootV2 = {
      ...this.root,
      updatedAt: Date.now(),
      servers: [...holders],
      tabs: this.root.tabs.map((tab, index) => {
        if (!touched.includes(index)) return tab;
        const at = indexOf.get(index);
        return { ...tab, cols: [...this.tabs[index].colIds], chunks: at === undefined ? [...this.tabs[index].chunks] : { index: uploaded[at].ref } };
      }),
    };
    Object.assign(this.root, root);
    this.dirty.clear();
    this.emit();
    return root;
  }
}

/** The cells with each position moved by `move` (null drops the cell). */
function moveCells(
  cells: Record<string, WorkbookCell> | undefined,
  move: (row: number, col: number) => [number, number] | null,
): Record<string, WorkbookCell> {
  const out: Record<string, WorkbookCell> = {};
  for (const [key, cell] of Object.entries(cells ?? {})) {
    const at = key.indexOf(":");
    const next = move(Number(key.slice(0, at)), Number(key.slice(at + 1)));
    if (next) out[`${next[0]}:${next[1]}`] = cell;
  }
  return out;
}
