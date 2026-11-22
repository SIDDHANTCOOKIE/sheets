import { describe, expect, it } from "vitest";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
import { calculateTyped, FormulaIndex } from "./bigFormulas";
import { editStructure } from "./bigStructure";
import { sortArea } from "./bigSort";
import { editLayout } from "./layoutEdit";
import { fakeBlossom, localSigner } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";

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

const START = { maxChunkBytes: 64 * 1024, headRawBytes: 100_000, rawCap: 1_000_000 };
const COLS = 8;

async function open(label: string) {
  const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: COLS, sharedStrings: false });
  const fake = fakeBlossom();
  const signer = localSigner();
  const servers = [`https://${label}.test`];
  const env = { fetch: fake.fetch, sleep: async () => {} };
  const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
    title: label, servers, signer, env, createWorker: inProcessWorker, start: START,
  });
  const store = new BigSheetStore(root, servers, env);
  await store.open();
  const reopen = async (from: BigSheetStore = store) => {
    const next = await from.save({ servers, signer, ...env });
    const again = new BigSheetStore(next, servers, env);
    await again.open();
    return again;
  };
  return { store, reopen };
}

/** Every cell's value, row by row (opens every chunk). */
async function snapshot(store: BigSheetStore, tab = 0): Promise<(unknown[] | null)[]> {
  const rows: (unknown[] | null)[] = [];
  for (let c = 0; c < store.tabs[tab].chunks.length; c++) await store.load(tab, c);
  for (let r = 0; r < store.tabs[tab].rows; r++) {
    const values = Array.from({ length: store.tabs[tab].cols }, (_, c) => store.cell(tab, r, c)?.value ?? null);
    rows.push(values.every((v) => v === null) ? null : values);
  }
  return rows;
}

describe("layout after inserting or deleting", () => {
  it("shifts sizes, hidden rows, merges and borders", () => {
    const layout = {
      rowHeights: { 2: 40, 10: 50 },
      hiddenRows: [3, 12],
      columnWidths: { 1: 200 },
      merges: [{ r: 8, c: 0, rs: 4, cs: 2 }],
      borders: [{ r1: 5, c1: 0, r2: 6, c2: 1, type: "all" as const }],
    };
    const inserted = editLayout(layout, { kind: "insert", axis: "row", index: 4, count: 3 });
    expect(inserted.rowHeights).toEqual({ 2: 40, 13: 50 });
    expect(inserted.hiddenRows).toEqual([3, 15]);
    expect(inserted.merges).toEqual([{ r: 11, c: 0, rs: 4, cs: 2 }]);
    expect(inserted.borders?.[0]).toMatchObject({ r1: 8, r2: 9 });
    expect(inserted.columnWidths).toEqual({ 1: 200 });
    // An insert inside a merged block makes it taller.
    expect(editLayout(layout, { kind: "insert", axis: "row", index: 10, count: 2 }).merges).toEqual([{ r: 8, c: 0, rs: 6, cs: 2 }]);
    // Deleting rows of a block shrinks it; deleting all of it removes it.
    expect(editLayout(layout, { kind: "delete", axis: "row", index: 9, count: 2 }).merges).toEqual([{ r: 8, c: 0, rs: 2, cs: 2 }]);
    expect(editLayout(layout, { kind: "delete", axis: "row", index: 8, count: 4 }).merges).toBeUndefined();
    // Deleting a row drops its height and moves the ones below up.
    expect(editLayout(layout, { kind: "delete", axis: "row", index: 2, count: 1 }).rowHeights).toEqual({ 9: 50 });
    const columns = editLayout(layout, { kind: "insert", axis: "col", index: 1, count: 1 });
    expect(columns.columnWidths).toEqual({ 2: 200 });
    expect(columns.merges).toEqual([{ r: 8, c: 0, rs: 4, cs: 3 }]);
  });
});

describe("inserting and deleting in a big sheet", () => {
  it("inserts rows in the middle, keeps everything below, and survives save and reload", async () => {
    const { store, reopen } = await open("struct-insert");
    expect(store.tabs[0].chunks.length).toBeGreaterThan(3);
    const before = await snapshot(store);
    const at = store.tabs[0].chunks[1].rowStart + 7;
    await store.setLayout(0, { rowHeights: { [at + 1]: 70 } });

    await store.applyEdit(0, { kind: "insert", axis: "row", index: at, count: 5 });
    expect(store.tabs[0].rows).toBe(before.length + 5);
    const expected = [...before.slice(0, at), ...Array(5).fill(null), ...before.slice(at)];
    expect(await snapshot(store)).toEqual(expected);
    expect(store.tabs[0].layout.rowHeights).toEqual({ [at + 6]: 70 });

    const again = await reopen();
    expect(again.tabs[0].rows).toBe(before.length + 5);
    expect(await snapshot(again)).toEqual(expected);
    expect(again.tabs[0].layout.rowHeights).toEqual({ [at + 6]: 70 });
  }, 120_000);

  it("deletes rows across chunks, including whole chunks, and survives save and reload", async () => {
    const { store, reopen } = await open("struct-delete");
    const chunks = store.tabs[0].chunks;
    expect(chunks.length).toBeGreaterThan(4);
    const before = await snapshot(store);
    const from = chunks[1].rowStart + 3;
    const to = chunks[3].rowStart + 4; // exclusive: chunk 2 is deleted whole
    await store.applyEdit(0, { kind: "delete", axis: "row", index: from, count: to - from });
    const expected = [...before.slice(0, from), ...before.slice(to)];
    expect(store.tabs[0].rows).toBe(expected.length);
    expect(await snapshot(store)).toEqual(expected);

    const again = await reopen();
    expect(again.tabs[0].rows).toBe(expected.length);
    expect(await snapshot(again)).toEqual(expected);
  }, 120_000);

  it("inserts and deletes columns, in loaded and unloaded chunks, through save and reload", async () => {
    const { store, reopen } = await open("struct-cols");
    const before = await snapshot(store);
    const cols0 = store.tabs[0].cols;
    await store.applyEdit(0, { kind: "insert", axis: "col", index: 2, count: 2 });
    const widen = (row: unknown[] | null) => (row ? [...row.slice(0, 2), null, null, ...row.slice(2)] : null);
    const inserted = before.map(widen);
    expect(store.tabs[0].cols).toBe(cols0 + 2);
    expect(await snapshot(store)).toEqual(inserted);

    const again = await reopen();
    expect(await snapshot(again)).toEqual(inserted);

    await again.applyEdit(0, { kind: "delete", axis: "col", index: 0, count: 1 });
    const trimmed = inserted.map((row) => (row ? row.slice(1) : null));
    expect(await snapshot(again)).toEqual(trimmed);
    // Read fresh from the saved sheet, the deleted column does not come back.
    const fresh = await reopen(again);
    expect(await snapshot(fresh)).toEqual(trimmed);
    expect(fresh.tabs[0].cols).toBe(cols0 + 1);
  }, 120_000);

  it("rewrites formulas, moves them with their rows, and gives #REF! for deleted cells", async () => {
    const { store, reopen } = await open("struct-formulas");
    const rowOf = (chunk: number, extra: number) => store.tabs[0].chunks[chunk].rowStart + extra;
    const source = rowOf(2, 10);
    const formulaRow = rowOf(1, 4);
    await store.load(0, store.chunkFor(0, source));
    await store.load(0, store.chunkFor(0, formulaRow));
    await store.setCell(0, source, 3, "10");
    const index = new FormulaIndex();
    const type = async (row: number, col: number, formula: string) => {
      const result = await calculateTyped(store, 0, row, col, formula);
      await store.edit(0, row, col, row, col, () => ({ formula, value: result.value }));
      await index.set(store, 0, row, col, formula);
    };
    await type(formulaRow, 5, `D${source + 1}*2`); // F = D(source)*2
    await type(formulaRow, 6, `SUM(D${source - 1}:D${source + 3})`); // a range around the source
    expect(store.cell(0, formulaRow, 5)?.value).toBe(20);

    // Insert 3 rows above the source: the formulas keep pointing at the same cells.
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: source - 3, count: 3 });
    expect(store.cell(0, source + 3, 3)?.value).toBe(10);
    expect(store.cell(0, formulaRow, 5)?.formula).toBe(`D${source + 4}*2`);
    expect(store.cell(0, formulaRow, 5)?.value).toBe(20);
    expect(store.cell(0, formulaRow, 6)?.formula).toBe(`SUM(D${source + 2}:D${source + 6})`);

    // Insert above the formula: it moves down and still reads the same cell.
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: formulaRow, count: 2 });
    expect(store.cell(0, formulaRow + 2, 5)?.formula).toBe(`D${source + 6}*2`);
    expect(store.cell(0, formulaRow + 2, 5)?.value).toBe(20);

    const again = await reopen();
    await again.load(0, again.chunkFor(0, formulaRow + 2));
    expect(again.cell(0, formulaRow + 2, 5)?.formula).toBe(`D${source + 6}*2`);

    // Delete the row the first formula reads: it becomes #REF!.
    await editStructure(store, index, 0, { kind: "delete", axis: "row", index: source + 5, count: 1 });
    const broken = store.cell(0, formulaRow + 2, 5);
    expect(broken?.formula).toContain("#REF!");
    expect(broken?.type).toBe("error");
  }, 180_000);
});

describe("sorting a selection", () => {
  it("sorts rows by a column, keeps empty cells last, and undoes in one step", async () => {
    const { store } = await open("struct-sort");
    await store.load(0, 0);
    const top = 10;
    const input = [["pear", 3], ["Apple", 10], [undefined, 1], ["fig", 2], ["banana", 9]] as const;
    for (let i = 0; i < input.length; i++) {
      const [name, n] = input[i];
      await store.setCell(0, top + i, 0, name ?? "");
      await store.setCell(0, top + i, 1, String(n));
    }
    const area = { r1: top, c1: 0, r2: top + 4, c2: 1 };
    const read = () => Array.from({ length: 5 }, (_, i) => [store.cell(0, top + i, 0)?.value ?? null, store.cell(0, top + i, 1)?.value]);

    expect(await sortArea(store, 0, area, 0, false)).toBe(true);
    expect(read().map((r) => r[0])).toEqual(["Apple", "banana", "fig", "pear", null]);
    expect(read().map((r) => r[1])).toEqual([10, 9, 2, 3, 1]); // rows moved whole

    await sortArea(store, 0, area, 1, true);
    expect(read().map((r) => r[1])).toEqual([10, 9, 3, 2, 1]);

    store.undo();
    expect(read().map((r) => r[0])).toEqual(["Apple", "banana", "fig", "pear", null]);

    await store.edit(0, top, 2, top, 2, () => ({ formula: "A1", value: 1 }));
    await expect(sortArea(store, 0, { ...area, c2: 2 }, 0, false)).rejects.toThrow(/formulas/);
    await expect(sortArea(store, 0, area, 0, false, [{ r: top, c: 0, rs: 2, cs: 1 }])).rejects.toThrow(/merged/);
  }, 120_000);
});

describe("links and comments", () => {
  it("keep their data through save and reload and move with inserted rows", async () => {
    const { store, reopen } = await open("struct-notes");
    await store.load(0, 0);
    await store.edit(0, 5, 2, 5, 2, () => ({ value: "site", link: { url: "https://example.com" }, comment: "check me", validation: { type: "list", options: ["site", "other"] } }));
    await store.applyEdit(0, { kind: "insert", axis: "row", index: 2, count: 2 });
    const again = await reopen();
    await again.load(0, 0);
    expect(again.cell(0, 7, 2)).toMatchObject({ value: "site", link: { url: "https://example.com" }, comment: "check me", validation: { type: "list", options: ["site", "other"] } });
    expect(again.cell(0, 5, 2)?.comment).toBeUndefined();
  }, 120_000);
});
