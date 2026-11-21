// Differential sweep: runs every FortuneSheet function name through its engine and ours with many
// argument shapes and writes a JSON report. Only runs when SWEEP_OUT is set (it is a tool, not a test).
import { createRequire } from "node:module";
import { readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, it } from "vitest";
import { api, defaultContext, ensureSheetIndex, initSheetIndex, type Context, type Sheet } from "@fortune-sheet/core";
import { calculate, FormulaError } from "../../../packages/sheets-formula/src";

const OUT = process.env.SWEEP_OUT ?? "sweep.json";

function ctxFor(): Context {
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
  const data = Array.from({ length: 60 }, () => Array.from({ length: 20 }, () => null));
  const sheets = [{ id: "s1", name: "Sheet1", order: 0, status: 1, row: 60, column: 20, data }] as unknown as Sheet[];
  ensureSheetIndex(sheets, () => "x");
  ctx.luckysheetfile = sheets;
  ctx.currentSheetId = "s1";
  initSheetIndex(ctx);
  return ctx;
}

const GRID: Record<string, number | string> = {};
const A = [3, 1, 4, 1, 5];
const B = [2, 7, 1, 8, 2];
const C = ["alpha", "beta", "gamma", "delta", "eps"];
A.forEach((v, r) => (GRID[`${r}:0`] = v));
B.forEach((v, r) => (GRID[`${r}:1`] = v));
C.forEach((v, r) => (GRID[`${r}:2`] = v));

const CANDIDATES = [
  "A1:A5", "A1:A5,B1:B5", "A1", "A1,B1", "0.5", "0.5,0.5", "2,3", "0.5,2,3", "5,0.05,3", '"text"', '"text","x"', "C1:C5",
  "DATE(2024,3,15)", "DATE(2024,3,15),DATE(2024,6,1)", "1,2,3,4", "A1:A5,2", "A1:A5,0.5", "A1:A5,B1:B5,2", "2,A1:A5",
  "1,2,3,TRUE", "10,5,2,TRUE", '"b",C1:C5,A1:A5', "3,A1:B5,2,FALSE", "TRUE", "0.05,12,-1000", "-1000,300,400,500",
  '"2024-03-15"', "10", "1,A1:A5", "0.1,A1:A5", '"a,b,c",","', '"hello",2', '"hello",2,3', "100,0.05,5,1000", "5,10,2",
  "1,2,3,4,5,6", "A1:A5,B1:B5,C1:C5", "0.3,5,10", "2,5,0.5,TRUE", "5,2,TRUE", '"x","y","z"',
];

const VOLATILE = new Set(["RAND", "RANDBETWEEN", "NOW", "TODAY"]);
const ERR = /^(#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A|#NULL!|Err:.*)$/;

type Out = { kind: "err"; text: string } | { kind: "val"; v: unknown };

function theirs(ctx: Context, formula: string): Out {
  try {
    api.setCellValue(ctx, 0, 10, formula, null);
    const cell = (ctx.luckysheetfile[0].data as { v?: unknown; m?: unknown }[][])[0][10];
    if (!cell) return { kind: "err", text: "empty" };
    const v = cell.v ?? cell.m;
    if (typeof v === "string" && ERR.test(v)) return { kind: "err", text: v };
    return { kind: "val", v };
  } catch (error) {
    return { kind: "err", text: `throw:${String(error).slice(0, 40)}` };
  }
}

function ours(formula: string): Out {
  try {
    const value = calculate(formula, {
      sheet: "Sheet1",
      getCell: (_s, r, c) => (GRID[`${r}:${c}`] as number | string | undefined) ?? null,
      extent: () => ({ rows: 5, cols: 3 }),
      now: () => new Date(Date.UTC(2024, 4, 15, 12)),
      random: () => 0.5,
    });
    if (value instanceof FormulaError) return { kind: "err", text: value.code };
    return { kind: "val", v: value };
  } catch (error) {
    return { kind: "err", text: `throw:${String(error).slice(0, 40)}` };
  }
}

function same(a: unknown, b: unknown): boolean {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  if (typeof a === "number" && typeof b === "string") return String(a) === b || Number(b) === a;
  if (typeof a === "string" && typeof b === "number") return same(b, a);
  if (a == null || a === "") return b == null || b === "";
  if (b == null || b === "") return false;
  return String(a) === String(b);
}

describe.skipIf(!process.env.SWEEP_OUT)("sweep", () => {
  it("every function", () => {
    const pnpm = path.resolve("node_modules/.pnpm");
    const dir = readdirSync(pnpm).find((name) => name.startsWith("@formulajs+formulajs@2.9.3")) as string;
    const old = createRequire(path.join(pnpm, dir, "node_modules/@formulajs/formulajs/package.json"))("@formulajs/formulajs");
    const names = Object.keys(old).filter((key) => typeof old[key] === "function" && /^[A-Z][A-Z0-9_.]*$/.test(key));
    const ctx = ctxFor();
    for (const [key, v] of Object.entries(GRID)) {
      const [r, c] = key.split(":").map(Number);
      api.setCellValue(ctx, r, c, v, null);
    }
    const report: Record<string, unknown> = {};
    for (const name of names) {
      if (VOLATILE.has(name)) continue;
      const disagree: unknown[] = [];
      const oursOnlyError: unknown[] = [];
      const theirsOnlyError: unknown[] = [];
      let agree = 0;
      let bothError = 0;
      for (const args of CANDIDATES) {
        const formula = `=${name}(${args})`;
        const t = theirs(ctx, formula);
        const o = ours(formula);
        if (t.kind === "err" && o.kind === "err") bothError++;
        else if (t.kind === "val" && o.kind === "val") {
          if (same(t.v, o.v)) agree++;
          else disagree.push({ args, theirs: t.v, ours: o.v });
        } else if (o.kind === "err") oursOnlyError.push({ args, theirs: (t as { v: unknown }).v, ours: o.text });
        else theirsOnlyError.push({ args, theirs: (t as { text: string }).text, ours: o.v });
      }
      report[name] = { agree, bothError, disagree, oursOnlyError, theirsOnlyError };
    }
    writeFileSync(OUT, JSON.stringify({ count: names.length, report }, null, 1));
  }, 600_000);
});
