// Differential parity: every scenario ran once on FortuneSheet's own engine (headless); its answers are
// recorded in reference.json. The big-sheet store must end each scenario with the same workbook.
// DIVERGENCES lists the places where FortuneSheet itself behaves oddly (or where the two models differ on
// purpose); each has a reason, and the big-sheet result is asserted too, so a divergence cannot silently change.
import { describe, expect, it } from "vitest";
import { runOnBigStore } from "./bigGrid";
import { sample, SCENARIOS, view, type View } from "./scenarios";
import reference from "./reference.json";

const REFERENCE = reference as unknown as Record<string, View>;

/** Scenarios where FortuneSheet and the big-sheet store differ, with the reason. The big-sheet answer is pinned. */
const withValue = (reference: View, cell: string, value: string | number): View => ({
  ...reference,
  cells: { ...reference.cells, [cell]: { ...reference.cells[cell], value } },
});

const DIVERGENCES: Record<string, { reason: string; big: (reference: View) => View }> = {
  // FortuneSheet's "both" freeze counts the row differently from its own "row" freeze (rows 1 for row index 1,
  // but columns 1 for column index 0). The big-sheet freeze always means "through this row/column".
  "freeze two rows and a column": {
    reason: "FortuneSheet's 'both' freeze is inconsistent with its 'row' freeze; ours freezes through the given row and column.",
    big: (reference) => ({ ...reference, frozen: { rows: 2, columns: 1 } }),
  },
  // FortuneSheet rewrites the formula text when rows are deleted but keeps the old RESULT (600 for
  // =#REF!*2). Excel and the big-sheet store recalculate: the result is the error.
  "a formula whose target row is deleted": {
    reason: "FortuneSheet leaves a stale result after a delete; we recalculate (#REF!).",
    big: (reference) => withValue(reference, "4:1", "#REF!"),
  },
  // Same cause: SUM(B2:B3) after the middle row went must be 1200 + 80... FortuneSheet still shows the old 1580.
  "a range shrinks when a row inside it is deleted": {
    reason: "FortuneSheet leaves a stale result after a delete; we recalculate (1280).",
    big: (reference) => withValue(reference, "7:1", 1280),
  },
};

// Where FortuneSheet's own answer is wrong or missing, ours follows Excel.
const FORTUNE_WRONG: Record<string, { reason: string; value: string | number }> = {
  '=VLOOKUP("Food",A2:B4,2,FALSE)': { reason: "FortuneSheet answers #NAME? for VLOOKUP with FALSE.", value: 300 },
  '=IFERROR(B2/0,"oops")': { reason: 'FortuneSheet returns "DIV/0" instead of the fallback.', value: "oops" },
  '=TEXT(B2,"0.00")': { reason: "FortuneSheet answers #ERROR! for TEXT with a plain number format.", value: "1200.00" },
  "=ROW(B3)": { reason: "FortuneSheet answers #NAME? for ROW(ref).", value: 3 },
  "=COLUMN(C1)": { reason: "FortuneSheet answers #N/A for COLUMN(ref).", value: 3 },
};
for (const [formula, { reason, value }] of Object.entries(FORTUNE_WRONG)) {
  DIVERGENCES[`formula ${formula}`] = { reason, big: (reference) => withValue(reference, "8:3", value) };
}

describe("parity: FortuneSheet's recorded answers vs the big-sheet store", () => {
  for (const [name, steps] of Object.entries(SCENARIOS)) {
    it(name, async () => {
      const expected = REFERENCE[name];
      expect(expected, `no recorded answer for "${name}"`).toBeDefined();
      const candidate = view(await runOnBigStore(sample(), steps));
      const divergence = DIVERGENCES[name];
      if (divergence) expect(candidate).toEqual(divergence.big(expected));
      else expect(candidate).toEqual(expected);
    }, 60_000);
  }
});
