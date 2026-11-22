// Inserting and deleting rows and columns in a big sheet, formulas included.
//
// Cells, ids and the saved look move in the store (BigSheetStore.applyEdit). Here the formulas follow:
// every formula in the workbook is rewritten so that it still points at the same cells (or at #REF!
// when they were deleted), its own cell moves with its row/column, and formulas that lost cells are
// recalculated. The formula index must know every formula first, so the first edit looks through
// every chunk (like the first formula edit does).
import type { BigSheetStore } from "./bigSheet";
import { FormulaIndex, refreshDependents, type ChangedRect, type RefreshProgress } from "./bigFormulas";
import { moveIndex, type StructEdit } from "./layoutEdit";

/** Formulas that change text sit in chunks that stay in memory until saved; this keeps that bounded. */
const MAX_REWRITE_CHUNKS = 24;

interface Plan {
  tab: number;
  row: number;
  col: number;
  newRow: number;
  newCol: number;
  formula: string;
  changed: boolean;
}

export async function editStructure(
  store: BigSheetStore,
  index: FormulaIndex,
  tab: number,
  edit: StructEdit,
  onProgress?: RefreshProgress,
): Promise<void> {
  const engine = await import("@formstr/sheets-formula");
  await index.ensureScanned(store, onProgress);
  const sheet = store.tabs[tab].name;
  const formulaEdit = { ...edit, sheet };

  // Work out what every formula becomes before anything moves.
  const plans: Plan[] = [];
  for (const entry of index.all()) {
    let formula = entry.formula;
    try {
      formula = engine.print(engine.rewriteForEdit(engine.parse(entry.formula), formulaEdit, store.tabs[entry.tab].name)).replace(/^=/, "");
    } catch {
      // text the engine cannot read is left as it is
    }
    let newRow = entry.row;
    let newCol = entry.col;
    if (entry.tab === tab) {
      const moved = moveIndex(edit.axis === "row" ? entry.row : entry.col, edit);
      if (moved === null) continue; // the formula's own cell is deleted
      if (edit.axis === "row") newRow = moved;
      else newCol = moved;
    }
    plans.push({ tab: entry.tab, row: entry.row, col: entry.col, newRow, newCol, formula, changed: formula !== entry.formula });
  }
  const touched = new Set(plans.filter((p) => p.changed).map((p) => `${p.tab}:${store.chunkFor(p.tab, p.row)}`));
  if (touched.size > MAX_REWRITE_CHUNKS) {
    throw new Error(
      "Formulas that point at the rows or columns being moved are spread over too many parts of this sheet to update at once. Try a smaller sheet, or move the formulas together first.",
    );
  }

  await store.applyEdit(tab, edit);

  for (const plan of plans) {
    if (!plan.changed) continue;
    await store.load(plan.tab, store.chunkFor(plan.tab, plan.newRow));
    const cell = store.cell(plan.tab, plan.newRow, plan.newCol);
    if (cell) await store.setDerived(plan.tab, plan.newRow, plan.newCol, { ...cell, formula: plan.formula });
  }

  // Every formula is known and now sits at its new place: rebuild the index from the plans.
  index.reset();
  for (const plan of plans) await index.set(store, plan.tab, plan.newRow, plan.newCol, plan.formula);

  // Formulas that lost cells (or whose range shrank) get new results.
  if (edit.kind === "delete") {
    const stale: ChangedRect[] = plans
      .filter((p) => p.changed)
      .map((p) => ({ tab: p.tab, r1: p.newRow, c1: p.newCol, r2: p.newRow, c2: p.newCol }));
    if (stale.length > 0) await refreshDependents(store, index, stale, onProgress, true);
  }
}
