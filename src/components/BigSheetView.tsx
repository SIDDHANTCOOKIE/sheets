// Grid for sheets too big for the main editor. The canvas grid (@formstr/sheets-grid) draws only the
// rows on screen and only the chunks holding them are downloaded (BigSheetStore), so an 11M-cell
// sheet scrolls in a normal browser tab. Editing is switched by BIG_SHEET_EDITING.
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { addBorder, columnName, Grid, type BorderType, type GridData, type Selection } from "@formstr/sheets-grid";
import { AlertTriangle, Loader2, Save } from "lucide-react";
import { BIG_SHEET_EDITING } from "../lib/constants";
import { BigSheetStore } from "../lib/v2/bigSheet";
import { calculateTyped, FormulaIndex, pointRect, refreshDependents, type ChangedRect } from "../lib/v2/bigFormulas";
import {
  copyBlock,
  copyBlockNow,
  fillArea,
  fromTsv,
  isLoadedArea,
  pasteBlock,
  pasteText,
  type Block,
} from "../lib/v2/bigClipboard";
import type { SheetRootV2 } from "../lib/v2/root";
import { uploadProgress } from "../lib/v2/progress";
import type { LibraryEntry } from "../lib/types";
import type { WorkbookCell } from "../lib/workbook";
import { cellDisplay } from "../lib/workbook/format";
import { BigSheetToolbar, shiftDecimals, type StylePatch } from "./BigSheetToolbar";
import { useLibrary } from "../providers/LibraryProvider";
import { useToast } from "../providers/ToastProvider";
import { cx } from "../lib/utils";

function shown(cell: WorkbookCell | null | undefined): string {
  return cell ? cellDisplay(cell) : "";
}

/** A cell after a style change: the cached display text is dropped so the new number format shows. */
function restyle(cell: WorkbookCell | undefined, patch: StylePatch): WorkbookCell {
  const { display: _display, ...rest } = cell ?? {};
  void _display;
  const style = { ...rest.style, ...patch };
  for (const key of Object.keys(style) as (keyof typeof style)[]) if (style[key] === undefined) delete style[key];
  return { ...rest, style };
}

/** Gives the grid what the tab saved about its look: column widths, row heights, hidden rows/columns, frozen panes. */
function applyLayout(grid: Grid, store: BigSheetStore, tab: number): void {
  const layout = store.tabs[tab]?.layout ?? {};
  const overrides = (sizes?: Record<number, number>, hidden?: number[]) => [
    ...Object.entries(sizes ?? {}).map(([index, size]) => ({ index: Number(index), size })),
    ...(hidden ?? []).map((index) => ({ index, size: 0 })),
  ];
  grid.setSizes("col", overrides(layout.columnWidths, layout.hiddenColumns));
  grid.setSizes("row", overrides(layout.rowHeights, layout.hiddenRows));
  grid.setFrozen(layout.frozen?.rows ?? 0, layout.frozen?.columns ?? 0);
  grid.setMerges(layout.merges ?? []);
  grid.setBorders(layout.borders ?? []);
}

/** The store's tab as the grid reads it. */
function tabData(store: BigSheetStore, tab: number): GridData {
  return {
    get rowCount() {
      return store.tabs[tab]?.rows ?? 0;
    },
    get colCount() {
      return store.tabs[tab]?.cols ?? 0;
    },
    cell(row, col) {
      if (!store.isLoaded(tab, row)) return undefined;
      const cell = store.cell(tab, row, col);
      return cell ? { text: shown(cell), style: cell.style } : null;
    },
    ensure(first, end) {
      const info = store.tabs[tab];
      if (!info || info.chunks.length === 0 || end <= first) return;
      const from = store.chunkFor(tab, first);
      const to = store.chunkFor(tab, Math.max(first, end - 1));
      for (let c = from; c <= to; c++) void store.load(tab, c).catch(() => {});
    },
    subscribe: (listener) => store.subscribe(listener),
  };
}

export function BigSheetView({ entry, root }: { entry: LibraryEntry; root: SheetRootV2 }) {
  const library = useLibrary();
  const { toast } = useToast();
  // Created once per root: recreating it (e.g. when the server list loads) would drop unsaved edits.
  const serversRef = useRef(library.blossomServers);
  const store = useMemo(() => new BigSheetStore(root, serversRef.current), [root]);
  const [ready, setReady] = useState(false);
  const [openError, setOpenError] = useState("");
  const [tab, setTab] = useState(0);
  const [, setVersion] = useState(0);
  const [selection, setSelection] = useState<Selection | null>(null);
  /** Style copied by the paint-format button, applied to the next selection. */
  const [painting, setPainting] = useState<WorkbookCell["style"] | null>(null);
  const [editing, setEditing] = useState<{ row: number; col: number; text: string } | null>(null);
  const [saving, setSaving] = useState(false);
  /** Right-click menu: where it is on the page. */
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  /** Progress of a formula that has to read several chunks, e.g. =SUM(A:A). */
  const [calculating, setCalculating] = useState<string | null>(null);
  const hostRef = useRef<HTMLDivElement>(null);
  const gridRef = useRef<Grid | null>(null);
  /** Set by Escape so the blur that follows the input unmounting does not save the edit. */
  const cancelled = useRef(false);
  /** The cells of the last copy, so pasting them keeps formulas and formats. */
  const clipboard = useRef<Block | null>(null);
  const canEdit = BIG_SHEET_EDITING && !!entry.editKey;
  // The grid's callbacks are created once, so they read the latest state through these refs.
  const live = useRef({ canEdit, editing, painting, tab });
  live.current = { canEdit, editing, painting, tab };

  useEffect(() => {
    let cancelled = false;
    setReady(false);
    store
      .open()
      .then(() => !cancelled && setReady(true))
      .catch((error) => !cancelled && setOpenError(error instanceof Error ? error.message : String(error)));
    // Toolbar state (undo/redo, dirty count) follows the store; the grid redraws itself.
    const off = store.subscribe(() => setVersion((n) => n + 1));
    return () => {
      cancelled = true;
      off();
    };
  }, [store]);

  const info = store.tabs[tab];

  // Which cells hold formulas, so a change can refresh the formulas that read it.
  const index = useMemo(() => new FormulaIndex(), [store]);
  // Refreshes run one after another, so two quick edits never recalculate the same cells at once.
  const refreshQueue = useRef<Promise<void>>(Promise.resolve());
  const refresh = useCallback(
    (areas: ChangedRect[], recalcInside = false): Promise<void> => {
      const job = refreshQueue.current.then(async () => {
        try {
          const count = await refreshDependents(
            store,
            index,
            areas,
            (done, total, phase) =>
              setCalculating(
                phase === "scan"
                  ? `Looking for formulas… ${done} of ${total} parts`
                  : total > 1
                    ? `Updating formulas… ${done.toLocaleString()} of ${total.toLocaleString()}`
                    : null,
              ),
            recalcInside,
          );
          if (count > 0) setVersion((n) => n + 1);
        } catch (error) {
          toast(error instanceof Error ? error.message : "Couldn't update formulas.", "error");
        } finally {
          setCalculating(null);
        }
      });
      refreshQueue.current = job;
      return job;
    },
    [index, store, toast],
  );
  const refreshRef = useRef(refresh);
  refreshRef.current = refresh;
  useEffect(() => {
    if (!ready) return;
    store.onChunkLoaded = (tabNo, chunk) => void index.scanChunk(store, tabNo, chunk);
    for (const [tabNo, chunk] of store.loadedChunks()) void index.scanChunk(store, tabNo, chunk);
    return () => {
      store.onChunkLoaded = undefined;
    };
  }, [ready, store, index]);

  const range = selection?.ranges[selection.ranges.length - 1] ?? null;
  const active = selection?.active ?? null;

  const editRange = useCallback(
    (change: (cell: WorkbookCell | undefined) => WorkbookCell | undefined, target?: Selection | null) => {
      const area = (target ?? selection)?.ranges[(target ?? selection)!.ranges.length - 1];
      if (!area || !canEdit) return;
      store
        .edit(tab, area.r1, area.c1, area.r2, area.c2, change)
        .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
    },
    [canEdit, selection, store, tab, toast],
  );

  // One grid for the life of the view; tabs swap its data.
  useEffect(() => {
    const host = hostRef.current;
    if (!ready || !host || !store.tabs[0]) return;
    /** Clears values (keeps formats), then refreshes the formulas that read the cleared cells. */
    const clearArea = (area: { r1: number; c1: number; r2: number; c2: number }) => {
      if (!live.current.canEdit) return;
      const tabNo = live.current.tab;
      store
        .edit(tabNo, area.r1, area.c1, area.r2, area.c2, (cell) => (cell?.style ? { style: cell.style } : undefined))
        .then(() => refreshRef.current([{ tab: tabNo, r1: area.r1, c1: area.c1, r2: area.r2, c2: area.c2 }]))
        .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
    };
    const grid = new Grid(host, {
      data: tabData(store, live.current.tab),
      onSelect: (next) => {
        setSelection(next);
        const style = live.current.painting;
        if (style !== null && live.current.canEdit) {
          setPainting(null);
          const area = next.ranges[next.ranges.length - 1];
          store
            .edit(live.current.tab, area.r1, area.c1, area.r2, area.c2, (cell) => ({ ...(cell ?? {}), style: style ? { ...style } : undefined }))
            .catch(() => {});
        }
      },
      onEdit: (cell, text) => {
        if (!live.current.canEdit || !store.isLoaded(live.current.tab, cell.row)) return;
        const existing = store.cell(live.current.tab, cell.row, cell.col);
        const initial = text ?? (existing?.formula ? `=${existing.formula.replace(/^=/, "")}` : shown(existing));
        cancelled.current = false;
        setEditing({ row: cell.row, col: cell.col, text: initial });
      },
      onClear: (next) => clearArea(next.ranges[next.ranges.length - 1]),
      onContextMenu: (info) => setMenu({ x: info.x, y: info.y }),
      onResize: (kind, index, size) => {
        // Resizing in a read-only sheet only lasts until the page is closed.
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        const layout = store.tabs[tabNo].layout;
        const patch = kind === "col" ? { columnWidths: { ...layout.columnWidths, [index]: size } } : { rowHeights: { ...layout.rowHeights, [index]: size } };
        store.setLayout(tabNo, patch).catch(() => {});
      },
      onCopy: (area, cut) => {
        if (cut && !live.current.canEdit) return undefined;
        const tabNo = live.current.tab;
        const finish = (block: Block) => {
          clipboard.current = block;
          if (cut) clearArea(area);
          return block.text;
        };
        try {
          // Everything in memory: answer inside the copy event. Otherwise download first.
          if (isLoadedArea(store, tabNo, area)) return finish(copyBlockNow(store, tabNo, area));
          return copyBlock(store, tabNo, area).then(finish, (error) => {
            toast(error instanceof Error ? error.message : String(error), "error");
            return undefined;
          });
        } catch (error) {
          toast(error instanceof Error ? error.message : String(error), "error");
          return undefined;
        }
      },
      onPaste: (text, next) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        const area = next.ranges[next.ranges.length - 1];
        const copied = clipboard.current;
        const normal = text.replace(/\r\n?/g, "\n");
        const job =
          copied && copied.text === normal
            ? pasteBlock(store, tabNo, area.r1, area.c1, copied, area)
            : pasteText(store, tabNo, area.r1, area.c1, fromTsv(normal), area);
        job
          .then(async (done) => {
            const written = { tab: tabNo, ...done };
            gridRef.current?.select({
              ranges: [done],
              active: { row: done.r1, col: done.c1 },
              anchor: { row: done.r1, col: done.c1 },
            });
            await refreshRef.current([written], true);
          })
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
      onFill: (direction, next) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        fillArea(store, tabNo, next.ranges[next.ranges.length - 1], direction)
          .then(async (done) => {
            if (done) await refreshRef.current([{ tab: tabNo, ...done }], true);
          })
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
    });
    gridRef.current = grid;
    applyLayout(grid, store, live.current.tab);
    return () => {
      grid.destroy();
      gridRef.current = null;
    };
    // The tab is read through `live`; switching tabs calls setData below, not a rebuild.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ready, store]);

  const switchTab = (index: number) => {
    setTab(index);
    setSelection(null);
    setEditing(null);
    gridRef.current?.setData(tabData(store, index));
    if (gridRef.current) applyLayout(gridRef.current, store, index);
  };

  /** Keeps the first rows/columns in view while scrolling; remembered with the sheet when it can be edited. */
  const freeze = (rows: number, columns: number) => {
    gridRef.current?.setFrozen(rows, columns);
    if (canEdit) store.setLayout(tab, { frozen: rows > 0 || columns > 0 ? { rows, columns } : undefined }).catch(() => {});
    setVersion((n) => n + 1);
  };
  /** Puts a border on the selection (or takes them off with "none"). */
  const border = (type: BorderType, color: string, style: string) => {
    if (!canEdit || !range) return;
    const layout = store.tabs[tab].layout;
    const next = addBorder(layout.borders ?? [], { ...range, type, color, style });
    store.setLayout(tab, { borders: next.length > 0 ? next : undefined }).then(() => gridRef.current && applyLayout(gridRef.current, store, tab)).catch(() => {});
  };

  /**
   * Merges the selection into one cell (it keeps the top-left value), or splits merged cells.
   * "all" merges everything, "rows" each row on its own, "columns" each column on its own.
   */
  const merge = async (how: "all" | "rows" | "columns" | "split") => {
    if (!canEdit || !range) return;
    const layout = store.tabs[tab].layout;
    const touches = (m: { r: number; c: number; rs: number; cs: number }) =>
      m.r <= range.r2 && m.r + m.rs - 1 >= range.r1 && m.c <= range.c2 && m.c + m.cs - 1 >= range.c1;
    const kept = (layout.merges ?? []).filter((m) => !touches(m));
    const added: { r: number; c: number; rs: number; cs: number }[] = [];
    if (how === "all") added.push({ r: range.r1, c: range.c1, rs: range.r2 - range.r1 + 1, cs: range.c2 - range.c1 + 1 });
    if (how === "rows") for (let r = range.r1; r <= range.r2; r++) added.push({ r, c: range.c1, rs: 1, cs: range.c2 - range.c1 + 1 });
    if (how === "columns") for (let c = range.c1; c <= range.c2; c++) added.push({ r: range.r1, c, rs: range.r2 - range.r1 + 1, cs: 1 });
    const blocks = added.filter((m) => m.rs * m.cs > 1);
    // Nothing to merge and nothing to split: leave the sheet as it is (no "unsaved changes").
    if (blocks.length === 0 && kept.length === (layout.merges ?? []).length) return;
    if (blocks.length > 20_000) {
      toast("That would merge too many cells at once.", "error");
      return;
    }
    try {
      // The other cells of a block give up their values; the top-left one is kept.
      for (const m of blocks) {
        const clear = (cell: WorkbookCell | undefined) => (cell?.style ? { style: cell.style } : undefined);
        if (m.rs > 1) await store.edit(tab, m.r + 1, m.c, m.r + m.rs - 1, m.c + m.cs - 1, clear);
        if (m.cs > 1) await store.edit(tab, m.r, m.c + 1, m.r, m.c + m.cs - 1, clear);
      }
      await store.setLayout(tab, { merges: [...kept, ...blocks].length > 0 ? [...kept, ...blocks] : undefined });
      if (gridRef.current) applyLayout(gridRef.current, store, tab);
      await refresh(blocks.map((m) => ({ tab, r1: m.r, c1: m.c, r2: m.r + m.rs - 1, c2: m.c + m.cs - 1 })));
    } catch (error) {
      toast(error instanceof Error ? error.message : String(error), "error");
    }
  };
  const mergedHere = !!range && !!store.tabs[tab]?.layout.merges?.some((m) => m.r <= range.r2 && m.r + m.rs - 1 >= range.r1 && m.c <= range.c2 && m.c + m.cs - 1 >= range.c1);

  /** Hides or shows the rows/columns of the selection (showing also reveals the hidden ones just beside it). */
  const setHidden = (axis: "rows" | "columns", hide: boolean) => {
    if (!canEdit || !range) return;
    const key = axis === "rows" ? "hiddenRows" : "hiddenColumns";
    const from = axis === "rows" ? range.r1 : range.c1;
    const to = axis === "rows" ? range.r2 : range.c2;
    const count = axis === "rows" ? store.tabs[tab].rows : store.tabs[tab].cols;
    const current = store.tabs[tab].layout[key] ?? [];
    let next: number[];
    if (hide) {
      if (to - from + 1 > 200_000) return toast("That is too many to hide at once.", "error");
      const set = new Set(current);
      for (let i = from; i <= to; i++) set.add(i);
      if (set.size >= count) return toast("Can't hide every " + (axis === "rows" ? "row" : "column") + ".", "error");
      next = [...set].sort((a, b) => a - b);
    } else {
      next = current.filter((i) => i < from - 1 || i > to + 1);
    }
    store
      .setLayout(tab, { [key]: next.length > 0 ? next : undefined })
      .then(() => gridRef.current && applyLayout(gridRef.current, store, tab))
      .catch(() => {});
  };
  /** Asks for a size and gives it to every selected row (or column). */
  const askSize = (axis: "rows" | "columns") => {
    if (!canEdit || !range) return;
    const rowsAxis = axis === "rows";
    const from = rowsAxis ? range.r1 : range.c1;
    const to = rowsAxis ? range.r2 : range.c2;
    const asked = window.prompt(rowsAxis ? "Row height in pixels (8–545):" : "Column width in pixels (8–2038):", String(rowsAxis ? 30 : 116));
    if (asked === null) return;
    const size = Math.round(Number(asked));
    const max = rowsAxis ? 545 : 2038;
    if (!Number.isFinite(size) || size < 8 || size > max) return toast(`Enter a number from 8 to ${max}.`, "error");
    if (to - from + 1 > 200_000) return toast("That is too many to resize at once.", "error");
    const key = rowsAxis ? "rowHeights" : "columnWidths";
    const next = { ...(store.tabs[tab].layout[key] ?? {}) };
    for (let i = from; i <= to; i++) next[i] = size;
    store
      .setLayout(tab, { [key]: next })
      .then(() => gridRef.current && applyLayout(gridRef.current, store, tab))
      .catch(() => {});
  };
  const hiddenNear = (axis: "rows" | "columns") => {
    if (!range) return false;
    const list = store.tabs[tab]?.layout[axis === "rows" ? "hiddenRows" : "hiddenColumns"] ?? [];
    const from = axis === "rows" ? range.r1 : range.c1;
    const to = axis === "rows" ? range.r2 : range.c2;
    return list.some((i) => i >= from - 1 && i <= to + 1);
  };

  const frozen = store.tabs[tab]?.layout.frozen ?? { rows: 0, columns: 0 };
  const freezeOptions = [
    { label: "No frozen rows or columns", rows: 0, columns: 0 },
    { label: "1 row", rows: 1, columns: 0 },
    { label: "2 rows", rows: 2, columns: 0 },
    { label: "1 column", rows: 0, columns: 1 },
    { label: "1 row and 1 column", rows: 1, columns: 1 },
    ...(active
      ? [
          { label: `Up to row ${active.row + 1}`, rows: active.row + 1, columns: frozen.columns },
          { label: `Up to column ${columnName(active.col)}`, rows: frozen.rows, columns: active.col + 1 },
        ]
      : []),
  ];

  const save = useCallback(async () => {
    if (store.dirtyCount === 0 || saving) return;
    setSaving(true);
    try {
      const target = library.bigSheetUpload();
      if (!target.signer || target.servers.length === 0) throw new Error("Sign in and tick a storage server to save.");
      const next = await store.save({
        servers: target.servers,
        replicas: target.replicas,
        signer: target.signer,
        onProgress: (event) =>
          uploadProgress.report({
            phase: "uploading",
            message: "Saving your changes…",
            tabs: [],
            bytesDone: event.bytesDone,
            bytesTotal: event.bytesTotal,
            blobsDone: event.blobsDone,
            blobsTotal: 0,
            servers: target.servers,
          }),
      });
      await library.publishBigRoot(entry, next);
      uploadProgress.dismiss();
      toast("Saved.", "success");
    } catch (error) {
      uploadProgress.dismiss();
      toast(error instanceof Error ? error.message : "Couldn't save.", "error");
    } finally {
      setSaving(false);
    }
  }, [entry, library, saving, store, toast]);

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void save();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [save]);

  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (store.dirtyCount > 0) event.preventDefault();
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [store]);

  const commit = async (move: "down" | "next" | "none") => {
    if (!editing) return;
    const { row, col, text } = editing;
    setEditing(null);
    if (text.trim().startsWith("=") && text.trim().length > 1) {
      const formula = text.trim().slice(1);
      try {
        setCalculating("Calculating…");
        const result = await calculateTyped(store, tab, row, col, formula, (done, total) =>
          setCalculating(total > 1 ? `Calculating… chunk ${done} of ${total}` : "Calculating…"),
        );
        await store.edit(tab, row, col, row, col, (cell) => ({
          ...(cell?.style ? { style: cell.style } : {}),
          formula,
          value: result.value,
          ...(result.error ? { type: "error" as const } : {}),
        }));
      } catch (error) {
        toast(error instanceof Error ? error.message : "Couldn't calculate that formula.", "error");
      } finally {
        setCalculating(null);
      }
    } else {
      await store.setCell(tab, row, col, text);
    }
    await refresh([pointRect({ tab, row, col })]);
    gridRef.current?.afterEdit(move);
    gridRef.current?.focus();
  };

  const undo = () => void refresh(store.undo().map(pointRect));
  const redo = () => void refresh(store.redo().map(pointRect));

  const applyStyle = (patch: StylePatch) => editRange((cell) => restyle(cell, patch));

  const selectionLabel = () => {
    if (!range) return "";
    const first = `${columnName(range.c1)}${range.r1 + 1}`;
    return range.r1 === range.r2 && range.c1 === range.c2 ? first : `${first}:${columnName(range.c2)}${range.r2 + 1}`;
  };

  if (openError) {
    return (
      <div className="editor-error editor-error-inline">
        <AlertTriangle size={24} />
        <h2>Couldn't open this sheet</h2>
        <p>{openError}</p>
      </div>
    );
  }
  if (!ready || !info) {
    return (
      <div className="editor-loading editor-loading-inline">
        <Loader2 size={22} className="spin" />
        <p>Loading the first rows…</p>
      </div>
    );
  }

  const selectedCell = active ? store.cell(tab, active.row, active.col) : undefined;
  const rows = info.rows;
  const grid = gridRef.current;
  const editRect = editing && grid ? grid.cellRect(editing) : null;

  return (
    <div
      className="big-sheet"
      tabIndex={-1}
      onKeyDown={(event) => {
        if (editing) return;
        const mod = event.ctrlKey || event.metaKey;
        const key = event.key.toLowerCase();
        if (mod && canEdit && (key === "z" || key === "y")) {
          event.preventDefault();
          if (key === "y" || event.shiftKey) redo();
          else undo();
          return;
        }
        if (mod && canEdit && selection && (key === "b" || key === "i" || key === "u")) {
          event.preventDefault();
          const flag = key === "b" ? "bold" : key === "i" ? "italic" : "underline";
          applyStyle({ [flag]: !selectedCell?.style?.[flag] });
        }
      }}
    >
      <BigSheetToolbar
        style={selectedCell?.style ?? {}}
        canEdit={canEdit && !!selection}
        canUndo={store.canUndo}
        canRedo={store.canRedo}
        painting={painting !== null}
        onUndo={undo}
        onRedo={redo}
        onPaint={() => setPainting(painting !== null ? null : { ...(selectedCell?.style ?? {}) })}
        onClear={() =>
          editRange((cell) =>
            cell ? { ...(cell.value !== undefined ? { value: cell.value } : {}), ...(cell.formula ? { formula: cell.formula } : {}) } : undefined,
          )
        }
        onStyle={applyStyle}
        onBorder={border}
        onMerge={(how) => void merge(how)}
        canSplit={mergedHere}
        freeze={{ rows: frozen.rows, columns: frozen.columns, options: freezeOptions, apply: freeze }}
        onFormat={(format) => applyStyle({ numberFormat: format })}
        onDecimals={(delta) => editRange((cell) => restyle(cell, { numberFormat: shiftDecimals(cell?.style?.numberFormat, delta) }))}
      />
      <div className="big-sheet-bar">
        <span className="big-sheet-ref">{selectionLabel()}</span>
        <span className="big-sheet-fx">fx</span>
        <span className="big-sheet-formula">
          {selectedCell?.formula ? `=${selectedCell.formula.replace(/^=/, "")}` : shown(selectedCell)}
        </span>
        <input
          className="big-sheet-goto"
          placeholder="Go to row"
          inputMode="numeric"
          onKeyDown={(event) => {
            if (event.key !== "Enter") return;
            const row = Math.max(0, Math.min(rows - 1, Number((event.target as HTMLInputElement).value.replace(/\D/g, "")) - 1));
            gridRef.current?.scrollToRow(row);
            gridRef.current?.select({ ranges: [{ r1: row, r2: row, c1: active?.col ?? 0, c2: active?.col ?? 0 }], active: { row, col: active?.col ?? 0 }, anchor: { row, col: active?.col ?? 0 } });
          }}
        />
        <span className="big-sheet-note">
          {calculating ?? `${rows.toLocaleString()} rows · large sheet mode${canEdit ? "" : " · read-only"}`}
        </span>
        {canEdit && (
          <button type="button" className="btn btn-primary btn-sm" disabled={saving || store.dirtyCount === 0} onClick={() => void save()}>
            {saving ? <Loader2 size={14} className="spin" /> : <Save size={14} />}
            {store.dirtyCount ? `Save ${store.dirtyCount} changed chunk${store.dirtyCount === 1 ? "" : "s"}` : "Saved"}
          </button>
        )}
      </div>
      <div className="big-sheet-scroll big-sheet-canvas">
        <div ref={hostRef} className="big-sheet-host" />
        {editing && editRect && (
          <input
            className="big-cell-input"
            style={{ top: editRect.y, left: editRect.x, width: Math.max(editRect.width, 120), height: editRect.height }}
            autoFocus
            value={editing.text}
            onChange={(event) => setEditing({ ...editing, text: event.target.value })}
            onBlur={() => {
              if (cancelled.current) cancelled.current = false;
              else void commit("none");
            }}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                void commit("down");
              } else if (event.key === "Tab") {
                event.preventDefault();
                void commit("next");
              } else if (event.key === "Escape") {
                cancelled.current = true;
                setEditing(null);
                gridRef.current?.focus();
              }
            }}
          />
        )}
      </div>
      {menu && range && (
        <>
          <div className="big-menu-backdrop" onPointerDown={() => setMenu(null)} onContextMenu={(event) => (event.preventDefault(), setMenu(null))} />
          <div className="big-menu" style={{ top: menu.y, left: menu.x }} role="menu">
            {(
              [
                [`Hide row${range.r2 > range.r1 ? "s" : ""} ${range.r1 + 1}${range.r2 > range.r1 ? "–" + (range.r2 + 1) : ""}`, () => setHidden("rows", true), true],
                [`Hide column${range.c2 > range.c1 ? "s" : ""} ${columnName(range.c1)}${range.c2 > range.c1 ? "–" + columnName(range.c2) : ""}`, () => setHidden("columns", true), true],
                ["Set row height…", () => askSize("rows"), true],
                ["Set column width…", () => askSize("columns"), true],
                ["Show hidden rows here", () => setHidden("rows", false), hiddenNear("rows")],
                ["Show hidden columns here", () => setHidden("columns", false), hiddenNear("columns")],
              ] as [string, () => void, boolean][]
            ).map(([label, run, enabled]) => (
              <button
                key={label}
                type="button"
                role="menuitem"
                disabled={!canEdit || !enabled}
                onClick={() => {
                  setMenu(null);
                  run();
                }}
              >
                {label}
              </button>
            ))}
          </div>
        </>
      )}
      <div className="big-sheet-tabs">
        {store.tabs.map((item, index) => (
          <button key={item.name + index} type="button" className={cx("big-sheet-tab", index === tab && "is-active")} onClick={() => switchTab(index)}>
            {item.name}
          </button>
        ))}
      </div>
    </div>
  );
}
