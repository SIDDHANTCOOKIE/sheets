// The grid itself: a canvas that draws the cells on screen, a scroll area that gives it a scrollbar,
// and the mouse/keyboard handling that drives the selection. Cell contents come from a lazy data
// source, so the same grid shows 10 cells or 25 million.
import {
  advance,
  click,
  isSelected,
  move,
  selectAll,
  single,
  wholeCol,
  wholeRow,
  type Cell,
  type Direction,
  type Selection,
} from "./selection";
import { SizeIndex } from "./sizes";
import { scrollExtent, toScrollPx, toSheetPx, visible } from "./viewport";

export interface GridCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
  fontSize?: number;
  fontFamily?: string;
  wrap?: boolean;
}

export interface GridCell {
  text: string;
  style?: GridCellStyle;
}

/** What the grid reads. `cell` returns undefined while the cell's data is still loading. */
export interface GridData {
  rowCount: number;
  colCount: number;
  cell(row: number, col: number): GridCell | null | undefined;
  /** Asks for the rows on screen to be loaded; called on every scroll. */
  ensure?(firstRow: number, endRow: number): void;
  /** Calls back when cells changed or finished loading. Returns an unsubscribe function. */
  subscribe?(listener: () => void): () => void;
}

export interface GridTheme {
  font: string;
  background: string;
  text: string;
  line: string;
  headBackground: string;
  headText: string;
  headSelected: string;
  selection: string;
  selectionFill: string;
  loading: string;
}

export const LIGHT_THEME: GridTheme = {
  font: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
  background: "#ffffff",
  text: "#1f2328",
  line: "#e1e4e8",
  headBackground: "#f6f8fa",
  headText: "#57606a",
  headSelected: "#dbeafe",
  selection: "#2563eb",
  selectionFill: "rgba(37, 99, 235, 0.12)",
  loading: "#f3f4f6",
};

export interface GridOptions {
  data: GridData;
  rowHeight?: number;
  colWidth?: number;
  headerWidth?: number;
  headerHeight?: number;
  theme?: GridTheme;
  /** Called when the user starts editing (Enter/F2/double-click/typing). `text` is set when typing replaces the cell. */
  onEdit?(cell: Cell, text: string | undefined): void;
  onSelect?(selection: Selection): void;
  /** Delete/Backspace on the selection. */
  onClear?(selection: Selection): void;
  /** Cells to look at for Ctrl+arrow (does the cell hold anything?). */
  hasData?(row: number, col: number): boolean;
  onResize?(kind: "row" | "col", index: number, size: number): void;
}

export function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

const MIN_SIZE = 8;

export class Grid {
  readonly rows: SizeIndex;
  readonly cols: SizeIndex;
  selection: Selection | null = null;

  private readonly options: GridOptions;
  private readonly theme: GridTheme;
  private headW: number;
  private readonly baseHeadW: number;
  private readonly headH: number;
  private readonly root: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly spacer: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;
  private width = 0;
  private height = 0;
  private frame = 0;
  private data: GridData;
  private unsubscribe?: () => void;
  private observer?: ResizeObserver;
  private drag: { kind: "cells" | "row" | "col"; from: number } | { kind: "resize"; axis: "row" | "col"; index: number; start: number; size: number } | null = null;

  constructor(parent: HTMLElement, options: GridOptions) {
    this.options = options;
    this.data = options.data;
    this.theme = options.theme ?? LIGHT_THEME;
    this.baseHeadW = options.headerWidth ?? 46;
    this.headW = this.baseHeadW;
    this.headH = options.headerHeight ?? 28;
    this.rows = new SizeIndex(options.data.rowCount, options.rowHeight ?? 30);
    this.cols = new SizeIndex(options.data.colCount, options.colWidth ?? 116);

    this.root = document.createElement("div");
    this.root.className = "sheets-grid";
    this.root.tabIndex = 0;
    this.root.style.cssText = "position:relative;width:100%;height:100%;overflow:hidden;outline:none;";
    this.scroller = document.createElement("div");
    this.scroller.style.cssText = "position:absolute;inset:0;overflow:auto;";
    this.spacer = document.createElement("div");
    this.spacer.style.cssText = "position:absolute;top:0;left:0;width:1px;height:1px;";
    this.canvas = document.createElement("canvas");
    this.canvas.style.cssText = "position:absolute;top:0;left:0;pointer-events:none;";
    this.scroller.append(this.spacer);
    this.root.append(this.canvas, this.scroller);
    parent.append(this.root);
    const ctx = this.canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas is not available in this browser.");
    this.ctx = ctx;

    this.scroller.addEventListener("scroll", this.schedule);
    this.scroller.addEventListener("pointerdown", this.onPointerDown);
    this.scroller.addEventListener("pointermove", this.onPointerMove);
    this.scroller.addEventListener("pointerup", this.onPointerUp);
    this.scroller.addEventListener("dblclick", this.onDoubleClick);
    this.root.addEventListener("keydown", this.onKeyDown);
    this.observer = new ResizeObserver(() => this.resize());
    this.observer.observe(this.root);
    this.attach(this.data);
    this.resize();
  }

  /** Switches to another data source (another tab). Selection and scroll reset. */
  setData(data: GridData): void {
    this.unsubscribe?.();
    this.data = data;
    this.rows.setLength(data.rowCount);
    this.cols.setLength(data.colCount);
    this.selection = null;
    this.scroller.scrollTop = 0;
    this.scroller.scrollLeft = 0;
    this.attach(data);
    this.updateSpacer();
    this.schedule();
  }

  private attach(data: GridData): void {
    this.unsubscribe = data.subscribe?.(this.schedule);
  }

  destroy(): void {
    this.unsubscribe?.();
    this.observer?.disconnect();
    cancelAnimationFrame(this.frame);
    this.root.remove();
  }

  focus(): void {
    this.root.focus({ preventScroll: true });
  }

  /** Redraws on the next animation frame (several calls in a row draw once). */
  readonly schedule = (): void => {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.draw();
    });
  };

  select(selection: Selection | null, reveal = true): void {
    this.selection = selection;
    if (selection && reveal) this.reveal(selection.active);
    if (selection) this.options.onSelect?.(selection);
    this.schedule();
  }

  /** Scrolls so `cell` is fully on screen. */
  reveal(cell: Cell): void {
    const viewW = this.width - this.headW;
    const viewH = this.height - this.headH;
    const rowsTotal = this.rows.total;
    const colsTotal = this.cols.total;
    const top = toSheetPx(this.scroller.scrollTop, rowsTotal, viewH);
    const left = toSheetPx(this.scroller.scrollLeft, colsTotal, viewW);
    const y = this.rows.offset(cell.row);
    const x = this.cols.offset(cell.col);
    if (y < top) this.scroller.scrollTop = toScrollPx(y, rowsTotal, viewH);
    else if (y + this.rows.size(cell.row) > top + viewH) {
      this.scroller.scrollTop = toScrollPx(y + this.rows.size(cell.row) - viewH, rowsTotal, viewH);
    }
    if (x < left) this.scroller.scrollLeft = toScrollPx(x, colsTotal, viewW);
    else if (x + this.cols.size(cell.col) > left + viewW) {
      this.scroller.scrollLeft = toScrollPx(x + this.cols.size(cell.col) - viewW, colsTotal, viewW);
    }
  }

  /** Where `cell` is on screen, relative to the grid's top-left corner (for placing an editor). */
  cellRect(cell: Cell): { x: number; y: number; width: number; height: number } {
    const top = toSheetPx(this.scroller.scrollTop, this.rows.total, this.height - this.headH);
    const left = toSheetPx(this.scroller.scrollLeft, this.cols.total, this.width - this.headW);
    return {
      x: this.headW + this.cols.offset(cell.col) - left,
      y: this.headH + this.rows.offset(cell.row) - top,
      width: this.cols.size(cell.col),
      height: this.rows.size(cell.row),
    };
  }

  scrollToRow(row: number): void {
    const target = Math.max(0, Math.min(this.rows.length - 1, row));
    this.scroller.scrollTop = toScrollPx(this.rows.offset(target), this.rows.total, this.height - this.headH);
    this.schedule();
  }

  // --- layout -----------------------------------------------------------------------------

  private resize(): void {
    // The canvas covers what the scroll area shows without its scrollbars.
    this.width = Math.max(1, this.scroller.clientWidth);
    this.height = Math.max(1, this.scroller.clientHeight);
    const dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.floor(this.width * dpr);
    this.canvas.height = Math.floor(this.height * dpr);
    this.canvas.style.width = `${this.width}px`;
    this.canvas.style.height = `${this.height}px`;
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.updateSpacer();
    this.schedule();
  }

  /** The scroll area is as long as the sheet (capped), offset by the headers the canvas draws over. */
  private updateSpacer(): void {
    // Wide enough for the biggest row number ("25,000,000").
    this.headW = Math.max(this.baseHeadW, Math.ceil(this.rows.length.toLocaleString().length * 7 + 16));
    const viewW = this.width - this.headW;
    const viewH = this.height - this.headH;
    this.spacer.style.height = `${scrollExtent(this.rows.total, viewH) + this.headH}px`;
    this.spacer.style.width = `${scrollExtent(this.cols.total, viewW) + this.headW}px`;
  }

  private view() {
    const viewW = this.width - this.headW;
    const viewH = this.height - this.headH;
    const top = toSheetPx(this.scroller.scrollTop, this.rows.total, viewH);
    const left = toSheetPx(this.scroller.scrollLeft, this.cols.total, viewW);
    return {
      viewW,
      viewH,
      rows: visible(this.rows, top, viewH),
      cols: visible(this.cols, left, viewW),
    };
  }

  // --- drawing ----------------------------------------------------------------------------

  private draw(): void {
    // The sheet may have grown (an edit past the last row) since the last frame.
    if (this.rows.length !== this.data.rowCount || this.cols.length !== this.data.colCount) {
      this.rows.setLength(this.data.rowCount);
      this.cols.setLength(this.data.colCount);
      this.updateSpacer();
    }
    const { ctx, theme } = this;
    const { viewW, viewH, rows, cols } = this.view();
    const fonts = (style: GridCellStyle | undefined) =>
      `${style?.italic ? "italic " : ""}${style?.bold ? "700 " : ""}${Math.min(style?.fontSize ?? 11, 24) * (96 / 72)}px ${style?.fontFamily ? `${style.fontFamily}, ` : ""}${theme.font}`;

    this.data.ensure?.(rows.first, rows.end);

    ctx.clearRect(0, 0, this.width, this.height);
    ctx.fillStyle = theme.background;
    ctx.fillRect(0, 0, this.width, this.height);

    // Positions of the visible rows/columns (hidden ones have size 0 and are skipped).
    const ys: number[] = [];
    let y = this.headH - rows.shift;
    for (let r = rows.first; r < rows.end; r++) {
      ys.push(y);
      y += this.rows.size(r);
    }
    const xs: number[] = [];
    let x = this.headW - cols.shift;
    for (let c = cols.first; c < cols.end; c++) {
      xs.push(x);
      x += this.cols.size(c);
    }

    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headW, this.headH, viewW, viewH);
    ctx.clip();
    ctx.textBaseline = "middle";
    for (let r = rows.first; r < rows.end; r++) {
      const h = this.rows.size(r);
      if (h === 0) continue;
      const top = ys[r - rows.first];
      for (let c = cols.first; c < cols.end; c++) {
        const w = this.cols.size(c);
        if (w === 0) continue;
        const left = xs[c - cols.first];
        const cell = this.data.cell(r, c);
        const style = cell?.style;
        if (cell === undefined) {
          ctx.fillStyle = theme.loading;
          ctx.fillRect(left, top, w, h);
        } else if (style?.background) {
          ctx.fillStyle = style.background;
          ctx.fillRect(left, top, w, h);
        }
        if (cell?.text) this.drawText(cell, left, top, w, h, fonts(style));
      }
    }
    // Grid lines.
    ctx.strokeStyle = theme.line;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let i = 0; i < ys.length; i++) {
      if (this.rows.size(rows.first + i) === 0) continue;
      const line = Math.round(ys[i] + this.rows.size(rows.first + i)) - 0.5;
      ctx.moveTo(this.headW, line);
      ctx.lineTo(this.headW + viewW, line);
    }
    for (let i = 0; i < xs.length; i++) {
      if (this.cols.size(cols.first + i) === 0) continue;
      const line = Math.round(xs[i] + this.cols.size(cols.first + i)) - 0.5;
      ctx.moveTo(line, this.headH);
      ctx.lineTo(line, this.headH + viewH);
    }
    ctx.stroke();
    this.drawSelection(rows, cols, ys, xs);
    ctx.restore();

    this.drawHeaders(rows, cols, ys, xs, viewW, viewH);
  }

  private drawText(cell: GridCell, left: number, top: number, w: number, h: number, font: string): void {
    const { ctx, theme } = this;
    const style = cell.style;
    ctx.save();
    ctx.beginPath();
    ctx.rect(left, top, w, h);
    ctx.clip();
    ctx.font = font;
    ctx.fillStyle = style?.color ?? theme.text;
    const align = style?.align ?? (/^[-+]?[\d.,]+%?$/.test(cell.text) ? "right" : "left");
    ctx.textAlign = align;
    const pad = 4;
    const tx = align === "left" ? left + pad : align === "right" ? left + w - pad : left + w / 2;
    const lines = style?.wrap ? wrapText(ctx, cell.text, w - pad * 2) : [cell.text.split("\n")[0]];
    const lineH = Math.min(style?.fontSize ?? 11, 24) * (96 / 72) * 1.3;
    const startY = style?.wrap ? top + lineH / 2 + 2 : top + h / 2;
    lines.forEach((line, i) => {
      const ty = startY + i * lineH;
      ctx.fillText(line, tx, ty);
      const width = ctx.measureText(line).width;
      const from = align === "left" ? tx : align === "right" ? tx - width : tx - width / 2;
      if (style?.underline) ctx.fillRect(from, ty + lineH / 2 - 2, width, 1);
      if (style?.strike) ctx.fillRect(from, ty, width, 1);
    });
    ctx.restore();
  }

  private drawSelection(
    rows: { first: number; end: number },
    cols: { first: number; end: number },
    ys: number[],
    xs: number[],
  ): void {
    const sel = this.selection;
    if (!sel) return;
    const { ctx, theme } = this;
    const rectOf = (r1: number, c1: number, r2: number, c2: number) => {
      // Rows/columns outside the window are clamped to its edge so a huge range still draws.
      const ra = Math.max(r1, rows.first);
      const rb = Math.min(r2, rows.end - 1);
      const ca = Math.max(c1, cols.first);
      const cb = Math.min(c2, cols.end - 1);
      if (ra > rb || ca > cb) return null;
      const top = ys[ra - rows.first];
      const left = xs[ca - cols.first];
      return {
        left,
        top,
        width: xs[cb - cols.first] + this.cols.size(cb) - left,
        height: ys[rb - rows.first] + this.rows.size(rb) - top,
      };
    };
    for (const range of sel.ranges) {
      const rect = rectOf(range.r1, range.c1, range.r2, range.c2);
      if (!rect) continue;
      ctx.fillStyle = theme.selectionFill;
      ctx.fillRect(rect.left, rect.top, rect.width, rect.height);
      ctx.strokeStyle = theme.selection;
      ctx.lineWidth = 2;
      ctx.strokeRect(rect.left + 1, rect.top + 1, rect.width - 2, rect.height - 2);
    }
    // The active cell is drawn without the tint so it reads as "where typing goes".
    const active = rectOf(sel.active.row, sel.active.col, sel.active.row, sel.active.col);
    if (active) {
      ctx.fillStyle = theme.background;
      const cell = this.data.cell(sel.active.row, sel.active.col);
      ctx.fillRect(active.left + 2, active.top + 2, active.width - 4, active.height - 4);
      if (cell?.text) {
        this.drawText(cell, active.left, active.top, active.width, active.height, `${cell.style?.bold ? "700 " : ""}${Math.min(cell.style?.fontSize ?? 11, 24) * (96 / 72)}px ${theme.font}`);
      }
    }
  }

  private drawHeaders(
    rows: { first: number; end: number },
    cols: { first: number; end: number },
    ys: number[],
    xs: number[],
    viewW: number,
    viewH: number,
  ): void {
    const { ctx, theme } = this;
    const sel = this.selection;
    const rowSelected = (r: number) => !!sel?.ranges.some((g) => r >= g.r1 && r <= g.r2);
    const colSelected = (c: number) => !!sel?.ranges.some((g) => c >= g.c1 && c <= g.c2);
    ctx.font = `12px ${theme.font}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    ctx.save();
    ctx.beginPath();
    ctx.rect(0, this.headH, this.headW, viewH);
    ctx.clip();
    for (let r = rows.first; r < rows.end; r++) {
      const h = this.rows.size(r);
      if (h === 0) continue;
      const top = ys[r - rows.first];
      ctx.fillStyle = rowSelected(r) ? theme.headSelected : theme.headBackground;
      ctx.fillRect(0, top, this.headW, h);
      ctx.fillStyle = theme.headText;
      ctx.fillText((r + 1).toLocaleString(), this.headW / 2, top + h / 2);
      ctx.fillStyle = theme.line;
      ctx.fillRect(0, Math.round(top + h) - 1, this.headW, 1);
    }
    ctx.restore();

    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headW, 0, viewW, this.headH);
    ctx.clip();
    for (let c = cols.first; c < cols.end; c++) {
      const w = this.cols.size(c);
      if (w === 0) continue;
      const left = xs[c - cols.first];
      ctx.fillStyle = colSelected(c) ? theme.headSelected : theme.headBackground;
      ctx.fillRect(left, 0, w, this.headH);
      ctx.fillStyle = theme.headText;
      ctx.fillText(columnName(c), left + w / 2, this.headH / 2);
      ctx.fillStyle = theme.line;
      ctx.fillRect(Math.round(left + w) - 1, 0, 1, this.headH);
    }
    ctx.restore();

    ctx.fillStyle = theme.headBackground;
    ctx.fillRect(0, 0, this.headW, this.headH);
    ctx.fillStyle = theme.line;
    ctx.fillRect(0, this.headH - 1, this.width, 1);
    ctx.fillRect(this.headW - 1, 0, 1, this.height);
  }

  // --- mouse ------------------------------------------------------------------------------

  /** Row/column under a point (in the grid's own coordinates), or the header zone it is in. */
  private hit(clientX: number, clientY: number) {
    const rect = this.root.getBoundingClientRect();
    const px = clientX - rect.left;
    const py = clientY - rect.top;
    const { rows, cols } = this.view();
    const top = toSheetPx(this.scroller.scrollTop, this.rows.total, this.height - this.headH);
    const left = toSheetPx(this.scroller.scrollLeft, this.cols.total, this.width - this.headW);
    void rows;
    void cols;
    const onColHead = py < this.headH;
    const onRowHead = px < this.headW;
    const row = this.rows.indexAt(top + Math.max(0, py - this.headH));
    const col = this.cols.indexAt(left + Math.max(0, px - this.headW));
    return { px, py, row, col, onColHead, onRowHead, top, left };
  }

  /** The edge of a row/column header the pointer is on (within 4px), for resizing. */
  private edge(hit: ReturnType<Grid["hit"]>): { axis: "row" | "col"; index: number } | null {
    if (hit.onColHead && !hit.onRowHead) {
      const x = hit.px - this.headW + hit.left;
      const end = this.cols.offset(hit.col) + this.cols.size(hit.col);
      const startEdge = this.cols.offset(hit.col);
      if (Math.abs(x - end) <= 4) return { axis: "col", index: hit.col };
      if (Math.abs(x - startEdge) <= 4 && hit.col > 0) return { axis: "col", index: this.cols.nextVisible(hit.col - 1, -1) };
    }
    if (hit.onRowHead && !hit.onColHead) {
      const y = hit.py - this.headH + hit.top;
      const end = this.rows.offset(hit.row) + this.rows.size(hit.row);
      const startEdge = this.rows.offset(hit.row);
      if (Math.abs(y - end) <= 4) return { axis: "row", index: hit.row };
      if (Math.abs(y - startEdge) <= 4 && hit.row > 0) return { axis: "row", index: this.rows.nextVisible(hit.row - 1, -1) };
    }
    return null;
  }

  private onPointerDown = (event: PointerEvent): void => {
    if (event.button !== 0) return;
    const hit = this.hit(event.clientX, event.clientY);
    // The scrollbars sit inside the scroller; ignore presses on them.
    if (hit.px >= this.width || hit.py >= this.height) return;
    this.focus();
    const mode = event.shiftKey ? "extend" : event.ctrlKey || event.metaKey ? "add" : "replace";
    const bounds = { rows: this.rows, cols: this.cols };
    const resize = this.edge(hit);
    if (resize) {
      const size = resize.axis === "col" ? this.cols.size(resize.index) : this.rows.size(resize.index);
      this.drag = { kind: "resize", axis: resize.axis, index: resize.index, start: resize.axis === "col" ? event.clientX : event.clientY, size };
    } else if (hit.onRowHead && hit.onColHead) {
      this.select(selectAll(bounds), false);
    } else if (hit.onColHead) {
      this.drag = { kind: "col", from: hit.col };
      this.select(wholeCol(bounds, hit.col), false);
    } else if (hit.onRowHead) {
      this.drag = { kind: "row", from: hit.row };
      this.select(wholeRow(bounds, hit.row), false);
    } else {
      this.drag = { kind: "cells", from: 0 };
      this.select(click(this.selection, { row: hit.row, col: hit.col }, mode), false);
    }
    this.scroller.setPointerCapture(event.pointerId);
    event.preventDefault();
  };

  private onPointerMove = (event: PointerEvent): void => {
    const hit = this.hit(event.clientX, event.clientY);
    const drag = this.drag;
    if (!drag) {
      const over = this.edge(hit);
      this.scroller.style.cursor = over ? (over.axis === "col" ? "col-resize" : "row-resize") : "";
      return;
    }
    const bounds = { rows: this.rows, cols: this.cols };
    if (drag.kind === "resize") {
      const delta = (drag.axis === "col" ? event.clientX : event.clientY) - drag.start;
      const size = Math.max(MIN_SIZE, Math.round(drag.size + delta));
      (drag.axis === "col" ? this.cols : this.rows).set(drag.index, size);
      this.updateSpacer();
      this.schedule();
    } else if (drag.kind === "col") {
      this.select(wholeCol(bounds, drag.from, hit.col), false);
    } else if (drag.kind === "row") {
      this.select(wholeRow(bounds, drag.from, hit.row), false);
    } else if (this.selection) {
      this.select(click(this.selection, { row: hit.row, col: hit.col }, "extend"), false);
    }
  };

  private onPointerUp = (event: PointerEvent): void => {
    const drag = this.drag;
    this.drag = null;
    if (this.scroller.hasPointerCapture(event.pointerId)) this.scroller.releasePointerCapture(event.pointerId);
    if (drag?.kind === "resize") {
      const size = drag.axis === "col" ? this.cols.size(drag.index) : this.rows.size(drag.index);
      this.options.onResize?.(drag.axis === "col" ? "col" : "row", drag.index, size);
    }
  };

  private onDoubleClick = (event: MouseEvent): void => {
    const hit = this.hit(event.clientX, event.clientY);
    if (hit.onColHead || hit.onRowHead) return;
    this.options.onEdit?.({ row: hit.row, col: hit.col }, undefined);
  };

  // --- keyboard ---------------------------------------------------------------------------

  private onKeyDown = (event: KeyboardEvent): void => {
    const sel = this.selection;
    const mod = event.ctrlKey || event.metaKey;
    const bounds = { rows: this.rows, cols: this.cols };
    const arrows: Record<string, Direction> = { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right" };
    if (mod && event.key.toLowerCase() === "a") {
      event.preventDefault();
      this.select(selectAll(bounds), false);
      return;
    }
    if (!sel) {
      if (arrows[event.key] || event.key === "Tab") {
        event.preventDefault();
        this.select(single({ row: 0, col: 0 }));
      }
      return;
    }
    const dir = arrows[event.key];
    if (dir) {
      event.preventDefault();
      this.select(move(bounds, sel, dir, { extend: event.shiftKey, jump: mod, hasData: this.options.hasData }));
      return;
    }
    const page = Math.max(1, Math.floor((this.height - this.headH) / (this.rows.size(sel.active.row) || 30)) - 1);
    switch (event.key) {
      case "PageDown":
      case "PageUp": {
        event.preventDefault();
        const dr = event.key === "PageDown" ? page : -page;
        const row = this.rows.nextVisible(Math.max(0, Math.min(this.rows.length - 1, sel.active.row + dr)), dr > 0 ? -1 : 1);
        this.select(event.shiftKey ? click(sel, { row, col: sel.active.col }, "extend") : single({ row, col: sel.active.col }));
        return;
      }
      case "Home":
      case "End": {
        event.preventDefault();
        const end = event.key === "End";
        const row = mod ? (end ? this.rows.length - 1 : 0) : sel.active.row;
        const col = end && !mod ? this.cols.length - 1 : mod && end ? this.cols.length - 1 : 0;
        this.select(single({ row: this.rows.nextVisible(row, end ? -1 : 1), col: this.cols.nextVisible(col, end ? -1 : 1) }));
        return;
      }
      case "Tab":
        event.preventDefault();
        this.select(advance(bounds, sel, event.shiftKey ? "previous" : "next"));
        return;
      case "Enter":
        event.preventDefault();
        if (this.options.onEdit && !event.shiftKey && (event.altKey || event.metaKey)) return;
        this.options.onEdit?.(sel.active, undefined);
        return;
      case "F2":
        event.preventDefault();
        this.options.onEdit?.(sel.active, undefined);
        return;
      case "Escape":
        return;
      case "Delete":
      case "Backspace":
        event.preventDefault();
        this.options.onClear?.(sel);
        return;
      default:
        if (!mod && !event.altKey && event.key.length === 1) {
          event.preventDefault();
          this.options.onEdit?.(sel.active, event.key);
        }
    }
  };

  /** After an edit commits: Enter moves down within the selection, Tab to the right. */
  afterEdit(direction: "down" | "next" | "none"): void {
    if (!this.selection || direction === "none") return;
    this.select(advance({ rows: this.rows, cols: this.cols }, this.selection, direction));
    this.focus();
  }

  isSelected(row: number, col: number): boolean {
    return !!this.selection && isSelected(this.selection, row, col);
  }
}

function wrapText(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split("\n")) {
    let line = "";
    for (const word of paragraph.split(" ")) {
      const next = line ? `${line} ${word}` : word;
      if (line && ctx.measureText(next).width > width) {
        lines.push(line);
        line = word;
      } else line = next;
    }
    lines.push(line);
  }
  return lines;
}
