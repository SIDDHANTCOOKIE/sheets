// One user-level action, described independently of any engine. A scenario is a list of steps.

/** One user-level action, described independently of either engine. */
export type Step =
  | { op: "format"; row: number; col: number; attr: "bl" | "it" | "un" | "cl" | "bg" | "fc" | "fs" | "ht" | "vt" | "tb"; value: number | string }
  | { op: "merge"; r1: number; c1: number; r2: number; c2: number; type?: "merge-all" | "merge-cancel" }
  | { op: "set"; row: number; col: number; value: string | number }
  | { op: "insertRow"; index: number; count: number; direction: "above" | "below" }
  | { op: "insertColumn"; index: number; count: number; direction: "left" | "right" }
  | { op: "deleteRow"; start: number; end: number }
  | { op: "deleteColumn"; start: number; end: number }
  | { op: "clear"; row: number; col: number }
  | { op: "rowHeight"; row: number; height: number }
  | { op: "columnWidth"; col: number; width: number }
  | { op: "hide"; axis: "row" | "column"; index: number }
  | { op: "show"; axis: "row" | "column"; index: number }
  | { op: "freeze"; type: "row" | "column" | "both"; row: number; col: number }
  | { op: "numberFormat"; row: number; col: number; format: string }
  | { op: "rotation"; row: number; col: number; value: number }
  | { op: "fill"; r1: number; c1: number; r2: number; c2: number; direction: "down" | "right"; count: number };
