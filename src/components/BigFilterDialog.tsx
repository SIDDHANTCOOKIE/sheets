// Filter dialog for big sheets: pick the values of a column to show; other rows are hidden.
import { useEffect, useMemo, useRef, useState } from "react";
import type { BigSheetStore } from "../lib/v2/bigSheet";
import { BLANK, rowsToHide, scanColumn, type ColumnFilter, type ColumnScan } from "../lib/v2/bigFilter";

export function BigFilterDialog({
  store,
  tab,
  col,
  columnLabel,
  current,
  onApply,
  onClear,
  onClose,
}: {
  store: BigSheetStore;
  tab: number;
  col: number;
  columnLabel: string;
  current?: ColumnFilter;
  onApply: (filter: ColumnFilter) => void;
  onClear: () => void;
  onClose: () => void;
}) {
  const [scan, setScan] = useState<ColumnScan | null>(null);
  const [progress, setProgress] = useState("Reading the column…");
  const [error, setError] = useState("");
  const [header, setHeader] = useState(current?.header ?? true);
  const [search, setSearch] = useState("");
  const [checked, setChecked] = useState<Set<string> | null>(null);
  const stop = useRef(false);

  useEffect(() => {
    stop.current = false;
    scanColumn(store, tab, col, (done, total) => !stop.current && setProgress(`Reading the column… ${done} of ${total} parts`), () => stop.current)
      .then((result) => {
        if (!result || stop.current) return;
        setScan(result);
        // Start from "everything not hidden by the current filter".
        const hidden = new Set(current?.hidden ?? []);
        const shown = new Set<string>();
        for (const { value } of result.values) shown.add(value);
        if (current) {
          const visible = new Set<string>();
          result.texts.forEach((text, row) => !hidden.has(row) && visible.add(text));
          for (let row = 0; row < store.tabs[tab].rows && !visible.has(BLANK); row++) {
            if (!result.texts.has(row) && !hidden.has(row) && row >= (current.header ? 1 : 0)) visible.add(BLANK);
          }
          setChecked(visible);
        } else setChecked(shown);
      })
      .catch((e) => setError(e instanceof Error ? e.message : "Couldn't read the column."));
    return () => {
      stop.current = true;
    };
    // The dialog is made for one column; it is re-created for another.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const listed = useMemo(
    () => (scan ? scan.values.filter((item) => item.value.toLowerCase().includes(search.toLowerCase())) : []),
    [scan, search],
  );

  const apply = () => {
    if (!scan || !checked) return;
    const keep = (text: string | null) => checked.has(text ?? BLANK);
    const hidden = rowsToHide(scan, store.tabs[tab].rows, header, keep);
    const everyListedChecked = scan.values.every((item) => checked.has(item.value));
    if (hidden.length === 0 && everyListedChecked) return onClear();
    onApply({
      col,
      header,
      hidden,
      summary: checked.size === 1 ? `is ${[...checked][0]}` : `is one of ${checked.size.toLocaleString()} values`,
    });
  };

  return (
    <div className="big-dialog-backdrop" onPointerDown={onClose}>
      <div
        className="big-dialog"
        role="dialog"
        aria-label="Filter"
        onPointerDown={(event) => event.stopPropagation()}
        onKeyDown={(event) => event.key === "Escape" && onClose()}
      >
        <h3>Filter column {columnLabel}</h3>
        {error ? (
          <p className="big-dialog-error">{error}</p>
        ) : !scan || !checked ? (
          <p className="big-dialog-sub">{progress}</p>
        ) : (
          <>
            <label className="big-dialog-check">
              <input type="checkbox" checked={header} onChange={(event) => setHeader(event.target.checked)} /> The first row is a header (never hidden)
            </label>
            <input placeholder="Search values" value={search} onChange={(event) => setSearch(event.target.value)} autoFocus />
            <div className="big-dialog-row">
              <button type="button" className="btn btn-sm" onClick={() => setChecked(new Set(scan.values.map((item) => item.value)))}>
                Select all
              </button>
              <button type="button" className="btn btn-sm" onClick={() => setChecked(new Set())}>
                Clear all
              </button>
            </div>
            <div className="big-filter-list">
              {listed.map((item) => (
                <label key={item.value} className="big-filter-item">
                  <input
                    type="checkbox"
                    checked={checked.has(item.value)}
                    onChange={(event) => {
                      const next = new Set(checked);
                      if (event.target.checked) next.add(item.value);
                      else next.delete(item.value);
                      setChecked(next);
                    }}
                  />
                  <span className="big-filter-value">{item.value}</span>
                  <span className="big-filter-count">{item.count.toLocaleString()}</span>
                </label>
              ))}
              {listed.length === 0 && <p className="big-dialog-sub">No values match.</p>}
            </div>
            {scan.truncated && <p className="big-dialog-sub">Only the 5,000 most common values are listed.</p>}
          </>
        )}
        <div className="big-dialog-actions">
          {current && (
            <button type="button" className="btn btn-sm" onClick={onClear}>
              Clear filter
            </button>
          )}
          <span style={{ flex: 1 }} />
          <button type="button" className="btn btn-sm" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="btn btn-primary btn-sm" disabled={!scan || !checked} onClick={apply}>
            Apply
          </button>
        </div>
      </div>
    </div>
  );
}
