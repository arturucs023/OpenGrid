import React, { useEffect, useRef } from "react";
import { useVirtualizer } from "@tanstack/react-virtual";
import { useStore } from "../store/workbookStore";
import { colToLetter } from "../model/types";
import type { CellStyle } from "../model/types";

const ROW_COUNT = 2000;
const COL_COUNT = 100;

function colWidth(meta: any, col: number) {
  const w = meta?.sheets?.[meta.activeSheet]?.colWidths?.[col];
  return w ? w : 88;
}
function rowHeight(meta: any, row: number) {
  const h = meta?.sheets?.[meta.activeSheet]?.rowHeights?.[row];
  return h ? h : 20;
}

function inRanges(r: number, c: number, ranges: string[]): boolean {
  for (const rng of ranges) {
    const parts = rng.split(":");
    const s = parseA1Local(parts[0]);
    const e = parts[1] ? parseA1Local(parts[1]) : s;
    if (!s || !e) continue;
    const r1 = Math.min(s.row, e.row), r2 = Math.max(s.row, e.row);
    const c1 = Math.min(s.col, e.col), c2 = Math.max(s.col, e.col);
    if (r >= r1 && r <= r2 && c >= c1 && c <= c2) return true;
  }
  return false;
}
function parseA1Local(ref: string) {
  const m = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  let n = 0;
  for (let i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
  return { col: n - 1, row: parseInt(m[2], 10) - 1 };
}

function displayValue(engine: any, meta: any, sheet: string, row: number, col: number): string {
  const raw = engine.getSerialized(sheet, row, col);
  if (raw === null || raw === undefined) return "";
  if (typeof raw === "string" && raw.startsWith("=")) {
    const v = engine.getCellValue(sheet, row, col);
    if (v && typeof v === "object" && "type" in v) return String((v as any).value ?? (v as any).message ?? "#ERR");
    return v === null || v === undefined ? "" : String(v);
  }
  // number format
  const style = meta.styles?.[`${sheet}!R${row}C${col}`] as CellStyle | undefined;
  const num = typeof raw === "number" ? raw : (typeof raw === "string" && raw !== "" && !isNaN(Number(raw)) ? Number(raw) : NaN);
  if (!isNaN(num) && style?.numberFormat) {
    // basic format handling
    const fmt = style.numberFormat;
    if (fmt === "0") return Math.round(num).toLocaleString("es-ES", { maximumFractionDigits: 0 });
    if (fmt === "0.00") return num.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (fmt === "#,##0") return num.toLocaleString("es-ES", { maximumFractionDigits: 0 });
    if (fmt === "#,##0.00") return num.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (fmt === "0%") return (num * 100).toFixed(0) + "%";
    if (fmt === "$#,##0.00") return "$" + num.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    try { return new Intl.NumberFormat("es-ES").format(num); } catch { return String(raw); }
  }
  return String(raw);
}

export const Grid: React.FC = () => {
  const store = useStore();
  const engine = store.engine;
  const meta = store.meta;
  const sheet = store.activeSheet;
  const bodyRef = useRef<HTMLDivElement>(null);
  const colHeaderRef = useRef<HTMLDivElement>(null);
  const rowHeaderRef = useRef<HTMLDivElement>(null);
  const syncing = useRef(false);

  const hiddenRows = new Set<number>(meta?.sheets?.[sheet]?.hiddenRows || []);
  const visibleRows: number[] = [];
  for (let i = 0; i < ROW_COUNT; i++) if (!hiddenRows.has(i)) visibleRows.push(i);
  const rowVirtualizer = useVirtualizer({
    count: visibleRows.length,
    getScrollElement: () => bodyRef.current,
    estimateSize: (i) => rowHeight(meta, visibleRows[i]),
    overscan: 8,
  });
  const colVirtualizer = useVirtualizer({
    count: COL_COUNT,
    getScrollElement: () => bodyRef.current,
    estimateSize: (i) => colWidth(meta, i),
    horizontal: true,
    overscan: 8,
  });

  // sync scroll
  useEffect(() => {
    const body = bodyRef.current;
    const ch = colHeaderRef.current;
    const rh = rowHeaderRef.current;
    if (!body) return;
    const onBody = () => {
      if (syncing.current) return;
      syncing.current = true;
      if (ch) ch.scrollLeft = body.scrollLeft;
      if (rh) rh.scrollTop = body.scrollTop;
      syncing.current = false;
    };
    body.addEventListener("scroll", onBody);
    return () => body.removeEventListener("scroll", onBody);
  }, []);

  const syncFromHeader = (el: HTMLDivElement | null, axis: "h" | "v") => {
    if (!el) return;
    el.addEventListener("scroll", () => {
      if (syncing.current) return;
      syncing.current = true;
      const body = bodyRef.current;
      if (body) {
        if (axis === "h") body.scrollLeft = el.scrollLeft;
        else body.scrollTop = el.scrollTop;
      }
      syncing.current = false;
    });
  };
  useEffect(() => { syncFromHeader(colHeaderRef.current, "h"); syncFromHeader(rowHeaderRef.current, "v"); }, []);

  const totalWidth = colVirtualizer.getTotalSize();
  const totalHeight = rowVirtualizer.getTotalSize();

  const active = store.selection.active;
  const ranges = store.selection.ranges;

  const handleCellClick = (r: number, c: number, e: React.MouseEvent) => {
    if (store.editing) store.commitEdit();
    store.selectCell(r, c, e.shiftKey, e.ctrlKey || e.metaKey);
  };
  const handleCellDblClick = (r: number, c: number) => {
    store.startEdit(r, c);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (store.editing) return;
    const { row, col } = store.selection.active;
    if (e.key === "ArrowUp") { e.preventDefault(); store.selectCell(Math.max(0, row - 1), col); }
    else if (e.key === "ArrowDown") { e.preventDefault(); store.selectCell(Math.min(ROW_COUNT - 1, row + 1), col); }
    else if (e.key === "ArrowLeft") { e.preventDefault(); store.selectCell(row, Math.max(0, col - 1)); }
    else if (e.key === "ArrowRight") { e.preventDefault(); store.selectCell(row, Math.min(COL_COUNT - 1, col + 1)); }
    else if (e.key === "Enter") { e.preventDefault(); store.startEdit(row, col); }
    else if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); store.setCellInput(sheet, row, col, null); }
    else if (e.key === "F2") { e.preventDefault(); store.startEdit(row, col); }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c") { e.preventDefault(); store.copy(); }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "x") { e.preventDefault(); store.cut(); }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "v") { e.preventDefault(); store.paste(row, col); }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") { e.preventDefault(); if (e.shiftKey) store.redo(); else store.undo(); }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") { e.preventDefault(); store.redo(); }
    else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      store.startEdit(row, col, e.key);
    }
  };

  const rows = rowVirtualizer.getVirtualItems();
  const cols = colVirtualizer.getVirtualItems();

  return (
    <div className="grid-outer" tabIndex={0} onKeyDown={handleKeyDown}>
      <div className="grid-corner" />
      <div className="grid-colheader" ref={colHeaderRef}>
        <div style={{ width: totalWidth, height: 20, position: "relative" }}>
          {cols.map((vc) => (
            <div key={vc.index} className="col-head" style={{ position: "absolute", left: vc.start, width: vc.size, top: 0, height: 20 }}>
              {colToLetter(vc.index)}
            </div>
          ))}
        </div>
      </div>
      <div className="grid-rowheader" ref={rowHeaderRef}>
        <div style={{ height: totalHeight, width: 40, position: "relative" }}>
          {rows.map((vr) => (
            <div key={vr.index} className="row-head" style={{ position: "absolute", top: vr.start, height: vr.size, left: 0, width: 40 }}>
              {visibleRows[vr.index] + 1}
            </div>
          ))}
        </div>
      </div>
      <div className="grid-body" ref={bodyRef}>
        <div style={{ width: totalWidth, height: totalHeight, position: "relative" }}>
          {rows.map((vr) => {
            const r = visibleRows[vr.index];
            return (
            <React.Fragment key={vr.index}>
              {cols.map((vc) => {
                const c = vc.index;
                const key = `${sheet}!R${r}C${c}`;
                const style = meta.styles?.[key] as CellStyle | undefined;
                const selected = inRanges(r, c, ranges);
                const isActive = active.row === r && active.col === c;
                const isEditing = store.editing && store.editing.row === r && store.editing.col === c;
                return (
                  <div
                    key={vc.index}
                    className={`cell ${selected ? "selected" : ""} ${isActive ? "active" : ""}`}
                    style={{
                      position: "absolute",
                      top: vr.start,
                      left: vc.start,
                      width: vc.size,
                      height: vr.size,
                      fontFamily: style?.fontFamily,
                      fontSize: style?.fontSize ? style.fontSize + "px" : undefined,
                      fontWeight: style?.bold ? "bold" : undefined,
                      fontStyle: style?.italic ? "italic" : undefined,
                      textDecoration: style?.underline ? "underline" : style?.strikethrough ? "line-through" : undefined,
                      color: style?.color,
                      background: style?.background,
                      textAlign: style?.hAlign,
                      border: undefined,
                    }}
                    onClick={(e) => handleCellClick(r, c, e)}
                    onDoubleClick={() => handleCellDblClick(r, c)}
                  >
                    {isEditing ? (
                      <input
                        className="cell-editor"
                        autoFocus
                        value={store.editing!.value}
                        onChange={(e) => store.updateEdit(e.target.value)}
                        onBlur={() => store.commitEdit()}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") { e.preventDefault(); store.commitEdit(); store.selectCell(Math.min(ROW_COUNT - 1, r + 1), c); }
                          if (e.key === "Tab") { e.preventDefault(); store.commitEdit(); store.selectCell(r, Math.min(COL_COUNT - 1, c + 1)); }
                          if (e.key === "Escape") { e.preventDefault(); store.cancelEdit(); }
                        }}
                      />
                    ) : (
                      <span className="cell-text">{displayValue(engine, meta, sheet, r, c)}</span>
                    )}
                  </div>
                );
              })}
            </React.Fragment>
          );
        })}
        </div>
      </div>
    </div>
  );
};


