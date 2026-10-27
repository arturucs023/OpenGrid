import { create } from "zustand";
import { FormulaEngine } from "../engine/hyper";
import type { WorkbookMeta, CellStyle } from "../model/types";
import { a1, parseA1 } from "../model/types";

const STORAGE_KEY = "excelclone.workbook.v1";
const ROW_COUNT_FILTER = 2000;

function defaultMeta(): WorkbookMeta {
  const sheets = { Hoja1: { name: "Hoja1", colWidths: {}, rowHeights: {} } };
  return { activeSheet: "Hoja1", styles: {}, sheets, zoom: 1, selection: { sheet: "Hoja1", ranges: ["A1"], active: "A1" } };
}

function cloneMeta(m: WorkbookMeta): WorkbookMeta {
  return structuredClone(m);
}

interface HistorySnap { sheets: Record<string, any[][]>; meta: WorkbookMeta; activeSheet: string; }

interface StoreState {
  workbookName: string;
  engine: FormulaEngine;
  meta: WorkbookMeta;
  selection: { ranges: string[]; active: { row: number; col: number }; anchor: { row: number; col: number } };
  editing: { row: number; col: number; value: string } | null;
  clipboard: { type: "copy" | "cut"; sheet: string; ranges: string[] } | null;
  zoom: number;
  past: HistorySnap[];
  future: HistorySnap[];

  // derived
  activeSheet: string;

  // actions
  init: () => void;
  pushHistory: () => void;
  undo: () => void;
  redo: () => void;

  setActiveSheet: (name: string) => void;
  addSheet: (name?: string) => void;
  removeSheet: (name: string) => void;
  renameSheet: (oldName: string, newName: string) => void;

  selectCell: (row: number, col: number, extend?: boolean, additive?: boolean) => void;
  selectRange: (r1: number, c1: number, r2: number, c2: number) => void;
  setSelectionActive: (row: number, col: number) => void;

  startEdit: (row: number, col: number, initial?: string) => void;
  updateEdit: (value: string) => void;
  commitEdit: () => void;
  cancelEdit: () => void;

  setCellInput: (sheet: string, row: number, col: number, input: any) => void;

  copy: () => void;
  cut: () => void;
  paste: (targetRow: number, targetCol: number) => void;

  setStyle: (style: Partial<CellStyle>, target?: { r1:number;c1:number;r2:number;c2:number }) => void;
  setColWidth: (col: number, width: number) => void;
  setRowHeight: (row: number, height: number) => void;
  setZoom: (z: number) => void;
  filterByActiveColumn: (text: string) => void;
  clearFilter: () => void;

  saveLocal: () => void;
  loadLocal: () => boolean;
  resetDemo: () => void;
}

function snap(engine: FormulaEngine, meta: WorkbookMeta): HistorySnap {
  return { sheets: engine.exportSheets(), meta: cloneMeta(meta), activeSheet: meta.activeSheet };
}
function restoreSnap(s: HistorySnap): { engine: FormulaEngine; meta: WorkbookMeta } {
  const meta = cloneMeta(s.meta);
  const engine = new FormulaEngine(meta, s.sheets);
  meta.activeSheet = s.activeSheet;
  return { engine, meta };
}

function ensureSheetMeta(meta: WorkbookMeta, name: string) {
  if (!meta.sheets[name]) meta.sheets[name] = { name, colWidths: {}, rowHeights: {} };
}

export const useStore = create<StoreState>((set, get) => ({
  workbookName: "Libro1",
  engine: new FormulaEngine(defaultMeta(), { Hoja1: [[]] }),
  meta: defaultMeta(),
  selection: { ranges: ["A1"], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } },
  editing: null,
  clipboard: null,
  zoom: 1,
  past: [],
  future: [],
  activeSheet: "Hoja1",

  init: () => {
    const meta = defaultMeta();
    const engine = new FormulaEngine(meta, { Hoja1: [[]] });
    set({ engine, meta, activeSheet: meta.activeSheet, selection: { ranges: ["A1"], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } }, past: [], future: [] });
  },

  pushHistory: () => {
    const { engine, meta, past } = get();
    const h = [...past, snap(engine, meta)];
    if (h.length > 50) h.shift();
    set({ past: h, future: [] });
  },

  undo: () => {
    const { past, future, engine, meta } = get();
    if (past.length === 0) return;
    future.push(snap(engine, meta));
    const s = past.pop()!;
    const { engine: ne, meta: nm } = restoreSnap(s);
    set({ engine: ne, meta: nm, activeSheet: nm.activeSheet, past: [...past], future: [...future] });
  },

  redo: () => {
    const { past, future, engine, meta } = get();
    if (future.length === 0) return;
    past.push(snap(engine, meta));
    const s = future.pop()!;
    const { engine: ne, meta: nm } = restoreSnap(s);
    set({ engine: ne, meta: nm, activeSheet: nm.activeSheet, past: [...past], future: [...future] });
  },

  setActiveSheet: (name) => {
    const { meta } = get();
    meta.activeSheet = name;
    ensureSheetMeta(meta, name);
    set({ meta: cloneMeta(meta), activeSheet: name });
  },

  addSheet: (name) => {
    get().pushHistory();
    const { engine, meta } = get();
    let n = name || `Hoja${Object.keys(meta.sheets).length + 1}`;
    // unique
    let base = n; let i = 2;
    while (meta.sheets[n] || engine.sheetNames().includes(n)) { n = `${base} (${i++})`; }
    engine.addSheet(n);
    ensureSheetMeta(meta, n);
    meta.sheets[n].name = n;
    meta.activeSheet = n;
    set({ engine, meta: cloneMeta(meta), activeSheet: n, selection: { ranges: ["A1"], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } } });
  },

  removeSheet: (name) => {
    const { meta } = get();
    if (Object.keys(meta.sheets).length <= 1) return;
    get().pushHistory();
    const { engine, meta: m } = get();
    engine.removeSheet(name);
    delete m.sheets[name];
    if (m.activeSheet === name) {
      const first = Object.keys(m.sheets)[0];
      m.activeSheet = first;
    }
    // remove styles for sheet
    for (const k of Object.keys(m.styles)) if (k.startsWith(name + "!")) delete m.styles[k];
    set({ engine, meta: cloneMeta(m), activeSheet: m.activeSheet });
  },

  renameSheet: (oldName, newName) => {
    if (!newName || oldName === newName) return;
    const { meta } = get();
    if (meta.sheets[newName]) return;
    get().pushHistory();
    const { engine, meta: m } = get();
    engine.renameSheet(oldName, newName);
    // rename meta + styles
    if (m.sheets[oldName]) { m.sheets[newName] = { ...m.sheets[oldName], name: newName }; delete m.sheets[oldName]; }
    const styles: Record<string, CellStyle> = {};
    for (const [k, v] of Object.entries(m.styles)) {
      if (k.startsWith(oldName + "!")) styles[newName + "!" + k.slice(oldName.length + 1)] = v; else styles[k] = v;
    }
    m.styles = styles;
    if (m.activeSheet === oldName) m.activeSheet = newName;
    set({ engine, meta: cloneMeta(m), activeSheet: m.activeSheet });
  },

  selectCell: (row, col, extend, additive) => {
    const { selection } = get();
    const addr = a1(row, col);
    if (extend) {
      const ranges = [...selection.ranges];
      const anchor = selection.anchor || selection.active;
      const nr1 = Math.min(anchor.row, row);
      const nc1 = Math.min(anchor.col, col);
      const nr2 = Math.max(anchor.row, row);
      const nc2 = Math.max(anchor.col, col);
      const newRange = `${a1(nr1, nc1)}:${a1(nr2, nc2)}`;
      ranges[ranges.length - 1] = newRange;
      set({ selection: { ranges, active: { row, col }, anchor } });
    } else if (additive) {
      set({ selection: { ranges: [...selection.ranges, addr], active: { row, col }, anchor: { row, col } } });
    } else {
      set({ selection: { ranges: [addr], active: { row, col }, anchor: { row, col } } });
    }
  },

  selectRange: (r1, c1, r2, c2) => {
    const rr1 = Math.min(r1, r2), cc1 = Math.min(c1, c2), rr2 = Math.max(r1, r2), cc2 = Math.max(c1, c2);
    const range = `${a1(rr1, cc1)}:${a1(rr2, cc2)}`;
    set({ selection: { ranges: [range], active: { row: rr2, col: cc2 }, anchor: { row: rr1, col: cc1 } } });
  },

  setSelectionActive: (row, col) => {
    const { selection } = get();
    set({ selection: { ...selection, active: { row, col } } });
  },

  startEdit: (row, col, initial) => {
    const { engine, activeSheet } = get();
    const raw = initial !== undefined ? initial : (engine.getSerialized(activeSheet, row, col) ?? "");
    const str = raw === null || raw === undefined ? "" : String(raw);
    set({ editing: { row, col, value: str } });
  },

  updateEdit: (value) => {
    const { editing } = get();
    if (!editing) return;
    set({ editing: { ...editing, value } });
  },

  commitEdit: () => {
    const { editing, activeSheet } = get();
    if (!editing) return;
    get().pushHistory();
    const { engine } = get();
    const { row, col, value } = editing;
    // parse input
    let input: any = value;
    if (value.startsWith("=")) input = value;
    else if (value === "") input = null;
    else if (!isNaN(Number(value)) && value.trim() !== "") input = Number(value);
    else if (value.toLowerCase() === "true") input = true;
    else if (value.toLowerCase() === "false") input = false;
    engine.setCell(activeSheet, row, col, input);
    set({ editing: null });
  },

  cancelEdit: () => set({ editing: null }),

  setCellInput: (sheet, row, col, input) => {
    get().pushHistory();
    const { engine } = get();
    engine.setCell(sheet, row, col, input);
    set({ engine: get().engine });
  },

  copy: () => {
    const { activeSheet, selection } = get();
    set({ clipboard: { type: "copy", sheet: activeSheet, ranges: [...selection.ranges] } });
  },
  cut: () => {
    const { activeSheet, selection } = get();
    set({ clipboard: { type: "cut", sheet: activeSheet, ranges: [...selection.ranges] } });
  },
  paste: (targetRow, targetCol) => {
    const { clipboard, engine, activeSheet, meta } = get();
    if (!clipboard) return;
    get().pushHistory();
    // For simplicity: copy serialized values from source range to target
    const srcSheet = clipboard.sheet;
    // parse first range
    const rangeStr = clipboard.ranges[0];
    const parts = rangeStr.split(":");
    const s = parseA1(parts[0])!;
    const e = parts[1] ? parseA1(parts[1])! : s;
    const rows = e.row - s.row + 1;
    const cols = e.col - s.col + 1;
    // read source data
    const data: any[][] = [];
    for (let r = 0; r < rows; r++) {
      const row: any[] = [];
      for (let c = 0; c < cols; c++) {
        row.push(engine.getSerialized(srcSheet, s.row + r, s.col + c));
      }
      data.push(row);
    }
    // apply with relative refs adjustment? HyperFormula setCellContents with formula string keeps refs absolute? 
    // We'll just set serialized (formulas as-is) -> refs won't shift. For MVP okay-ish; improve by adjusting.
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const val = data[r][c];
        engine.setCell(activeSheet, targetRow + r, targetCol + c, val === null ? null : val);
      }
    }
    if (clipboard.type === "cut") {
      for (let r = 0; r < rows; r++)
        for (let c = 0; c < cols; c++)
          engine.setCell(srcSheet, s.row + r, s.col + c, null);
      set({ clipboard: null });
    }
    set({ engine: get().engine, meta: cloneMeta(meta) });
  },

  setStyle: (style, target) => {
    get().pushHistory();
    const { meta, selection, activeSheet } = get();
    let r1, c1, r2, c2;
    if (target) { r1=target.r1; c1=target.c1; r2=target.r2; c2=target.c2; }
    else {
      // from selection first range
      const rangeStr = selection.ranges[0];
      const parts = rangeStr.split(":");
      const s = parseA1(parts[0])!;
      const e = parts[1] ? parseA1(parts[1])! : s;
      r1 = s.row; c1 = s.col; r2 = e.row; c2 = e.col;
    }
    const sheet = activeSheet;
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const key = `${sheet}!R${r}C${c}`;
        const prev = meta.styles[key] || {};
        meta.styles[key] = { ...prev, ...style };
      }
    }
    set({ meta: cloneMeta(meta) });
  },

  setColWidth: (col, width) => {
    get().pushHistory();
    const { meta, activeSheet } = get();
    ensureSheetMeta(meta, activeSheet);
    meta.sheets[activeSheet].colWidths[col] = width;
    set({ meta: cloneMeta(meta) });
  },
  setRowHeight: (row, height) => {
    get().pushHistory();
    const { meta, activeSheet } = get();
    ensureSheetMeta(meta, activeSheet);
    meta.sheets[activeSheet].rowHeights[row] = height;
    set({ meta: cloneMeta(meta) });
  },
  setZoom: (z) => set({ zoom: z, meta: cloneMeta(get().meta) }),
  filterByActiveColumn: (text) => {
    get().pushHistory();
    const { meta, activeSheet, selection } = get();
    ensureSheetMeta(meta, activeSheet);
    const col = selection.active.col;
    const hidden: number[] = [];
    const t = text.toLowerCase();
    for (let r = 1; r < ROW_COUNT_FILTER; r++) {
      const v = get().engine.getCellValue(activeSheet, r, col);
      const s = v === null || v === undefined ? "" : String(v).toLowerCase();
      if (!s.includes(t)) hidden.push(r);
    }
    meta.sheets[activeSheet].hiddenRows = hidden;
    set({ meta: cloneMeta(meta) });
  },
  clearFilter: () => {
    get().pushHistory();
    const { meta, activeSheet } = get();
    ensureSheetMeta(meta, activeSheet);
    meta.sheets[activeSheet].hiddenRows = [];
    set({ meta: cloneMeta(meta) });
  },

  saveLocal: () => {
    const { engine, meta, workbookName } = get();
    const payload = { name: workbookName, sheets: engine.exportSheets(), meta };
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); } catch {}
  },
  loadLocal: () => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      const payload = JSON.parse(raw);
      const meta = payload.meta as WorkbookMeta;
      const engine = new FormulaEngine(meta, payload.sheets);
      set({ workbookName: payload.name || "Libro1", engine, meta, activeSheet: meta.activeSheet, past: [], future: [] });
      return true;
    } catch { return false; }
  },
  resetDemo: () => {
    // will be filled by demo module
  },
}));
