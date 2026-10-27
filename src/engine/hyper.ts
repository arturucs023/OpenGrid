import { HyperFormula } from "hyperformula";
import type { WorkbookMeta } from "../model/types";

export const HF_OPTIONS = { licenseKey: "gpl-v3" as const };

export class FormulaEngine {
  private hf: HyperFormula;
  private meta: WorkbookMeta;

  constructor(meta: WorkbookMeta, sheets: Record<string, (string | number | boolean | null)[][]>) {
    this.meta = meta;
    this.hf = HyperFormula.buildFromSheets(sheets, HF_OPTIONS);
    // ensure sheet names/order
  }

  static empty(meta: WorkbookMeta): FormulaEngine {
    const sheets: Record<string, any[][]> = {};
    for (const name of Object.keys(meta.sheets)) sheets[name] = [[]];
    if (Object.keys(sheets).length === 0) sheets["Hoja1"] = [[]];
    return new FormulaEngine(meta, sheets);
  }

  get hfInstance() { return this.hf; }

  sheetNames() { return this.hf.getSheetNames(); }
  sheetId(name: string) { return this.hf.getSheetId(name); }

  getCellValue(sheet: string, row: number, col: number) {
    const id = this.hf.getSheetId(sheet);
    if (id === undefined) return null;
    try { return this.hf.getCellValue({ sheet: id, row, col }); } catch { return null; }
  }
  getCellFormula(sheet: string, row: number, col: number): string | undefined {
    const id = this.hf.getSheetId(sheet);
    if (id === undefined) return undefined;
    try { return this.hf.getCellFormula({ sheet: id, row, col }) ?? undefined; } catch { return undefined; }
  }
  getSerialized(sheet: string, row: number, col: number) {
    const id = this.hf.getSheetId(sheet);
    if (id === undefined) return null;
    try { return this.hf.getCellSerialized({ sheet: id, row, col }); } catch { return null; }
  }

  setCell(sheet: string, row: number, col: number, value: string | number | boolean | null) {
    const id = this.hf.getSheetId(sheet);
    if (id === undefined) return;
    this.hf.setCellContents({ sheet: id, row, col }, value);
  }

  setRange(sheet: string, r1: number, c1: number, _r2: number, _c2: number, value: any) {
    const id = this.hf.getSheetId(sheet);
    if (id === undefined) return;
    this.hf.setCellContents({ sheet: id, row: r1, col: c1 }, value);
  }

  clearCell(sheet: string, row: number, col: number) {
    this.setCell(sheet, row, col, null);
  }

  addSheet(name: string) {
    const n = this.hf.addSheet(name);
    if (!this.meta.sheets[n]) this.meta.sheets[n] = { name: n, colWidths: {}, rowHeights: {} };
    return n;
  }
  removeSheet(name: string) {
    const id = this.hf.getSheetId(name);
    if (id === undefined) return;
    this.hf.removeSheet(id);
    delete this.meta.sheets[name];
  }
  renameSheet(oldName: string, newName: string) {
    const id = this.hf.getSheetId(oldName);
    if (id === undefined) return;
    this.hf.renameSheet(id, newName);
    if (this.meta.sheets[oldName]) {
      this.meta.sheets[newName] = { ...this.meta.sheets[oldName], name: newName };
      delete this.meta.sheets[oldName];
    }
    if (this.meta.activeSheet === oldName) this.meta.activeSheet = newName;
  }

  getSheetValues(sheet: string) {
    const id = this.hf.getSheetId(sheet);
    if (id === undefined) return [];
    return this.hf.getSheetValues(id);
  }
  getAllSerialized() { return this.hf.getAllSheetsSerialized(); }

  undo() { this.hf.undo(); }
  redo() { this.hf.redo(); }
  canUndo() { return this.hf.isThereSomethingToUndo(); }
  canRedo() { return this.hf.isThereSomethingToRedo(); }

  exportSheets() { return this.hf.getAllSheetsSerialized(); }
  importSheets(sheets: Record<string, any[][]>) {
    this.hf = HyperFormula.buildFromSheets(sheets, HF_OPTIONS);
  }
}
