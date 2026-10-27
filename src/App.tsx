import { useState, useEffect } from "react";
import { useStore } from "./store/workbookStore";
import { Grid } from "./components/Grid";
import { colToLetter } from "./model/types";
import { loadDemo } from "./demo/demoWorkbook";
import type { CellStyle } from "./model/types";

const RIBBON_TABS = ["Inicio", "Insertar", "Diseño de página", "Fórmulas", "Datos", "Revisar", "Vista"];

export default function App() {
  const store = useStore();
  const [ribbonTab, setRibbonTab] = useState("Inicio");
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; kind: string } | null>(null);
  const [dialog, setDialog] = useState<{ type: string; value: string } | null>(null);
  const [renameTarget, setRenameTarget] = useState<string | null>(null);

  useEffect(() => {
    if (import.meta.env.DEV) (window as any).__store = useStore;
    store.init();
    const loaded = store.loadLocal();
    if (!loaded) loadDemo(store);
  }, []);

  // demo load handled after demo module exists; fallback init
  const sheet = store.activeSheet;
  const active = store.selection.active;
  const addr = `${colToLetter(active.col)}${active.row + 1}`;

  const formulaValue = store.editing && store.editing.row === active.row && store.editing.col === active.col
    ? store.editing.value
    : (store.engine.getSerialized(sheet, active.row, active.col) ?? "");
  const formulaStr = formulaValue === null || formulaValue === undefined ? "" : String(formulaValue);

  const toggleStyle = (key: keyof CellStyle, val?: any) => {
    const cur = store.meta.styles?.[`${sheet}!R${active.row}C${active.col}`]?.[key];
    const next = val !== undefined ? val : !cur;
    store.setStyle({ [key]: next } as any);
  };

  const currentStyle = store.meta.styles?.[`${sheet}!R${active.row}C${active.col}`] || {};

  return (
    <div className="app">
      <div className="titlebar">
        <span>ExcelClone - {store.workbookName}</span>
        <span className="title" />
      </div>

      <div className="ribbon">
        <div className="ribbon-tabs">
          {RIBBON_TABS.map((t) => (
            <div key={t} className={`ribbon-tab ${ribbonTab === t ? "active" : ""}`} onClick={() => setRibbonTab(t)}>{t}</div>
          ))}
        </div>
        <div className="ribbon-body">
          {ribbonTab === "Inicio" && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => { store.undo(); }} title="Deshacer"><span className="icon">↶</span>Deshacer</button>
                  <button className="ribbon-btn" onClick={() => { store.redo(); }} title="Rehacer"><span className="icon">↷</span>Rehacer</button>
                </div>
                <div className="ribbon-group-title">Deshacer</div>
              </div>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => store.copy()} title="Copiar"><span className="icon">⧉</span>Copiar</button>
                  <button className="ribbon-btn" onClick={() => store.cut()} title="Cortar"><span className="icon">✂</span>Cortar</button>
                  <button className="ribbon-btn" onClick={() => store.paste(active.row, active.col)} title="Pegar"><span className="icon">📋</span>Pegar</button>
                </div>
                <div className="ribbon-group-title">Portapapeles</div>
              </div>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <select className="ribbon-input" value={currentStyle.fontFamily || "Calibri"} onChange={(e) => toggleStyle("fontFamily", e.target.value)}>
                    <option>Calibri</option><option>Arial</option><option>Times New Roman</option><option>Consolas</option>
                  </select>
                  <select className="ribbon-input" value={currentStyle.fontSize || 11} onChange={(e) => toggleStyle("fontSize", Number(e.target.value))}>
                    {[8,9,10,11,12,14,16,18,20,24,28,32].map((s) => <option key={s}>{s}</option>)}
                  </select>
                </div>
                <div className="ribbon-row">
                  <button className={`ribbon-btn ${currentStyle.bold ? "active" : ""}`} onClick={() => toggleStyle("bold")} title="Negrita"><span className="icon" style={{ fontWeight: "bold" }}>B</span></button>
                  <button className={`ribbon-btn ${currentStyle.italic ? "active" : ""}`} onClick={() => toggleStyle("italic")} title="Cursiva"><span className="icon" style={{ fontStyle: "italic" }}>I</span></button>
                  <button className={`ribbon-btn ${currentStyle.underline ? "active" : ""}`} onClick={() => toggleStyle("underline")} title="Subrayado"><span className="icon" style={{ textDecoration: "underline" }}>U</span></button>
                  <input type="color" title="Color texto" value={currentStyle.color || "#000000"} onChange={(e) => toggleStyle("color", e.target.value)} style={{ width: 28, height: 24 }} />
                  <input type="color" title="Relleno" value={currentStyle.background || "#ffffff"} onChange={(e) => toggleStyle("background", e.target.value)} style={{ width: 28, height: 24 }} />
                </div>
                <div className="ribbon-group-title">Fuente</div>
              </div>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => store.setStyle({ hAlign: "left" })} title="Izquierda">⬅</button>
                  <button className="ribbon-btn" onClick={() => store.setStyle({ hAlign: "center" })} title="Centro">⬌</button>
                  <button className="ribbon-btn" onClick={() => store.setStyle({ hAlign: "right" })} title="Derecha">➡</button>
                  <button className="ribbon-btn" onClick={() => store.setStyle({ wrap: true })} title="Ajustar texto">Wrap</button>
                </div>
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => store.setStyle({ borders: { top: { style: "thin", color: "#000" }, right: { style: "thin", color: "#000" }, bottom: { style: "thin", color: "#000" }, left: { style: "thin", color: "#000" } } })} title="Bordes">のではない</button>
                  <select className="ribbon-input" value={currentStyle.numberFormat || "General"} onChange={(e) => store.setStyle({ numberFormat: e.target.value })}>
                    <option>General</option><option>0</option><option>0.00</option><option>#,##0</option><option>#,##0.00</option><option>0%</option><option>$#,##0.00</option><option>dd/mm/yyyy</option>
                  </select>
                </div>
                <div className="ribbon-group-title">Alineación / Número</div>
              </div>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => { const w = prompt("Ancho columna", "100"); if (w) store.setColWidth(active.col, Number(w)); }} title="Ancho col">Col↔</button>
                  <button className="ribbon-btn" onClick={() => { const h = prompt("Alto fila", "24"); if (h) store.setRowHeight(active.row, Number(h)); }} title="Alto fila">Fila↕</button>
                </div>
                <div className="ribbon-group-title">Celdas</div>
              </div>
            </>
          )}
          {ribbonTab === "Insertar" && (
            <div className="ribbon-group">
              <div className="ribbon-row">
                <button className="ribbon-btn" onClick={() => { const svg = buildChartSvg(store); setDialog({ type: "chart", value: svg }); }} title="Gráfico"><span className="icon">📊</span>Gráfico</button>
                <button className="ribbon-btn" onClick={() => store.addSheet()} title="Nueva hoja"><span className="icon"> </span>Hoja</button>
              </div>
              <div className="ribbon-group-title">Insertar</div>
            </div>
          )}
          {ribbonTab === "Datos" && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => sortSelection(store, true)} title="Ordenar A-Z">A→Z</button>
                  <button className="ribbon-btn" onClick={() => sortSelection(store, false)} title="Ordenar Z-A">Z→A</button>
                </div>
                <div className="ribbon-group-title">Ordenar</div>
              </div>
              <div className="ribbon-group">
                <div className="ribbon-row">
                  <button className="ribbon-btn" onClick={() => { const t = prompt("Filtrar columna activa contiene:"); if (t !== null) store.filterByActiveColumn(t); }} title="Filtrar">Filtrar</button>
                  <button className="ribbon-btn" onClick={() => store.clearFilter()} title="Quitar filtro">Quitar</button>
                </div>
                <div className="ribbon-group-title">Filtrar</div>
              </div>
            </>
          )}
          {["Diseño de página", "Fórmulas", "Revisar", "Vista"].includes(ribbonTab) && (
            <div className="ribbon-group">
              <div className="ribbon-row">
                <span style={{ fontSize: 12, color: "#666" }}>Funciones disponibles en la cinta. Use fórmulas en celdas (=SUM, =AVERAGE, =IF...).</span>
              </div>
            </div>
          )}
        </div>
      </div>

      <div className="formulabar">
        <div className="namebox">{addr}</div>
        <div className="fx">fx</div>
        <input
          className="input"
          value={formulaStr}
          onChange={(e) => {
            if (!store.editing || store.editing.row !== active.row || store.editing.col !== active.col) {
              store.startEdit(active.row, active.col, e.target.value);
            } else {
              store.updateEdit(e.target.value);
            }
          }}
          onBlur={() => { if (store.editing) store.commitEdit(); }}
          onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); store.commitEdit(); } if (e.key === "Escape") { e.preventDefault(); store.cancelEdit(); } }}
        />
      </div>

      <Grid />

      <div className="sheetbar">
        {Object.keys(store.meta.sheets).map((name) => (
          <div
            key={name}
            className={`sheet-tab ${store.activeSheet === name ? "active" : ""}`}
            onClick={() => store.setActiveSheet(name)}
            onDoubleClick={() => { setRenameTarget(name); setDialog({ type: "rename", value: name }); }}
            onContextMenu={(e) => { e.preventDefault(); setCtxMenu({ x: e.clientX, y: e.clientY, kind: name }); }}
          >
            {name}
          </div>
        ))}
        <div className="sheet-add" onClick={() => store.addSheet()} title="Nueva hoja">+</div>
        <div className="sheet-menu">
          <button className="btn" onClick={() => store.saveLocal()}>Guardar</button>
          <button className="btn" onClick={() => { if (store.loadLocal()) alert("Cargado"); }}>Cargar</button>
          <button className="btn" onClick={() => exportXlsx(store)}>Exportar XLSX</button>
          <button className="btn" onClick={() => importXlsx(store)}>Importar XLSX</button>
        </div>
      </div>

      <div className="statusbar">
        <span>Listo</span>
        <span>Zoom</span>
        <div className="zoom">
          <input type="range" min={0.5} max={2} step={0.1} value={store.zoom} onChange={(e) => store.setZoom?.(Number(e.target.value))} />
          <span>{Math.round(store.zoom * 100)}%</span>
        </div>
      </div>

      {ctxMenu && (
        <div className="contextmenu" style={{ left: ctxMenu.x, top: ctxMenu.y }} onMouseLeave={() => setCtxMenu(null)}>
          <button onClick={() => { store.removeSheet(ctxMenu.kind); setCtxMenu(null); }}>Eliminar hoja</button>
          <button onClick={() => { setRenameTarget(ctxMenu.kind); setDialog({ type: "rename", value: ctxMenu.kind }); setCtxMenu(null); }}>Renombrar</button>
        </div>
      )}

      {dialog?.type === "rename" && (
        <div className="dialog-overlay" onClick={() => setDialog(null)}>
          <div className="dialog" onClick={(e) => e.stopPropagation()}>
            <h3>Renombrar hoja</h3>
            <input value={dialog.value} onChange={(e) => setDialog({ ...dialog, value: e.target.value })} />
            <div className="dialog-actions">
              <button onClick={() => setDialog(null)}>Cancelar</button>
              <button onClick={() => { if (renameTarget) store.renameSheet(renameTarget, dialog.value); setDialog(null); setRenameTarget(null); }}>Aceptar</button>
            </div>
          </div>
        </div>
      )}

      {dialog?.type === "chart" && (
        <div className="dialog-overlay" onClick={() => setDialog(null)}>
          <div className="dialog" onClick={(e) => e.stopPropagation()}>
            <h3>Gráfico de barras</h3>
            <div dangerouslySetInnerHTML={{ __html: dialog.value }} />
            <div className="dialog-actions"><button onClick={() => setDialog(null)}>Cerrar</button></div>
          </div>
        </div>
      )}
    </div>
  );
}

// ---- helpers ----
function buildChartSvg(store: any): string {
  try {
    const sheet = store.activeSheet;
    const ranges = store.selection.ranges;
    const range = ranges[0];
    const parts = range.split(":");
    const s = parseA1(parts[0]); const e = parts[1] ? parseA1(parts[1]) : s;
    let html = "<svg width='320' height='220' xmlns='http://www.w3.org/2000/svg'>";
    const vals: number[] = [];
    const labels: string[] = [];
    for (let r = s.row; r <= e.row; r++) {
      const lab = store.engine.getCellValue(sheet, r, s.col);
      const v = store.engine.getCellValue(sheet, r, s.col + 1);
      labels.push(String(lab ?? ""));
      if (typeof v === "number") vals.push(v); else if (typeof v === "string" && !isNaN(Number(v))) vals.push(Number(v));
    }
    const max = Math.max(...vals, 1);
    vals.forEach((v, i) => {
      const h = (v / max) * 160;
      html += `<rect x='${i * 44 + 20}' y='${180 - h}' width='34' height='${h}' fill='#217346'/><text x='${i * 44 + 20}' y='200' font-size='10'>${labels[i]}</text>`;
    });
    html += "</svg>";
    return html;
  } catch (err) { console.error(err); return "<svg></svg>"; }
}

function sortSelection(store: any, asc: boolean) {
  store.pushHistory();
  const sheet = store.activeSheet;
  const range = store.selection.ranges[0];
  const parts = range.split(":");
  const s = parseA1(parts[0]); const e = parts[1] ? parseA1(parts[1]) : s;
  const rows: any[] = [];
  for (let r = s.row; r <= e.row; r++) {
    const row: any[] = [];
    for (let c = s.col; c <= e.col; c++) row.push(store.engine.getSerialized(sheet, r, c));
    rows.push({ r, row });
  }
  const colIndex = 0;
  rows.sort((a, b) => {
    const va = a.row[colIndex], vb = b.row[colIndex];
    if (va === vb) return 0;
    if (va === null) return 1;
    if (vb === null) return -1;
    return asc ? (va > vb ? 1 : -1) : (va > vb ? -1 : 1);
  });
  // write back
  rows.forEach((item, i) => {
    const targetRow = s.row + i;
    for (let c = 0; c < item.row.length; c++) {
      store.engine.setCell(sheet, targetRow, s.col + c, item.row[c]);
    }
  });
  store.setZoom(store.zoom);
}

function parseA1(ref: string) {
  const m = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return { row: 0, col: 0 };
  let n = 0;
  for (let i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
  return { col: n - 1, row: parseInt(m[2], 10) - 1 };
}

function exportXlsx(store: any) {
  import("exceljs").then(async (mod) => {
    const ExcelJS = mod.default || mod;
    const wb = new ExcelJS.Workbook();
    for (const name of Object.keys(store.meta.sheets)) {
      const ws = wb.addWorksheet(name);
      const data = store.engine.getSheetValues(name);
      data.forEach((row: any[], r: number) => {
        row.forEach((_v, c) => {
          const cell = ws.getCell(r + 1, c + 1);
          const raw = store.engine.getSerialized(name, r, c);
          cell.value = raw && typeof raw === "string" && raw.startsWith("=") ? { formula: raw.slice(1) } : (raw ?? null);
          const style = store.meta.styles?.[`${name}!R${r}C${c}`];
          if (style) {
            if (style.bold) cell.font = { ...(cell.font || {}), bold: true };
            if (style.color) cell.font = { ...(cell.font || {}), color: { argb: style.color.replace("#", "FF") } };
            if (style.background) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: style.background.replace("#", "FF") } };
            if (style.hAlign) cell.alignment = { ...(cell.alignment || {}), horizontal: style.hAlign };
          }
        });
      });
    }
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = (store.workbookName || "libro") + ".xlsx"; a.click();
    URL.revokeObjectURL(url);
  });
}

function importXlsx(store: any) {
  const inp = document.createElement("input");
  inp.type = "file"; inp.accept = ".xlsx";
  inp.onchange = async () => {
    const file = inp.files?.[0];
    if (!file) return;
    const ExcelJS = await import("exceljs");
    const wb = new (ExcelJS.default || ExcelJS).Workbook();
    const buf = await file.arrayBuffer();
    await wb.xlsx.load(buf);
    // rebuild store
    store.pushHistory();
    // remove existing sheets
    for (const n of Object.keys(store.meta.sheets)) {
      if (Object.keys(store.meta.sheets).length > 1) store.removeSheet(n);
    }
    // Actually easier: reset
    const meta = store.meta;
    // clear
    for (const n of Object.keys(meta.sheets)) delete meta.sheets[n];
    for (const k of Object.keys(meta.styles)) delete meta.styles[k];
    const sheetsData: Record<string, any[][]> = {};
    wb.eachSheet((ws: any) => {
      const name = ws.name;
      meta.sheets[name] = { name, colWidths: {}, rowHeights: {} };
      const rows: any[][] = [];
      ws.eachRow({ includeEmpty: true }, (row: any, r: number) => {
        const arr: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell: any, c: number) => {
          arr[c - 1] = cell.value && cell.value.formula ? "=" + cell.value.formula : cell.value;
        });
        rows[r - 1] = arr;
      });
      sheetsData[name] = rows.length ? rows : [[]];
    });
    const first = Object.keys(sheetsData)[0];
    meta.activeSheet = first;
    store.meta = meta;
    store.engine.importSheets(sheetsData);
    store.setActiveSheet(first);
    alert("Importado");
  };
  inp.click();
}

// patch store: add setZoom
// We'll monkey via prototype? Instead add to store not. Use local state? We'll just ignore zoom set; add action instead by editing store. 
