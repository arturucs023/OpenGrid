export function loadDemo(store: any) {
  store.init();
  store.addSheet("Ventas");
  store.addSheet("Resumen");
  try { store.removeSheet("Hoja1"); } catch {}
  store.setActiveSheet("Ventas");

  const v = "Ventas";
  const headers = ["Fecha", "Región", "Producto", "Cantidad", "Precio", "Total"];
  headers.forEach((h, i) => store.setCellInput(v, 0, i, h));
  store.setStyle({ bold: true, background: "#217346", color: "#ffffff", hAlign: "center" }, { r1: 0, c1: 0, r2: 0, c2: 5 });

  const regions = ["Norte", "Sur", "Este", "Oeste"];
  const products = ["Laptop", "Mouse", "Teclado", "Monitor"];
  const prices: Record<string, number> = { Laptop: 1200, Mouse: 25, Teclado: 80, Monitor: 300 };
  let r = 1;
  for (let i = 0; i < 24; i++) {
    const prod = products[i % products.length];
    const qty = 1 + (i % 5) + (i % 3);
    const price = prices[prod];
    store.setCellInput(v, r, 0, `2026-${String((i % 12) + 1).padStart(2, "0")}-${String((i % 28) + 1).padStart(2, "0")}`);
    store.setCellInput(v, r, 1, regions[i % regions.length]);
    store.setCellInput(v, r, 2, prod);
    store.setCellInput(v, r, 3, qty);
    store.setCellInput(v, r, 4, price);
    store.setCellInput(v, r, 5, `=D${r + 1}*E${r + 1}`);
    store.setStyle({ numberFormat: "#,##0.00" }, { r1: r, c1: 4, r2: r, c2: 5 });
    store.setStyle({ numberFormat: "0" }, { r1: r, c1: 3, r2: r, c2: 3 });
    r++;
  }
  const last = r - 1;
  // totals row
  store.setCellInput(v, r, 2, "Total");
  store.setCellInput(v, r, 5, `=SUM(F2:F${last + 1})`);
  store.setStyle({ bold: true }, { r1: r, c1: 2, r2: r, c2: 5 });
  store.setStyle({ numberFormat: "#,##0.00" }, { r1: r, c1: 5, r2: r, c2: 5 });

  // borders for table
  store.setStyle({ borders: { top: { style: "thin", color: "#bbb" }, right: { style: "thin", color: "#bbb" }, bottom: { style: "thin", color: "#bbb" }, left: { style: "thin", color: "#bbb" } } }, { r1: 0, c1: 0, r2: last + 1, c2: 5 });

  store.setActiveSheet("Resumen");
  // Resumen
  const s = "Resumen";
  store.setCellInput(s, 0, 0, "Métrica");
  store.setCellInput(s, 0, 1, "Valor");
  store.setStyle({ bold: true, background: "#217346", color: "#fff" }, { r1: 0, c1: 0, r2: 0, c2: 1 });
  const rows = [
    ["Total ventas", `=SUM(Ventas!F2:F${last + 1})`],
    ["Promedio venta", `=AVERAGE(Ventas!F2:F${last + 1})`],
    ["Venta mínima", `=MIN(Ventas!F2:F${last + 1})`],
    ["Venta máxima", `=MAX(Ventas!F2:F${last + 1})`],
    ["Cantidad registros", `=COUNT(Ventas!F2:F${last + 1})`],
    ["Ventas > 100", `=SUMIF(Ventas!F2:F${last + 1},">100")`],
  ];
  rows.forEach(([label, formula], i) => {
    store.setCellInput(s, i + 1, 0, label);
    store.setCellInput(s, i + 1, 1, formula);
    store.setStyle({ numberFormat: "#,##0.00" }, { r1: i + 1, c1: 1, r2: i + 1, c2: 1 });
  });

  store.setActiveSheet("Ventas");
  try { store.saveLocal(); } catch {}
}
