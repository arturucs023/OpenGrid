export type CellAddress = { sheet: string; row: number; col: number };

export interface BorderSpec {
  style?: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";
  color?: string;
}

export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  background?: string;
  hAlign?: "left" | "center" | "right" | "justify";
  vAlign?: "top" | "middle" | "bottom";
  wrap?: boolean;
  numberFormat?: string;
  borders?: {
    top?: BorderSpec;
    right?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
  };
}

export interface SheetMeta {
  name: string;
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  frozenRows?: number;
  frozenCols?: number;
  hiddenRows?: number[];
  hiddenCols?: number[];
  merged?: { r1: number; c1: number; r2: number; c2: number }[];
}

export interface WorkbookMeta {
  activeSheet: string;
  styles: Record<string, CellStyle>; // key: sheet!A1
  sheets: Record<string, SheetMeta>;
  zoom: number;
  selection?: { sheet: string; ranges: string[]; active: string };
}

export interface Workbook {
  name: string;
  meta: WorkbookMeta;
}

export const cellKey = (sheet: string, row: number, col: number) =>
  `${sheet}!R${row}C${col}`;

export const colToLetter = (col: number): string => {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

export const letterToCol = (letters: string): number => {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
};

export const a1 = (row: number, col: number): string =>
  `${colToLetter(col)}${row + 1}`;

export const parseA1 = (ref: string): { row: number; col: number } | null => {
  const m = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return { col: letterToCol(m[1].toUpperCase()), row: parseInt(m[2], 10) - 1 };
};

export const rangeA1 = (r1: number, c1: number, r2: number, c2: number): string =>
  `${a1(r1, c1)}:${a1(r2, c2)}`;
