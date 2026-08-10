import { emptyCell, defaultStyle, deepClone, cellId, clamp } from "./utils.js";
import { evaluateFormula, formatDisplay, detectUrl, normalizeUrl, isHyperlinkValue } from "./formulas.js";
import { History } from "./history.js";

const DEFAULT_ROWS = 100;
const DEFAULT_COLS = 26;

function createSheet(name = "Sheet1") {
  return {
    id: crypto.randomUUID(),
    name,
    rows: DEFAULT_ROWS,
    cols: DEFAULT_COLS,
    cells: {},
    colWidths: {},
  };
}

export function createWorkbook(title = "Untitled spreadsheet") {
  const sheet = createSheet("Sheet1");
  return {
    title,
    activeSheetId: sheet.id,
    sheets: [sheet],
  };
}

export class SpreadsheetApp {
  constructor(workbook) {
    this.workbook = workbook || createWorkbook();
    this.history = new History();
    this.selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
    this.active = { row: 0, col: 0 };
    this.clipboard = null;
    this.dirty = false;
    this._valueCache = new Map();
    this.pushHistory();
  }

  get sheet() {
    return this.workbook.sheets.find((s) => s.id === this.workbook.activeSheetId);
  }

  snapshot() {
    return {
      workbook: deepClone(this.workbook),
      selection: { ...this.selection },
      active: { ...this.active },
    };
  }

  pushHistory() {
    this.history.push(this.snapshot());
  }

  restore(snap) {
    this.workbook = snap.workbook;
    this.selection = snap.selection;
    this.active = snap.active;
    this.invalidate();
    this.dirty = true;
  }

  undo() {
    const snap = this.history.undo();
    if (snap) this.restore(snap);
    return !!snap;
  }

  redo() {
    const snap = this.history.redo();
    if (snap) this.restore(snap);
    return !!snap;
  }

  invalidate() {
    this._valueCache.clear();
  }

  cellKey(row, col) {
    return `${row},${col}`;
  }

  getCell(row, col) {
    const key = this.cellKey(row, col);
    return this.sheet.cells[key] || emptyCell();
  }

  ensureCell(row, col) {
    const key = this.cellKey(row, col);
    if (!this.sheet.cells[key]) {
      this.sheet.cells[key] = emptyCell();
    }
    return this.sheet.cells[key];
  }

  setRaw(row, col, raw, { record = true } = {}) {
    if (record) this.pushHistory();
    const cell = this.ensureCell(row, col);
    cell.raw = raw ?? "";
    if (cell.raw === "" && !hasStyle(cell.style)) {
      delete this.sheet.cells[this.cellKey(row, col)];
    }
    this.invalidate();
    this.dirty = true;
  }

  setStyle(row, col, patch, { record = true } = {}) {
    if (record) this.pushHistory();
    const cell = this.ensureCell(row, col);
    cell.style = { ...defaultStyle(), ...cell.style, ...patch };
    this.dirty = true;
  }

  applyStyleToSelection(patch) {
    this.pushHistory();
    const { r1, c1, r2, c2 } = normalized(this.selection);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = this.ensureCell(r, c);
        cell.style = { ...defaultStyle(), ...cell.style, ...patch };
      }
    }
    this.dirty = true;
  }

  getValue(row, col) {
    const key = this.cellKey(row, col);
    if (this._valueCache.has(key)) return this._valueCache.get(key);

    const evaluating = new Set();
    const compute = (r, c) => {
      const k = this.cellKey(r, c);
      if (this._valueCache.has(k)) return this._valueCache.get(k);
      if (evaluating.has(k)) return "#CYCLE!";
      evaluating.add(k);
      const cell = this.getCell(r, c);
      const raw = cell.raw;
      let value;
      if (typeof raw === "string" && raw.startsWith("=")) {
        value = evaluateFormula(raw, compute, evaluating);
      } else {
        value = evaluateFormula(raw, compute, evaluating);
      }
      evaluating.delete(k);
      this._valueCache.set(k, value);
      return value;
    };

    return compute(row, col);
  }

  getDisplay(row, col) {
    const cell = this.getCell(row, col);
    return formatDisplay(this.getValue(row, col), cell.style);
  }

  getHyperlink(row, col) {
    const cell = this.getCell(row, col);
    if (cell.link) return normalizeUrl(cell.link);
    const value = this.getValue(row, col);
    if (isHyperlinkValue(value) && value.url) return normalizeUrl(value.url);
    const fromRaw = detectUrl(cell.raw);
    if (fromRaw) return normalizeUrl(fromRaw);
    const display = this.getDisplay(row, col);
    const fromDisplay = detectUrl(display);
    return fromDisplay ? normalizeUrl(fromDisplay) : null;
  }

  setHyperlink(row, col, url, label) {
    const href = normalizeUrl(url);
    if (!href) return;
    this.pushHistory();
    const text = label != null && String(label) !== "" ? String(label) : href;
    const escapedUrl = href.replace(/"/g, '""');
    const escapedLabel = text.replace(/"/g, '""');
    const cell = this.ensureCell(row, col);
    cell.raw = `=HYPERLINK("${escapedUrl}","${escapedLabel}")`;
    delete cell.link;
    this.invalidate();
    this.dirty = true;
  }

  removeHyperlink(row, col) {
    this.pushHistory();
    const cell = this.ensureCell(row, col);
    const value = this.getValue(row, col);
    if (isHyperlinkValue(value)) {
      cell.raw = value.label;
    }
    delete cell.link;
    this.invalidate();
    this.dirty = true;
  }

  select(row, col, extend = false) {
    const sheet = this.sheet;
    row = clamp(row, 0, sheet.rows - 1);
    col = clamp(col, 0, sheet.cols - 1);
    if (!extend) {
      this.active = { row, col };
      this.selection = { r1: row, c1: col, r2: row, c2: col };
    } else {
      this.selection = {
        r1: this.active.row,
        c1: this.active.col,
        r2: row,
        c2: col,
      };
    }
  }

  selectRange(r1, c1, r2, c2) {
    this.active = { row: r1, col: c1 };
    this.selection = { r1, c1, r2, c2 };
  }

  clearSelectionContents() {
    this.pushHistory();
    const { r1, c1, r2, c2 } = normalized(this.selection);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const key = this.cellKey(r, c);
        const cell = this.sheet.cells[key];
        if (!cell) continue;
        cell.raw = "";
        if (!hasStyle(cell.style)) delete this.sheet.cells[key];
      }
    }
    this.invalidate();
    this.dirty = true;
  }

  copySelection(cut = false) {
    const { r1, c1, r2, c2 } = normalized(this.selection);
    const data = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        row.push(deepClone(this.getCell(r, c)));
      }
      data.push(row);
    }
    this.clipboard = { data, cut, r1, c1, r2, c2 };
    if (cut) {
      // mark only; actual clear on paste
    }
    return this.clipboardToTSV();
  }

  clipboardToTSV() {
    if (!this.clipboard) return "";
    return this.clipboard.data
      .map((row) => row.map((c) => String(c.raw ?? "").replace(/\t/g, " ").replace(/\n/g, " ")).join("\t"))
      .join("\n");
  }

  paste(tsv = null) {
    this.pushHistory();
    const start = this.active;
    if (tsv != null && typeof tsv === "string" && (!this.clipboard || tsv !== this.clipboardToTSV())) {
      const rows = tsv.replace(/\r/g, "").split("\n").filter((l, i, arr) => !(i === arr.length - 1 && l === ""));
      for (let r = 0; r < rows.length; r++) {
        const cols = rows[r].split("\t");
        for (let c = 0; c < cols.length; c++) {
          const rr = start.row + r;
          const cc = start.col + c;
          if (rr >= this.sheet.rows || cc >= this.sheet.cols) continue;
          this.ensureCell(rr, cc).raw = cols[c];
        }
      }
    } else if (this.clipboard) {
      const { data, cut, r1, c1, r2, c2 } = this.clipboard;
      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < data[r].length; c++) {
          const rr = start.row + r;
          const cc = start.col + c;
          if (rr >= this.sheet.rows || cc >= this.sheet.cols) continue;
          this.sheet.cells[this.cellKey(rr, cc)] = deepClone(data[r][c]);
        }
      }
      if (cut) {
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            const inPaste =
              r >= start.row &&
              r < start.row + data.length &&
              c >= start.col &&
              c < start.col + data[0].length;
            if (!inPaste) delete this.sheet.cells[this.cellKey(r, c)];
          }
        }
        this.clipboard.cut = false;
      }
    }
    this.invalidate();
    this.dirty = true;
  }

  fillDown() {
    const { r1, c1, r2, c2 } = normalized(this.selection);
    if (r2 <= r1) return;
    this.pushHistory();
    for (let c = c1; c <= c2; c++) {
      const src = this.getCell(r1, c);
      for (let r = r1 + 1; r <= r2; r++) {
        this.sheet.cells[this.cellKey(r, c)] = deepClone(src);
      }
    }
    this.invalidate();
    this.dirty = true;
  }

  insertRow(at, below = false) {
    this.pushHistory();
    const sheet = this.sheet;
    const row = below ? at + 1 : at;
    const next = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [r, c] = key.split(",").map(Number);
      if (r >= row) next[`${r + 1},${c}`] = cell;
      else next[key] = cell;
    }
    sheet.cells = next;
    sheet.rows += 1;
    this.invalidate();
    this.dirty = true;
  }

  deleteRow(at) {
    if (this.sheet.rows <= 1) return;
    this.pushHistory();
    const sheet = this.sheet;
    const next = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [r, c] = key.split(",").map(Number);
      if (r === at) continue;
      if (r > at) next[`${r - 1},${c}`] = cell;
      else next[key] = cell;
    }
    sheet.cells = next;
    sheet.rows -= 1;
    this.select(Math.min(this.active.row, sheet.rows - 1), this.active.col);
    this.invalidate();
    this.dirty = true;
  }

  insertCol(at, right = false) {
    this.pushHistory();
    const sheet = this.sheet;
    const col = right ? at + 1 : at;
    const next = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [r, c] = key.split(",").map(Number);
      if (c >= col) next[`${r},${c + 1}`] = cell;
      else next[key] = cell;
    }
    sheet.cells = next;
    const widths = {};
    for (const [k, w] of Object.entries(sheet.colWidths || {})) {
      const ci = Number(k);
      if (ci >= col) widths[ci + 1] = w;
      else widths[ci] = w;
    }
    sheet.colWidths = widths;
    sheet.cols += 1;
    this.invalidate();
    this.dirty = true;
  }

  deleteCol(at) {
    if (this.sheet.cols <= 1) return;
    this.pushHistory();
    const sheet = this.sheet;
    const next = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [r, c] = key.split(",").map(Number);
      if (c === at) continue;
      if (c > at) next[`${r},${c - 1}`] = cell;
      else next[key] = cell;
    }
    sheet.cells = next;
    const widths = {};
    for (const [k, w] of Object.entries(sheet.colWidths || {})) {
      const ci = Number(k);
      if (ci === at) continue;
      if (ci > at) widths[ci - 1] = w;
      else widths[ci] = w;
    }
    sheet.colWidths = widths;
    sheet.cols -= 1;
    this.select(this.active.row, Math.min(this.active.col, sheet.cols - 1));
    this.invalidate();
    this.dirty = true;
  }

  setColWidth(col, width) {
    this.sheet.colWidths = this.sheet.colWidths || {};
    this.sheet.colWidths[col] = clamp(width, 40, 400);
    this.dirty = true;
  }

  getColWidth(col) {
    return this.sheet.colWidths?.[col] ?? 100;
  }

  addSheet() {
    this.pushHistory();
    const n = this.workbook.sheets.length + 1;
    const sheet = createSheet(`Sheet${n}`);
    this.workbook.sheets.push(sheet);
    this.workbook.activeSheetId = sheet.id;
    this.select(0, 0);
    this.invalidate();
    this.dirty = true;
  }

  switchSheet(id) {
    if (id === this.workbook.activeSheetId) return;
    this.pushHistory();
    this.workbook.activeSheetId = id;
    this.select(0, 0);
    this.invalidate();
  }

  renameSheet(id, name) {
    const sheet = this.workbook.sheets.find((s) => s.id === id);
    if (!sheet) return;
    this.pushHistory();
    sheet.name = name.trim() || sheet.name;
    this.dirty = true;
  }

  deleteSheet(id) {
    if (this.workbook.sheets.length <= 1) return false;
    this.pushHistory();
    this.workbook.sheets = this.workbook.sheets.filter((s) => s.id !== id);
    if (this.workbook.activeSheetId === id) {
      this.workbook.activeSheetId = this.workbook.sheets[0].id;
    }
    this.select(0, 0);
    this.invalidate();
    this.dirty = true;
    return true;
  }

  clearActiveSheet() {
    this.pushHistory();
    this.sheet.cells = {};
    this.invalidate();
    this.dirty = true;
  }

  activeCellId() {
    return cellId(this.active.row, this.active.col);
  }

  toCSV() {
    const sheet = this.sheet;
    const lines = [];
    for (let r = 0; r < sheet.rows; r++) {
      const row = [];
      let last = -1;
      for (let c = 0; c < sheet.cols; c++) {
        const raw = this.getCell(r, c).raw;
        if (raw !== "") last = c;
        row.push(raw);
      }
      if (last >= 0) {
        lines.push(
          row
            .slice(0, last + 1)
            .map((v) => {
              const s = String(v ?? "");
              if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
              return s;
            })
            .join(",")
        );
      }
    }
    return lines.join("\n");
  }

  toJSON() {
    return JSON.stringify(this.workbook, null, 2);
  }

  loadFromCSV(rows, { title } = {}) {
    this.pushHistory();
    const sheet = this.sheet;
    sheet.cells = {};
    const maxR = Math.max(rows.length, DEFAULT_ROWS);
    const maxC = Math.max(DEFAULT_COLS, ...rows.map((r) => r.length), 1);
    sheet.rows = Math.min(Math.max(maxR, DEFAULT_ROWS), 500);
    sheet.cols = Math.min(Math.max(maxC, DEFAULT_COLS), 100);
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < (rows[r] || []).length; c++) {
        const raw = rows[r][c];
        if (raw === "" || raw == null) continue;
        this.ensureCell(r, c).raw = String(raw);
      }
    }
    if (title) this.workbook.title = title;
    this.select(0, 0);
    this.invalidate();
    this.dirty = true;
  }

  /** Load one or more sheets from parsed XLSX `{ name, rows, links? }[]` */
  loadFromXlsxSheets(xlsxSheets, { title } = {}) {
    if (!xlsxSheets?.length) throw new Error("XLSX has no sheets");
    this.pushHistory();
    const sheets = xlsxSheets.map((s, i) => {
      const sheet = createSheet(s.name || `Sheet${i + 1}`);
      const rows = s.rows || [];
      const maxR = Math.max(rows.length, DEFAULT_ROWS);
      const maxC = Math.max(DEFAULT_COLS, ...rows.map((r) => r.length || 0), 1);
      sheet.rows = Math.min(Math.max(maxR, DEFAULT_ROWS), 500);
      sheet.cols = Math.min(Math.max(maxC, DEFAULT_COLS), 100);
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < (rows[r] || []).length; c++) {
          const raw = rows[r][c];
          if (raw === "" || raw == null) continue;
          const key = `${r},${c}`;
          sheet.cells[key] = emptyCell();
          sheet.cells[key].raw = String(raw);
        }
      }
      // Apply Excel hyperlinks as =HYPERLINK formulas
      for (const [key, info] of Object.entries(s.links || {})) {
        const url = info?.url || info;
        if (!url || String(url).startsWith("#")) continue;
        const [r, c] = key.split(",").map(Number);
        const label = info?.display || sheet.cells[key]?.raw || url;
        const cell = sheet.cells[key] || emptyCell();
        const esc = (t) => String(t).replace(/"/g, '""');
        cell.raw = `=HYPERLINK("${esc(url)}","${esc(label)}")`;
        sheet.cells[key] = cell;
      }
      return sheet;
    });
    this.workbook.sheets = sheets;
    this.workbook.activeSheetId = sheets[0].id;
    if (title) this.workbook.title = title;
    this.select(0, 0);
    this.history.clear();
    this.pushHistory();
    this.invalidate();
    this.dirty = true;
  }

  /** Export all sheets as AOA for XLSX writer */
  toXlsxSheets() {
    return this.workbook.sheets.map((sheet) => {
      let maxR = 0;
      let maxC = 0;
      for (const key of Object.keys(sheet.cells)) {
        const [r, c] = key.split(",").map(Number);
        if (sheet.cells[key]?.raw) {
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      }
      const rows = [];
      for (let r = 0; r <= maxR; r++) {
        const row = [];
        for (let c = 0; c <= maxC; c++) {
          row.push(sheet.cells[`${r},${c}`]?.raw ?? "");
        }
        rows.push(row);
      }
      return { name: sheet.name, rows };
    });
  }

  loadWorkbookData(data, { title } = {}) {
    if (!data || !Array.isArray(data.sheets) || !data.sheets.length) {
      throw new Error("Invalid Gridia workbook file");
    }
    this.pushHistory();
    this.workbook = deepClone(data);
    if (title) this.workbook.title = title;
    if (!this.workbook.sheets.find((s) => s.id === this.workbook.activeSheetId)) {
      this.workbook.activeSheetId = this.workbook.sheets[0].id;
    }
    this.select(0, 0);
    this.history.clear();
    this.pushHistory();
    this.invalidate();
    this.dirty = false;
  }

  seedDemo() {
    const set = (r, c, raw, style) => {
      const cell = this.ensureCell(r, c);
      cell.raw = raw;
      if (style) cell.style = { ...defaultStyle(), ...style };
    };
    set(0, 0, "Item", { bold: true, bg: "#e6f4ea" });
    set(0, 1, "Qty", { bold: true, bg: "#e6f4ea", align: "right" });
    set(0, 2, "Price", { bold: true, bg: "#e6f4ea", align: "right" });
    set(0, 3, "Total", { bold: true, bg: "#e6f4ea", align: "right" });
    set(1, 0, "Apples");
    set(1, 1, "12");
    set(1, 2, "1.5");
    set(1, 3, "=B2*C2");
    set(2, 0, "Oranges");
    set(2, 1, "8");
    set(2, 2, "2.25");
    set(2, 3, "=B3*C3");
    set(3, 0, "Bananas");
    set(3, 1, "20");
    set(3, 2, "0.75");
    set(3, 3, "=B4*C4");
    set(5, 0, "Subtotal", { bold: true });
    set(5, 3, "=SUM(D2:D4)", { bold: true, format: "currency" });
    set(6, 0, "Tax (8%)");
    set(6, 3, "=D6*0.08", { format: "currency" });
    set(7, 0, "Grand Total", { bold: true });
    set(7, 3, "=D6+D7", { bold: true, format: "currency", bg: "#fff3cd" });
    this.invalidate();
    this.dirty = true;
    this.pushHistory();
  }
}

function normalized(sel) {
  return {
    r1: Math.min(sel.r1, sel.r2),
    c1: Math.min(sel.c1, sel.c2),
    r2: Math.max(sel.r1, sel.r2),
    c2: Math.max(sel.c1, sel.c2),
  };
}

function hasStyle(style) {
  if (!style) return false;
  const d = defaultStyle();
  return Object.keys(d).some((k) => style[k] !== d[k] && !(k === "bg" && !style.bg));
}

export { normalized, createSheet, DEFAULT_ROWS, DEFAULT_COLS };
