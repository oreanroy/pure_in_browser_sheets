import { SpreadsheetApp, createWorkbook, normalized } from "./spreadsheet.js";
import { colToLetter, cellId, debounce, clamp } from "./utils.js";
import {
  loadWorkbook,
  saveWorkbook,
  listFiles,
  loadNamedWorkbook,
  saveNamedWorkbook,
  saveFileMeta,
  loadFileMeta,
  fileNameFromLocation,
  setFileInUrl,
  defaultFileName,
  guessKind,
  removeFileFromIndex,
  migrateLegacyWorkbook,
} from "./storage.js";
import {
  openLocalFile,
  saveAsLocalFile,
  saveWithHandle,
  parseCSV,
  titleFromFilename,
  safeFilename,
  supportsFileSystemAccess,
} from "./files.js";
import { readXlsx, writeXlsx } from "./xlsx.js";

const ROW_H = 24;
const HEADER_W = 46;
const HEADER_H = 24;

/** @type {{ handle: FileSystemFileHandle | null, name: string | null, kind: 'csv' | 'json' | 'xlsx' | null }} */
let linkedFile = { handle: null, name: null, kind: null };

const els = {
  gridScroll: document.getElementById("grid-scroll"),
  sheetInner: document.getElementById("sheet-inner"),
  grid: document.getElementById("grid"),
  colHeaders: document.getElementById("col-headers"),
  rowHeaders: document.getElementById("row-headers"),
  editor: document.getElementById("cell-editor"),
  formula: document.getElementById("formula-input"),
  nameBox: document.getElementById("name-box"),
  title: document.getElementById("doc-title"),
  saveStatus: document.getElementById("save-status"),
  status: document.getElementById("status-bar"),
  tabs: document.getElementById("sheet-tabs"),
  contextMenu: document.getElementById("context-menu"),
  dropdown: document.getElementById("dropdown-menu"),
  textColor: document.getElementById("text-color"),
  fillColor: document.getElementById("fill-color"),
  textColorBar: document.getElementById("text-color-bar"),
  fillColorBar: document.getElementById("fill-color-bar"),
  fontSize: document.getElementById("font-size"),
  fontFamily: document.getElementById("font-family"),
  fileSwitcher: document.getElementById("file-switcher"),
};

let app;
let editing = false;
let editOrigin = null;
let selecting = false;
let overlayEl = null;
let fillHandleEl = null;
let suppressFormulaSync = false;

function init() {
  migrateLegacyWorkbook();
  const saved = loadWorkbook();
  if (saved?.data) {
    app = new SpreadsheetApp(saved.data);
    linkedFile = {
      handle: null,
      name: saved.name,
      kind: saved.kind || guessKind(saved.name),
    };
    setFileInUrl(saved.name, { replace: true });
  } else {
    app = new SpreadsheetApp(createWorkbook());
    app.seedDemo();
    app.dirty = false;
    const name = defaultFileName(app.workbook.title, "xlsx");
    linkedFile = { handle: null, name, kind: "xlsx" };
    saveNamedWorkbook(name, app.workbook, { kind: "xlsx" });
    saveFileMeta(name, { kind: "xlsx" });
    setFileInUrl(name, { replace: true });
  }

  els.title.value = app.workbook.title || "Untitled spreadsheet";
  document.title = `${linkedFile.name || "Gridia"} — Gridia`;
  buildGrid();
  renderAll();
  renderFileSwitcher();
  bindEvents();
  updateSaveStatus();
  els.gridScroll.focus({ preventScroll: true });
}

function buildGrid() {
  const sheet = app.sheet;
  const widths = [];
  for (let c = 0; c < sheet.cols; c++) widths.push(`${app.getColWidth(c)}px`);

  els.colHeaders.innerHTML = "";
  els.colHeaders.style.gridTemplateColumns = widths.join(" ");
  for (let c = 0; c < sheet.cols; c++) {
    const h = document.createElement("div");
    h.className = "col-header";
    h.dataset.col = c;
    h.textContent = colToLetter(c);
    const resizer = document.createElement("div");
    resizer.className = "col-resizer";
    resizer.dataset.col = c;
    h.appendChild(resizer);
    els.colHeaders.appendChild(h);
  }

  els.rowHeaders.innerHTML = "";
  els.rowHeaders.style.gridTemplateRows = `repeat(${sheet.rows}, ${ROW_H}px)`;
  for (let r = 0; r < sheet.rows; r++) {
    const h = document.createElement("div");
    h.className = "row-header";
    h.dataset.row = r;
    h.textContent = String(r + 1);
    els.rowHeaders.appendChild(h);
  }

  els.grid.innerHTML = "";
  els.grid.style.gridTemplateColumns = widths.join(" ");
  els.grid.style.gridTemplateRows = `repeat(${sheet.rows}, ${ROW_H}px)`;

  const frag = document.createDocumentFragment();
  for (let r = 0; r < sheet.rows; r++) {
    for (let c = 0; c < sheet.cols; c++) {
      const cell = document.createElement("div");
      cell.className = "cell";
      cell.dataset.row = r;
      cell.dataset.col = c;
      frag.appendChild(cell);
    }
  }
  els.grid.appendChild(frag);

  if (!overlayEl) {
    overlayEl = document.createElement("div");
    overlayEl.className = "selection-overlay";
    els.grid.appendChild(overlayEl);
  } else {
    els.grid.appendChild(overlayEl);
  }

  if (!fillHandleEl) {
    fillHandleEl = document.createElement("div");
    fillHandleEl.className = "fill-handle";
    fillHandleEl.title = "Drag to fill";
  }
  els.grid.appendChild(fillHandleEl);
}

function cellEl(row, col) {
  return els.grid.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
}

function renderCells() {
  const sheet = app.sheet;
  const nodes = els.grid.querySelectorAll(".cell");
  let i = 0;
  for (let r = 0; r < sheet.rows; r++) {
    for (let c = 0; c < sheet.cols; c++) {
      const el = nodes[i++];
      const cell = app.getCell(r, c);
      const display = app.getDisplay(r, c);
      const value = app.getValue(r, c);
      const link = app.getHyperlink(r, c);
      el.textContent = display;
      el.classList.toggle("formula-error", typeof value === "string" && value.startsWith("#"));
      el.classList.toggle("cell-link", !!link);
      if (link) {
        el.title = `${link}\n⌘/Ctrl+click to open`;
      } else {
        el.removeAttribute("title");
      }
      applyCellStyle(el, cell.style, !!link);
    }
  }
}

function applyCellStyle(el, style, isLink = false) {
  el.style.fontWeight = style.bold ? "700" : "";
  el.style.fontStyle = style.italic ? "italic" : "";
  el.style.textDecoration = style.underline || isLink ? "underline" : "";
  el.style.textAlign = style.align || "left";
  el.style.color = isLink ? "#1155cc" : style.color || "";
  el.style.background = style.bg || "";
  el.style.fontFamily = quoteFont(style.fontFamily || "Roboto");
  el.style.fontSize = `${style.fontSize || 12}px`;
}

function quoteFont(name) {
  const n = String(name || "Roboto");
  return /\s/.test(n) ? `"${n}", sans-serif` : `${n}, sans-serif`;
}

function renderSelection() {
  const { r1, c1, r2, c2 } = normalized(app.selection);
  const { row: ar, col: ac } = app.active;

  els.grid.querySelectorAll(".cell.selected, .cell.active").forEach((el) => {
    el.classList.remove("selected", "active");
  });
  els.colHeaders.querySelectorAll(".active").forEach((el) => el.classList.remove("active"));
  els.rowHeaders.querySelectorAll(".active").forEach((el) => el.classList.remove("active"));

  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const el = cellEl(r, c);
      if (!el) continue;
      if (r === ar && c === ac) el.classList.add("active");
      else el.classList.add("selected");
    }
  }

  for (let c = c1; c <= c2; c++) {
    els.colHeaders.querySelector(`[data-col="${c}"]`)?.classList.add("active");
  }
  for (let r = r1; r <= r2; r++) {
    els.rowHeaders.querySelector(`[data-row="${r}"]`)?.classList.add("active");
  }

  const left = colOffset(c1);
  const top = r1 * ROW_H;
  const width = colOffset(c2 + 1) - left;
  const height = (r2 - r1 + 1) * ROW_H;
  overlayEl.style.left = `${left}px`;
  overlayEl.style.top = `${top}px`;
  overlayEl.style.width = `${width}px`;
  overlayEl.style.height = `${height}px`;
  overlayEl.style.display = r1 === r2 && c1 === c2 ? "none" : "block";

  fillHandleEl.style.left = `${colOffset(c2) + app.getColWidth(c2) - 5}px`;
  fillHandleEl.style.top = `${(r2 + 1) * ROW_H - 5}px`;

  els.nameBox.textContent = rangeLabel();
  if (!editing && !suppressFormulaSync) {
    els.formula.value = app.getCell(ar, ac).raw;
  }
  syncToolbar();
  updateStatus();
}

function colOffset(col) {
  let x = 0;
  for (let c = 0; c < col; c++) x += app.getColWidth(c);
  return x;
}

function rangeLabel() {
  const { r1, c1, r2, c2 } = normalized(app.selection);
  const a = cellId(r1, c1);
  if (r1 === r2 && c1 === c2) return a;
  return `${a}:${cellId(r2, c2)}`;
}

function renderTabs() {
  els.tabs.innerHTML = "";
  for (const sheet of app.workbook.sheets) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "sheet-tab" + (sheet.id === app.workbook.activeSheetId ? " active" : "");
    btn.dataset.id = sheet.id;

    const name = document.createElement("span");
    name.textContent = sheet.name;
    btn.appendChild(name);

    if (app.workbook.sheets.length > 1) {
      const close = document.createElement("button");
      close.type = "button";
      close.className = "tab-close";
      close.dataset.id = sheet.id;
      close.title = "Delete sheet";
      close.textContent = "×";
      btn.appendChild(close);
    }

    btn.addEventListener("click", (e) => {
      if (e.target.classList.contains("tab-close")) return;
      if (sheet.id !== app.workbook.activeSheetId) {
        commitEdit(true);
        app.switchSheet(sheet.id);
        buildGrid();
        renderAll();
        scheduleSave();
      }
    });

    btn.addEventListener("dblclick", () => {
      const next = prompt("Rename sheet", sheet.name);
      if (next != null) {
        app.renameSheet(sheet.id, next);
        renderTabs();
        scheduleSave();
      }
    });

    els.tabs.appendChild(btn);
  }

  els.tabs.querySelectorAll(".tab-close").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (confirm("Delete this sheet?")) {
        app.deleteSheet(btn.dataset.id);
        buildGrid();
        renderAll();
        scheduleSave();
      }
    });
  });
}

function renderAll() {
  renderCells();
  renderSelection();
  renderTabs();
}

function syncToolbar() {
  const style = app.getCell(app.active.row, app.active.col).style;
  document.getElementById("btn-bold").classList.toggle("active", !!style.bold);
  document.getElementById("btn-italic").classList.toggle("active", !!style.italic);
  document.getElementById("btn-underline").classList.toggle("active", !!style.underline);
  document.getElementById("btn-align-left").classList.toggle("active", style.align === "left");
  document.getElementById("btn-align-center").classList.toggle("active", style.align === "center");
  document.getElementById("btn-align-right").classList.toggle("active", style.align === "right");
  els.fontSize.value = String(style.fontSize || 12);
  if (els.fontFamily) {
    const fam = style.fontFamily || "Roboto";
    if (![...els.fontFamily.options].some((o) => o.value === fam)) {
      const opt = document.createElement("option");
      opt.value = fam;
      opt.textContent = fam;
      els.fontFamily.appendChild(opt);
    }
    els.fontFamily.value = fam;
    els.fontFamily.style.fontFamily = quoteFont(fam);
  }
  els.textColor.value = style.color || "#202124";
  els.textColorBar.style.background = style.color || "#202124";
  els.fillColor.value = style.bg || "#ffffff";
  els.fillColorBar.style.background = style.bg || "#ffffff";
}

function updateStatus() {
  const { r1, c1, r2, c2 } = normalized(app.selection);
  const count = (r2 - r1 + 1) * (c2 - c1 + 1);
  if (count === 1) {
    const link = app.getHyperlink(app.active.row, app.active.col);
    els.status.textContent = link
      ? `${app.activeCellId()} · link · ⌘/Ctrl+click to open`
      : app.activeCellId();
    return;
  }
  let sum = 0;
  let n = 0;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const v = app.getValue(r, c);
      if (typeof v === "number" && Number.isFinite(v)) {
        sum += v;
        n++;
      }
    }
  }
  els.status.textContent = n ? `Sum: ${formatNum(sum)}  ·  Count: ${count}  ·  Numeric: ${n}` : `${count} cells`;
}

function formatNum(n) {
  return Number.isInteger(n) ? String(n) : String(Math.round(n * 1000) / 1000);
}

function updateSaveStatus() {
  if (app.dirty) {
    const target = linkedFile.name ? `Unsaved · ${linkedFile.name}` : "Unsaved changes";
    els.saveStatus.textContent = target;
    els.saveStatus.classList.add("dirty");
    els.saveStatus.title = target;
    return;
  }
  if (linkedFile.name) {
    els.saveStatus.textContent = `Saved · ${linkedFile.name}`;
    els.saveStatus.title = linkedFile.name;
  } else {
    els.saveStatus.textContent = "Saved in browser";
    els.saveStatus.title = "Autosaved to this browser";
  }
  els.saveStatus.classList.remove("dirty");
}

const scheduleSave = debounce(() => {
  app.workbook.title = els.title.value;
  if (linkedFile.name) {
    saveWorkbook(app.workbook, linkedFile.name, linkedFile.kind);
  } else {
    saveWorkbook(app.workbook);
  }
  app.dirty = false;
  updateSaveStatus();
  renderFileSwitcher();
}, 400);

function renderFileSwitcher() {
  if (!els.fileSwitcher) return;
  const files = listFiles();
  const current = linkedFile.name;
  if (current && !files.some((f) => f.name === current)) {
    files.unshift({ name: current, kind: linkedFile.kind });
  }
  els.fileSwitcher.innerHTML = files
    .map(
      (f) =>
        `<option value="${escapeAttr(f.name)}" ${f.name === current ? "selected" : ""}>${escapeHtml(f.name)}</option>`
    )
    .join("");
  if (!files.length) {
    els.fileSwitcher.innerHTML = `<option value="">Untitled</option>`;
  }
}

function escapeAttr(s) {
  return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function switchToFile(name, { fromUrl = false } = {}) {
  if (!name || name === linkedFile.name) {
    if (!fromUrl && name) setFileInUrl(name, { replace: true });
    return;
  }
  commitEdit(true);
  if (app.dirty && linkedFile.name) {
    saveWorkbook(app.workbook, linkedFile.name, linkedFile.kind);
    app.dirty = false;
  }

  const data = loadNamedWorkbook(name);
  if (!data) {
    els.status.textContent = `File not in browser library: ${name}`;
    if (linkedFile.name) setFileInUrl(linkedFile.name, { replace: true });
    renderFileSwitcher();
    return;
  }

  const meta = loadFileMeta(name);
  app = new SpreadsheetApp(data);
  linkedFile = {
    handle: null,
    name,
    kind: meta?.kind || guessKind(name),
  };
  els.title.value = app.workbook.title || titleFromFilename(name);
  document.title = `${name} — Gridia`;
  if (!fromUrl) setFileInUrl(name);
  buildGrid();
  renderAll();
  renderFileSwitcher();
  updateSaveStatus();
  els.status.textContent = `Switched to ${name}`;
}

function hitCell(clientX, clientY) {
  const gridRect = els.grid.getBoundingClientRect();
  const x = clientX - gridRect.left;
  const y = clientY - gridRect.top;
  if (x < 0 || y < 0) return null;
  let col = -1;
  let acc = 0;
  for (let c = 0; c < app.sheet.cols; c++) {
    acc += app.getColWidth(c);
    if (x < acc) {
      col = c;
      break;
    }
  }
  const row = Math.floor(y / ROW_H);
  if (row < 0 || row >= app.sheet.rows || col < 0) return null;
  return { row, col };
}

function startEdit(initial = null, fromFormula = false) {
  if (editing) return;
  editing = true;
  editOrigin = { ...app.active };
  const cell = app.getCell(app.active.row, app.active.col);
  const el = cellEl(app.active.row, app.active.col);
  if (!el) return;

  const left = colOffset(app.active.col) + HEADER_W;
  const top = app.active.row * ROW_H + HEADER_H;
  els.editor.style.left = `${left}px`;
  els.editor.style.top = `${top}px`;
  els.editor.style.width = `${Math.max(app.getColWidth(app.active.col), 80)}px`;
  els.editor.style.height = `${ROW_H}px`;
  els.editor.style.fontFamily = quoteFont(cell.style?.fontFamily || "Roboto");
  els.editor.style.fontSize = `${cell.style?.fontSize || 12}px`;
  els.editor.style.fontWeight = cell.style?.bold ? "700" : "";
  els.editor.style.fontStyle = cell.style?.italic ? "italic" : "";
  els.editor.style.color = cell.style?.color || "";
  els.editor.style.textAlign = cell.style?.align || "left";
  els.editor.value = initial != null ? initial : cell.raw;
  els.editor.classList.add("visible");
  if (!fromFormula) {
    els.editor.focus();
    if (initial != null) {
      els.editor.setSelectionRange(els.editor.value.length, els.editor.value.length);
    } else {
      els.editor.select();
    }
  }
  suppressFormulaSync = true;
  els.formula.value = els.editor.value;
  el.textContent = "";
}

function commitEdit(save) {
  if (!editing) return;
  const raw = els.editor.value;
  const { row, col } = editOrigin;
  editing = false;
  editOrigin = null;
  els.editor.classList.remove("visible");
  suppressFormulaSync = false;
  if (save) {
    app.setRaw(row, col, raw);
    scheduleSave();
  }
  renderCells();
  renderSelection();
  updateSaveStatus();
  els.gridScroll.focus({ preventScroll: true });
}

function cancelEdit() {
  commitEdit(false);
  renderCells();
  renderSelection();
}

function bindEvents() {
  els.gridScroll.tabIndex = 0;

  els.grid.addEventListener("mousedown", (e) => {
    if (e.target.classList.contains("fill-handle")) return;
    const hit = hitCell(e.clientX, e.clientY);
    if (!hit) return;
    // Ctrl/Cmd+click opens hyperlink
    if (e.metaKey || e.ctrlKey) {
      const url = app.getHyperlink(hit.row, hit.col);
      if (url) {
        e.preventDefault();
        e.stopPropagation();
        window.open(url, "_blank", "noopener,noreferrer");
        els.status.textContent = `Opened ${url}`;
        return;
      }
    }
    e.preventDefault();
    hideMenus();
    if (editing) commitEdit(true);
    selecting = true;
    app.select(hit.row, hit.col, e.shiftKey);
    renderSelection();
    els.gridScroll.focus({ preventScroll: true });
  });

  window.addEventListener("mousemove", (e) => {
    if (!selecting) return;
    const hit = hitCell(e.clientX, e.clientY);
    if (!hit) return;
    app.select(hit.row, hit.col, true);
    renderSelection();
  });

  window.addEventListener("mouseup", () => {
    selecting = false;
  });

  els.grid.addEventListener("dblclick", (e) => {
    const hit = hitCell(e.clientX, e.clientY);
    if (!hit) return;
    app.select(hit.row, hit.col);
    renderSelection();
    startEdit();
  });

  fillHandleEl.addEventListener("mousedown", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const start = normalized(app.selection);
    const onMove = (ev) => {
      const hit = hitCell(ev.clientX, ev.clientY);
      if (!hit) return;
      app.selection = {
        r1: start.r1,
        c1: start.c1,
        r2: Math.max(start.r2, hit.row),
        c2: Math.max(start.c2, hit.col),
      };
      renderSelection();
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      app.fillDown();
      renderCells();
      renderSelection();
      scheduleSave();
      updateSaveStatus();
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  });

  // Column resize
  els.colHeaders.addEventListener("mousedown", (e) => {
    if (!e.target.classList.contains("col-resizer")) return;
    e.preventDefault();
    e.stopPropagation();
    const col = Number(e.target.dataset.col);
    const startX = e.clientX;
    const startW = app.getColWidth(col);
    const onMove = (ev) => {
      app.setColWidth(col, startW + (ev.clientX - startX));
      rebuildColumnsOnly();
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      scheduleSave();
      updateSaveStatus();
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  });

  els.colHeaders.addEventListener("click", (e) => {
    const header = e.target.closest(".col-header");
    if (!header || e.target.classList.contains("col-resizer")) return;
    const col = Number(header.dataset.col);
    app.selectRange(0, col, app.sheet.rows - 1, col);
    renderSelection();
  });

  els.rowHeaders.addEventListener("click", (e) => {
    const header = e.target.closest(".row-header");
    if (!header) return;
    const row = Number(header.dataset.row);
    app.selectRange(row, 0, row, app.sheet.cols - 1);
    renderSelection();
  });

  els.editor.addEventListener("input", () => {
    els.formula.value = els.editor.value;
  });

  els.editor.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      commitEdit(true);
      app.select(app.active.row + (e.shiftKey ? -1 : 1), app.active.col);
      renderSelection();
      scrollActiveIntoView();
    } else if (e.key === "Tab") {
      e.preventDefault();
      commitEdit(true);
      app.select(app.active.row, app.active.col + (e.shiftKey ? -1 : 1));
      renderSelection();
      scrollActiveIntoView();
    } else if (e.key === "Escape") {
      e.preventDefault();
      cancelEdit();
    }
  });

  els.formula.addEventListener("focus", () => {
    if (!editing) startEdit(null, true);
  });

  els.formula.addEventListener("input", () => {
    if (editing) els.editor.value = els.formula.value;
  });

  els.formula.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      if (editing) els.editor.value = els.formula.value;
      commitEdit(true);
    } else if (e.key === "Escape") {
      e.preventDefault();
      cancelEdit();
    }
  });

  els.gridScroll.addEventListener("keydown", onGridKeyDown);

  els.title.addEventListener("input", () => {
    app.workbook.title = els.title.value;
    app.dirty = true;
    updateSaveStatus();
    scheduleSave();
  });
  els.title.addEventListener("change", () => {
    // Keep URL file name stable; title is document name inside the file.
    document.title = `${linkedFile.name || els.title.value} — Gridia`;
  });

  document.getElementById("btn-undo").addEventListener("click", () => {
    commitEdit(true);
    if (app.undo()) {
      buildGrid();
      renderAll();
      scheduleSave();
      updateSaveStatus();
    }
  });
  document.getElementById("btn-redo").addEventListener("click", () => {
    commitEdit(true);
    if (app.redo()) {
      buildGrid();
      renderAll();
      scheduleSave();
      updateSaveStatus();
    }
  });

  document.getElementById("btn-bold").addEventListener("click", () => toggleStyle("bold"));
  document.getElementById("btn-italic").addEventListener("click", () => toggleStyle("italic"));
  document.getElementById("btn-underline").addEventListener("click", () => toggleStyle("underline"));
  document.getElementById("btn-link").addEventListener("click", () => insertOrEditLink());
  document.getElementById("btn-align-left").addEventListener("click", () => setStyle({ align: "left" }));
  document.getElementById("btn-align-center").addEventListener("click", () => setStyle({ align: "center" }));
  document.getElementById("btn-align-right").addEventListener("click", () => setStyle({ align: "right" }));

  els.textColor.addEventListener("input", () => {
    setStyle({ color: els.textColor.value });
    els.textColorBar.style.background = els.textColor.value;
  });
  els.fillColor.addEventListener("input", () => {
    setStyle({ bg: els.fillColor.value });
    els.fillColorBar.style.background = els.fillColor.value;
  });
  els.fontSize.addEventListener("change", () => {
    setStyle({ fontSize: Number(els.fontSize.value) });
  });
  els.fontFamily.addEventListener("change", () => {
    setStyle({ fontFamily: els.fontFamily.value });
    els.fontFamily.style.fontFamily = quoteFont(els.fontFamily.value);
  });

  document.getElementById("btn-pct").addEventListener("click", () => setStyle({ format: "percent" }));
  document.getElementById("btn-currency").addEventListener("click", () => setStyle({ format: "currency" }));
  document.getElementById("btn-decimal-dec").addEventListener("click", () => {
    const d = Math.max(0, (app.getCell(app.active.row, app.active.col).style.decimals ?? 2) - 1);
    setStyle({ format: "number", decimals: d });
  });
  document.getElementById("btn-decimal-inc").addEventListener("click", () => {
    const d = Math.min(8, (app.getCell(app.active.row, app.active.col).style.decimals ?? 2) + 1);
    setStyle({ format: "number", decimals: d });
  });

  document.getElementById("btn-add-sheet").addEventListener("click", () => {
    commitEdit(true);
    app.addSheet();
    buildGrid();
    renderAll();
    scheduleSave();
  });

  document.getElementById("btn-open").addEventListener("click", () => openFile());
  document.getElementById("btn-save").addEventListener("click", () => saveFile(false));
  document.getElementById("btn-save-as").addEventListener("click", () => saveFile(true));

  els.fileSwitcher?.addEventListener("change", () => {
    const name = els.fileSwitcher.value;
    if (name) switchToFile(name);
  });

  window.addEventListener("hashchange", () => {
    const name = fileNameFromLocation();
    if (name) switchToFile(name, { fromUrl: true });
  });
  window.addEventListener("popstate", () => {
    const name = fileNameFromLocation();
    if (name) switchToFile(name, { fromUrl: true });
  });

  document.getElementById("btn-clear").addEventListener("click", () => {
    if (confirm("Clear all cells on this sheet?")) {
      app.clearActiveSheet();
      renderAll();
      scheduleSave();
      updateSaveStatus();
    }
  });

  // Context menu
  els.gridScroll.addEventListener("contextmenu", (e) => {
    const hit = hitCell(e.clientX, e.clientY);
    if (!hit) return;
    e.preventDefault();
    if (editing) commitEdit(true);
    const sel = normalized(app.selection);
    const inside = hit.row >= sel.r1 && hit.row <= sel.r2 && hit.col >= sel.c1 && hit.col <= sel.c2;
    if (!inside) {
      app.select(hit.row, hit.col);
      renderSelection();
    }
    showContextMenu(e.clientX, e.clientY);
  });

  els.contextMenu.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;
    runContextAction(btn.dataset.action);
    hideMenus();
  });

  document.querySelectorAll(".menu-item").forEach((item) => {
    item.addEventListener("click", (e) => {
      e.stopPropagation();
      openDropdown(item);
    });
  });

  window.addEventListener("click", () => hideMenus());
  window.addEventListener("blur", () => hideMenus());

  window.addEventListener("keydown", (e) => {
    const meta = e.metaKey || e.ctrlKey;
    if (!meta) return;
    const key = e.key.toLowerCase();
    if (key === "s") {
      e.preventDefault();
      saveFile(e.shiftKey);
    } else if (key === "o") {
      e.preventDefault();
      openFile();
    } else if (key === "k") {
      e.preventDefault();
      insertOrEditLink();
    }
  });

  document.addEventListener("copy", (e) => {
    if (editing || isTypingTarget(e.target)) return;
    const tsv = app.copySelection(false);
    e.clipboardData.setData("text/plain", tsv);
    e.preventDefault();
    els.status.textContent = "Copied";
  });

  document.addEventListener("cut", (e) => {
    if (editing || isTypingTarget(e.target)) return;
    const tsv = app.copySelection(true);
    e.clipboardData.setData("text/plain", tsv);
    e.preventDefault();
    app.clearSelectionContents();
    renderCells();
    renderSelection();
    scheduleSave();
    updateSaveStatus();
  });

  document.addEventListener("paste", (e) => {
    if (editing || isTypingTarget(e.target)) return;
    const text = e.clipboardData.getData("text/plain");
    e.preventDefault();
    app.paste(text);
    renderCells();
    renderSelection();
    scheduleSave();
    updateSaveStatus();
  });
}

function rebuildColumnsOnly() {
  const sheet = app.sheet;
  const widths = [];
  for (let c = 0; c < sheet.cols; c++) widths.push(`${app.getColWidth(c)}px`);
  els.colHeaders.style.gridTemplateColumns = widths.join(" ");
  els.grid.style.gridTemplateColumns = widths.join(" ");
  renderSelection();
}

function toggleStyle(key) {
  const cur = app.getCell(app.active.row, app.active.col).style[key];
  app.applyStyleToSelection({ [key]: !cur });
  renderCells();
  renderSelection();
  scheduleSave();
  updateSaveStatus();
}

function setStyle(patch) {
  app.applyStyleToSelection(patch);
  renderCells();
  renderSelection();
  scheduleSave();
  updateSaveStatus();
}

function isTypingTarget(el) {
  return el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable);
}

function onGridKeyDown(e) {
  if (editing) return;
  if (isTypingTarget(e.target) && e.target !== els.gridScroll) return;

  const meta = e.metaKey || e.ctrlKey;
  const { row, col } = app.active;

  if (meta && (e.key.toLowerCase() === "s" || e.key.toLowerCase() === "o")) {
    // handled globally
    return;
  }

  if (meta && e.key.toLowerCase() === "k") {
    e.preventDefault();
    insertOrEditLink();
    return;
  }

  if (meta && e.key.toLowerCase() === "z") {
    e.preventDefault();
    if (e.shiftKey) {
      if (app.redo()) {
        buildGrid();
        renderAll();
        scheduleSave();
      }
    } else if (app.undo()) {
      buildGrid();
      renderAll();
      scheduleSave();
    }
    updateSaveStatus();
    return;
  }
  if (meta && e.key.toLowerCase() === "y") {
    e.preventDefault();
    if (app.redo()) {
      buildGrid();
      renderAll();
      scheduleSave();
      updateSaveStatus();
    }
    return;
  }
  if (meta && e.key.toLowerCase() === "b") {
    e.preventDefault();
    toggleStyle("bold");
    return;
  }
  if (meta && e.key.toLowerCase() === "i") {
    e.preventDefault();
    toggleStyle("italic");
    return;
  }
  if (meta && e.key.toLowerCase() === "u") {
    e.preventDefault();
    toggleStyle("underline");
    return;
  }

  if (e.key === "Delete" || e.key === "Backspace") {
    e.preventDefault();
    app.clearSelectionContents();
    renderCells();
    renderSelection();
    scheduleSave();
    updateSaveStatus();
    return;
  }

  if (e.key === "F2") {
    e.preventDefault();
    startEdit();
    return;
  }

  if (e.key === "Enter") {
    e.preventDefault();
    app.select(row + (e.shiftKey ? -1 : 1), col);
    renderSelection();
    scrollActiveIntoView();
    return;
  }

  if (e.key === "Tab") {
    e.preventDefault();
    app.select(row, col + (e.shiftKey ? -1 : 1));
    renderSelection();
    scrollActiveIntoView();
    return;
  }

  const move = {
    ArrowUp: [-1, 0],
    ArrowDown: [1, 0],
    ArrowLeft: [0, -1],
    ArrowRight: [0, 1],
  };
  if (move[e.key]) {
    e.preventDefault();
    const [dr, dc] = move[e.key];
    app.select(row + dr, col + dc, e.shiftKey);
    renderSelection();
    scrollActiveIntoView();
    return;
  }

  if (e.key === "Home") {
    e.preventDefault();
    app.select(row, 0, e.shiftKey);
    renderSelection();
    scrollActiveIntoView();
    return;
  }
  if (e.key === "End") {
    e.preventDefault();
    app.select(row, app.sheet.cols - 1, e.shiftKey);
    renderSelection();
    scrollActiveIntoView();
    return;
  }

  if (e.key.length === 1 && !meta && !e.altKey) {
    e.preventDefault();
    startEdit(e.key);
  }
}

function scrollActiveIntoView() {
  const el = cellEl(app.active.row, app.active.col);
  el?.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function showContextMenu(x, y) {
  const menu = els.contextMenu;
  menu.classList.remove("hidden");
  const pad = 8;
  const w = menu.offsetWidth;
  const h = menu.offsetHeight;
  menu.style.left = `${clamp(x, pad, window.innerWidth - w - pad)}px`;
  menu.style.top = `${clamp(y, pad, window.innerHeight - h - pad)}px`;
}

function hideMenus() {
  els.contextMenu.classList.add("hidden");
  els.dropdown.classList.add("hidden");
  document.querySelectorAll(".menu-item.open").forEach((el) => el.classList.remove("open"));
}

function runContextAction(action) {
  const { row, col } = app.active;
  switch (action) {
    case "cut": {
      const tsv = app.copySelection(true);
      navigator.clipboard?.writeText(tsv);
      app.clearSelectionContents();
      renderCells();
      renderSelection();
      scheduleSave();
      break;
    }
    case "copy": {
      const tsv = app.copySelection(false);
      navigator.clipboard?.writeText(tsv);
      els.status.textContent = "Copied";
      break;
    }
    case "paste":
      navigator.clipboard.readText().then((text) => {
        app.paste(text);
        renderCells();
        renderSelection();
        scheduleSave();
        updateSaveStatus();
      }).catch(() => {
        els.status.textContent = "Clipboard permission needed";
      });
      break;
    case "insert-row-above":
      app.insertRow(row, false);
      buildGrid();
      renderAll();
      scheduleSave();
      break;
    case "insert-row-below":
      app.insertRow(row, true);
      buildGrid();
      renderAll();
      scheduleSave();
      break;
    case "delete-row":
      app.deleteRow(row);
      buildGrid();
      renderAll();
      scheduleSave();
      break;
    case "insert-col-left":
      app.insertCol(col, false);
      buildGrid();
      renderAll();
      scheduleSave();
      break;
    case "insert-col-right":
      app.insertCol(col, true);
      buildGrid();
      renderAll();
      scheduleSave();
      break;
    case "delete-col":
      app.deleteCol(col);
      buildGrid();
      renderAll();
      scheduleSave();
      break;
    case "clear-contents":
      app.clearSelectionContents();
      renderCells();
      renderSelection();
      scheduleSave();
      break;
    case "open-link": {
      const url = app.getHyperlink(row, col);
      if (url) window.open(url, "_blank", "noopener,noreferrer");
      else els.status.textContent = "No link in cell";
      break;
    }
    case "insert-link":
      insertOrEditLink();
      break;
    case "remove-link":
      app.removeHyperlink(row, col);
      renderCells();
      renderSelection();
      scheduleSave();
      updateSaveStatus();
      break;
  }
  updateSaveStatus();
}

function insertOrEditLink() {
  commitEdit(true);
  const { row, col } = app.active;
  const existing = app.getHyperlink(row, col) || "";
  const currentLabel = app.getDisplay(row, col) || "";
  const url = prompt("Link URL", existing || "https://");
  if (url == null) return;
  const trimmed = url.trim();
  if (!trimmed) {
    app.removeHyperlink(row, col);
    renderCells();
    renderSelection();
    scheduleSave();
    updateSaveStatus();
    return;
  }
  const label = prompt("Link text (label)", currentLabel || trimmed);
  if (label == null) return;
  app.setHyperlink(row, col, trimmed, label);
  renderCells();
  renderSelection();
  scheduleSave();
  updateSaveStatus();
}

function openDropdown(item) {
  const menu = item.dataset.menu;
  hideMenus();
  item.classList.add("open");
  const rect = item.getBoundingClientRect();
  const items = menuItems(menu);
  els.dropdown.innerHTML = items
    .map((it) =>
      it.sep
        ? `<div class="menu-sep"></div>`
        : `<button type="button" data-cmd="${it.cmd}"><span>${it.label}</span>${
            it.hint ? `<span class="hint">${it.hint}</span>` : ""
          }</button>`
    )
    .join("");
  els.dropdown.classList.remove("hidden");
  els.dropdown.style.left = `${rect.left}px`;
  els.dropdown.style.top = `${rect.bottom + 4}px`;
  els.dropdown.onclick = (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-cmd]");
    if (!btn) return;
    runMenuCmd(btn.dataset.cmd);
    hideMenus();
  };
}

function menuItems(menu) {
  if (menu === "file") {
    return [
      { label: "Open…", cmd: "open", hint: "⌘O" },
      { label: "Save", cmd: "save", hint: "⌘S" },
      { label: "Save As…", cmd: "save-as", hint: "⇧⌘S" },
      { sep: true },
      { label: "Remove from browser library", cmd: "remove-library" },
      { sep: true },
      { label: "Export Excel (.xlsx)", cmd: "export-xlsx" },
      { label: "Export CSV download", cmd: "export-csv" },
      { label: "Export Gridia JSON", cmd: "export-json" },
      { sep: true },
      { label: "Clear sheet", cmd: "clear" },
      { label: "New blank sheet tab", cmd: "add-sheet" },
    ];
  }
  if (menu === "edit") {
    return [
      { label: "Undo", cmd: "undo", hint: "⌘Z" },
      { label: "Redo", cmd: "redo", hint: "⌘Y" },
      { sep: true },
      { label: "Cut", cmd: "cut", hint: "⌘X" },
      { label: "Copy", cmd: "copy", hint: "⌘C" },
      { label: "Paste", cmd: "paste", hint: "⌘V" },
      { sep: true },
      { label: "Delete contents", cmd: "clear-contents", hint: "Del" },
    ];
  }
  if (menu === "view") {
    return [{ label: "Scroll to A1", cmd: "goto-a1" }];
  }
  if (menu === "insert") {
    return [
      { label: "Insert link…", cmd: "insert-link", hint: "⌘K" },
      { label: "Insert row above", cmd: "insert-row-above" },
      { label: "Insert row below", cmd: "insert-row-below" },
      { label: "Insert column left", cmd: "insert-col-left" },
      { label: "Insert column right", cmd: "insert-col-right" },
      { sep: true },
      { label: "New sheet", cmd: "add-sheet" },
    ];
  }
  if (menu === "format") {
    return [
      { label: "Bold", cmd: "bold", hint: "⌘B" },
      { label: "Italic", cmd: "italic", hint: "⌘I" },
      { label: "Underline", cmd: "underline", hint: "⌘U" },
      { sep: true },
      { label: "Number", cmd: "fmt-number" },
      { label: "Currency", cmd: "fmt-currency" },
      { label: "Percent", cmd: "fmt-percent" },
    ];
  }
  return [];
}

function runMenuCmd(cmd) {
  switch (cmd) {
    case "open":
      openFile();
      break;
    case "save":
      saveFile(false);
      break;
    case "save-as":
      saveFile(true);
      break;
    case "remove-library":
      if (linkedFile.name && confirm(`Remove ${linkedFile.name} from browser library?`)) {
        const removed = linkedFile.name;
        removeFileFromIndex(removed);
        const rest = listFiles();
        if (rest[0]) switchToFile(rest[0].name);
        else {
          app = new SpreadsheetApp(createWorkbook());
          const name = defaultFileName("Untitled spreadsheet", "xlsx");
          linkedFile = { handle: null, name, kind: "xlsx" };
          saveNamedWorkbook(name, app.workbook, { kind: "xlsx" });
          setFileInUrl(name);
          els.title.value = app.workbook.title;
          buildGrid();
          renderAll();
          renderFileSwitcher();
          updateSaveStatus();
        }
        els.status.textContent = `Removed ${removed}`;
      }
      break;
    case "export-xlsx":
      downloadExport("xlsx");
      break;
    case "export-csv":
      downloadExport("csv");
      break;
    case "export-json":
      downloadExport("json");
      break;
    case "clear":
      document.getElementById("btn-clear").click();
      break;
    case "add-sheet":
      document.getElementById("btn-add-sheet").click();
      break;
    case "undo":
      document.getElementById("btn-undo").click();
      break;
    case "redo":
      document.getElementById("btn-redo").click();
      break;
    case "cut":
      runContextAction("cut");
      break;
    case "copy":
      runContextAction("copy");
      break;
    case "paste":
      runContextAction("paste");
      break;
    case "clear-contents":
      runContextAction("clear-contents");
      break;
    case "goto-a1":
      app.select(0, 0);
      renderSelection();
      scrollActiveIntoView();
      break;
    case "insert-link":
      insertOrEditLink();
      break;
    case "insert-row-above":
      runContextAction("insert-row-above");
      break;
    case "insert-row-below":
      runContextAction("insert-row-below");
      break;
    case "insert-col-left":
      runContextAction("insert-col-left");
      break;
    case "insert-col-right":
      runContextAction("insert-col-right");
      break;
    case "bold":
      toggleStyle("bold");
      break;
    case "italic":
      toggleStyle("italic");
      break;
    case "underline":
      toggleStyle("underline");
      break;
    case "fmt-number":
      setStyle({ format: "number" });
      break;
    case "fmt-currency":
      setStyle({ format: "currency" });
      break;
    case "fmt-percent":
      setStyle({ format: "percent" });
      break;
  }
}

async function openFile() {
  try {
    commitEdit(true);
    if (app.dirty && linkedFile.name) {
      saveWorkbook(app.workbook, linkedFile.name, linkedFile.kind);
    }
    const opened = await openLocalFile();
    const title = titleFromFilename(opened.name);
    if (opened.kind === "json") {
      const data = JSON.parse(opened.text);
      app.loadWorkbookData(data, { title });
    } else if (opened.kind === "xls") {
      throw new Error("Old .xls format is not supported. Re-save as .xlsx in Excel/Numbers, then Open again.");
    } else if (opened.kind === "xlsx") {
      const parsed = await readXlsx(opened.buffer);
      app.loadFromXlsxSheets(parsed.sheets, { title });
    } else {
      const rows = parseCSV(opened.text);
      app.loadFromCSV(rows, { title });
    }
    linkedFile = { handle: opened.handle, name: opened.name, kind: opened.kind };
    els.title.value = app.workbook.title;
    document.title = `${opened.name} — Gridia`;
    saveNamedWorkbook(opened.name, app.workbook, { kind: opened.kind });
    saveFileMeta(opened.name, { kind: opened.kind });
    setFileInUrl(opened.name);
    app.dirty = false;
    buildGrid();
    renderAll();
    renderFileSwitcher();
    updateSaveStatus();
    els.status.textContent = `Opened ${opened.name}`;
  } catch (err) {
    if (String(err?.message || err).includes("cancel") || err?.name === "AbortError") return;
    console.error(err);
    alert(`Could not open file: ${err.message || err}`);
  }
}

async function filePayload(kind) {
  if (kind === "json") {
    return {
      contents: app.toJSON(),
      mime: "application/json",
      suggestedName: linkedFile.name?.match(/\.json$/i)
        ? linkedFile.name
        : safeFilename(app.workbook.title, "gridia.json"),
      kind: "json",
    };
  }
  if (kind === "xlsx") {
    const bytes = await writeXlsx(app.toXlsxSheets());
    return {
      contents: bytes,
      mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      suggestedName: linkedFile.name?.match(/\.xlsx$/i)
        ? linkedFile.name
        : safeFilename(app.workbook.title, "xlsx"),
      kind: "xlsx",
    };
  }
  return {
    contents: app.toCSV(),
    mime: "text/csv;charset=utf-8",
    suggestedName: linkedFile.name?.match(/\.csv$/i)
      ? linkedFile.name
      : safeFilename(app.workbook.title, "csv"),
    kind: "csv",
  };
}

async function saveFile(forceSaveAs) {
  try {
    commitEdit(true);
    app.workbook.title = els.title.value;
    const kind = linkedFile.kind || "xlsx";
    let payload = await filePayload(kind);

    if (!forceSaveAs && linkedFile.handle && supportsFileSystemAccess()) {
      await saveWithHandle(linkedFile.handle, payload.contents, payload.mime);
      app.dirty = false;
      saveWorkbook(app.workbook, linkedFile.name, linkedFile.kind);
      setFileInUrl(linkedFile.name, { replace: true });
      document.title = `${linkedFile.name} — Gridia`;
      renderFileSwitcher();
      updateSaveStatus();
      els.status.textContent = `Saved ${linkedFile.name}`;
      return;
    }

    if (supportsFileSystemAccess() && (forceSaveAs || !linkedFile.handle)) {
      const handle = await window.showSaveFilePicker({
        suggestedName: payload.suggestedName,
        types: [
          {
            description: "Excel workbook",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            },
          },
          { description: "CSV spreadsheet", accept: { "text/csv": [".csv"] } },
          {
            description: "Gridia workbook",
            accept: { "application/json": [".gridia.json", ".json"] },
          },
        ],
      });
      const name = handle.name || "";
      const chosenKind = /\.xlsx$/i.test(name) ? "xlsx" : /\.json$/i.test(name) ? "json" : "csv";
      payload = await filePayload(chosenKind);
      await saveWithHandle(handle, payload.contents, payload.mime);
      linkedFile = { handle, name: handle.name, kind: chosenKind };
      app.dirty = false;
      saveWorkbook(app.workbook, linkedFile.name, linkedFile.kind);
      setFileInUrl(linkedFile.name);
      document.title = `${linkedFile.name} — Gridia`;
      renderFileSwitcher();
      updateSaveStatus();
      els.status.textContent = `Saved ${handle.name}`;
      return;
    }

    let chosenKind = kind;
    if (forceSaveAs || !linkedFile.name) {
      const pick = prompt("Save as format: type xlsx, csv, or json", chosenKind);
      if (pick == null) return;
      const p = String(pick).toLowerCase();
      chosenKind = p.includes("json") ? "json" : p.includes("csv") ? "csv" : "xlsx";
    }
    payload = await filePayload(chosenKind);
    const result = await saveAsLocalFile(payload);
    linkedFile = {
      handle: result.handle,
      name: result.name,
      kind: chosenKind,
    };
    app.dirty = false;
    saveWorkbook(app.workbook, linkedFile.name, linkedFile.kind);
    setFileInUrl(linkedFile.name);
    document.title = `${linkedFile.name} — Gridia`;
    renderFileSwitcher();
    updateSaveStatus();
    els.status.textContent = `Downloaded ${result.name}`;
  } catch (err) {
    if (String(err?.message || err).includes("cancel") || err?.name === "AbortError") return;
    console.error(err);
    alert(`Could not save file: ${err.message || err}`);
  }
}

async function downloadExport(kind) {
  commitEdit(true);
  const payload = await filePayload(kind);
  const blob =
    payload.contents instanceof Blob
      ? payload.contents
      : new Blob([payload.contents], { type: payload.mime });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = payload.suggestedName;
  a.click();
  URL.revokeObjectURL(a.href);
  els.status.textContent = `Downloaded ${payload.suggestedName}`;
}

init();
