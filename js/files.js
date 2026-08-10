/** Local file open/save helpers (File System Access API + download fallback). */

const JSON_TYPE = {
  description: "Gridia workbook",
  accept: { "application/json": [".gridia.json", ".json"] },
};

const CSV_TYPE = {
  description: "CSV spreadsheet",
  accept: { "text/csv": [".csv"], "text/plain": [".csv", ".txt"] },
};

const XLSX_TYPE = {
  description: "Excel workbook",
  accept: {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
    "application/vnd.ms-excel": [".xls"],
  },
};

export function supportsFileSystemAccess() {
  return typeof window.showOpenFilePicker === "function" && typeof window.showSaveFilePicker === "function";
}

export function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let i = 0;
  let inQuotes = false;
  const s = String(text).replace(/^\uFEFF/, "");

  while (i < s.length) {
    const ch = s[i];
    if (inQuotes) {
      if (ch === '"') {
        if (s[i + 1] === '"') {
          cell += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      cell += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === ",") {
      row.push(cell);
      cell = "";
      i++;
      continue;
    }
    if (ch === "\n" || (ch === "\r" && s[i + 1] === "\n")) {
      row.push(cell);
      cell = "";
      rows.push(row);
      row = [];
      i += ch === "\r" ? 2 : 1;
      continue;
    }
    if (ch === "\r") {
      row.push(cell);
      cell = "";
      rows.push(row);
      row = [];
      i++;
      continue;
    }
    cell += ch;
    i++;
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === "") {
    rows.pop();
  }
  return rows;
}

export function titleFromFilename(name) {
  return (
    String(name || "Untitled")
      .replace(/\.(gridia\.json|json|csv|txt|xlsx|xls)$/i, "")
      .replace(/_/g, " ")
      .trim() || "Untitled spreadsheet"
  );
}

export function detectKind(filename, text) {
  const lower = (filename || "").toLowerCase();
  if (lower.endsWith(".xlsx")) return "xlsx";
  if (lower.endsWith(".xls")) return "xls";
  if (lower.endsWith(".csv") || lower.endsWith(".txt")) return "csv";
  if (lower.endsWith(".json") || lower.endsWith(".gridia.json")) return "json";
  const t = (text || "").trim();
  if (t.startsWith("{") && /"sheets"\s*:/.test(t)) return "json";
  return "csv";
}

export function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

export async function openLocalFile() {
  if (supportsFileSystemAccess()) {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [XLSX_TYPE, CSV_TYPE, JSON_TYPE],
      excludeAcceptAllOption: false,
    });
    const file = await handle.getFile();
    const kind = detectKind(file.name);
    if (kind === "xlsx") {
      const buffer = await file.arrayBuffer();
      return { handle, name: file.name, buffer, kind };
    }
    if (kind === "xls") {
      return { handle, name: file.name, buffer: null, kind: "xls" };
    }
    const text = await file.text();
    return { handle, name: file.name, text, kind: detectKind(file.name, text) };
  }

  return new Promise((resolve, reject) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept =
      ".xlsx,.xls,.csv,.txt,.json,.gridia.json,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv,application/json";
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) {
        reject(new Error("cancelled"));
        return;
      }
      const kind = detectKind(file.name);
      if (kind === "xlsx") {
        resolve({ handle: null, name: file.name, buffer: await file.arrayBuffer(), kind });
        return;
      }
      if (kind === "xls") {
        resolve({ handle: null, name: file.name, buffer: null, kind: "xls" });
        return;
      }
      const text = await file.text();
      resolve({ handle: null, name: file.name, text, kind: detectKind(file.name, text) });
    };
    input.oncancel = () => reject(new Error("cancelled"));
    input.click();
  });
}

export async function saveWithHandle(handle, contents, mime) {
  const writable = await handle.createWritable();
  await writable.write(contents instanceof Blob ? contents : new Blob([contents], { type: mime }));
  await writable.close();
}

export async function saveAsLocalFile({ contents, mime, suggestedName, kind }) {
  if (supportsFileSystemAccess()) {
    const type =
      kind === "json" ? JSON_TYPE : kind === "xlsx" ? XLSX_TYPE : CSV_TYPE;
    const handle = await window.showSaveFilePicker({
      suggestedName,
      types: [type],
    });
    await saveWithHandle(handle, contents, mime);
    return { handle, name: handle.name, kind };
  }
  downloadBlob(contents instanceof Blob ? contents : new Blob([contents], { type: mime }), suggestedName);
  return { handle: null, name: suggestedName, kind };
}

export function safeFilename(title, ext) {
  const base = (title || "Untitled_spreadsheet").replace(/[^\w\-]+/g, "_").replace(/_+/g, "_");
  return `${base}.${ext}`;
}

export { XLSX_TYPE, CSV_TYPE, JSON_TYPE };
