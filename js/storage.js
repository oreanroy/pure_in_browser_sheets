const INDEX_KEY = "gridia.files.index.v1";
const FILE_PREFIX = "gridia.file.v1:";
const LEGACY_KEY = "gridia.workbook.v1";
const META_PREFIX = "gridia.meta.v1:";

function safeKey(name) {
  return encodeURIComponent(String(name || "Untitled").trim() || "Untitled");
}

export function listFiles() {
  try {
    const raw = localStorage.getItem(INDEX_KEY);
    const list = raw ? JSON.parse(raw) : [];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

function writeIndex(list) {
  localStorage.setItem(INDEX_KEY, JSON.stringify(list));
}

export function touchFileInIndex(name, { kind } = {}) {
  const files = listFiles().filter((f) => f.name !== name);
  files.unshift({
    name,
    kind: kind || guessKind(name),
    updatedAt: Date.now(),
  });
  writeIndex(files.slice(0, 40));
}

export function removeFileFromIndex(name) {
  writeIndex(listFiles().filter((f) => f.name !== name));
  localStorage.removeItem(FILE_PREFIX + safeKey(name));
  localStorage.removeItem(META_PREFIX + safeKey(name));
}

export function loadNamedWorkbook(name) {
  try {
    const raw = localStorage.getItem(FILE_PREFIX + safeKey(name));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export function saveNamedWorkbook(name, data, { kind } = {}) {
  localStorage.setItem(FILE_PREFIX + safeKey(name), JSON.stringify(data));
  touchFileInIndex(name, { kind });
  // Keep legacy single-slot in sync for older builds
  localStorage.setItem(LEGACY_KEY, JSON.stringify(data));
}

export function loadFileMeta(name) {
  try {
    const raw = localStorage.getItem(META_PREFIX + safeKey(name));
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function saveFileMeta(name, meta) {
  localStorage.setItem(META_PREFIX + safeKey(name), JSON.stringify(meta || {}));
}

/** Migrate old single workbook into named library once */
export function migrateLegacyWorkbook() {
  const files = listFiles();
  if (files.length) return;
  try {
    const raw = localStorage.getItem(LEGACY_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    const name = `${data.title || "Untitled spreadsheet"}.gridia.json`;
    saveNamedWorkbook(name, data, { kind: "json" });
  } catch {
    /* ignore */
  }
}

export function guessKind(name) {
  const lower = String(name || "").toLowerCase();
  if (lower.endsWith(".xlsx")) return "xlsx";
  if (lower.endsWith(".csv") || lower.endsWith(".txt")) return "csv";
  if (lower.endsWith(".json") || lower.endsWith(".gridia.json")) return "json";
  return "xlsx";
}

export function defaultFileName(title, kind = "xlsx") {
  const base = String(title || "Untitled spreadsheet")
    .replace(/[^\w\- ]+/g, "")
    .trim()
    .replace(/\s+/g, "_") || "Untitled";
  const ext = kind === "json" ? "gridia.json" : kind === "csv" ? "csv" : "xlsx";
  return `${base}.${ext}`;
}

/** Parse `#/files/name.xlsx` or `#files/name.xlsx` or `?file=` */
export function fileNameFromLocation() {
  const params = new URLSearchParams(location.search);
  if (params.get("file")) return decodeURIComponent(params.get("file"));

  const hash = location.hash.replace(/^#/, "");
  const m = hash.match(/^\/?files\/(.+)$/i);
  if (m) return decodeURIComponent(m[1]);
  return null;
}

export function setFileInUrl(name, { replace = false } = {}) {
  const encoded = encodeURIComponent(name);
  const next = `#/files/${encoded}`;
  if (location.hash === next) return;
  if (replace) history.replaceState(null, "", next);
  else history.pushState(null, "", next);
}

// Back-compat exports used by older call sites
export function loadWorkbook() {
  migrateLegacyWorkbook();
  const fromUrl = fileNameFromLocation();
  if (fromUrl) {
    const data = loadNamedWorkbook(fromUrl);
    if (data) return { name: fromUrl, data, kind: loadFileMeta(fromUrl)?.kind || guessKind(fromUrl) };
  }
  const files = listFiles();
  if (files[0]) {
    const data = loadNamedWorkbook(files[0].name);
    if (data) return { name: files[0].name, data, kind: files[0].kind || guessKind(files[0].name) };
  }
  return null;
}

export function saveWorkbook(data, name, kind) {
  if (!name) {
    localStorage.setItem(LEGACY_KEY, JSON.stringify(data));
    return;
  }
  saveNamedWorkbook(name, data, { kind });
  saveFileMeta(name, { kind: kind || guessKind(name) });
}
