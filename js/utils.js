/** Column index (0-based) ↔ letter (A, B, …, AA) */
export function colToLetter(col) {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function letterToCol(letter) {
  let col = 0;
  const s = letter.toUpperCase();
  for (let i = 0; i < s.length; i++) {
    col = col * 26 + (s.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function cellId(row, col) {
  return `${colToLetter(col)}${row + 1}`;
}

export function parseCellId(id) {
  const m = String(id).trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  return { row: parseInt(m[2], 10) - 1, col: letterToCol(m[1]) };
}

export function parseRange(ref) {
  const parts = String(ref).toUpperCase().split(":");
  if (parts.length === 1) {
    const a = parseCellId(parts[0]);
    return a ? { r1: a.row, c1: a.col, r2: a.row, c2: a.col } : null;
  }
  const a = parseCellId(parts[0]);
  const b = parseCellId(parts[1]);
  if (!a || !b) return null;
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

export function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

export function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

export function defaultStyle() {
  return {
    bold: false,
    italic: false,
    underline: false,
    align: "left",
    color: "#202124",
    bg: "",
    fontFamily: "Roboto",
    fontSize: 12,
    format: "general",
    decimals: 2,
  };
}

export function emptyCell() {
  return { raw: "", style: defaultStyle() };
}

export function debounce(fn, ms) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}
