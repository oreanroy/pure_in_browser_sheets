import { parseCellId, parseRange, letterToCol } from "./utils.js";

/** Display + URL pair from =HYPERLINK() */
export class HyperlinkValue {
  constructor(url, label) {
    this.url = String(url ?? "");
    this.label = label == null || label === "" ? this.url : String(label);
  }
  toString() {
    return this.label;
  }
  valueOf() {
    return this.label;
  }
}

export function isHyperlinkValue(v) {
  return v instanceof HyperlinkValue;
}

export function detectUrl(text) {
  if (text == null) return null;
  const s = String(text).trim();
  if (!s) return null;
  if (/^https?:\/\/\S+$/i.test(s)) return s;
  if (/^mailto:\S+$/i.test(s)) return s;
  if (/^www\.\S+$/i.test(s)) return `https://${s}`;
  return null;
}

export function normalizeUrl(url) {
  const s = String(url || "").trim();
  if (!s) return "";
  if (/^(https?:|mailto:)/i.test(s)) return s;
  if (/^www\./i.test(s)) return `https://${s}`;
  // bare domain-ish
  if (/^[a-z0-9.-]+\.[a-z]{2,}([/:].*)?$/i.test(s)) return `https://${s}`;
  return s;
}

const FUNCTIONS = {
  SUM: (args) => args.reduce((a, b) => a + (toNum(b) || 0), 0),
  AVERAGE: (args) => {
    const nums = args.map(toNum).filter((n) => n !== null);
    return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
  },
  COUNT: (args) => args.map(toNum).filter((n) => n !== null).length,
  COUNTA: (args) => args.filter((v) => v !== "" && v !== null && v !== undefined).length,
  MIN: (args) => {
    const nums = args.map(toNum).filter((n) => n !== null);
    return nums.length ? Math.min(...nums) : 0;
  },
  MAX: (args) => {
    const nums = args.map(toNum).filter((n) => n !== null);
    return nums.length ? Math.max(...nums) : 0;
  },
  ABS: (args) => Math.abs(toNum(args[0]) || 0),
  ROUND: (args) => {
    const n = toNum(args[0]) || 0;
    const d = toNum(args[1]) ?? 0;
    const f = 10 ** d;
    return Math.round(n * f) / f;
  },
  IF: (args) => (truthy(args[0]) ? args[1] : args[2]),
  CONCAT: (args) => args.map((a) => (a == null ? "" : String(a))).join(""),
  CONCATENATE: (args) => args.map((a) => (a == null ? "" : String(a))).join(""),
  LEN: (args) => String(args[0] ?? "").length,
  UPPER: (args) => String(args[0] ?? "").toUpperCase(),
  LOWER: (args) => String(args[0] ?? "").toLowerCase(),
  TRIM: (args) => String(args[0] ?? "").trim(),
  HYPERLINK: (args) => {
    const url = args[0] == null ? "" : String(unwrap(args[0]));
    const label = args.length > 1 ? unwrap(args[1]) : url;
    return new HyperlinkValue(url, label == null ? url : String(label));
  },
  TRUE: () => true,
  FALSE: () => false,
  AND: (args) => args.every(truthy),
  OR: (args) => args.some(truthy),
  NOT: (args) => !truthy(args[0]),
};

function unwrap(v) {
  if (isHyperlinkValue(v)) return v.label;
  return v;
}

function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  if (isHyperlinkValue(v)) return toNum(v.label);
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "boolean") return v ? 1 : 0;
  const n = Number(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}

function truthy(v) {
  if (isHyperlinkValue(v)) return truthy(v.label);
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v === "" || v == null) return false;
  const n = toNum(v);
  if (n !== null) return n !== 0;
  return Boolean(v);
}

function tokenize(expr) {
  const tokens = [];
  let i = 0;
  while (i < expr.length) {
    const ch = expr[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if ("(),:+-*/^&%<>=!".includes(ch)) {
      if ((ch === "<" || ch === ">" || ch === "!" || ch === "=") && i + 1 < expr.length) {
        const two = ch + expr[i + 1];
        if (["<=", ">=", "<>", "==", "!="].includes(two)) {
          tokens.push({ type: "op", value: two === "==" ? "=" : two === "!=" ? "<>" : two });
          i += 2;
          continue;
        }
      }
      if (ch === "=" && tokens.length > 0) {
        tokens.push({ type: "op", value: "=" });
        i++;
        continue;
      }
      tokens.push({ type: "op", value: ch });
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let s = "";
      while (j < expr.length && expr[j] !== '"') {
        if (expr[j] === "\\" && j + 1 < expr.length) {
          s += expr[j + 1];
          j += 2;
        } else {
          s += expr[j++];
        }
      }
      tokens.push({ type: "string", value: s });
      i = j + 1;
      continue;
    }
    if (/[0-9.]/.test(ch)) {
      let j = i;
      while (j < expr.length && /[0-9.]/.test(expr[j])) j++;
      tokens.push({ type: "number", value: parseFloat(expr.slice(i, j)) });
      i = j;
      continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i;
      while (j < expr.length && /[A-Za-z0-9_]/.test(expr[j])) j++;
      // range like A1:B2 or cell A1
      if (j < expr.length && expr[j] === ":" && /[A-Za-z]/.test(expr[j + 1] || "")) {
        let k = j + 1;
        while (k < expr.length && /[A-Za-z0-9]/.test(expr[k])) k++;
        tokens.push({ type: "ref", value: expr.slice(i, k).toUpperCase() });
        i = k;
        continue;
      }
      const word = expr.slice(i, j);
      // cell ref like A1 / AA12
      if (/^[A-Za-z]+\d+$/.test(word)) {
        tokens.push({ type: "ref", value: word.toUpperCase() });
      } else {
        tokens.push({ type: "ident", value: word.toUpperCase() });
      }
      i = j;
      continue;
    }
    throw new Error(`Unexpected '${ch}'`);
  }
  return tokens;
}

class Parser {
  constructor(tokens, getCellValue, evaluating) {
    this.tokens = tokens;
    this.i = 0;
    this.getCellValue = getCellValue;
    this.evaluating = evaluating;
  }

  peek() {
    return this.tokens[this.i];
  }

  next() {
    return this.tokens[this.i++];
  }

  expect(value) {
    const t = this.next();
    if (!t || t.value !== value) throw new Error(`Expected ${value}`);
    return t;
  }

  parse() {
    const v = this.parseComparison();
    if (this.i < this.tokens.length) throw new Error("Unexpected token");
    return v;
  }

  parseComparison() {
    let left = this.parseConcat();
    while (this.peek() && ["=", "<>", "<", ">", "<=", ">="].includes(this.peek().value)) {
      const op = this.next().value;
      const right = this.parseConcat();
      const ln = toNum(left);
      const rn = toNum(right);
      if (ln !== null && rn !== null) {
        left =
          op === "="
            ? ln === rn
            : op === "<>"
              ? ln !== rn
              : op === "<"
                ? ln < rn
                : op === ">"
                  ? ln > rn
                  : op === "<="
                    ? ln <= rn
                    : ln >= rn;
      } else {
        const ls = String(left ?? "");
        const rs = String(right ?? "");
        left =
          op === "="
            ? ls === rs
            : op === "<>"
              ? ls !== rs
              : op === "<"
                ? ls < rs
                : op === ">"
                  ? ls > rs
                  : op === "<="
                    ? ls <= rs
                    : ls >= rs;
      }
    }
    return left;
  }

  parseConcat() {
    let left = this.parseAdd();
    while (this.peek() && this.peek().value === "&") {
      this.next();
      const right = this.parseAdd();
      left = `${left ?? ""}${right ?? ""}`;
    }
    return left;
  }

  parseAdd() {
    let left = this.parseMul();
    while (this.peek() && (this.peek().value === "+" || this.peek().value === "-")) {
      const op = this.next().value;
      const right = this.parseMul();
      const ln = toNum(left) ?? 0;
      const rn = toNum(right) ?? 0;
      left = op === "+" ? ln + rn : ln - rn;
    }
    return left;
  }

  parseMul() {
    let left = this.parsePow();
    while (this.peek() && (this.peek().value === "*" || this.peek().value === "/" || this.peek().value === "%")) {
      const op = this.next().value;
      const right = this.parsePow();
      const ln = toNum(left) ?? 0;
      const rn = toNum(right) ?? 0;
      if (op === "*") left = ln * rn;
      else if (op === "/") left = rn === 0 ? "#DIV/0!" : ln / rn;
      else left = ln % rn;
    }
    return left;
  }

  parsePow() {
    let left = this.parseUnary();
    while (this.peek() && this.peek().value === "^") {
      this.next();
      const right = this.parseUnary();
      left = (toNum(left) ?? 0) ** (toNum(right) ?? 0);
    }
    return left;
  }

  parseUnary() {
    if (this.peek() && (this.peek().value === "+" || this.peek().value === "-")) {
      const op = this.next().value;
      const v = this.parseUnary();
      const n = toNum(v) ?? 0;
      return op === "-" ? -n : n;
    }
    return this.parsePrimary();
  }

  parsePrimary() {
    const t = this.peek();
    if (!t) throw new Error("Unexpected end");

    if (t.type === "number") {
      this.next();
      return t.value;
    }
    if (t.type === "string") {
      this.next();
      return t.value;
    }
    if (t.type === "ref") {
      this.next();
      return this.resolveRef(t.value);
    }
    if (t.type === "ident") {
      this.next();
      if (this.peek() && this.peek().value === "(") {
        this.next();
        const args = [];
        if (!(this.peek() && this.peek().value === ")")) {
          args.push(...this.parseArgList());
        }
        this.expect(")");
        const fn = FUNCTIONS[t.value];
        if (!fn) throw new Error(`Unknown function ${t.value}`);
        return fn(args);
      }
      if (t.value === "TRUE") return true;
      if (t.value === "FALSE") return false;
      throw new Error(`Unknown name ${t.value}`);
    }
    if (t.value === "(") {
      this.next();
      const v = this.parseComparison();
      this.expect(")");
      return v;
    }
    throw new Error(`Unexpected ${t.value}`);
  }

  parseArgList() {
    const flat = [];
    const pushArg = (v) => {
      if (Array.isArray(v)) flat.push(...v);
      else flat.push(v);
    };
    pushArg(this.parseComparison());
    while (this.peek() && this.peek().value === ",") {
      this.next();
      pushArg(this.parseComparison());
    }
    return flat;
  }

  resolveRef(ref) {
    const range = parseRange(ref);
    if (!range) throw new Error(`Bad ref ${ref}`);
    const values = [];
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const key = `${r},${c}`;
        if (this.evaluating.has(key)) throw new Error("#CYCLE!");
        values.push(this.getCellValue(r, c));
      }
    }
    if (range.r1 === range.r2 && range.c1 === range.c2) return values[0];
    return values;
  }
}

export function evaluateFormula(raw, getCellValue, evaluating = new Set()) {
  if (typeof raw !== "string" || !raw.startsWith("=")) {
    return coerceLiteral(raw);
  }
  const expr = raw.slice(1).trim();
  if (!expr) return "";
  try {
    const tokens = tokenize(expr);
    const parser = new Parser(tokens, getCellValue, evaluating);
    return parser.parse();
  } catch (e) {
    const msg = e.message || "ERROR";
    if (msg.startsWith("#")) return msg;
    return `#ERROR!`;
  }
}

function coerceLiteral(raw) {
  if (raw == null || raw === "") return "";
  if (typeof raw === "number") return raw;
  const s = String(raw);
  if (/^-?\d+(\.\d+)?$/.test(s.trim())) return Number(s.trim());
  return s;
}

export function formatDisplay(value, style) {
  if (value === null || value === undefined || value === "") return "";
  if (isHyperlinkValue(value)) return value.label;
  if (typeof value === "string" && value.startsWith("#")) return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";

  const format = style?.format || "general";
  const decimals = style?.decimals ?? 2;
  const n = typeof value === "number" ? value : toNum(value);

  if (n !== null && format === "percent") {
    return `${(n * 100).toFixed(decimals)}%`;
  }
  if (n !== null && format === "currency") {
    return n.toLocaleString(undefined, {
      style: "currency",
      currency: "USD",
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    });
  }
  if (n !== null && format === "number") {
    return n.toLocaleString(undefined, {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    });
  }
  if (typeof value === "number") {
    if (Number.isInteger(value)) return String(value);
    return String(Math.round(value * 1e10) / 1e10);
  }
  return String(value);
}

export { toNum, letterToCol, parseCellId };
