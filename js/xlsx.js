/**
 * Pure-browser XLSX (OOXML) read/write — no Node, no CDN.
 */
import { unzip, zip } from "./zip.js";

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFF_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseXml(bytes) {
  const text = new TextDecoder("utf-8").decode(bytes);
  return new DOMParser().parseFromString(text, "application/xml");
}

function q(el, local, ns = NS) {
  return el.getElementsByTagNameNS(ns, local);
}

function cellRefToRC(ref) {
  const m = String(ref).toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10) - 1, col: col - 1 };
}

function rcToCellRef(row, col) {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `${s}${row + 1}`;
}

function excelSerialToDate(serial) {
  // Excel epoch 1899-12-30 (with 1900 leap bug)
  const ms = Math.round((Number(serial) - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(serial);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  // Prefer date-only when time is midnight
  if (d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0) {
    return `${yyyy}-${mm}-${dd}`;
  }
  return d.toISOString().slice(0, 19).replace("T", " ");
}

function readSharedStrings(files) {
  const raw = files.get("xl/sharedStrings.xml");
  if (!raw) return [];
  const doc = parseXml(raw);
  const out = [];
  for (const si of q(doc, "si")) {
    const texts = [...si.getElementsByTagNameNS(NS, "t")].map((t) => t.textContent || "");
    out.push(texts.join(""));
  }
  return out;
}

function readSheetRows(sheetBytes, sharedStrings) {
  const doc = parseXml(sheetBytes);
  const rows = [];
  for (const rowEl of q(doc, "row")) {
    const rIdx = parseInt(rowEl.getAttribute("r") || "0", 10) - 1;
    while (rows.length <= rIdx) rows.push([]);
    for (const c of q(rowEl, "c")) {
      const ref = c.getAttribute("r");
      const rc = ref ? cellRefToRC(ref) : null;
      if (!rc) continue;
      const t = c.getAttribute("t") || "";
      const vEl = q(c, "v")[0];
      const isEl = q(c, "is")[0];
      let val = "";
      if (t === "inlineStr" && isEl) {
        val = [...isEl.getElementsByTagNameNS(NS, "t")].map((x) => x.textContent || "").join("");
      } else if (vEl) {
        const raw = vEl.textContent || "";
        if (t === "s") val = sharedStrings[parseInt(raw, 10)] ?? "";
        else if (t === "b") val = raw === "1" ? "TRUE" : "FALSE";
        else if (t === "str") val = raw;
        else {
          const n = Number(raw);
          if (Number.isFinite(n) && n > 20000 && n < 60000 && raw.includes(".")) {
            if (Number.isInteger(n) && n >= 30000 && n <= 60000) {
              val = excelSerialToDate(n);
            } else {
              val = String(n);
            }
          } else if (Number.isFinite(n)) {
            if (Number.isInteger(n) && n >= 30000 && n <= 60000) {
              val = excelSerialToDate(n);
            } else {
              val = String(n);
            }
          } else {
            val = raw;
          }
        }
      }
      const row = rows[rc.row];
      while (row.length <= rc.col) row.push("");
      row[rc.col] = val;
    }
  }
  return rows;
}

function readSheetHyperlinks(sheetBytes, sheetRelsBytes) {
  const links = {};
  if (!sheetBytes) return links;
  const doc = parseXml(sheetBytes);
  const relMap = new Map();
  if (sheetRelsBytes) {
    const rels = parseXml(sheetRelsBytes);
    for (const rel of rels.getElementsByTagNameNS(REL_NS, "Relationship")) {
      relMap.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
    }
  }
  for (const h of q(doc, "hyperlink")) {
    const ref = h.getAttribute("ref");
    if (!ref) continue;
    // Only single-cell refs for now (A1); skip ranges
    const rc = cellRefToRC(ref.split(":")[0]);
    if (!rc) continue;
    const rid = h.getAttributeNS(OFF_REL, "id") || h.getAttribute("r:id");
    const location = h.getAttribute("location");
    const display = h.getAttribute("display");
    let url = rid ? relMap.get(rid) : null;
    if (!url && location) url = `#${location}`;
    if (!url) continue;
    links[`${rc.row},${rc.col}`] = { url, display: display || null };
  }
  return links;
}

/**
 * @param {ArrayBuffer|Uint8Array} buffer
 * @returns {Promise<{ sheets: { name: string, rows: string[][], links: Record<string,{url:string,display:string|null}> }[] }>}
 */
export async function readXlsx(buffer) {
  const files = await unzip(buffer);
  const shared = readSharedStrings(files);

  const wb = parseXml(files.get("xl/workbook.xml"));
  const rels = parseXml(files.get("xl/_rels/workbook.xml.rels"));
  const relMap = new Map();
  for (const rel of rels.getElementsByTagNameNS(REL_NS, "Relationship")) {
    relMap.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
  }

  const sheets = [];
  for (const sheet of q(wb, "sheet")) {
    const name = sheet.getAttribute("name") || `Sheet${sheets.length + 1}`;
    const rid = sheet.getAttributeNS(OFF_REL, "id") || sheet.getAttribute("r:id");
    let target = relMap.get(rid);
    if (!target) continue;
    if (target.startsWith("/")) target = target.slice(1);
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    const data = files.get(target);
    if (!data) continue;
    const relPath = target.replace(/worksheets\/([^/]+)$/, "worksheets/_rels/$1.rels");
    const sheetRels = files.get(relPath);
    const rows = readSheetRows(data, shared);
    const links = readSheetHyperlinks(data, sheetRels);
    sheets.push({ name, rows, links });
  }

  if (!sheets.length) throw new Error("No worksheets found in XLSX");
  return { sheets };
}

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function sheetToXml(rows) {
  const rowXml = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const cells = [];
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (val === "" || val == null) continue;
      const ref = rcToCellRef(r, c);
      const s = String(val);
      if (/^-?\d+(\.\d+)?$/.test(s) && !s.startsWith("=")) {
        cells.push(`<c r="${ref}"><v>${s}</v></c>`);
      } else if (s.startsWith("=")) {
        cells.push(`<c r="${ref}" t="str"><v>${xmlEscape(s)}</v></c>`);
      } else {
        cells.push(`<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(s)}</t></is></c>`);
      }
    }
    if (cells.length) rowXml.push(`<row r="${r + 1}">${cells.join("")}</row>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}"><sheetData>${rowXml.join("")}</sheetData></worksheet>`;
}

/**
 * @param {{ name: string, rows: string[][] }[]} sheets
 * @returns {Promise<Uint8Array>}
 */
export async function writeXlsx(sheets) {
  const entries = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheets
    .map(
      (_, i) =>
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    )
    .join("")}
</Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="${OFF_REL}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS}" xmlns:r="${OFF_REL}">
  <sheets>
    ${sheets
      .map(
        (s, i) =>
          `<sheet name="${xmlEscape(s.name).slice(0, 31)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`
      )
      .join("")}
  </sheets>
</workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  ${sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="${OFF_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
    )
    .join("")}
</Relationships>`,
  };

  sheets.forEach((s, i) => {
    entries[`xl/worksheets/sheet${i + 1}.xml`] = sheetToXml(s.rows || []);
  });

  return zip(entries);
}
