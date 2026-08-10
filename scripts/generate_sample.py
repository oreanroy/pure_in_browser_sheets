#!/usr/bin/env python3
"""Generate the complex Gridia sample workbook + README preview PNGs."""

from __future__ import annotations

import json
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
PREVIEWS = SAMPLES / "previews"


def style(**kwargs):
    base = {
        "bold": False,
        "italic": False,
        "underline": False,
        "align": "left",
        "color": "#202124",
        "bg": "",
        "fontFamily": "Roboto",
        "fontSize": 12,
        "format": "general",
        "decimals": 2,
    }
    base.update(kwargs)
    return base


def cell(raw, **st):
    return {"raw": str(raw), "style": style(**st)}


def sheet(name: str, cells: dict, col_widths: dict | None = None, rows=40, cols=12):
    return {
        "id": str(uuid.uuid4()),
        "name": name,
        "rows": rows,
        "cols": cols,
        "cells": cells,
        "colWidths": col_widths or {},
    }


def build_workbook():
    # —— Dashboard ——
    dash = {}
    dash["0,0"] = cell("Gridia Ops HQ", bold=True, fontSize=20, fontFamily="Montserrat", color="#1a7340")
    dash["1,0"] = cell("Q3 performance snapshot · sample workbook", italic=True, color="#5f6368", fontSize=11)
    dash["0,5"] = cell("=HYPERLINK(\"https://github.com/oreanroy/pure_in_browser_sheets\",\"Open repo\")", color="#1155cc", underline=True)

    headers = ["Metric", "Actual", "Target", "Variance", "Status"]
    for i, h in enumerate(headers):
        dash[f"3,{i}"] = cell(h, bold=True, bg="#1a7340", color="#ffffff", align="center" if i else "left")

    metrics = [
        ("Revenue", 248500, 240000),
        ("Gross margin %", 0.42, 0.40),
        ("Active customers", 1860, 1800),
        ("NPS", 62, 55),
        ("Churn %", 0.028, 0.035),
    ]
    for r, (label, actual, target) in enumerate(metrics, start=4):
        dash[f"{r},0"] = cell(label, bold=True)
        if "%" in label:
            dash[f"{r},1"] = cell(str(actual), format="percent", decimals=1, align="right")
            dash[f"{r},2"] = cell(str(target), format="percent", decimals=1, align="right")
            dash[f"{r},3"] = cell(f"=B{r+1}-C{r+1}", format="percent", decimals=1, align="right")
        elif label == "Revenue":
            dash[f"{r},1"] = cell(str(actual), format="currency", align="right")
            dash[f"{r},2"] = cell(str(target), format="currency", align="right")
            dash[f"{r},3"] = cell(f"=B{r+1}-C{r+1}", format="currency", align="right")
        else:
            dash[f"{r},1"] = cell(str(actual), align="right", format="number", decimals=0)
            dash[f"{r},2"] = cell(str(target), align="right", format="number", decimals=0)
            dash[f"{r},3"] = cell(f"=B{r+1}-C{r+1}", align="right", format="number", decimals=0)
        # Churn: lower is better
        if label.startswith("Churn"):
            ok = actual <= target
            dash[f"{r},4"] = cell(
                "On track" if ok else "Watch",
                bold=True,
                align="center",
                bg="#e6f4ea" if ok else "#fce8e6",
                color="#1a7340" if ok else "#c5221f",
            )
        else:
            ok = actual >= target
            dash[f"{r},4"] = cell(
                f'=IF(D{r+1}>=0,"On track","Watch")',
                bold=True,
                align="center",
                bg="#e6f4ea" if ok else "#fce8e6",
                color="#1a7340" if ok else "#c5221f",
            )

    dash["10,0"] = cell("Pipeline health", bold=True, fontSize=14, fontFamily="Montserrat")
    dash["11,0"] = cell("Open deals")
    dash["11,1"] = cell("10", align="right")
    dash["12,0"] = cell("Weighted pipeline")
    dash["12,1"] = cell("523650", format="currency", align="right", bold=True)
    dash["13,0"] = cell("Avg deal size")
    dash["13,1"] = cell("103200", format="currency", align="right")
    dash["14,0"] = cell("Coverage vs target")
    dash["14,1"] = cell("=B12/240000", format="percent", decimals=0, align="right")

    dash["10,3"] = cell("Notes", bold=True, bg="#fff3cd")
    dash["11,3"] = cell("Targets locked after board review (Aug).", fontSize=11)
    dash["12,3"] = cell("Churn improving vs Q2 (−40 bps).", fontSize=11)
    dash["13,3"] = cell("=HYPERLINK(\"https://example.com/ops-runbook\",\"Ops runbook\")", color="#1155cc", underline=True)

    # —— Budget ——
    budget = {}
    budget["0,0"] = cell("Department budget · FY2026", bold=True, fontSize=16, fontFamily="Montserrat", color="#1a7340")
    for i, h in enumerate(["Department", "Q1", "Q2", "Q3", "Q4", "Annual", "% of total"]):
        budget[f"2,{i}"] = cell(h, bold=True, bg="#e8f0fe", color="#174ea6", align="right" if i else "left")

    depts = [
        ("Engineering", 120000, 125000, 130000, 140000),
        ("Product", 45000, 48000, 52000, 55000),
        ("Sales", 90000, 95000, 110000, 125000),
        ("Marketing", 60000, 62000, 70000, 80000),
        ("People", 35000, 36000, 38000, 40000),
        ("G&A", 28000, 28000, 30000, 32000),
    ]
    for r, (name, q1, q2, q3, q4) in enumerate(depts, start=3):
        budget[f"{r},0"] = cell(name)
        budget[f"{r},1"] = cell(q1, format="currency", align="right")
        budget[f"{r},2"] = cell(q2, format="currency", align="right")
        budget[f"{r},3"] = cell(q3, format="currency", align="right")
        budget[f"{r},4"] = cell(q4, format="currency", align="right")
        budget[f"{r},5"] = cell(f"=SUM(B{r+1}:E{r+1})", format="currency", align="right", bold=True)
        # Annual / total annual (F10) — use plain refs (engine has no $ absolutes)
        budget[f"{r},6"] = cell(f"=F{r+1}/F10", format="percent", decimals=1, align="right")

    budget["9,0"] = cell("TOTAL", bold=True, bg="#1a7340", color="#ffffff")
    for c in range(1, 6):
        col = chr(65 + c)
        budget[f"9,{c}"] = cell(f"=SUM({col}4:{col}9)", format="currency", align="right", bold=True, bg="#1a7340", color="#ffffff")
    budget["9,6"] = cell("1", format="percent", decimals=0, align="right", bold=True, bg="#1a7340", color="#ffffff")

    budget["11,0"] = cell("Headcount plan", bold=True, fontSize=14)
    for i, h in enumerate(["Role", "Current", "Open reqs", "Fully loaded cost"]):
        budget[f"12,{i}"] = cell(h, bold=True, bg="#f3f3f3")
    roles = [
        ("Staff engineer", 4, 1, 210000),
        ("Product manager", 3, 1, 165000),
        ("AE (enterprise)", 6, 2, 145000),
        ("Designer", 2, 1, 125000),
    ]
    for r, (role, cur, open_, cost) in enumerate(roles, start=13):
        budget[f"{r},0"] = cell(role)
        budget[f"{r},1"] = cell(cur, align="right")
        budget[f"{r},2"] = cell(open_, align="right", bg="#fff3cd")
        budget[f"{r},3"] = cell(cost, format="currency", align="right")
    budget["17,0"] = cell("Hiring budget impact")
    budget["17,3"] = cell("=(B14+C14)*D14+(B15+C15)*D15+(B16+C16)*D16+(B17+C17)*D17", format="currency", align="right", bold=True)

    # —— Deals ——
    deals = {}
    deals["0,0"] = cell("Enterprise pipeline", bold=True, fontSize=16, fontFamily="Montserrat", color="#1a7340")
    deals["1,0"] = cell("Filter: stage ≠ Lost · amounts in USD", italic=True, color="#5f6368", fontSize=11)
    for i, h in enumerate(["Account", "Owner", "Stage", "Amount", "Prob %", "Weighted", "Close", "Link"]):
        deals[f"3,{i}"] = cell(h, bold=True, bg="#202124", color="#ffffff", align="center" if i > 2 else "left")

    rows = [
        ("Northwind Labs", "Asha", "Negotiation", 82000, 0.7, "2026-09-15", "https://example.com/deals/nw"),
        ("Blue River Health", "Jon", "Proposal", 54000, 0.45, "2026-10-01", "https://example.com/deals/br"),
        ("Cedar Robotics", "Asha", "Discovery", 120000, 0.25, "2026-11-12", "https://example.com/deals/cr"),
        ("Helios Energy", "Priya", "Negotiation", 210000, 0.6, "2026-09-30", "https://example.com/deals/he"),
        ("Orbit Freight", "Jon", "Closed Won", 67000, 1.0, "2026-08-01", "https://example.com/deals/of"),
        ("Pixel & Form", "Maya", "Proposal", 38000, 0.4, "2026-10-20", "https://example.com/deals/pf"),
        ("Summit Analytics", "Priya", "Discovery", 95000, 0.2, "2026-12-05", "https://example.com/deals/sa"),
        ("Kite Payments", "Maya", "Negotiation", 150000, 0.55, "2026-09-22", "https://example.com/deals/kp"),
        ("Lumen Schools", "Asha", "Closed Won", 41000, 1.0, "2026-07-18", "https://example.com/deals/ls"),
        ("Nova Grid", "Jon", "Proposal", 175000, 0.35, "2026-11-01", "https://example.com/deals/ng"),
    ]
    for r, (acct, owner, stage, amt, prob, close, url) in enumerate(rows, start=4):
        bg = "#e6f4ea" if stage == "Closed Won" else ("#e8f0fe" if stage == "Negotiation" else "")
        deals[f"{r},0"] = cell(acct, bg=bg)
        deals[f"{r},1"] = cell(owner, bg=bg)
        deals[f"{r},2"] = cell(stage, bg=bg, align="center")
        deals[f"{r},3"] = cell(amt, format="currency", align="right", bg=bg)
        deals[f"{r},4"] = cell(prob, format="percent", decimals=0, align="right", bg=bg)
        deals[f"{r},5"] = cell(f"=D{r+1}*E{r+1}", format="currency", align="right", bold=True, bg=bg)
        deals[f"{r},6"] = cell(close, align="center", bg=bg, fontFamily="IBM Plex Mono", fontSize=11)
        deals[f"{r},7"] = cell(f'=HYPERLINK("{url}","Open")', color="#1155cc", underline=True, align="center", bg=bg)

    deals["15,0"] = cell("Totals", bold=True)
    deals["15,3"] = cell("=SUM(D5:D14)", format="currency", align="right", bold=True, bg="#f3f3f3")
    deals["15,5"] = cell("=SUM(F5:F14)", format="currency", align="right", bold=True, bg="#e6f4ea")
    deals["16,0"] = cell("Best case (100%)")
    deals["16,5"] = cell("=SUM(D5:D14)", format="currency", align="right")
    deals["17,0"] = cell("Commit (≥60% prob)")
    deals["17,5"] = cell("=F5+F8+F12+F9", format="currency", align="right", bold=True)

    # —— Hiring ——
    hiring = {}
    hiring["0,0"] = cell("Hiring scorecard", bold=True, fontSize=16, fontFamily="Montserrat", color="#1a7340")
    for i, h in enumerate(["Candidate", "Role", "Location", "Stage", "Score", "Decision", "Portfolio"]):
        hiring[f"2,{i}"] = cell(h, bold=True, bg="#fce8e6", color="#c5221f")
    cands = [
        ("Riley Chen", "Staff Eng", "Remote", "Onsite", 4.6, "Advance", "https://example.com/r"),
        ("Sam Okonkwo", "PM", "NYC", "Offer", 4.8, "Offer out", "https://example.com/s"),
        ("Jules Park", "Designer", "SF", "Screen", 3.9, "Hold", "https://example.com/j"),
        ("Alex Rivera", "AE", "Austin", "Onsite", 4.2, "Advance", "https://example.com/a"),
        ("Morgan Lee", "Staff Eng", "Remote", "Reject", 2.8, "Pass", "https://example.com/m"),
        ("Casey Brooks", "PM", "Remote", "Screen", 4.0, "Advance", "https://example.com/c"),
    ]
    for r, row in enumerate(cands, start=3):
        for c, val in enumerate(row):
            if c == 6:
                hiring[f"{r},{c}"] = cell(f'=HYPERLINK("{val}","Link")', color="#1155cc", underline=True)
            elif c == 4:
                hiring[f"{r},{c}"] = cell(val, align="right", format="number", decimals=1, bold=True)
            elif c == 5:
                color = "#1a7340" if val in ("Advance", "Offer out") else ("#c5221f" if val == "Pass" else "#b06000")
                hiring[f"{r},{c}"] = cell(val, align="center", bold=True, color=color)
            else:
                hiring[f"{r},{c}"] = cell(val)
    hiring["10,0"] = cell("Average score")
    hiring["10,4"] = cell("=AVERAGE(E4:E9)", format="number", decimals=2, bold=True, align="right")
    hiring["11,0"] = cell("Interview load (onsite+offer)")
    hiring["11,4"] = cell("=2+1", align="right")  # illustrative
    hiring["11,4"] = cell("3", align="right", bold=True)

    wb = {
        "title": "Gridia Complex Sample — Ops HQ",
        "activeSheetId": None,
        "sheets": [
            sheet("Dashboard", dash, {"0": 160, "1": 110, "2": 110, "3": 110, "4": 100, "5": 140}, rows=50, cols=10),
            sheet("Budget", budget, {"0": 140, "1": 100, "2": 100, "3": 100, "4": 100, "5": 110, "6": 100}, rows=50, cols=10),
            sheet("Deals", deals, {"0": 150, "1": 80, "2": 110, "3": 100, "4": 80, "5": 110, "6": 110, "7": 70}, rows=50, cols=10),
            sheet("Hiring", hiring, {"0": 130, "1": 100, "2": 90, "3": 90, "4": 70, "5": 100, "6": 80}, rows=40, cols=10),
        ],
    }
    wb["activeSheetId"] = wb["sheets"][0]["id"]
    return wb


def excel_escape(s: str) -> str:
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def col_letter(c: int) -> str:
    n = c + 1
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def sheet_to_xlsx_xml(cells: dict) -> str:
    # group by row
    by_row: dict[int, list] = {}
    for key, cell in cells.items():
        r, c = map(int, key.split(","))
        by_row.setdefault(r, []).append((c, cell))
    rows_xml = []
    for r in sorted(by_row):
        parts = []
        for c, cell in sorted(by_row[r], key=lambda x: x[0]):
            ref = f"{col_letter(c)}{r + 1}"
            raw = cell.get("raw", "")
            if raw.startswith("="):
                # store formula display text as string for interoperability
                parts.append(f'<c r="{ref}" t="str"><v>{excel_escape(raw)}</v></c>')
            else:
                try:
                    float(raw)
                    parts.append(f'<c r="{ref}"><v>{excel_escape(raw)}</v></c>')
                except ValueError:
                    parts.append(f'<c r="{ref}" t="inlineStr"><is><t>{excel_escape(raw)}</t></is></c>')
        rows_xml.append(f'<row r="{r + 1}">{"".join(parts)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(rows_xml)}</sheetData></worksheet>'
    )


def write_xlsx(path: Path, wb: dict):
    sheets = wb["sheets"]
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        overrides = [
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        ]
        for i in range(len(sheets)):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides)
            + "</Types>",
        )
        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )
        sheet_tags = []
        rels = []
        for i, s in enumerate(sheets):
            sheet_tags.append(
                f'<sheet name="{excel_escape(s["name"])[:31]}" sheetId="{i+1}" r:id="rId{i+1}"/>'
            )
            rels.append(
                f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
            )
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", sheet_to_xlsx_xml(s["cells"]))
        z.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets>{"".join(sheet_tags)}</sheets></workbook>',
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels)
            + "</Relationships>",
        )
    path.write_bytes(buf.getvalue())


def try_font(size: int, bold: bool = False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/Library/Fonts/Arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except Exception:
            continue
    return ImageFont.load_default()


def format_display(raw: str, st: dict) -> str:
    if raw.startswith("="):
        # Show a friendly computed-looking value for known patterns in sample
        return raw
    fmt = st.get("format", "general")
    try:
        n = float(raw)
    except ValueError:
        return raw
    if fmt == "currency":
        return f"${n:,.0f}" if float(n).is_integer() else f"${n:,.2f}"
    if fmt == "percent":
        return f"{n*100:.1f}%".replace(".0%", "%")
    if fmt == "number":
        d = st.get("decimals", 2)
        return f"{n:.{d}f}"
    if float(n).is_integer():
        return str(int(n))
    return str(n)


# Precomputed display values for preview (formulas resolved for the sample)
PREVIEW_OVERRIDES = {
    ("Dashboard", "4,3"): "$8,500",
    ("Dashboard", "4,4"): "On track",
    ("Dashboard", "5,3"): "2.0%",
    ("Dashboard", "5,4"): "On track",
    ("Dashboard", "6,3"): "60",
    ("Dashboard", "6,4"): "On track",
    ("Dashboard", "7,3"): "7",
    ("Dashboard", "7,4"): "On track",
    ("Dashboard", "8,3"): "-0.7%",
    ("Dashboard", "8,4"): "On track",
    ("Dashboard", "0,5"): "Open repo",
    ("Dashboard", "11,1"): "10",
    ("Dashboard", "12,1"): "$523,650",
    ("Dashboard", "13,1"): "$103,200",
    ("Dashboard", "14,1"): "43%",
    ("Dashboard", "13,3"): "Ops runbook",
    ("Budget", "3,5"): "$515,000",
    ("Budget", "3,6"): "31.2%",
    ("Budget", "4,5"): "$200,000",
    ("Budget", "4,6"): "12.1%",
    ("Budget", "5,5"): "$420,000",
    ("Budget", "5,6"): "25.5%",
    ("Budget", "6,5"): "$272,000",
    ("Budget", "6,6"): "16.5%",
    ("Budget", "7,5"): "$149,000",
    ("Budget", "7,6"): "9.0%",
    ("Budget", "8,5"): "$118,000",
    ("Budget", "8,6"): "7.2%",
    ("Budget", "9,1"): "$378,000",
    ("Budget", "9,2"): "$394,000",
    ("Budget", "9,3"): "$430,000",
    ("Budget", "9,4"): "$472,000",
    ("Budget", "9,5"): "$1,674,000",
    ("Budget", "9,6"): "100%",
    ("Budget", "17,3"): "$2,815,000",
    ("Deals", "4,5"): "$57,400",
    ("Deals", "5,5"): "$24,300",
    ("Deals", "6,5"): "$30,000",
    ("Deals", "7,5"): "$126,000",
    ("Deals", "8,5"): "$67,000",
    ("Deals", "9,5"): "$15,200",
    ("Deals", "10,5"): "$19,000",
    ("Deals", "11,5"): "$82,500",
    ("Deals", "12,5"): "$41,000",
    ("Deals", "13,5"): "$61,250",
    ("Deals", "4,7"): "Open",
    ("Deals", "5,7"): "Open",
    ("Deals", "6,7"): "Open",
    ("Deals", "7,7"): "Open",
    ("Deals", "8,7"): "Open",
    ("Deals", "9,7"): "Open",
    ("Deals", "10,7"): "Open",
    ("Deals", "11,7"): "Open",
    ("Deals", "12,7"): "Open",
    ("Deals", "13,7"): "Open",
    ("Deals", "15,3"): "$1,032,000",
    ("Deals", "15,5"): "$523,650",
    ("Deals", "16,5"): "$1,032,000",
    ("Deals", "17,5"): "$333,900",
    ("Hiring", "3,6"): "Link",
    ("Hiring", "4,6"): "Link",
    ("Hiring", "5,6"): "Link",
    ("Hiring", "6,6"): "Link",
    ("Hiring", "7,6"): "Link",
    ("Hiring", "8,6"): "Link",
    ("Hiring", "10,4"): "4.05",
}


def render_sheet_png(sheet: dict, out: Path, max_rows=18, max_cols=8):
    cells = sheet["cells"]
    widths = []
    for c in range(max_cols):
        widths.append(int(sheet.get("colWidths", {}).get(str(c), 110)))
    header_w, header_h = 36, 24
    row_h = 26
    width = header_w + sum(widths) + 1
    height = header_h + max_rows * row_h + 48

    img = Image.new("RGB", (width, height), "#f0f2f4")
    draw = ImageDraw.Draw(img)

    # top chrome
    draw.rectangle([0, 0, width, 40], fill="#ffffff")
    draw.text((12, 10), f"Gridia  ·  {sheet['name']}", fill="#1a7340", font=try_font(14, True))
    draw.line([0, 40, width, 40], fill="#e2e4e8")

    origin_y = 40
    # corner
    draw.rectangle([0, origin_y, header_w, origin_y + header_h], fill="#f5f6f7", outline="#c0c0c0")
    # col headers
    x = header_w
    for c in range(max_cols):
        draw.rectangle([x, origin_y, x + widths[c], origin_y + header_h], fill="#f5f6f7", outline="#e0e0e0")
        label = col_letter(c)
        draw.text((x + widths[c] / 2 - 4, origin_y + 5), label, fill="#5f6368", font=try_font(11, True))
        x += widths[c]

    font = try_font(12)
    font_bold = try_font(12, True)

    for r in range(max_rows):
        y = origin_y + header_h + r * row_h
        draw.rectangle([0, y, header_w, y + row_h], fill="#f5f6f7", outline="#e0e0e0")
        draw.text((10, y + 6), str(r + 1), fill="#5f6368", font=try_font(10, True))
        x = header_w
        for c in range(max_cols):
            key = f"{r},{c}"
            cell = cells.get(key)
            bg = "#ffffff"
            fg = "#202124"
            text = ""
            use_bold = False
            align = "left"
            if cell:
                st = cell.get("style", {})
                bg = st.get("bg") or "#ffffff"
                if not bg:
                    bg = "#ffffff"
                raw = cell.get("raw", "")
                text = PREVIEW_OVERRIDES.get((sheet["name"], key), format_display(raw, st))
                if text.startswith("=HYPERLINK"):
                    text = PREVIEW_OVERRIDES.get((sheet["name"], key), "Link")
                elif text.startswith("="):
                    text = PREVIEW_OVERRIDES.get((sheet["name"], key), "…")
                fg = st.get("color") or "#202124"
                if "HYPERLINK" in raw or text in ("Open", "Open repo", "Ops runbook", "Link"):
                    fg = "#1155cc"
                use_bold = bool(st.get("bold"))
                align = st.get("align") or "left"
            draw.rectangle([x, y, x + widths[c], y + row_h], fill=bg, outline="#e2e4e8")
            if text:
                fnt = font_bold if use_bold else font
                # truncate
                display = text
                while fnt.getlength(display) > widths[c] - 10 and len(display) > 1:
                    display = display[:-2] + "…"
                tw = fnt.getlength(display)
                if align == "right":
                    tx = x + widths[c] - 6 - tw
                elif align == "center":
                    tx = x + (widths[c] - tw) / 2
                else:
                    tx = x + 5
                draw.text((tx, y + 6), display, fill=fg, font=fnt)
            x += widths[c]

    # selection hint on A1 area
    img.save(out, "PNG", optimize=True)


def main():
    SAMPLES.mkdir(parents=True, exist_ok=True)
    PREVIEWS.mkdir(parents=True, exist_ok=True)
    wb = build_workbook()

    json_path = SAMPLES / "ops_hq_complex_sample.gridia.json"
    json_path.write_text(json.dumps(wb, indent=2), encoding="utf-8")
    print("wrote", json_path)

    xlsx_path = SAMPLES / "ops_hq_complex_sample.xlsx"
    write_xlsx(xlsx_path, wb)
    print("wrote", xlsx_path)

    for s in wb["sheets"]:
        out = PREVIEWS / f"{s['name'].lower()}.png"
        max_cols = 8 if s["name"] == "Deals" else 7
        max_rows = 18 if s["name"] in ("Budget", "Deals") else 16
        render_sheet_png(s, out, max_rows=max_rows, max_cols=max_cols)
        print("wrote", out)

    # collage hero
    images = [Image.open(PREVIEWS / f"{n}.png") for n in ("dashboard", "budget", "deals", "hiring")]
    pad = 16
    # 2x2 grid
    w = max(images[0].width, images[1].width) + max(images[2].width, images[3].width) + pad * 3
    # normalize heights per row
    row1_h = max(images[0].height, images[1].height)
    row2_h = max(images[2].height, images[3].height)
    h = row1_h + row2_h + pad * 3
    hero = Image.new("RGB", (w, h), "#e8eee9")
    positions = [
        (pad, pad),
        (pad * 2 + images[0].width, pad),
        (pad, pad * 2 + row1_h),
        (pad * 2 + images[2].width, pad * 2 + row1_h),
    ]
    for im, pos in zip(images, positions):
        hero.paste(im, pos)
    hero_path = PREVIEWS / "hero_collage.png"
    hero.save(hero_path, "PNG", optimize=True)
    print("wrote", hero_path)


if __name__ == "__main__":
    main()
