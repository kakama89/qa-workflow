#!/usr/bin/env python3
"""Read-only inspection and validation for test-case XLSX workbooks.

Works with any team template: the Test Case sheet, header row and columns
are detected from header aliases (English and Vietnamese), and every guess
can be overridden on the command line. Neither command modifies a workbook.

Usage:
  python validate_workbook.py profile  WORKBOOK [options]
  python validate_workbook.py profile  --profile NAME
  python validate_workbook.py validate WORKBOOK [--template TEMPLATE] [options]
  python validate_workbook.py validate WORKBOOK --profile NAME

--profile loads a tester profile saved with qa_profiles.py (template and
layout); options given explicitly override the profile.

On Windows use "py -3" (or "python"). Run with -h for all options.

Exit codes: 0 = no errors, 1 = errors found (validate only), 2 = usage or
file error.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.utils.cell import range_boundaries
except ImportError:
    sys.stderr.write(
        "ERROR: openpyxl is required. Install it with: python -m pip install openpyxl\n"
    )
    raise SystemExit(2)

HEADER_ALIASES = {
    "id": ["test case id", "testcase id", "test case no", "tc id", "tc_id", "tc no", "id",
           "test id", "case id", "mã test case", "mã testcase", "mã tc", "ma tc", "stt"],
    "title": ["test case title", "testcase title", "test case name", "title", "test scenario",
              "scenario", "summary", "tên test case", "tiêu đề", "kịch bản", "mô tả"],
    "precondition": ["pre-conditions", "pre-condition", "preconditions", "precondition",
                     "pre condition", "điều kiện tiên quyết", "tiền điều kiện", "điều kiện"],
    "steps": ["test steps", "test step", "steps", "step", "procedure", "các bước thực hiện",
              "bước thực hiện", "các bước"],
    "data": ["test datas", "test data", "data", "input data", "dữ liệu test",
             "dữ liệu kiểm thử", "dữ liệu"],
    "expected": ["expected result", "expected results", "expected output", "expected",
                 "kết quả mong muốn", "kết quả mong đợi", "kết quả dự kiến"],
    "environment": ["test environment", "environment", "env", "môi trường"],
    "priority": ["priority", "mức độ ưu tiên", "độ ưu tiên", "ưu tiên"],
    "automation": ["automation status", "automation", "automated", "auto"],
    "result": ["test results", "test result", "result", "execution result", "execution status",
               "actual status", "kết quả kiểm thử", "kết quả thực tế", "kết quả"],
    "bug": ["bugid", "bug id", "bug", "defect id", "defect", "jira id"],
    "notes": ["notes", "note", "remarks", "remark", "comment", "comments", "ghi chú"],
    "requirement": ["requirement id", "requirement", "req id", "req", "mã yêu cầu", "yêu cầu"],
    "ac": ["acceptance criteria id", "acceptance criteria", "ac id", "ac"],
    "review": ["review status", "trạng thái review"],
}

ENUM_ROLES = ("environment", "priority", "automation", "result")
NOT_EXECUTED = {"", "untested", "not run", "not executed", "chưa test", "chưa thực hiện"}

SECRET_KEY_RE = re.compile(
    r"(?i)(?<![A-Za-z0-9])([A-Za-z0-9_-]*(?:token|api[_-]?key|secret|password|passwd|pwd|"
    r"cookie|authorization|session[_-]?id|private[_-]?key)[A-Za-z0-9_-]*)"
    r"[\"']?\s*[:=]\s*[\"']?([^\r\n\"']+)"
)
SECRET_VALUE_PATTERNS = [
    ("JWT", re.compile(r"\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}")),
    ("Bearer token", re.compile(r"(?i)\bbearer\s+[A-Za-z0-9._~+/=-]{16,}")),
    ("Basic auth", re.compile(r"(?i)\bbasic\s+[A-Za-z0-9+/=]{16,}")),
    ("Private key", re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----")),
    ("AWS access key", re.compile(r"\bAKIA[0-9A-Z]{16}\b")),
]
PLACEHOLDER_RE = re.compile(
    r"(?i)^(<[^>]*>|\{\{.*\}\}|\$\{.*\}|\[[^\]]*\]|x{3,}|\*{3,}|\.{3,}|redacted|masked|"
    r"placeholder|hidden|n/?a|null|none|tbd|empty|blank)$"
)
RANGE_RE = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?"
    r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)"
)


class UsageError(Exception):
    pass


def norm(value) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFC", str(value)).strip().lower()
    return re.sub(r"\s+", " ", text)


ALIAS_TO_ROLE = {norm(alias): role for role, aliases in HEADER_ALIASES.items() for alias in aliases}
# Earlier aliases are more specific: "Mã TC" beats "STT" when both columns exist.
ALIAS_RANK = {norm(alias): i for aliases in HEADER_ALIASES.values() for i, alias in enumerate(aliases)}


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def formula_text(value):
    """Return the formula string for formula cells (incl. array formulas), else None."""
    if isinstance(value, str) and value.startswith("="):
        return value
    text = getattr(value, "text", None)  # openpyxl ArrayFormula
    if isinstance(text, str):
        return text if text.startswith("=") else "=" + text
    return None


def compress_rows(rows, limit=12) -> str:
    """Render [15,16,17,21] as '15-17, 21'."""
    spans = []
    for r in sorted(set(rows)):
        if spans and r == spans[-1][1] + 1:
            spans[-1][1] = r
        else:
            spans.append([r, r])
    parts = [str(a) if a == b else f"{a}-{b}" for a, b in spans]
    if len(parts) > limit:
        parts = parts[:limit] + [f"... (+{len(parts) - limit} more)"]
    return ", ".join(parts)


class Findings:
    def __init__(self):
        self.items = []

    def add(self, level, code, message, location=""):
        self.items.append({"level": level, "code": code, "location": location, "message": message})

    def count(self, level) -> int:
        return sum(1 for f in self.items if f["level"] == level)


class Sheet:
    """Worksheet wrapper that resolves merged cells to their anchor value."""

    def __init__(self, ws, ws_values=None):
        self.ws = ws
        self.ws_values = ws_values
        self.anchor = {}
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        self.anchor[(r, c)] = (rng.min_row, rng.min_col)

    def value(self, r, c):
        r, c = self.anchor.get((r, c), (r, c))
        return self.ws.cell(r, c).value

    def cached(self, r, c):
        if self.ws_values is None:
            return None
        r, c = self.anchor.get((r, c), (r, c))
        return self.ws_values.cell(r, c).value


def roles_in_row(ws, row, max_cols=100):
    best = {}
    for c in range(1, min(ws.max_column, max_cols) + 1):
        key = norm(ws.cell(row, c).value)
        role = ALIAS_TO_ROLE.get(key)
        if role and (role not in best or ALIAS_RANK[key] < best[role][0]):
            best[role] = (ALIAS_RANK[key], c)
    return {role: c for role, (_, c) in best.items()}


def header_score(roles) -> int:
    return len(roles) + 2 * ("steps" in roles) + 2 * ("expected" in roles) + ("id" in roles)


def find_header(ws, max_rows=60):
    best = None
    for r in range(1, min(ws.max_row, max_rows) + 1):
        roles = roles_in_row(ws, r)
        if len(roles) >= 3 and ("steps" in roles or "expected" in roles):
            if best is None or header_score(roles) > header_score(best[1]):
                best = (r, roles)
    return best


def apply_column_overrides(ws, header_row, cols, overrides):
    for item in overrides or []:
        if "=" not in item:
            raise UsageError(f"--col expects ROLE=COLUMN, got {item!r}")
        role, target = (part.strip() for part in item.split("=", 1))
        if role not in HEADER_ALIASES:
            raise UsageError(f"Unknown role {role!r}. Known roles: {', '.join(HEADER_ALIASES)}")
        col = None
        for c in range(1, ws.max_column + 1):
            if norm(ws.cell(header_row, c).value) == norm(target):
                col = c
                break
        if col is None and re.fullmatch(r"[A-Za-z]{1,3}", target):
            col = column_index_from_string(target.upper())
        if col is None:
            raise UsageError(f"--col {item!r}: no header {target!r} in row {header_row}")
        for other in [r for r, c in cols.items() if c == col and r != role]:
            del cols[other]  # an explicit mapping replaces the auto-detected role of that column
        cols[role] = col


def locate(wb, wb_values, args):
    """Find the Test Case sheet, header row, column roles and data rows."""
    if args.sheet:
        if args.sheet not in wb.sheetnames:
            raise UsageError(f"Sheet {args.sheet!r} not found. Sheets: {wb.sheetnames}")
        candidates = [wb[args.sheet]]
    else:
        candidates = wb.worksheets
    best = None
    for ws in candidates:
        if args.header_row:
            hit = (args.header_row, roles_in_row(ws, args.header_row))
        else:
            hit = find_header(ws)
        if not hit:
            continue
        key = (header_score(hit[1]), "test" in norm(ws.title))
        if best is None or key > best[0]:
            best = (key, ws, hit[0], hit[1])
    if best is None:
        return None
    _, ws, header_row, cols = best
    apply_column_overrides(ws, header_row, cols, args.col)
    sheet = Sheet(ws, wb_values[ws.title] if wb_values is not None else None)

    data_start = args.data_start_row or detect_data_start(sheet, header_row, cols)
    content_roles = [r for r in ("steps", "expected") if r in cols] or [r for r in cols if r != "id"]
    other_cols = [c for role, c in cols.items() if role != "id"]
    last = 0
    for r in range(data_start, ws.max_row + 1):
        if any(not is_blank(ws.cell(r, c).value) for c in other_cols):
            last = r
    id_col = cols.get("id")
    if id_col:
        for r in range(data_start, ws.max_row + 1):
            v = ws.cell(r, id_col).value
            if not is_blank(v) and not formula_text(v):
                last = max(last, r)
    tc_rows, section_rows, id_only_rows = [], [], []
    for r in range(data_start, last + 1):
        if any(not is_blank(sheet.value(r, cols[role])) for role in content_roles
               if (r, cols[role]) not in sheet.anchor):
            tc_rows.append(r)
        elif any(not is_blank(ws.cell(r, c).value) for c in other_cols):
            section_rows.append(r)
        elif id_col and not is_blank(ws.cell(r, id_col).value) and not formula_text(ws.cell(r, id_col).value):
            id_only_rows.append(r)
    return {
        "sheet": sheet,
        "header_row": header_row,
        "cols": cols,
        "data_start": data_start,
        "tc_rows": tc_rows,
        "section_rows": section_rows,
        "id_only_rows": id_only_rows,
    }


def detect_data_start(sheet, header_row, cols) -> int:
    """Skip instruction/guide rows that templates often place under the header."""
    r = header_row + 1
    id_col = cols.get("id")
    for _ in range(3):
        if r > sheet.ws.max_row or not id_col:
            break
        v = sheet.ws.cell(r, id_col).value
        if isinstance(v, str) and not formula_text(v) and (len(v) > 40 or "\n" in v):
            r += 1
            continue
        break
    return r


def col_label(layout, role) -> str:
    c = layout["cols"][role]
    header = layout["sheet"].ws.cell(layout["header_row"], c).value
    return f"{get_column_letter(c)} ({str(header).strip()})"


def check_ids(layout, findings, info):
    if "id" not in layout["cols"]:
        findings.add("warning", "id-column-missing", "No Test Case ID column detected; use --col id=<column>.")
        return
    sheet, title = layout["sheet"], layout["sheet"].ws.title
    id_col = layout["cols"]["id"]
    letter = get_column_letter(id_col)
    missing, formulas, ids, uncached = [], {}, {}, 0
    for r in layout["tc_rows"]:
        raw = sheet.value(r, id_col)
        if is_blank(raw):
            missing.append(r)
            continue
        f = formula_text(raw)
        value = sheet.cached(r, id_col) if f else raw
        if f:
            formulas[r] = f
        if is_blank(value):
            uncached += bool(f)
            continue
        ids.setdefault(norm(value), []).append((r, str(value).strip()))
    info["id_mode"] = "-" if not layout["tc_rows"] else "formula" if formulas and len(formulas) == len(layout["tc_rows"]) - len(missing) \
        else ("mixed" if formulas else "literal")
    if formulas:
        info["id_formula_example"] = formulas[min(formulas)]
    if missing:
        findings.add("error", "id-missing", f"Test case rows without an ID: {compress_rows(missing)}",
                     f"{title}!{letter}")
    if layout["id_only_rows"]:
        findings.add("error", "id-without-content",
                     f"Rows with an ID but no steps/expected result: {compress_rows(layout['id_only_rows'])}",
                     f"{title}!{letter}")
    for key, hits in ids.items():
        if len(hits) > 1:
            findings.add("error", "id-duplicate",
                         f"Duplicate Test Case ID {hits[0][1]!r} in rows {compress_rows([h[0] for h in hits])}",
                         f"{title}!{letter}")
    if uncached:
        findings.add("info", "id-not-evaluated",
                     f"{uncached} ID formulas have no cached value, so duplicate IDs cannot be checked. "
                     "Open and save the file in Excel/LibreOffice, then validate again.",
                     f"{title}!{letter}")
    if len(formulas) > 1:
        first = min(formulas)
        groups = {}
        for r, f in formulas.items():
            try:
                base = Translator(f, origin=f"{letter}{r}").translate_formula(f"{letter}{first}")
            except Exception:
                base = f
            groups.setdefault(base, []).append(r)
        if len(groups) > 1:
            main = max(groups.values(), key=len)
            odd = [r for rows in groups.values() if rows is not main for r in rows]
            findings.add("warning", "id-formula-inconsistent",
                         f"ID formula differs from the majority pattern in rows {compress_rows(odd)}",
                         f"{title}!{letter}")


def check_required(layout, findings, required):
    sheet, title = layout["sheet"], layout["sheet"].ws.title
    for role in required:
        if role == "id":
            continue
        if role not in layout["cols"]:
            findings.add("warning", "required-column-missing", f"Required column {role!r} not found in header.")
            continue
        c = layout["cols"][role]
        rows = [r for r in layout["tc_rows"] if is_blank(sheet.value(r, c))]
        if rows:
            findings.add("error", "required-empty",
                         f"{col_label(layout, role)} is empty in rows {compress_rows(rows)}. "
                         "Fill it or write an explicit flag such as TBD.",
                         f"{title}!{get_column_letter(c)}")


def list_items(wb, dv):
    """Return the allowed values of a list data validation, or None if not resolvable."""
    f = (dv.formula1 or "").strip()
    if f.startswith('"') and f.endswith('"'):
        return [s.strip() for s in f[1:-1].split(",") if s.strip()]
    m = re.fullmatch(r"=?(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", f)
    if m:
        name = (m.group(1) or m.group(2)).replace("''", "'")
        if name in wb.sheetnames:
            ws = wb[name]
            min_col, min_row, max_col, max_row = range_boundaries(f"{m.group(3)}{m.group(4)}:{m.group(5)}{m.group(6)}")
            return [str(ws.cell(r, c).value).strip() for r in range(min_row, max_row + 1)
                    for c in range(min_col, max_col + 1) if not is_blank(ws.cell(r, c).value)]
    return None


def check_validations(wb, layout, findings, info):
    sheet, title = layout["sheet"], layout["sheet"].ws.title
    tc_rows = layout["tc_rows"]
    col_role = {c: role for role, c in layout["cols"].items()}
    covered = set()
    info["data_validations"] = []
    for dv in sheet.ws.data_validations.dataValidation:
        items = list_items(wb, dv) if dv.type == "list" else None
        info["data_validations"].append({"range": str(dv.sqref), "type": dv.type,
                                         "formula": dv.formula1, "items": items})
        for rng in str(dv.sqref).split():
            min_col, min_row, max_col, max_row = range_boundaries(rng)
            min_row, max_row = min_row or 1, max_row or 1048576
            for c in range(min_col or 1, (max_col or 16384) + 1):
                role = col_role.get(c)
                if not role or max_row <= layout["header_row"]:
                    continue
                covered.add(role)
                if tc_rows and max_row < tc_rows[-1]:
                    findings.add("error", "validation-range-short",
                                 f"Dropdown/validation on {col_label(layout, role)} covers {rng} "
                                 f"but test cases continue to row {tc_rows[-1]}.", f"{title}!{rng}")
                if tc_rows and min_row > tc_rows[0]:
                    findings.add("warning", "validation-range-late",
                                 f"Validation on {col_label(layout, role)} starts at row {min_row}, "
                                 f"after the first test case row {tc_rows[0]}.", f"{title}!{rng}")
                if items:
                    allowed = {norm(i) for i in items}
                    bad = [r for r in tc_rows if min_row <= r <= max_row
                           and not is_blank(sheet.value(r, c)) and not formula_text(sheet.value(r, c))
                           and norm(sheet.value(r, c)) not in allowed]
                    if bad:
                        findings.add("warning", "validation-value",
                                     f"{col_label(layout, role)} has values outside the dropdown list "
                                     f"{items} in rows {compress_rows(bad)}.", f"{title}!{get_column_letter(c)}")
    for role in ENUM_ROLES:
        if role in layout["cols"] and role not in covered:
            findings.add("info", "validation-missing", f"{col_label(layout, role)} has no dropdown/validation.")


def iter_formulas(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                f = formula_text(cell.value)
                if f:
                    yield ws, cell, f


def check_formulas(wb, layout, findings, info):
    tc_ws = layout["sheet"].ws
    tc_rows = layout["tc_rows"]
    cols = set(layout["cols"].values())
    id_col = layout["cols"].get("id")
    first_data = tc_rows[0] if tc_rows else layout["data_start"]
    refs = []
    for ws, cell, f in iter_formulas(wb):
        loc = f"{ws.title}!{cell.coordinate}"
        if "#REF!" in f.upper():
            findings.add("error", "formula-ref-error", "Formula contains #REF! (broken reference).", loc)
        if ws is tc_ws and cell.row >= layout["data_start"]:
            continue  # per-row formulas such as ID generators
        for m in RANGE_RE.finditer(f):
            ref_sheet = (m.group(1) or "").replace("''", "'") or m.group(2) or ws.title
            if ref_sheet != tc_ws.title:
                continue
            c1, r1 = column_index_from_string(m.group(3)), int(m.group(4))
            c2, r2 = column_index_from_string(m.group(5)), int(m.group(6))
            if c1 != c2 or c1 not in cols or not (layout["header_row"] < r1 <= first_data):
                continue
            refs.append({"cell": loc, "formula": f, "range": m.group(0), "end": r2})
            if tc_rows and r2 < tc_rows[-1]:
                findings.add("error", "formula-range-short",
                             f"Range {m.group(0)} stops at row {r2} but test cases continue to row {tc_rows[-1]}.",
                             loc)
            if c1 == id_col and "COUNTA(" in f.upper() and info.get("id_mode") in ("formula", "mixed"):
                findings.add("warning", "counta-formula-ids",
                             f"COUNTA over formula-generated IDs also counts formulas returning \"\"; "
                             f"consider COUNTIF({m.group(0)},\"?*\").", loc)
    ends = sorted({r["end"] for r in refs})
    if len(ends) > 1:
        findings.add("warning", "formula-range-inconsistent",
                     f"Summary formulas over {tc_ws.title} use different end rows {ends}; "
                     "align them so every metric counts the same rows.")
    info["summary_formulas"] = [{"cell": r["cell"], "formula": r["formula"]} for r in refs]

    names = wb.defined_names
    pairs = list(names.items()) if hasattr(names, "items") else [(d.name, d) for d in getattr(names, "definedName", [])]
    workbook_level = len(pairs)
    for ws in wb.worksheets:
        scoped = getattr(ws, "defined_names", None)
        if scoped and hasattr(scoped, "items"):
            pairs += [(f"{ws.title}!{n}", d) for n, d in scoped.items()]
    broken = [name for name, d in pairs if "#REF!" in str(getattr(d, "attr_text", "") or "").upper()]
    info["defined_names"] = len(pairs)
    if broken:
        sample = ", ".join(broken[:8]) + (", ..." if len(broken) > 8 else "")
        findings.add("warning", "defined-name-broken",
                     f"{len(broken)} of {len(pairs)} defined names ({workbook_level} workbook-level, "
                     f"{len(pairs) - workbook_level} sheet-level) point to #REF! ({sample}).")


def check_execution(layout, findings, info, stage):
    sheet, title = layout["sheet"], layout["sheet"].ws.title
    counts = {}
    if "result" in layout["cols"]:
        c = layout["cols"]["result"]
        executed = []
        for r in layout["tc_rows"]:
            v = sheet.value(r, c)
            if formula_text(v):
                continue
            key = "" if is_blank(v) else str(v).strip()
            counts[key or "(blank)"] = counts.get(key or "(blank)", 0) + 1
            if norm(key) not in NOT_EXECUTED:
                executed.append(r)
        if executed and stage == "design":
            findings.add("error", "execution-result-in-design",
                         f"{col_label(layout, 'result')} has execution results in rows {compress_rows(executed)}. "
                         "A design export must leave results blank or Untested.",
                         f"{title}!{get_column_letter(c)}")
    info["result_values"] = counts
    if "bug" in layout["cols"] and stage == "design":
        c = layout["cols"]["bug"]
        rows = [r for r in layout["tc_rows"] if not is_blank(sheet.value(r, c))]
        if rows:
            findings.add("warning", "bug-id-in-design",
                         f"{col_label(layout, 'bug')} is filled in rows {compress_rows(rows)} during test design.",
                         f"{title}!{get_column_letter(c)}")


def secret_kinds(text):
    kinds = [kind for kind, rx in SECRET_VALUE_PATTERNS if rx.search(text)]
    for key, value in SECRET_KEY_RE.findall(text):
        value = value.strip().rstrip("\\,;}").strip()
        if not value or PLACEHOLDER_RE.match(value):
            continue
        min_len = 3 if re.search(r"(?i)pass|pwd", key) else 8
        if len(value) >= min_len:
            kinds.append(f"{key} value")
    return kinds


def check_secrets(wb, findings):
    hits = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    kinds = secret_kinds(cell.value)
                    if kinds:
                        loc = f"{ws.title}!{cell.coordinate}"
                        hits.append(loc)
                        # Never print the value itself.
                        findings.add("error", "possible-secret",
                                     f"Possible credential ({', '.join(sorted(set(kinds)))}); value withheld. "
                                     "Replace with a placeholder such as <TOKEN>.", loc)
    return hits


def header_texts(layout):
    ws, r = layout["sheet"].ws, layout["header_row"]
    return [str(ws.cell(r, c).value).strip() for c in range(1, ws.max_column + 1)
            if not is_blank(ws.cell(r, c).value)]


def compare_template(wb, layout, template_path, output_path, args, findings):
    try:
        if Path(template_path).resolve() == Path(output_path).resolve():
            findings.add("error", "output-overwrites-template", "Output file is the template itself; export to a new file.")
            return
    except OSError:
        pass
    twb = open_workbook(template_path, data_only=False)
    t_sheets, o_sheets = twb.sheetnames, wb.sheetnames
    for name in t_sheets:
        if name not in o_sheets:
            findings.add("warning", "template-sheet-removed", f"Sheet {name!r} from the template is missing. "
                         "Confirm it was approved in the Template Change Review.")
    for name in o_sheets:
        if name not in t_sheets:
            findings.add("info", "template-sheet-added", f"Sheet {name!r} was added (not in template).")
    t_args = argparse.Namespace(sheet=None, header_row=args.header_row, col=args.col,
                                data_start_row=args.data_start_row)
    if layout["sheet"].ws.title in t_sheets:
        t_args.sheet = layout["sheet"].ws.title
    t_layout = locate(twb, None, t_args)
    if not t_layout:
        findings.add("info", "template-header-unknown", "Could not locate the Test Case header in the template.")
        return
    t_head, o_head = header_texts(t_layout), header_texts(layout)
    removed = [h for h in t_head if norm(h) not in {norm(x) for x in o_head}]
    added = [h for h in o_head if norm(h) not in {norm(x) for x in t_head}]
    if removed:
        findings.add("warning", "template-column-removed", f"Template columns missing in output: {removed}.")
    if added:
        findings.add("info", "template-column-added", f"Columns added compared to template: {added}.")

    def dv_signature(wb_, lay):
        ws = lay["sheet"].ws
        sig = set()
        for dv in ws.data_validations.dataValidation:
            for rng in str(dv.sqref).split():
                min_col, _, max_col, _ = range_boundaries(rng)
                for c in range(min_col or 1, (max_col or min_col or 1) + 1):
                    head = norm(ws.cell(lay["header_row"], c).value)
                    if head:
                        sig.add((head, tuple(norm(i) for i in (list_items(wb_, dv) or [dv.formula1]))))
        return sig

    for head, items in sorted(dv_signature(twb, t_layout) - dv_signature(wb, layout)):
        findings.add("warning", "template-validation-changed",
                     f"Dropdown on column {head!r} {list(items)} from the template is missing or changed.")


def open_workbook(path, data_only):
    p = Path(path)
    if not p.is_file():
        raise UsageError(f"File not found: {p}")
    if p.suffix.lower() == ".xls":
        raise UsageError("Legacy .xls is not supported; save the file as .xlsx first.")
    try:
        return load_workbook(p, data_only=data_only)
    except Exception as exc:
        raise UsageError(f"Cannot open workbook {p.name}: {type(exc).__name__}: {exc}")


def run(args) -> int:
    wb = open_workbook(args.workbook, data_only=False)
    wb_values = open_workbook(args.workbook, data_only=True)
    findings = Findings()
    info = {
        "workbook": str(Path(args.workbook)),
        "sheets": [{"name": ws.title, "state": ws.sheet_state, "dimensions": ws.dimensions}
                   for ws in wb.worksheets],
    }
    layout = locate(wb, wb_values, args)
    if layout is None:
        findings.add("error", "testcase-sheet-not-found",
                     "Could not find a Test Case header row (needs >= 3 known headers incl. steps or expected). "
                     "Use --sheet, --header-row and --col ROLE=COLUMN.")
    else:
        ws = layout["sheet"].ws
        info["testcase_sheet"] = ws.title
        info["header_row"] = layout["header_row"]
        info["data_start_row"] = layout["data_start"]
        info["columns"] = {role: col_label(layout, role) for role in sorted(layout["cols"], key=layout["cols"].get)}
        mapped = set(layout["cols"].values())
        info["unmapped_headers"] = [f"{get_column_letter(c)} ({str(ws.cell(layout['header_row'], c).value).strip()})"
                                    for c in range(1, ws.max_column + 1)
                                    if c not in mapped and not is_blank(ws.cell(layout["header_row"], c).value)]
        info["testcase_rows"] = len(layout["tc_rows"])
        info["testcase_row_span"] = compress_rows(layout["tc_rows"])
        info["section_rows"] = compress_rows(layout["section_rows"])
        if ws.max_column > max(mapped) + 50:
            findings.add("info", "sheet-dimension-bloat",
                         f"{ws.title} reports {ws.max_column} columns ({ws.dimensions}); "
                         "stray formatting beyond the table may slow tools down.")
        check_ids(layout, findings, info)
        required = [r.strip() for r in (args.required or "id,steps,expected").split(",") if r.strip()]
        check_required(layout, findings, required)
        check_validations(wb, layout, findings, info)
        check_formulas(wb, layout, findings, info)
        check_execution(layout, findings, info, args.stage)
        if args.command == "validate" and args.template:
            compare_template(wb, layout, args.template, args.workbook, args, findings)
    info["possible_secret_cells"] = check_secrets(wb, findings)

    if args.json:
        print(json.dumps({"info": info, "findings": findings.items}, ensure_ascii=False, indent=2, default=str))
    else:
        print_report(args.command, info, findings)
    if args.command == "profile":
        return 2 if layout is None else 0
    return 1 if findings.count("error") else 0


def print_report(command, info, findings):
    print(f"Workbook: {info['workbook']}")
    print("Sheets: " + ", ".join(f"{s['name']} [{s['state']}, {s['dimensions']}]" for s in info["sheets"]))
    if "testcase_sheet" in info:
        print(f"Test Case sheet: {info['testcase_sheet']} | header row {info['header_row']} | "
              f"data from row {info['data_start_row']}")
        print("Columns: " + "; ".join(f"{k}={v}" for k, v in info["columns"].items()))
        if info["unmapped_headers"]:
            print("Unmapped headers: " + "; ".join(info["unmapped_headers"]))
        print(f"Test case rows: {info['testcase_rows']} ({info['testcase_row_span'] or '-'}); "
              f"section rows: {info['section_rows'] or '-'}")
        print(f"ID mode: {info.get('id_mode', '-')}" +
              (f" | example: {info['id_formula_example']}" if info.get("id_formula_example") else ""))
        if command == "profile":
            print(f"Result values: {info.get('result_values') or '-'}")
            for dv in info.get("data_validations", []):
                print(f"  validation {dv['range']}: {dv['items'] if dv['items'] is not None else dv['formula']}")
            for sf in info.get("summary_formulas", []):
                print(f"  summary {sf['cell']}: {sf['formula']}")
    print("\nFindings:")
    if not findings.items:
        print("- None detected by these checks.")
    order = {"error": 0, "warning": 1, "info": 2}
    for f in sorted(findings.items, key=lambda x: order[x["level"]]):
        loc = f" {f['location']}" if f["location"] else ""
        print(f"- {f['level'].upper()} [{f['code']}]{loc}: {f['message']}")
    print(f"\nSummary: {findings.count('error')} error(s), {findings.count('warning')} warning(s), "
          f"{findings.count('info')} info. Secret matches may be false positives; review them manually.")
    print("Note: formulas are not recalculated; IDs and metrics are checked structurally.")


def build_parser():
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = p.add_subparsers(dest="command", required=True)
    for name, help_text in (("profile", "describe a template/workbook layout (always exit 0 when readable)"),
                            ("validate", "validate an exported workbook (exit 1 on errors)")):
        sp = sub.add_parser(name, help=help_text)
        sp.add_argument("workbook", nargs="?" if name == "profile" else None,
                        help="path to .xlsx/.xlsm" + (" (default: the profile's template)" if name == "profile" else ""))
        sp.add_argument("--profile", metavar="NAME",
                        help="tester profile from qa_profiles.py ('default' = default profile)")
        sp.add_argument("--home", help="profile folder (see qa_profiles.py)")
        sp.add_argument("--project-dir", help="project root for project-scope profiles")
        sp.add_argument("--sheet", help="Test Case sheet name (default: auto-detect)")
        sp.add_argument("--header-row", type=int, help="header row number (default: auto-detect)")
        sp.add_argument("--data-start-row", type=int, help="first data row (default: auto, skips guide rows)")
        sp.add_argument("--col", action="append", metavar="ROLE=COLUMN",
                        help=f"map a role to a column letter or header text; roles: {', '.join(HEADER_ALIASES)}")
        sp.add_argument("--required",
                        help="comma-separated roles that must be filled on every test case row "
                             "(default: id,steps,expected)")
        sp.add_argument("--stage", choices=("design", "execution"), default="design",
                        help="design (default): execution results must be blank/Untested")
        sp.add_argument("--json", action="store_true", help="machine-readable output")
        if name == "validate":
            sp.add_argument("--template", help="original template, to report sheet/column/dropdown changes "
                                               "(default: the profile's template)")
    return p


def resolve_profile(args):
    """Apply --profile settings to args; explicit options win."""
    if not getattr(args, "profile", None):
        return
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from qa_profiles import apply_profile, load_profile  # sibling module; imported lazily

    profile = load_profile(args.profile, args.home, args.project_dir)
    apply_profile(args, profile)
    if args.command == "profile" and not args.workbook:
        args.workbook = profile["template_path"]
    if args.command == "validate" and not args.template:
        args.template = profile["template_path"]


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    argv = list(sys.argv[1:] if argv is None else argv)
    # Backward compatible: "validate_workbook.py file.xlsx" means validate.
    if argv and argv[0] not in ("profile", "validate", "-h", "--help"):
        argv.insert(0, "validate")
    args = build_parser().parse_args(argv)
    try:
        resolve_profile(args)
        if not args.workbook:
            raise UsageError("Pass a workbook path or --profile NAME.")
        return run(args)
    except UsageError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
