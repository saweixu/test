# -*- coding: utf-8 -*-
# app_eorigin_t1_online.py
# Streamlit E-origin processor: invoices + T1 PDF -> checks -> T1 match -> modified invoices ZIP.

import io
import csv
import re
import time
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Callable, List, Tuple
from xml.etree import ElementTree as ET

import pandas as pd
import pdfplumber
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries


APP_DIR = Path(__file__).resolve().parent
APP_TITLE = "E-Origin T1 Invoice Processor"
LOGO_CANDIDATES = (
    APP_DIR / "logo.png",
    APP_DIR.parent / "logo.png",
    Path.cwd() / "logo.png",
)

MIN_T1_SCORE = 0.75
DESTINATION = "Rotterdam"
VAT_MODE = "SEA MT"
VAT_BASE_AMOUNT = Decimal("1650")
VAT_FIRST_ADDITION = Decimal("60")
VAT_OTHER_ADDITION = Decimal("40")
VAT_CSV_NAME = "VAT_charge.csv"
EORIGIN_API_BASE_URL = "https://athinalogistics.eorigin.eu/api/v1/external"
EORIGIN_UPLOAD_TIMEOUT = 90

COUNTRY_CODE_RE = re.compile(r"^[A-Z]{2}$")
OOXML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OOXML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = {"main": OOXML_NS}

ET.register_namespace("", OOXML_NS)
ET.register_namespace("r", OOXML_REL_NS)

NORMAL_INVOICE_NAME_RE = re.compile(r"^[A-Z0-9]+-\d+$", re.I)
HBL_RE = re.compile(r"(?:\(\s*)?HBL\s*(\d+)\s*(?:\))?", re.I)
CONTAINER_RE = re.compile(r"\b([A-Z]{4}\d{7})\b", re.I)


@dataclass
class ParseResult:
    name: str
    rows: List[Tuple[int, int]]
    score: float
    notes: str


def find_logo_path():
    for logo_path in LOGO_CANDIDATES:
        if logo_path.exists():
            return logo_path
    return None


def configure_page():
    logo_path = find_logo_path()
    page_config = {
        "page_title": "Athina Logistics Tool",
        "layout": "wide",
    }
    if logo_path:
        page_config["page_icon"] = str(logo_path)

    st.set_page_config(**page_config)

    if logo_path:
        st.sidebar.image(str(logo_path), width=200)

    st.sidebar.markdown("### Athina Logistics")
    st.sidebar.caption("Global Access")


def natural_key(path_or_name):
    stem = Path(str(path_or_name)).stem
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", stem)]


def sanitize_filename_part(value):
    text = str(value or "").strip()
    text = re.sub(r'[<>:"/\\|?*]+', "-", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[-\s]+$", "", text)
    text = re.sub(r"^[-\s]+", "", text)
    return text


def normalize_uploaded_invoice_name(file_name):
    path = Path(str(file_name))
    stem = path.stem.strip()
    suffix = path.suffix or ".xlsx"

    if NORMAL_INVOICE_NAME_RE.match(stem):
        return path.name, False, "Already normal"

    hbl_match = HBL_RE.search(stem)
    if not hbl_match:
        return path.name, False, "HBL number not found; original name kept"

    container_match = CONTAINER_RE.search(stem)
    if container_match:
        prefix = container_match.group(1).upper()
    else:
        prefix = stem.split(" - ")[0].strip()
        prefix = sanitize_filename_part(prefix)

    if not prefix:
        return path.name, False, "Container/prefix not found; original name kept"

    normalized = f"{prefix}-{int(hbl_match.group(1))}{suffix}"
    return normalized, normalized != path.name, f"Renamed from {path.name}"


def build_invoice_entries(invoice_files):
    entries = []
    seen = {}

    for uploaded_file in invoice_files:
        normalized_name, renamed, note = normalize_uploaded_invoice_name(uploaded_file.name)
        key = normalized_name.lower()
        if key in seen:
            raise ValueError(
                f"Two uploaded files become the same name after rename: {seen[key]} and {uploaded_file.name} -> {normalized_name}"
            )

        seen[key] = uploaded_file.name
        entries.append(
            {
                "name": normalized_name,
                "original_name": uploaded_file.name,
                "bytes": uploaded_file.getvalue(),
                "renamed": renamed,
                "rename_note": note,
            }
        )

    return sorted(entries, key=lambda item: natural_key(item["name"]))


def build_rename_frame(invoice_entries):
    return pd.DataFrame(
        [
            {
                "Original file": entry["original_name"],
                "Used as": entry["name"],
                "Renamed": "YES" if entry["renamed"] else "NO",
                "Note": entry["rename_note"],
            }
            for entry in invoice_entries
        ]
    )


def to_decimal(value):
    if value is None:
        return None
    text = str(value).replace(",", ".").strip()
    if not text:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def to_int(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(round(value))
    text = str(value).strip().replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        return int(round(float(text)))
    except ValueError:
        return None


def sheet_by_name_ci(wb, wanted):
    norm = wanted.strip().lower()
    for name in wb.sheetnames:
        if name.strip().lower() == norm:
            return wb[name]
    return None


def find_sum_row(ws, start_row=1, label_col="B"):
    if ws is None:
        return None
    pat = re.compile(r"^\s*SUM\s*[:\uFF1A]?\s*$", re.IGNORECASE)
    for row in range(start_row, ws.max_row + 1):
        value = ws[f"{label_col}{row}"].value
        if value is None:
            continue
        text = str(value).strip()
        if pat.match(text) or text.upper().startswith("SUM"):
            return row
    return None


def is_cell_in_merged(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return True
    return False


def merged_range_for_cell(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng
    return None


def get_merged_value(ws, cell_ref):
    cell = ws[cell_ref]
    if cell.value is not None:
        return cell.value

    row, col = coordinate_to_tuple(cell_ref)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None


def get_effective_cell_value(ws, row, col):
    value = ws.cell(row=row, column=col).value
    if value is not None:
        return value

    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None


def contains_chinese(text):
    return isinstance(text, str) and re.search(r"[\u4e00-\u9fff]", text) is not None


REF_RE = re.compile(
    r"^\s*=\s*(?:(?P<sheet>'[^']+'|[A-Za-z0-9 _.-]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)\s*$"
)
CELL_REF_RE = re.compile(
    r"^\s*(?:(?:'(?P<sheet_q>[^']+)'|(?P<sheet_u>[^!]+))!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)\s*$",
    re.IGNORECASE,
)
RANGE_REF_RE = re.compile(
    r"^\s*(?:(?:'(?P<sheet_q>[^']+)'|(?P<sheet_u>[^!]+))!)?(?P<start>\$?[A-Z]{1,3}\$?\d+):(?P<end>\$?[A-Z]{1,3}\$?\d+)\s*$",
    re.IGNORECASE,
)
CELL_TOKEN_RE = re.compile(
    r"(?:'[^']+'|[A-Za-z0-9_. -]+)!\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}\$?\d+",
    re.IGNORECASE,
)


def col_to_idx(col):
    number = 0
    for char in col:
        number = number * 26 + (ord(char) - 64)
    return number


def is_blank(value):
    return value is None or str(value).strip() == ""


def get_effective_cell(ws, row, col):
    if ws is None:
        return None

    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col)

    return ws.cell(row=row, column=col)


def sheet_pair(wb_values, wb_formulas, current_values, current_formula, sheet_name):
    if sheet_name:
        return sheet_by_name_ci(wb_values, sheet_name), sheet_by_name_ci(wb_formulas, sheet_name)
    return current_values, current_formula


def parse_cell_ref(ref):
    match = CELL_REF_RE.match(str(ref).strip())
    if not match:
        return None

    sheet_name = match.group("sheet_q") or match.group("sheet_u")
    return {
        "sheet": sheet_name.strip() if sheet_name else None,
        "row": int(match.group("row")),
        "col": col_to_idx(match.group("col").upper()),
    }


def parse_range_ref(ref):
    match = RANGE_REF_RE.match(str(ref).strip())
    if not match:
        return None

    sheet_name = match.group("sheet_q") or match.group("sheet_u")
    clean_range = f"{match.group('start')}:{match.group('end')}".replace("$", "")
    min_col, min_row, max_col, max_row = range_boundaries(clean_range)
    return {
        "sheet": sheet_name.strip() if sheet_name else None,
        "min_col": min_col,
        "min_row": min_row,
        "max_col": max_col,
        "max_row": max_row,
    }


def split_formula_args(text):
    args = []
    current = []
    depth = 0
    in_string = False

    for char in str(text):
        if char == '"':
            in_string = not in_string
            current.append(char)
            continue

        if not in_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
            elif char in {",", ";"} and depth == 0:
                args.append("".join(current).strip())
                current = []
                continue

        current.append(char)

    args.append("".join(current).strip())
    return args


def safe_eval_numeric(expr):
    import ast
    import operator

    operators = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.Pow: operator.pow,
        ast.USub: operator.neg,
        ast.UAdd: operator.pos,
    }

    def eval_node(node):
        if isinstance(node, ast.Expression):
            return eval_node(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value
        if isinstance(node, ast.Num):
            return node.n
        if isinstance(node, ast.BinOp) and type(node.op) in operators:
            return operators[type(node.op)](eval_node(node.left), eval_node(node.right))
        if isinstance(node, ast.UnaryOp) and type(node.op) in operators:
            return operators[type(node.op)](eval_node(node.operand))
        raise ValueError("Unsupported formula expression")

    parsed = ast.parse(expr, mode="eval")
    return eval_node(parsed)


def resolve_range_sum(wb_values, wb_formulas, current_values, current_formula, ref, depth=0, seen=None):
    parsed = parse_range_ref(ref)
    if not parsed:
        return None

    ws_values, ws_formula = sheet_pair(
        wb_values,
        wb_formulas,
        current_values,
        current_formula,
        parsed["sheet"],
    )
    if ws_values is None and ws_formula is None:
        return None

    total = Decimal("0")
    found = False
    for row in range(parsed["min_row"], parsed["max_row"] + 1):
        for col in range(parsed["min_col"], parsed["max_col"] + 1):
            value = resolve_cell_value(
                wb_values,
                wb_formulas,
                ws_values,
                ws_formula,
                row,
                col,
                depth + 1,
                seen,
            )
            dec = to_decimal(value)
            if dec is not None:
                total += dec
                found = True

    return total if found else None


def evaluate_condition(wb_values, wb_formulas, ws_values, ws_formula, expr, depth=0, seen=None):
    for op in (">=", "<=", "<>", "=", ">", "<"):
        if op not in expr:
            continue
        left, right = expr.split(op, 1)
        left_value = evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, left, depth + 1, seen)
        right_value = evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, right, depth + 1, seen)

        left_dec = to_decimal(left_value)
        right_dec = to_decimal(right_value)
        if left_dec is not None and right_dec is not None:
            left_value = left_dec
            right_value = right_dec
        else:
            left_value = str(left_value or "")
            right_value = str(right_value or "")

        if op == ">=":
            return left_value >= right_value
        if op == "<=":
            return left_value <= right_value
        if op == "<>":
            return left_value != right_value
        if op == "=":
            return left_value == right_value
        if op == ">":
            return left_value > right_value
        if op == "<":
            return left_value < right_value

    value = evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, expr, depth + 1, seen)
    dec = to_decimal(value)
    return bool(dec) if dec is not None else bool(value)


def evaluate_formula_function(wb_values, wb_formulas, ws_values, ws_formula, name, args_text, depth=0, seen=None):
    name = name.upper()
    args = split_formula_args(args_text)

    if name == "SUM":
        total = Decimal("0")
        found = False
        for arg in args:
            range_sum = resolve_range_sum(wb_values, wb_formulas, ws_values, ws_formula, arg, depth + 1, seen)
            value = range_sum if range_sum is not None else evaluate_formula_arg(
                wb_values, wb_formulas, ws_values, ws_formula, arg, depth + 1, seen
            )
            dec = to_decimal(value)
            if dec is not None:
                total += dec
                found = True
        return total if found else None

    if name == "SUBTOTAL":
        if len(args) < 2:
            return None
        return evaluate_formula_function(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            "SUM",
            ",".join(args[1:]),
            depth + 1,
            seen,
        )

    if name == "ROUND":
        if len(args) < 2:
            return None
        value = to_decimal(evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, args[0], depth + 1, seen))
        digits = to_int(evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, args[1], depth + 1, seen))
        if value is None or digits is None:
            return None
        quantum = Decimal("1").scaleb(-digits)
        return value.quantize(quantum, rounding=ROUND_HALF_UP)

    if name == "IFERROR":
        if len(args) < 2:
            return None
        value = evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, args[0], depth + 1, seen)
        if value is None:
            return evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, args[1], depth + 1, seen)
        return value

    if name == "IF":
        if len(args) < 3:
            return None
        condition = evaluate_condition(wb_values, wb_formulas, ws_values, ws_formula, args[0], depth + 1, seen)
        return evaluate_formula_arg(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            args[1] if condition else args[2],
            depth + 1,
            seen,
        )

    return None


def evaluate_formula_arg(wb_values, wb_formulas, ws_values, ws_formula, arg, depth=0, seen=None):
    text = str(arg).strip()
    if text.startswith("="):
        return evaluate_excel_formula(wb_values, wb_formulas, ws_values, ws_formula, text, depth + 1, seen)
    if text.startswith('"') and text.endswith('"'):
        return text[1:-1]

    range_sum = resolve_range_sum(wb_values, wb_formulas, ws_values, ws_formula, text, depth + 1, seen)
    if range_sum is not None:
        return range_sum

    parsed_cell = parse_cell_ref(text)
    if parsed_cell:
        target_values, target_formula = sheet_pair(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            parsed_cell["sheet"],
        )
        return resolve_cell_value(
            wb_values,
            wb_formulas,
            target_values,
            target_formula,
            parsed_cell["row"],
            parsed_cell["col"],
            depth + 1,
            seen,
        )

    dec = to_decimal(text)
    if dec is not None:
        return dec

    return evaluate_excel_formula(wb_values, wb_formulas, ws_values, ws_formula, "=" + text, depth + 1, seen)


def evaluate_excel_formula(wb_values, wb_formulas, ws_values, ws_formula, formula, depth=0, seen=None):
    if depth > 30:
        return None
    if not isinstance(formula, str):
        return formula

    expr = formula.strip()
    if expr.startswith("="):
        expr = expr[1:].strip()
    if expr.startswith("+"):
        expr = expr[1:].strip()
    if not expr:
        return None
    if expr.startswith('"') and expr.endswith('"'):
        return expr[1:-1]

    range_sum = resolve_range_sum(wb_values, wb_formulas, ws_values, ws_formula, expr, depth + 1, seen)
    if range_sum is not None:
        return range_sum

    parsed_cell = parse_cell_ref(expr)
    if parsed_cell:
        target_values, target_formula = sheet_pair(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            parsed_cell["sheet"],
        )
        return resolve_cell_value(
            wb_values,
            wb_formulas,
            target_values,
            target_formula,
            parsed_cell["row"],
            parsed_cell["col"],
            depth + 1,
            seen,
        )

    outer_func = re.match(r"^([A-Z][A-Z0-9.]*)\((.*)\)$", expr, re.IGNORECASE)
    if outer_func:
        value = evaluate_formula_function(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            outer_func.group(1),
            outer_func.group(2),
            depth + 1,
            seen,
        )
        if value is not None:
            return value

    numeric_expr = expr.replace("$", "").replace("^", "**")
    function_pattern = re.compile(r"([A-Z][A-Z0-9.]*)\(([^()]*)\)", re.IGNORECASE)

    for _ in range(20):
        match = function_pattern.search(numeric_expr)
        if not match:
            break
        value = evaluate_formula_function(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            match.group(1),
            match.group(2),
            depth + 1,
            seen,
        )
        if value is None:
            return None
        numeric_expr = numeric_expr[: match.start()] + str(value) + numeric_expr[match.end() :]

    def replace_cell(match):
        parsed = parse_cell_ref(match.group(0))
        if not parsed:
            return "0"
        target_values, target_formula = sheet_pair(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            parsed["sheet"],
        )
        value = resolve_cell_value(
            wb_values,
            wb_formulas,
            target_values,
            target_formula,
            parsed["row"],
            parsed["col"],
            depth + 1,
            seen,
        )
        dec = to_decimal(value)
        return str(dec if dec is not None else 0)

    numeric_expr = CELL_TOKEN_RE.sub(replace_cell, numeric_expr)
    numeric_expr = numeric_expr.replace(",", ".")
    if not re.fullmatch(r"[0-9+\-*/().\s]+", numeric_expr):
        return None

    try:
        return Decimal(str(safe_eval_numeric(numeric_expr)))
    except Exception:
        return None


def resolve_cell_value(wb_values, wb_formulas, ws_values, ws_formula, row, col, depth=0, seen=None):
    if depth > 30:
        return None
    if seen is None:
        seen = set()

    sheet_title = ws_formula.title if ws_formula is not None else (ws_values.title if ws_values is not None else "")
    key = (sheet_title, row, col)
    if key in seen:
        return None
    seen.add(key)

    value_cell = get_effective_cell(ws_values, row, col)
    cached_value = value_cell.value if value_cell is not None else None
    if not is_blank(cached_value):
        seen.discard(key)
        return cached_value

    formula_cell = get_effective_cell(ws_formula, row, col)
    formula_value = formula_cell.value if formula_cell is not None else None
    if isinstance(formula_value, str) and formula_value.strip().startswith("="):
        resolved = evaluate_excel_formula(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            formula_value,
            depth + 1,
            seen,
        )
        seen.discard(key)
        return resolved

    seen.discard(key)
    return formula_value if not is_blank(formula_value) else cached_value


def resolve_cell_ref_value(wb_values, wb_formulas, ws_values, ws_formula, ref):
    parsed = parse_cell_ref(ref)
    if not parsed:
        return None
    target_values, target_formula = sheet_pair(
        wb_values,
        wb_formulas,
        ws_values,
        ws_formula,
        parsed["sheet"],
    )
    return resolve_cell_value(
        wb_values,
        wb_formulas,
        target_values,
        target_formula,
        parsed["row"],
        parsed["col"],
    )


def find_sum_row_resolved(wb_values, wb_formulas, ws_values, ws_formula, start_row=1, label_col="B"):
    if ws_values is None and ws_formula is None:
        return None

    col = col_to_idx(label_col.upper())
    max_row = max(
        ws_values.max_row if ws_values is not None else 0,
        ws_formula.max_row if ws_formula is not None else 0,
    )
    pat = re.compile(r"^\s*SUM\s*[:\uFF1A]?\s*$", re.IGNORECASE)

    for row in range(start_row, max_row + 1):
        value = resolve_cell_value(wb_values, wb_formulas, ws_values, ws_formula, row, col)
        if value is None:
            continue
        text = str(value).strip()
        if pat.match(text) or text.upper().startswith("SUM"):
            return row

    return None


def resolve_simple_formula(wb_formula, ws_current, formula, depth=0):
    if depth > 10:
        return None
    if not isinstance(formula, str) or not formula.strip().startswith("="):
        return None

    match = REF_RE.match(formula)
    if not match:
        return None

    sheet = match.group("sheet")
    col = match.group("col")
    row = int(match.group("row"))

    if sheet:
        sheet = sheet.strip().strip("'")
        if sheet not in wb_formula.sheetnames:
            return None
        ws = wb_formula[sheet]
    else:
        ws = ws_current

    value = ws.cell(row=row, column=col_to_idx(col)).value
    if isinstance(value, str) and value.strip().startswith("="):
        return resolve_simple_formula(wb_formula, ws, value, depth + 1)
    return value


def header_value(wb_formula, ws_data, ws_formula, ref):
    value = ws_data[ref].value if ws_data else None
    if value is not None and str(value).strip() != "":
        return value
    if ws_formula is None:
        return value

    formula_value = ws_formula[ref].value
    if isinstance(formula_value, str) and formula_value.strip().startswith("="):
        return resolve_simple_formula(wb_formula, ws_formula, formula_value)
    return formula_value


def find_uncached_formula(wb_values, wb_formulas):
    for sheet_name in wb_formulas.sheetnames:
        if sheet_name not in wb_values.sheetnames:
            continue
        ws_formula = wb_formulas[sheet_name]
        ws_values = wb_values[sheet_name]
        for row in ws_formula.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    if ws_values[cell.coordinate].value is None:
                        return f"{sheet_name}!{cell.coordinate}"
    return None


def detect_asian_date_text(value):
    if not value:
        return False, None, None
    text = str(value).strip()
    match = re.match(r"^(\d{3,5})[./-](\d{1,2})[./-](\d{1,2})$", text)
    if not match:
        return False, None, None

    year, month, day = match.groups()
    if len(year) == 4 and year.startswith("25"):
        year = "20" + year[2:]
    elif len(year) == 3:
        year = "2" + year

    return True, text, f"{int(day)}/{int(month)}/{year}"


def final_check_invoice(file_name, file_bytes):
    errors = []
    warnings = []
    cartons = Decimal(0)
    gross = Decimal(0)

    try:
        wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True)
        wb_formulas = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception as exc:
        return {
            "File": file_name,
            "Status": "ERROR",
            "Errors": [f"Cannot open file: {exc}"],
            "Warnings": [],
            "Cartons": 0.0,
            "Gross Weight": 0.0,
        }

    ws_inv = sheet_by_name_ci(wb_values, "INVOICE")
    ws_inv_f = sheet_by_name_ci(wb_formulas, "INVOICE")
    ws_pack = sheet_by_name_ci(wb_values, "PACKING LIST")
    ws_pack_f = sheet_by_name_ci(wb_formulas, "PACKING LIST")

    if ws_inv is None:
        return {
            "File": file_name,
            "Status": "ERROR",
            "Errors": ["INVOICE sheet missing"],
            "Warnings": [],
            "Cartons": 0.0,
            "Gross Weight": 0.0,
        }

    if ws_pack is None:
        warnings.append("PACKING LIST sheet missing")

    file_stem = Path(file_name).stem

    def inv(ref):
        return resolve_cell_ref_value(wb_values, wb_formulas, ws_inv, ws_inv_f, ref)

    def pack(ref):
        return resolve_cell_ref_value(wb_values, wb_formulas, ws_pack, ws_pack_f, ref)

    inv_a2 = inv("A2")
    inv_c4 = inv("C4")
    if inv_a2 != inv_c4:
        errors.append(f"INVOICE A2 != C4 ({inv_a2} / {inv_c4})")

    if ws_pack is not None:
        pack_a2 = pack("A2")
        if not (inv_a2 == inv_c4 == pack_a2):
            errors.append(f"Inconsistent headers A2/C4/PACK A2 ({inv_a2} / {inv_c4} / {pack_a2})")

    inv_c5 = str(inv("C5") or "").strip()
    inv_j4 = str(inv("J4") or "").strip()
    pack_b4 = str(pack("B4") or "").strip() if ws_pack else ""
    if not (file_stem == inv_c5 == inv_j4 == pack_b4):
        errors.append(f"File name != C5/J4/PACK B4 ({file_stem}, {inv_c5}, {inv_j4}, {pack_b4})")

    changed, old_date, new_date = detect_asian_date_text(inv("J5"))
    if changed:
        warnings.append(f"Asian date format detected in J5 ({old_date}); output will write {new_date}.")

    j13 = str(inv("J13") or "").strip().upper()
    if j13 != "EUR":
        errors.append(f"J13 must be EUR ({inv('J13')})")

    j14 = str(inv("J14") or "").strip().upper()
    if j14 != "CIF":
        warnings.append(f"J14 must be CIF ({inv('J14')})")

    j16 = to_decimal(inv("J16"))
    if j16 != Decimal(4200):
        errors.append(f"J16 must be 4200 ({inv('J16')})")

    for row in range(11, 16):
        value = inv(f"C{row}")
        if value is None or str(value).strip() == "":
            errors.append(f"C{row} is empty")

    for row in (16, 17):
        value = inv(f"C{row}")
        text = str(value or "").strip().upper()
        if not COUNTRY_CODE_RE.match(text):
            errors.append(f"C{row} must be a 2-letter country code ({value})")

    for row in (11, 13, 15):
        value = inv(f"J{row}")
        if value is None or str(value).strip() == "":
            errors.append(f"J{row} is empty")

    c9 = str(inv("C9") or "").strip().upper()
    if c9 != "CN":
        errors.append(f"C9 must be CN ({inv('C9')})")

    inv_sum = find_sum_row_resolved(wb_values, wb_formulas, ws_inv, ws_inv_f, 19, "B")
    pack_sum = find_sum_row_resolved(wb_values, wb_formulas, ws_pack, ws_pack_f, 6, "B") if ws_pack else None

    if inv_sum is None:
        errors.append("INVOICE SUM row not found")
    if ws_pack and pack_sum is None:
        errors.append("PACKING LIST SUM row not found")

    if inv_sum:
        for row in range(20, inv_sum):
            if contains_chinese(inv(f"B{row}")):
                errors.append(f"Chinese character found in INVOICE B{row}")
                break

    if ws_pack and pack_sum:
        for row in range(6, pack_sum):
            if contains_chinese(pack(f"B{row}")):
                errors.append(f"Chinese character found in PACKING B{row}")
                break

    if inv_sum:
        merged = []
        for row in range(20, inv_sum):
            if is_cell_in_merged(ws_inv, row, 10):
                merged.append(f"J{row}")
            if is_cell_in_merged(ws_inv, row, 11):
                merged.append(f"K{row}")
        if merged:
            errors.append("INVOICE forbidden merged cells in J/K: " + ", ".join(merged[:30]))

    if ws_pack and pack_sum:
        merged = []
        for row in range(6, pack_sum):
            if is_cell_in_merged(ws_pack, row, 9):
                merged.append(f"I{row}")
            if is_cell_in_merged(ws_pack, row, 10):
                merged.append(f"J{row}")
        if merged:
            errors.append("PACKING forbidden merged cells in I/J: " + ", ".join(merged[:30]))

    if inv_sum:
        if any(isinstance(inv(f"G{row}"), str) and len(inv(f"G{row}").strip()) > 48 for row in range(20, inv_sum)):
            errors.append("INVOICE column G contains a value longer than 48 characters")

        bad_di = []
        for row in range(20, inv_sum):
            for col_letter, col_index in (("D", 4), ("I", 9)):
                if is_cell_in_merged(ws_inv, row, col_index):
                    continue
                value = inv(f"{col_letter}{row}")
                dec = to_decimal(value)
                if value is None or str(value).strip() == "" or dec is None or dec == 0:
                    bad_di.append(f"{col_letter}{row}")
        if bad_di:
            errors.append("INVOICE D/I empty, zero, or non-numeric: " + ", ".join(bad_di[:30]))

        text_weight = []
        bad_weight = []
        for row in range(20, inv_sum):
            j_value = inv(f"J{row}")
            k_value = inv(f"K{row}")
            if isinstance(j_value, str) or isinstance(k_value, str):
                text_weight.append(row)
                continue
            j_dec = to_decimal(j_value)
            k_dec = to_decimal(k_value)
            if j_dec is not None and k_dec is not None and j_dec >= k_dec:
                bad_weight.append(row)
        if text_weight:
            errors.append("INVOICE J/K text values on rows: " + ", ".join(map(str, text_weight[:30])))
        if bad_weight:
            errors.append("INVOICE net weight >= gross weight on rows: " + ", ".join(map(str, bad_weight[:30])))

    if inv_sum and ws_pack and pack_sum:
        inv_pieces = to_decimal(inv(f"H{inv_sum}"))
        inv_net = to_decimal(inv(f"J{inv_sum}"))
        inv_gross = to_decimal(inv(f"K{inv_sum}"))
        pack_pieces = to_decimal(pack(f"H{pack_sum}"))
        pack_net = to_decimal(pack(f"I{pack_sum}"))
        pack_gross = to_decimal(pack(f"J{pack_sum}"))
        pack_cartons = to_decimal(pack(f"G{pack_sum}"))

        if inv_pieces != pack_pieces:
            errors.append(f"Total pieces differ INV/PACK ({inv_pieces}/{pack_pieces})")
        if inv_net != pack_net:
            errors.append(f"Total net weight differs INV/PACK ({inv_net}/{pack_net})")
        if inv_gross != pack_gross:
            errors.append(f"Total gross weight differs INV/PACK ({inv_gross}/{pack_gross})")
        if inv_net is not None and inv_gross is not None and inv_net > inv_gross:
            errors.append(f"Total net > gross ({inv_net}>{inv_gross})")

        if pack_cartons is None:
            errors.append(f"Carton count missing in PACK G{pack_sum}")
        else:
            cartons = pack_cartons

        if pack_gross is not None:
            gross = pack_gross

        inv_b = [str(inv(f"B{row}") or "").strip() for row in range(20, inv_sum)]
        pack_b = [str(pack(f"B{row}") or "").strip() for row in range(6, pack_sum)]
        if inv_b != pack_b:
            errors.append(f"Column B descriptions differ INV/PACK ({len(inv_b)} rows / {len(pack_b)} rows)")

        line_errors = []
        for inv_row in range(20, inv_sum):
            pack_row = inv_row - 14
            if pack_row < 6 or pack_row >= pack_sum:
                continue
            inv_p = to_decimal(inv(f"H{inv_row}"))
            inv_n = to_decimal(inv(f"J{inv_row}"))
            inv_g = to_decimal(inv(f"K{inv_row}"))
            pack_p = to_decimal(pack(f"H{pack_row}"))
            pack_n = to_decimal(pack(f"I{pack_row}"))
            pack_g = to_decimal(pack(f"J{pack_row}"))
            q = Decimal("0.01")
            inv_n_q = inv_n.quantize(q) if inv_n is not None else None
            pack_n_q = pack_n.quantize(q) if pack_n is not None else None
            inv_g_q = inv_g.quantize(q) if inv_g is not None else None
            pack_g_q = pack_g.quantize(q) if pack_g is not None else None
            if inv_p != pack_p or inv_n_q != pack_n_q or inv_g_q != pack_g_q:
                line_errors.append(inv_row)
        if line_errors:
            errors.append("Line-by-line differences INV/PACK: " + ", ".join(map(str, line_errors[:30])))

    if ws_pack and pack_sum:
        bad_g = []
        for row in range(6, pack_sum):
            raw = pack(f"G{row}")
            dec = to_decimal(raw)
            if raw is None or str(raw).strip() == "" or dec is None or dec == 0:
                bad_g.append(f"G{row}")
        if bad_g:
            errors.append("PACKING G empty, zero, or non-numeric: " + ", ".join(bad_g[:30]))

    status = "ERROR" if errors else ("WARNING" if warnings else "OK")
    return {
        "File": file_name,
        "Status": status,
        "Errors": errors,
        "Warnings": warnings,
        "Cartons": float(cartons),
        "Gross Weight": float(gross),
    }


def normalize_line(text):
    text = text.replace("\u00a0", " ")
    text = text.replace("\t", " ")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def text_to_lines(text):
    out = []
    for line in (text or "").splitlines():
        line = normalize_line(line)
        if line:
            out.append(line)
    return out


def words_to_lines(page):
    words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    if not words:
        return []

    buckets = {}
    for word in words:
        key = int(round(word["top"] / 2.0) * 2)
        buckets.setdefault(key, []).append(word)

    lines = []
    for key in sorted(buckets):
        row = sorted(buckets[key], key=lambda item: item["x0"])
        text = " ".join(word["text"] for word in row if word.get("text"))
        text = normalize_line(text)
        if text:
            lines.append(text)
    return lines


def candidate_quality(lines):
    if not lines:
        return (0, 0, 0, 0, 0)

    kw_rx = re.compile(r"\bCT\b|\bPK\b|CTNO|\bKarton\b|Packung/Packst.ck|\bPA\b|\bPaket\b|\bBX\b|\bBN\b|B/N", re.I)
    item_rx = re.compile(r"\b\d+\s+\d+\s+\d+\s+\d+\s+(?:CT|PK)\s*(?:B/N|BN)\b", re.I)
    garbage_rx = re.compile(r"^[0\s]{4,}[0-9]{6,}\s+\d{2,3}$")

    return (
        sum(1 for line in lines if item_rx.search(line)),
        sum(1 for line in lines if kw_rx.search(line)),
        sum(1 for line in lines if 3 <= len(line) <= 120),
        -sum(1 for line in lines if garbage_rx.match(line)),
        len(lines),
    )


def extract_text_lines_from_pdf(pdf_bytes):
    all_lines = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            candidates = []
            try:
                candidates.append(text_to_lines(page.extract_text() or ""))
            except Exception:
                candidates.append([])
            try:
                candidates.append(text_to_lines(page.extract_text(layout=True) or ""))
            except Exception:
                candidates.append([])
            try:
                candidates.append(words_to_lines(page))
            except Exception:
                candidates.append([])

            all_lines.extend(max(candidates, key=candidate_quality))
    return all_lines


def dedup_and_sort(pairs):
    dedup = {}
    for pos, cartons in pairs:
        if pos not in dedup:
            dedup[pos] = cartons
    return sorted(dedup.items(), key=lambda item: item[0])


def auto_renum_if_all_same_pos(rows):
    if not rows:
        return rows
    if len({pos for pos, _cartons in rows}) == 1:
        return [(index + 1, cartons) for index, (_pos, cartons) in enumerate(rows)]
    return rows


def score_rows(rows, prefer_sequential_positions=True):
    if not rows:
        return 0.0, "0 row"

    count = len(rows)
    positions = [pos for pos, _cartons in rows]
    cartons = [qty for _pos, qty in rows]

    plausible_ratio = sum(1 for qty in cartons if 1 <= qty <= 5000) / max(1, count)
    unique_ratio = len(set(positions)) / max(1, count)
    pos_sorted = sorted(positions)
    span = pos_sorted[-1] - pos_sorted[0] + 1
    continuity_ratio = len(set(pos_sorted)) / max(1, span)
    starts_at_1 = 1.0 if pos_sorted[0] == 1 else 0.0
    qty_score = 1.0 if count >= 5 else (0.85 if count >= 2 else 0.65)
    seq_bonus = 0.07 if prefer_sequential_positions and pos_sorted[0] == 1 and continuity_ratio >= 0.95 else 0.0

    score = (
        0.35 * qty_score
        + 0.25 * plausible_ratio
        + 0.25 * continuity_ratio
        + 0.10 * unique_ratio
        + 0.05 * starts_at_1
        + seq_bonus
    )
    score = max(0.0, min(1.0, score))
    notes = (
        f"n={count}, plausible={plausible_ratio:.2f}, cont={continuity_ratio:.2f}, "
        f"unique={unique_ratio:.2f}, minpos={pos_sorted[0]}, maxpos={pos_sorted[-1]}"
    )
    return score, notes


ParserFn = Callable[[List[str]], List[Tuple[int, int]]]


def p_pk_nm_polish(lines):
    out = []
    pat_spaced = re.compile(r"\b(\d+)\s+(\d+)\s+\d+\s+(\d{1,9})\s*PK\s*NM\b", re.I)
    pat_glued = re.compile(r"\b(\d+)\s+(\d+)\s+(\d+)PKNM\b", re.I)
    blob = " \n ".join(str(item) for item in lines if item)
    for source in list(lines) + [blob]:
        for match in pat_spaced.finditer(str(source)):
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
        for match in pat_glued.finditer(str(source)):
            raw = match.group(3)
            cartons = int(raw[1:] if raw.startswith("1") and len(raw) > 1 else raw)
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_english_list_items_pk_sequential(lines):
    text = "\n".join(str(item) for item in lines if item)
    if not re.search(r"TRANSIT\s+LIST\s+OF\s+ITEMS|Decl\s+goods\s+it\.\s*Nr\.?", text, re.I):
        return []

    out = []
    seq = 0
    pat = re.compile(r"^\s*1\s+PK\s+(\d{1,9})(?:\s*-|\s*$)", re.I)
    for line in lines:
        match = pat.match(str(line).strip())
        if match:
            cartons = int(match.group(1))
            if cartons > 0:
                seq += 1
                out.append((seq, cartons))
    return out


def p_pk_bn_polish(lines):
    out = []
    pat_spaced = re.compile(r"\b(\d+)\s+(\d+)\s+\d+\s+(\d{1,9})\s*PK\s*(?:B/N|BN)\b", re.I)
    pat_glued = re.compile(r"\b(\d+)\s+(\d+)\s+(\d+)PK(?:B/N|BN)\b", re.I)
    blob = " \n ".join(str(item) for item in lines if item)
    for source in list(lines) + [blob]:
        for match in pat_spaced.finditer(str(source)):
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
        for match in pat_glued.finditer(str(source)):
            raw = match.group(3)
            cartons = int(raw[1:] if raw.startswith("1") and len(raw) > 1 else raw)
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_ct_bn_format(lines):
    out = []
    pat_spaced = re.compile(r"\b(\d+)\s+(\d+)\s+\d+\s+(\d+)\s+CT\s*B/N\b", re.I)
    pat_glued = re.compile(r"\b(\d+)\s+(\d+)\s+(\d+)CTB/N\b", re.I)
    for line in lines:
        text = str(line)
        for match in pat_spaced.finditer(text):
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
        for match in pat_glued.finditer(text):
            raw = match.group(3)
            cartons = int(raw[1:] if raw.startswith("1") and len(raw) > 1 else raw)
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_paket_before(lines):
    out = []
    pat = re.compile(r"\b(\d+)\s+(\d+)\b.*?\b(\d{1,9})\s*Paket\b", re.I)
    blob = " \n ".join(str(item) for item in lines if item)
    for source in list(lines) + [blob]:
        for match in pat.finditer(str(source)):
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_pa_after(lines):
    out = []
    pat = re.compile(r"\b(\d+)\s+(\d+)\s+\d+\s*PA\s*(\d{1,9})\b", re.I)
    blob = " \n ".join(str(item) for item in lines if item)
    for source in list(lines) + [blob]:
        for match in pat.finditer(str(source)):
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_ct_nm_packstuecke(lines):
    out = []
    pat = re.compile(r"\b(\d+)\s+(\d+)\s+\d+\s*CT\s*(\d{1,9})\s*NM(?=\b|(?=[A-Z]))", re.I)
    blob = " \n ".join(str(item) for item in lines if item)
    for source in list(lines) + [blob]:
        for match in pat.finditer(str(source)):
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_ctno_glued_or_spaced(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s+(\d+)\s+(\d+)\s*(CT|PK)\s*NO\b", re.I)
    for line in lines:
        match = pat.match(str(line))
        if not match:
            continue
        raw = match.group(3)
        cartons = int(raw[1:] if raw.startswith("1") and len(raw) > 1 else raw)
        out.append((int(match.group(2)), cartons))
    return dedup_and_sort(out)


def p_ct_spaced(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s+(\d+)\s+1\s+(\d+)\s*(CT|PK)\b(?:\s+[A-Z]{2})?\b\.?\s*$", re.I)
    for line in lines:
        match = pat.match(str(line))
        if match:
            out.append((int(match.group(2)), int(match.group(3))))
    return dedup_and_sort(out)


def p_ct_glued_with_leading_one(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s+(\d+)\s+1(\d+)\s*(CT|PK)\b", re.I)
    for line in lines:
        match = pat.match(str(line))
        if match:
            out.append((int(match.group(2)), int(match.group(3))))
    return dedup_and_sort(out)


def p_pkct_after(lines):
    out = []
    pat_range = re.compile(r"^\s*(\d+)\s+(\d+)\s+1\s*(PK|CT)\s*([0-9]+)\s*-\s*([0-9]+)\b", re.I)
    pat_main = re.compile(r"^\s*(\d+)\s+(\d+)\s+1\s*(PK|CT)\s*([0-9]+)\b", re.I)
    pat_alt = re.compile(r"^\s*(\d+)\s+1\s*(PK|CT)\s*([0-9]+)\b", re.I)

    for line in lines:
        text = str(line).strip()
        match = pat_range.match(text)
        if match:
            pos = int(match.group(2))
            left = int(match.group(4))
            right = int(match.group(5))
            cartons = right if str(match.group(4)).endswith("1") and int(str(match.group(4))[:-1] or 0) == right else left
            if cartons > 0:
                out.append((pos, cartons))
            continue

        match = pat_main.match(text)
        if match:
            out.append((int(match.group(2)), int(match.group(4))))
            continue

        match = pat_alt.match(text)
        if match:
            out.append((int(match.group(1)), int(match.group(3))))
    return dedup_and_sort(out)


def p_karton_inline(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s+(\d+)\s+(\d+)\s+Karton\b", re.I)
    for line in lines:
        match = pat.match(str(line))
        if match:
            out.append((int(match.group(2)), int(match.group(3))))
    return dedup_and_sort(out)


def p_packung_inline(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s+(\d+)\s+(\d+)\s+Packung/Packst.ck\b", re.I)
    for line in lines:
        match = pat.match(str(line))
        if match:
            out.append((int(match.group(2)), int(match.group(3))))
    return dedup_and_sort(out)


def p_dash_format(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s+(\d+)\s+(\d+)\s*-\s*")
    for line in lines:
        match = pat.match(str(line))
        if match:
            out.append((int(match.group(2)), int(match.group(3))))
    return dedup_and_sort(out)


def p_block_fallback(lines):
    out = []
    pat_pos_only = re.compile(r"^\s*(\d+)\s+(\d+)\s*$")
    pat_karton = re.compile(r"\b(\d+)\s*Karton\b", re.I)
    pat_packung = re.compile(r"\b(\d+)\s+Packung/Packst.ck\b", re.I)
    current_pos = None
    already = set()
    for line in lines:
        text = str(line)
        match_pos = pat_pos_only.match(text)
        if match_pos:
            current_pos = int(match_pos.group(2))
            continue
        if current_pos is None or current_pos in already:
            continue
        match = pat_karton.search(text) or pat_packung.search(text)
        if match:
            out.append((current_pos, int(match.group(1))))
            already.add(current_pos)
    return dedup_and_sort(out)


def p_eu_items_ct_after(lines):
    out = []
    pat = re.compile(r"^\s*(\d+)\s*(CT|PK)\s*(\d+)\s*(?:-|$)", re.I)
    seq = 0
    for line in lines:
        match = pat.match(str(line))
        if match:
            cartons = int(match.group(3))
            if cartons > 0:
                seq += 1
                out.append((seq, cartons))
    return dedup_and_sort(out)


def p_pk_bm_phase5(lines):
    out = []
    pat = re.compile(r"^\s*\d+\s+\d+\s+\d+\s+PK\s+(\d+)\s+B/M\b", re.I)
    seq = 0
    for line in lines:
        match = pat.match(str(line))
        if match:
            cartons = int(match.group(1))
            if cartons > 0:
                seq += 1
                out.append((seq, cartons))
    return dedup_and_sort(out)


def p_addr_pk_sequential(lines):
    out = []
    pat = re.compile(r"^\s*(?:\d+\s+)?ADDR\s*-\s*(\d{1,9})\s*PK\b", re.I)
    seq = 0
    for line in lines:
        text = str(line).strip()
        if re.search(r"^\s*32\s+Item\s+No\b", text, re.I):
            continue
        match = pat.match(text)
        if match:
            cartons = int(match.group(1))
            if cartons > 0:
                seq += 1
                out.append((seq, cartons))
    return dedup_and_sort(out)


def p_addr_pk(lines):
    out = []
    pat = re.compile(r"^\s*(\d{1,6})\b.*?\bADDR\s*-\s*(\d{1,9})\s*PK\b", re.I)
    for line in lines:
        match = pat.search(str(line))
        if match:
            cartons = int(match.group(2))
            if cartons > 0:
                out.append((int(match.group(1)), cartons))
    return dedup_and_sort(out)


def p_bx_support(lines):
    out = []
    pat_compact = re.compile(r"^\s*(\d+)\s+(\d+)\s+1(\d+)\s*BX\b\.?", re.I)
    pat_spaced = re.compile(r"^\s*(\d+)\s+(\d+)\s+BX\b\.?", re.I)
    for line in lines:
        text = str(line).strip()
        match = pat_compact.match(text)
        if match:
            cartons = int(match.group(3))
            if cartons > 0:
                out.append((int(match.group(2)), cartons))
            continue
        match = pat_spaced.match(text)
        if match:
            cartons = int(match.group(2))
            if cartons > 0:
                out.append((int(match.group(1)), cartons))
    if not out:
        return []
    out = auto_renum_if_all_same_pos(out)
    if len({pos for pos, _cartons in out}) == len(out):
        return out
    return dedup_and_sort(out)


def p_list_of_items_carton(lines):
    out = []
    seen = set()
    pat = re.compile(r"^\s*(\d{1,6})\s+(\d{1,6})\s+(\d{1,8})\s+Carton\b", re.I)
    for line in lines:
        match = pat.match(str(line).strip())
        if not match:
            continue
        pos = int(match.group(1))
        cartons = int(match.group(3))
        if cartons <= 0 or pos in seen:
            continue
        seen.add(pos)
        out.append((pos, cartons))
    return dedup_and_sort(out)


def p_transit_blocks(lines):
    raw = []
    pat_goods_it = re.compile(r"\[11\s*03\]\s*(\d{1,6})", re.I)
    pat_pkg = re.compile(r"^\s*\d+\s+(CT|PK)\s+(\d{1,9})\b", re.I)
    current_pos = None
    in_items_table = False
    for line in lines:
        text = str(line)
        match_goods = pat_goods_it.search(text)
        if match_goods:
            in_items_table = True
            current_pos = int(match_goods.group(1))
            continue
        if not in_items_table:
            continue
        match_pkg = pat_pkg.match(text.strip())
        if match_pkg:
            cartons = int(match_pkg.group(2))
            if cartons > 0:
                raw.append((current_pos if current_pos is not None else 1, cartons))
    if not raw:
        return []
    raw = auto_renum_if_all_same_pos(raw)
    if len({pos for pos, _cartons in raw}) == len(raw):
        return raw
    return dedup_and_sort(raw)


def p_package_rows(lines):
    out = []
    seen = set()
    pat_packages = re.compile(r"^\s*(\d{1,6})\s+(\d{1,6})\s+(\d{1,9})\s+Packages?\b", re.I)
    pat_pakunek = re.compile(r"^\s*(\d{1,6})\s+(\d{1,6})\s+(\d{1,9})\s+Pakunek\b", re.I)
    for line in lines:
        match = pat_packages.match(str(line).strip()) or pat_pakunek.match(str(line).strip())
        if not match:
            continue
        pos = int(match.group(1))
        cartons = int(match.group(3))
        if cartons <= 0 or pos in seen:
            continue
        seen.add(pos)
        out.append((pos, cartons))
    return dedup_and_sort(out)


PARSERS: List[Tuple[str, ParserFn, bool]] = [
    ("PK NM/nm Polish", p_pk_nm_polish, False),
    ("English TRANSIT LIST OF ITEMS", p_english_list_items_pk_sequential, True),
    ("PK BN/B/N Polish", p_pk_bn_polish, False),
    ("CT B/N format", p_ct_bn_format, False),
    ("PHASE5 PK B/M", p_pk_bm_phase5, True),
    ("CT NM", p_ct_nm_packstuecke, False),
    ("ADDR - X PK sequential", p_addr_pk_sequential, True),
    ("ADDR - X PK", p_addr_pk, False),
    ("BX support", p_bx_support, False),
    ("TRANSIT BLOCKS", p_transit_blocks, False),
    ("PACKAGE ROWS", p_package_rows, False),
    ("LIST OF ITEMS Carton", p_list_of_items_carton, False),
    ("CTNO glued/spaced", p_ctno_glued_or_spaced, False),
    ("CT spaced", p_ct_spaced, False),
    ("CT glued with leading 1", p_ct_glued_with_leading_one, False),
    ("PK/CT after", p_pkct_after, False),
    ("Karton inline", p_karton_inline, False),
    ("Packung inline", p_packung_inline, False),
    ("Dash format", p_dash_format, False),
    ("Block fallback", p_block_fallback, False),
    ("EU items CT/PK after", p_eu_items_ct_after, True),
    ("PA after", p_pa_after, False),
    ("Paket before", p_paket_before, False),
]


def choose_best_t1_parser(lines):
    results = []
    for name, parser, prefer_seq in PARSERS:
        rows = parser(lines)
        score, notes = score_rows(rows, prefer_sequential_positions=prefer_seq)
        results.append(ParseResult(name=name, rows=rows, score=score, notes=notes))
    best = max(results, key=lambda item: item.score) if results else None
    return best, results


def extract_t1_number(lines):
    text = "\n".join(lines)
    patterns = [
        r"\bMRN\b\s*[:\-]?\s*([0-9]{2}[A-Z]{2}[A-Z0-9]{14})\b",
        r"\b([0-9]{2}[A-Z]{2}[A-Z0-9]{14})\b",
        r"\bT1\s*(?:No\.?|Number|Nr\.?)?\s*[:\-]?\s*([A-Z0-9]{8,25})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return match.group(1).upper()
    return ""


def normalize_t1_date(raw_date):
    match = re.match(r"^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\s*$", str(raw_date))
    if not match:
        return ""

    day, month, year = match.groups()
    if len(year) == 2:
        year = "20" + year

    return f"{int(day):02d}/{int(month):02d}/{int(year):04d}"


def extract_t1_date(lines):
    date_re = re.compile(r"\b(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\b")
    context_terms = (
        "reeweg",
        "simplified proc",
        "vereenvoudigde",
        "kantoor van vertrek",
        "nl000510",
    )

    candidates = []
    for index, line in enumerate(lines):
        matches = date_re.findall(str(line))
        if not matches:
            continue

        window = " ".join(str(item).lower() for item in lines[max(0, index - 2) : index + 3])
        score = 0
        for term in context_terms:
            if term in window:
                score += 10
        if "uiterste datum" in window or "termijn" in window:
            score -= 6

        for raw_date in matches:
            date_value = normalize_t1_date(raw_date)
            if date_value:
                candidates.append((score, index, date_value))

    if not candidates:
        return ""

    candidates.sort(key=lambda item: (-item[0], item[1]))
    return candidates[0][2]


def extract_t1_from_pdf(pdf_bytes):
    lines = extract_text_lines_from_pdf(pdf_bytes)
    best, all_results = choose_best_t1_parser(lines)
    t1_no = extract_t1_number(lines)
    t1_date = extract_t1_date(lines)

    if best is None:
        return {
            "ok": False,
            "rows": [],
            "parser": "",
            "score": 0.0,
            "notes": "No parser result",
            "t1_no": t1_no,
            "t1_date": t1_date,
            "parser_scores": [],
        }

    return {
        "ok": bool(best.rows) and best.score >= MIN_T1_SCORE,
        "rows": best.rows,
        "parser": best.name,
        "score": best.score,
        "notes": best.notes,
        "t1_no": t1_no,
        "t1_date": t1_date,
        "parser_scores": all_results,
    }


def packing_has_merged_in_g(ws_pack, sum_row):
    if not sum_row:
        return False
    for rng in ws_pack.merged_cells.ranges:
        if rng.min_col <= 7 <= rng.max_col:
            if not (rng.max_row < 6 or rng.min_row > max(6, sum_row - 1)):
                return True
    return False


def read_packing_lines(invoice_entries):
    lines = []
    file_info = {}
    notes = []
    line_id = 0

    for entry in sorted(invoice_entries, key=lambda item: natural_key(item["name"])):
        file_name = entry["name"]
        try:
            wb_values = load_workbook(io.BytesIO(entry["bytes"]), data_only=True)
            wb_formulas = load_workbook(io.BytesIO(entry["bytes"]), data_only=False)
        except Exception as exc:
            notes.append({"File": file_name, "Type": "ERROR", "Message": f"Cannot read workbook: {exc}"})
            continue

        ws_pack = sheet_by_name_ci(wb_values, "PACKING LIST")
        ws_pack_f = sheet_by_name_ci(wb_formulas, "PACKING LIST")
        if ws_pack is None:
            notes.append({"File": file_name, "Type": "ERROR", "Message": "PACKING LIST sheet missing"})
            continue

        sum_row = find_sum_row_resolved(wb_values, wb_formulas, ws_pack, ws_pack_f, start_row=6, label_col="B")
        if not sum_row:
            notes.append({"File": file_name, "Type": "ERROR", "Message": "PACKING LIST SUM row not found"})
            continue

        merged_in_g = packing_has_merged_in_g(ws_pack, sum_row)
        cartons_list = []

        for row in range(6, sum_row):
            merged_g = merged_range_for_cell(ws_pack, row, 7)
            if merged_g is not None and (row != merged_g.min_row or 7 != merged_g.min_col):
                continue

            cartons_value = resolve_cell_value(wb_values, wb_formulas, ws_pack, ws_pack_f, row, 7)
            cartons = to_int(cartons_value)
            if cartons is None or cartons <= 0:
                continue
            cartons_list.append(cartons)
            lines.append(
                {
                    "id": line_id,
                    "file": file_name,
                    "sheet": ws_pack.title,
                    "row": row,
                    "cartons": cartons,
                }
            )
            line_id += 1

        file_info[file_name] = {
            "split_allowed": not merged_in_g,
            "merged_in_g": merged_in_g,
            "cartons_list": cartons_list,
        }

    return lines, file_info, notes


def subset_dp_find(target, lines, available_ids):
    dp = {0: None}
    for idx in available_ids:
        value = lines[idx]["cartons"]
        if value <= 0:
            continue
        for subtotal in sorted(list(dp.keys()), reverse=True):
            new_sum = subtotal + value
            if new_sum > target or new_sum in dp:
                continue
            dp[new_sum] = (subtotal, idx)
            if new_sum == target:
                out = []
                current = new_sum
                while current != 0:
                    previous, used_idx = dp[current]
                    out.append(used_idx)
                    current = previous
                out.reverse()
                return out
    return None


def auto_timeout(t1_count):
    if t1_count <= 80:
        return 30
    if t1_count <= 150:
        return 20
    if t1_count <= 300:
        return 15
    return 10


def solve_global_exact(counts_by_value, targets, timeout_s=None, combo_limit_per_target=400):
    start = time.time()
    values_desc = sorted([value for value in counts_by_value.keys() if counts_by_value[value] > 0], reverse=True)

    def timed_out():
        return timeout_s is not None and (time.time() - start) > timeout_s

    def pack_counts(counter):
        return tuple(counter.get(value, 0) for value in values_desc)

    dead = set()

    def gen_combos_for_target(target, counter):
        out = []
        if timed_out():
            return out

        max_possible = sum(value * counter.get(value, 0) for value in values_desc if value <= target)
        if max_possible < target:
            return out

        combo = []

        def dfs(index, remaining):
            if timed_out() or len(out) >= combo_limit_per_target:
                return
            if remaining == 0:
                out.append(combo.copy())
                return
            if index >= len(values_desc):
                return

            value = values_desc[index]
            count = counter.get(value, 0)
            if count <= 0 or value > remaining:
                dfs(index + 1, remaining)
                return

            max_here = sum(v * counter.get(v, 0) for v in values_desc[index:] if v <= remaining)
            if max_here < remaining:
                return

            max_take = min(count, remaining // value)
            for take in range(max_take, -1, -1):
                if timed_out():
                    return
                if take:
                    combo.extend([value] * take)
                    counter[value] -= take
                dfs(index + 1, remaining - value * take)
                if take:
                    counter[value] += take
                    del combo[-take:]

        dfs(0, target)
        out.sort(key=lambda values: (len(values), [-value for value in values]))
        return out

    targets_sorted = sorted(targets, key=lambda item: item["cartons"], reverse=True)
    assignment = {}

    def rec(position, counter):
        if timed_out():
            return False, "timeout"

        key = (position, pack_counts(counter))
        if key in dead:
            return False, "no_solution"

        if position >= len(targets_sorted):
            return True, "ok"

        target_row = targets_sorted[position]
        target = target_row["cartons"]
        combos = gen_combos_for_target(target, counter)

        if timed_out():
            return False, "timeout"
        if not combos:
            dead.add(key)
            return False, "no_solution"

        for combo in combos:
            if timed_out():
                return False, "timeout"

            ok = True
            for value in combo:
                if counter.get(value, 0) <= 0:
                    ok = False
                    break
                counter[value] -= 1

            if not ok:
                for value in combo:
                    counter[value] += 1
                continue

            assignment[target_row["t1_num"]] = combo
            good, status = rec(position + 1, counter)
            if good:
                return True, "ok"
            if status == "timeout":
                return False, "timeout"

            del assignment[target_row["t1_num"]]
            for value in combo:
                counter[value] += 1

        dead.add(key)
        return False, "no_solution"

    counter_start = Counter(counts_by_value)
    good, status = rec(0, counter_start)
    if good:
        return assignment, "ok"
    return None, status


def split_stem_prefix_num(filename):
    match = re.match(r"^(?P<prefix>.+)-(?P<num>\d+)$", Path(filename).stem)
    if not match:
        return None
    return match.group("prefix"), int(match.group("num"))


def build_suggestion(file_info, lines, unused, failed):
    failed_values = [item["cartons"] for item in failed]
    failed_sum = sum(failed_values)
    max_failed = max(failed_values) if failed_values else 0
    unused_by_file = defaultdict(int)
    for idx in unused:
        unused_by_file[lines[idx]["file"]] += lines[idx]["cartons"]

    rows = []
    for file_name, info in file_info.items():
        cartons_list = info.get("cartons_list") or []
        if not cartons_list or not info.get("split_allowed"):
            continue

        total_file = sum(cartons_list)
        max_file = max(cartons_list)
        unused_file = unused_by_file.get(file_name, 0)
        cover_count = sum(1 for target in failed_values if target <= total_file)
        can_cover_all = failed_sum > 0 and total_file >= failed_sum

        score = 0
        if can_cover_all:
            score += 10_000_000
            reason = f"Can cover failed total {failed_sum}"
        else:
            score += cover_count * 2_000_000
            if max_failed and total_file >= max_failed:
                score += 500_000
            score += total_file * 1000
            score += max_file * 200
            score += unused_file * 5000
            reason = f"Split candidate: total={total_file}, unused={unused_file}, covers={cover_count}/{len(failed_values)}"

        rows.append(
            {
                "File": file_name,
                "SplitAllowed": "YES",
                "Reason": reason,
                "MergedInG": "YES" if info.get("merged_in_g") else "NO",
                "Score": score,
            }
        )

    rows.sort(key=lambda item: (-item["Score"], natural_key(item["File"])))
    return pd.DataFrame(rows[:10])


def analyze_t1_match(t1_rows, invoice_entries):
    t1_lines = [{"t1_num": int(pos), "cartons": int(cartons)} for pos, cartons in t1_rows]
    if not t1_lines:
        return {"success": False, "message": "No T1 rows to analyze."}

    lines, file_info, read_notes = read_packing_lines(invoice_entries)
    total_pack = sum(line["cartons"] for line in lines)
    total_t1 = sum(line["cartons"] for line in t1_lines)
    totals_match = total_pack == total_t1

    items_by_value = defaultdict(list)
    counts_by_value = Counter()
    for idx, item in enumerate(lines):
        value = item["cartons"]
        counts_by_value[value] += 1
        items_by_value[value].append(idx)

    assignments = []
    failed = []
    used_indices = set()
    status = "not_run"

    if totals_match:
        pre_assignment = {}
        remaining_targets = []
        tmp_counts = Counter(counts_by_value)

        for target in sorted(t1_lines, key=lambda item: item["cartons"], reverse=True):
            value = target["cartons"]
            if tmp_counts.get(value, 0) > 0:
                tmp_counts[value] -= 1
                pre_assignment[target["t1_num"]] = [value]
            else:
                remaining_targets.append(target)

        solution, status = solve_global_exact(
            counts_by_value=tmp_counts,
            targets=remaining_targets,
            timeout_s=auto_timeout(len(t1_lines)),
            combo_limit_per_target=400,
        )

        assignment_values = None
        if status == "ok":
            assignment_values = dict(pre_assignment)
            assignment_values.update(solution)

        if assignment_values is not None:
            items_pool = {value: ids.copy() for value, ids in items_by_value.items()}
            ok_all = True

            for target in t1_lines:
                values = assignment_values.get(target["t1_num"])
                if not values:
                    ok_all = False
                    failed.append(target)
                    assignments.append({"t1_num": target["t1_num"], "cartons": target["cartons"], "lines": None})
                    continue

                picked = []
                for value in values:
                    if not items_pool.get(value):
                        ok_all = False
                        break
                    picked.append(items_pool[value].pop())

                if not ok_all:
                    failed.append(target)
                    assignments.append({"t1_num": target["t1_num"], "cartons": target["cartons"], "lines": None})
                    continue

                for picked_idx in picked:
                    used_indices.add(picked_idx)
                assignments.append({"t1_num": target["t1_num"], "cartons": target["cartons"], "lines": picked})
        else:
            status = status or "no_solution"

    if not totals_match or status != "ok":
        used_indices = set()
        assignments = []
        failed = []
        for target in sorted(t1_lines, key=lambda item: item["cartons"]):
            available = [idx for idx in range(len(lines)) if idx not in used_indices]
            picked = subset_dp_find(target["cartons"], lines, available)
            if picked is None:
                failed.append(target)
                assignments.append({"t1_num": target["t1_num"], "cartons": target["cartons"], "lines": None})
                continue
            for picked_idx in picked:
                used_indices.add(picked_idx)
            assignments.append({"t1_num": target["t1_num"], "cartons": target["cartons"], "lines": picked})

    unused = [idx for idx in range(len(lines)) if idx not in used_indices]
    success = bool(lines) and totals_match and not failed and not unused and status == "ok"

    result_rows = []
    for assignment in assignments:
        if not assignment.get("lines"):
            continue
        for line_idx in assignment["lines"]:
            line = lines[line_idx]
            result_rows.append(
                {
                    "T1_num": assignment["t1_num"],
                    "Cartons_T1": assignment["cartons"],
                    "File": line["file"],
                    "Sheet": line["sheet"],
                    "Row": line["row"],
                    "Cartons_Line": line["cartons"],
                }
            )

    unused_rows = [
        {
            "File": lines[idx]["file"],
            "Sheet": lines[idx]["sheet"],
            "Row": lines[idx]["row"],
            "Cartons": lines[idx]["cartons"],
        }
        for idx in unused
    ]
    failed_rows = [{"T1_num": item["t1_num"], "Cartons": item["cartons"]} for item in failed]

    suggest_df = build_suggestion(file_info, lines, unused, failed)
    message = "T1 and packing cartons match." if success else "T1 and packing cartons do not fully match."
    if not totals_match:
        message = f"Total cartons differ: packing={total_pack}, T1={total_t1}."

    return {
        "success": success,
        "message": message,
        "status": status,
        "total_pack": total_pack,
        "total_t1": total_t1,
        "result_df": pd.DataFrame(result_rows),
        "unused_df": pd.DataFrame(unused_rows),
        "failed_df": pd.DataFrame(failed_rows),
        "suggest_df": suggest_df,
        "read_notes_df": pd.DataFrame(read_notes),
    }


def set_general_format(ws, columns, start_row, end_row_excluded):
    for row in range(start_row, end_row_excluded):
        for col in columns:
            ws[f"{col}{row}"].number_format = "General"


def defuse_packing_g(wb):
    ws = sheet_by_name_ci(wb, "PACKING LIST")
    if ws is None:
        return 0

    sum_row = find_sum_row(ws, start_row=6, label_col="B")
    if not sum_row:
        return 0

    changed = 0
    for merged_range in list(ws.merged_cells.ranges):
        if not (merged_range.min_col <= 7 <= merged_range.max_col):
            continue
        if merged_range.max_row < 6 or merged_range.min_row >= sum_row:
            continue

        start_row = max(merged_range.min_row, 6)
        end_row = min(merged_range.max_row, sum_row - 1)
        if start_row > end_row:
            continue

        original_value = get_effective_cell_value(ws, start_row, 7)
        ws.unmerge_cells(str(merged_range))
        ws.cell(row=start_row, column=7).value = original_value

        for row in range(start_row + 1, end_row + 1):
            ws.cell(row=row, column=7).value = 0

        changed += 1

    return changed


def workbook_sheet_parts(xlsx_bytes):
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
        workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_target = {}
    for rel in rels_root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if not rid or not target:
            continue
        if target.startswith("/"):
            part = target.lstrip("/")
        else:
            part = "xl/" + target.lstrip("/")
        rid_to_target[rid] = part

    sheet_parts = {}
    for sheet in workbook_root.findall("main:sheets/main:sheet", XML_NS):
        name = sheet.attrib.get("name", "").strip()
        rid = sheet.attrib.get(f"{{{OOXML_REL_NS}}}id")
        part = rid_to_target.get(rid)
        if name and part:
            sheet_parts[name] = part

    return sheet_parts


def formula_cached_values_by_part(xlsx_bytes):
    sheet_parts = workbook_sheet_parts(xlsx_bytes)
    cached_by_part = {}

    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
        for _sheet_name, part in sheet_parts.items():
            if part not in zf.namelist():
                continue

            root = ET.fromstring(zf.read(part))
            cached = {}
            for cell in root.findall(".//main:c", XML_NS):
                formula = cell.find("main:f", XML_NS)
                value = cell.find("main:v", XML_NS)
                cell_ref = cell.attrib.get("r")
                if formula is None or value is None or value.text in (None, "") or not cell_ref:
                    continue

                cached[cell_ref] = {
                    "value": value.text,
                    "type": cell.attrib.get("t"),
                }

            cached_by_part[part] = cached

    return sheet_parts, cached_by_part


def restore_formula_cached_values(source_xlsx_bytes, modified_xlsx_bytes):
    try:
        source_sheet_parts, source_cached_by_part = formula_cached_values_by_part(source_xlsx_bytes)
        target_sheet_parts = workbook_sheet_parts(modified_xlsx_bytes)
        target_part_to_source_part = {}

        for sheet_name, target_part in target_sheet_parts.items():
            source_part = source_sheet_parts.get(sheet_name)
            if source_part:
                target_part_to_source_part[target_part] = source_part

        if not target_part_to_source_part:
            return modified_xlsx_bytes, 0, "No matching worksheet parts found for formula cache restore."

        restored = 0
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(modified_xlsx_bytes), "r") as zin:
            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    source_part = target_part_to_source_part.get(item.filename)

                    if source_part:
                        cached_values = source_cached_by_part.get(source_part, {})
                        if cached_values:
                            root = ET.fromstring(data)
                            changed = 0

                            for cell in root.findall(".//main:c", XML_NS):
                                formula = cell.find("main:f", XML_NS)
                                cell_ref = cell.attrib.get("r")
                                cached = cached_values.get(cell_ref)
                                if formula is None or not cached:
                                    continue

                                value = cell.find("main:v", XML_NS)
                                if value is None:
                                    value = ET.SubElement(cell, f"{{{OOXML_NS}}}v")
                                value.text = cached["value"]

                                if cached["type"]:
                                    cell.attrib["t"] = cached["type"]
                                elif cell.attrib.get("t") == "str":
                                    cell.attrib.pop("t", None)

                                changed += 1

                            if changed:
                                restored += changed
                                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

                    zout.writestr(item, data)

        return output.getvalue(), restored, ""
    except Exception as exc:
        return modified_xlsx_bytes, 0, f"Formula cache restore skipped: {exc}"


def apply_t1_numbers(wb, file_name, result_df):
    if result_df.empty:
        return 0

    file_rows = result_df[result_df["File"] == file_name]
    if file_rows.empty:
        return 0

    ws_pack = sheet_by_name_ci(wb, "PACKING LIST")
    ws_inv = sheet_by_name_ci(wb, "INVOICE")
    if ws_pack is None or ws_inv is None:
        return 0

    writes = 0
    for _, row in file_rows.iterrows():
        t1_num = int(row["T1_num"])
        pack_row = int(row["Row"])
        row_start = pack_row
        row_end = pack_row
        for rng in ws_pack.merged_cells.ranges:
            if rng.min_row <= pack_row <= rng.max_row and rng.min_col <= 7 <= rng.max_col:
                row_start = int(rng.min_row)
                row_end = int(rng.max_row)
                break

        for rr in range(row_start, row_end + 1):
            ws_pack.cell(row=rr, column=1).value = t1_num
            inv_row = 20 + (rr - 6)
            ws_inv.cell(row=inv_row, column=1).value = t1_num
            writes += 1

    return writes


def apply_add_info(wb, file_name, mrn, date_info, ship_name):
    ws = sheet_by_name_ci(wb, "INVOICE")
    if ws is None:
        return ["INVOICE sheet missing; add-info skipped."]

    notes = []
    sum_row_invoice = find_sum_row(ws, start_row=19, label_col="B")
    if sum_row_invoice:
        set_general_format(ws, ["D", "I", "J", "K"], 20, sum_row_invoice)
    else:
        notes.append("INVOICE SUM row not found for number format.")

    ws_pl = sheet_by_name_ci(wb, "PACKING LIST")
    if ws_pl:
        sum_row_pl = find_sum_row(ws_pl, start_row=6, label_col="B")
        if sum_row_pl:
            set_general_format(ws_pl, ["C", "F", "I", "J"], 6, sum_row_pl)
        else:
            notes.append("PACKING LIST SUM row not found for number format.")
    else:
        notes.append("PACKING LIST sheet missing for number format.")

    old_j15 = ws["J15"].value
    if old_j15 is not None:
        ws["L15"] = old_j15
    ws["J15"] = DESTINATION

    ws["L4"] = mrn
    ws["L5"] = date_info
    ws["C12"] = ship_name
    ws["J17"] = "accountant@athinalogi.com"
    ws["J4"] = ws["C5"].value

    return notes


def process_modified_invoices(invoice_entries, analysis, mrn, date_info, ship_name, thc):
    result_df = analysis["result_df"]
    vat_result = compute_vat_charge(invoice_entries, thc)
    modified_files = []
    modification_rows = []

    for entry in sorted(invoice_entries, key=lambda item: natural_key(item["name"])):
        file_name = entry["name"]

        try:
            wb = load_workbook(io.BytesIO(entry["bytes"]))
            t1_writes = apply_t1_numbers(wb, file_name, result_df)
            defused = defuse_packing_g(wb)
            notes = apply_add_info(wb, file_name, mrn, date_info, ship_name)

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            file_bytes, restored_count, restore_note = restore_formula_cached_values(entry["bytes"], output.getvalue())
            if restored_count:
                notes.append(f"Formula cached values restored: {restored_count}")
            if restore_note:
                notes.append(restore_note)

            modified_files.append((file_name, file_bytes))
            modification_rows.append(
                {
                    "File": file_name,
                    "Status": "Modified",
                    "Message": "; ".join(notes) if notes else "",
                    "T1 writes": t1_writes,
                    "Defused G ranges": defused,
                    "MRN": mrn,
                    "Date": date_info,
                    "Ship Name": ship_name,
                    "Destination": DESTINATION,
                }
            )
        except Exception as exc:
            modification_rows.append({"File": file_name, "Status": "Error", "Message": str(exc)})

    return {
        "modified_files": modified_files,
        "modification_df": pd.DataFrame(modification_rows),
        **vat_result,
    }


def parse_decimal_input(label, value):
    text = str(value or "").strip().replace(",", ".")
    if not text:
        raise ValueError(f"{label} is required.")
    try:
        return Decimal(text)
    except InvalidOperation:
        raise ValueError(f"{label} must be a valid number.")


def read_gross_from_invoice(file_bytes):
    try:
        wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True)
        wb_formulas = load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws_values = sheet_by_name_ci(wb_values, "INVOICE")
        ws_formula = sheet_by_name_ci(wb_formulas, "INVOICE")
        if ws_values is None:
            return None

        sum_row = find_sum_row_resolved(
            wb_values,
            wb_formulas,
            ws_values,
            ws_formula,
            start_row=19,
            label_col="B",
        )
        if not sum_row:
            return None

        gross = to_decimal(
            resolve_cell_ref_value(wb_values, wb_formulas, ws_values, ws_formula, f"K{sum_row}")
        )
        if gross is not None:
            return gross

        calculated = Decimal("0")
        found = 0
        for row in range(20, sum_row):
            value = to_decimal(
                resolve_cell_ref_value(wb_values, wb_formulas, ws_values, ws_formula, f"K{row}")
            )
            if value is not None:
                calculated += value
                found += 1
        return calculated if found else None
    except Exception:
        return None


def build_vat_csv(rows_for_csv):
    output = io.StringIO()
    writer = csv.writer(output, delimiter=";", lineterminator="\n")
    for file_name, vat_value in rows_for_csv:
        writer.writerow([file_name, vat_value])
    return output.getvalue().encode("utf-8")


def compute_vat_charge(invoice_entries, thc):
    base = VAT_BASE_AMOUNT + thc
    gross_by_file = {}
    total_gross = Decimal("0")
    skipped_rows = []

    for entry in sorted(invoice_entries, key=lambda item: natural_key(item["name"])):
        gross = read_gross_from_invoice(entry["bytes"])
        if gross is None:
            skipped_rows.append(
                {
                    "File": entry["name"],
                    "Mode": VAT_MODE,
                    "Gross Weight": "",
                    "Coefficient": "",
                    "Addition": "",
                    "VAT Charge": "",
                    "Status": "Skipped",
                    "Message": "Gross weight not found in INVOICE SUM row.",
                }
            )
            continue
        gross_by_file[entry["name"]] = gross
        total_gross += gross

    if total_gross == 0:
        raise ValueError("Total gross weight is 0. VAT charge cannot be calculated.")

    coefficient = (base / total_gross).quantize(Decimal("0.000001"))
    vat_rows = []
    rows_for_csv = []

    ordered_items = [(name, gross_by_file[name]) for name in sorted(gross_by_file.keys(), key=natural_key)]
    for index, (file_name, gross) in enumerate(ordered_items, start=1):
        addition = VAT_FIRST_ADDITION if index == 1 else VAT_OTHER_ADDITION
        vat = (gross * coefficient).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) + addition
        vat_text = f"{vat:.2f}"
        rows_for_csv.append((file_name, vat_text))
        vat_rows.append(
            {
                "File": file_name,
                "Mode": VAT_MODE,
                "Gross Weight": str(gross),
                "Coefficient": str(coefficient),
                "Addition": str(addition),
                "VAT Charge": vat_text,
                "Status": "OK",
                "Message": "",
            }
        )

    return {
        "vat_df": pd.DataFrame(vat_rows + skipped_rows),
        "vat_csv_bytes": build_vat_csv(rows_for_csv),
        "vat_base": base,
        "vat_coefficient": coefficient,
        "vat_total_gross": total_gross,
    }


def build_report_workbook(final_check_df, issue_df, t1_df, analysis, processing=None, rename_df=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if rename_df is not None and not rename_df.empty:
            rename_df.to_excel(writer, index=False, sheet_name="File_Rename")
        final_check_df.to_excel(writer, index=False, sheet_name="Final_Check")
        issue_df.to_excel(writer, index=False, sheet_name="Issues")
        t1_df.to_excel(writer, index=False, sheet_name="T1_Extract")
        analysis.get("result_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="T1_Result")
        analysis.get("failed_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="T1_Failed")
        analysis.get("unused_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="Packing_Unused")
        analysis.get("suggest_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="T1_Suggest")
        if processing:
            processing.get("vat_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="VAT_Charge")
            processing.get("modification_df", pd.DataFrame()).to_excel(writer, index=False, sheet_name="Modifications")
    output.seek(0)
    return output.getvalue()


def build_zip(modified_files, report_bytes, vat_csv_bytes=None):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, data in modified_files:
            zf.writestr(f"modified_invoices/{file_name}", data)
        if vat_csv_bytes:
            zf.writestr(VAT_CSV_NAME, vat_csv_bytes)
        zf.writestr("EORIGIN_processing_report.xlsx", report_bytes)
    output.seek(0)
    return output.getvalue()


def read_secret(name, default=""):
    try:
        value = st.secrets.get(name, default)
    except Exception:
        return default
    return default if value is None else str(value).strip()


def response_error_message(response):
    try:
        payload = response.json()
        if isinstance(payload, dict):
            return payload.get("reason") or payload.get("message") or str(payload)
        return str(payload)
    except ValueError:
        return response.text.strip()[:500] or response.reason


def normalize_bearer_token(token):
    token = str(token or "").strip()
    if token.lower().startswith("bearer "):
        return token.split(None, 1)[1].strip()
    return token


def authenticate_eorigin(email, password):
    response = requests.post(
        f"{EORIGIN_API_BASE_URL}/authenticate",
        json={"email": email, "password": password},
        timeout=30,
    )
    if not response.ok:
        raise RuntimeError(f"Authentication failed: {response_error_message(response)}")

    try:
        payload = response.json()
    except ValueError as exc:
        raise RuntimeError("Authentication response is not valid JSON.") from exc

    token = ""
    if isinstance(payload, dict):
        token = payload.get("token") or payload.get("accessToken") or payload.get("access_token") or ""
    token = normalize_bearer_token(token)
    if not token:
        raise RuntimeError("Authentication succeeded, but no token was returned.")
    return token


def build_eorigin_upload_file(modified_files):
    if len(modified_files) == 1:
        file_name, file_bytes = modified_files[0]
        return (
            file_name,
            file_bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, file_bytes in modified_files:
            zf.writestr(file_name, file_bytes)
    output.seek(0)
    return "EORIGIN_modified_invoices_upload.zip", output.getvalue(), "application/zip"


def upload_modified_invoices_to_eorigin(modified_files, batch_name, customer_id, template_id, token, email, password):
    if not modified_files:
        raise ValueError("No modified invoice available to upload.")

    token = normalize_bearer_token(token)
    if not token:
        token = authenticate_eorigin(email, password)

    upload_name, upload_bytes, upload_mime = build_eorigin_upload_file(modified_files)
    response = requests.post(
        f"{EORIGIN_API_BASE_URL}/upload-batch",
        headers={"Authorization": f"Bearer {token}"},
        data={
            "name": batch_name,
            "customer": customer_id,
            "template": template_id,
        },
        files={"file": (upload_name, upload_bytes, upload_mime)},
        timeout=EORIGIN_UPLOAD_TIMEOUT,
    )
    if not response.ok:
        raise RuntimeError(f"E-Origin upload failed: {response_error_message(response)}")

    message = "Upload started on E-Origin."
    try:
        payload = response.json()
        if isinstance(payload, dict):
            message = payload.get("message") or message
        elif payload:
            message = str(payload)
    except ValueError:
        if response.text.strip():
            message = response.text.strip()[:500]

    return {"message": message, "file_name": upload_name}


def render_eorigin_upload_box(processing, default_batch_name):
    st.subheader("E-Origin Upload")
    st.caption("Uploads the modified invoice file(s) with the E-Origin external API.")

    with st.expander("Upload modified invoices to E-Origin"):
        st.caption(
            "For Streamlit Cloud, prefer setting EORIGIN_TOKEN or EORIGIN_EMAIL / "
            "EORIGIN_PASSWORD plus EORIGIN_CUSTOMER_ID and EORIGIN_TEMPLATE_ID in app secrets."
        )

        batch_name = st.text_input("Batch name", value=default_batch_name, key="eorigin_batch_name")

        c1, c2 = st.columns(2)
        customer_id_input = c1.text_input(
            "Customer ID",
            placeholder="Configured in secrets" if read_secret("EORIGIN_CUSTOMER_ID") else "Required",
            key="eorigin_customer_id",
        )
        template_id_input = c2.text_input(
            "Template ID",
            placeholder="Configured in secrets" if read_secret("EORIGIN_TEMPLATE_ID") else "Required",
            key="eorigin_template_id",
        )

        token_input = st.text_input(
            "Bearer token",
            type="password",
            placeholder="Configured in secrets or leave empty to use email/password",
            key="eorigin_token",
        )
        c1, c2 = st.columns(2)
        email_input = c1.text_input(
            "E-Origin email",
            placeholder="Configured in secrets" if read_secret("EORIGIN_EMAIL") else "Required if no token",
            key="eorigin_email",
        )
        password_input = c2.text_input(
            "E-Origin password",
            type="password",
            placeholder="Configured in secrets" if read_secret("EORIGIN_PASSWORD") else "Required if no token",
            key="eorigin_password",
        )

        customer_id = customer_id_input.strip() or read_secret("EORIGIN_CUSTOMER_ID")
        template_id = template_id_input.strip() or read_secret("EORIGIN_TEMPLATE_ID")
        token = token_input.strip() or read_secret("EORIGIN_TOKEN")
        email = email_input.strip() or read_secret("EORIGIN_EMAIL")
        password = password_input or read_secret("EORIGIN_PASSWORD")

        if st.button("Upload modified invoices to E-Origin", type="primary"):
            missing = []
            if not batch_name.strip():
                missing.append("Batch name")
            if not customer_id:
                missing.append("Customer ID")
            if not template_id:
                missing.append("Template ID")
            if not token and (not email or not password):
                missing.append("Bearer token or E-Origin email/password")

            if missing:
                st.error("Missing: " + ", ".join(missing))
            else:
                try:
                    with st.spinner("Uploading to E-Origin..."):
                        upload_result = upload_modified_invoices_to_eorigin(
                            modified_files=processing["modified_files"],
                            batch_name=batch_name.strip(),
                            customer_id=customer_id,
                            template_id=template_id,
                            token=token,
                            email=email,
                            password=password,
                        )
                    st.session_state["eorigin_upload_result"] = upload_result
                    st.success(f"{upload_result['message']} File sent: {upload_result['file_name']}")
                except requests.RequestException as exc:
                    st.error(f"E-Origin connection error: {exc}")
                except Exception as exc:
                    st.error(str(exc))

        upload_result = st.session_state.get("eorigin_upload_result")
        if upload_result:
            st.success(f"Last upload: {upload_result['message']}")


def build_summary_frames(check_results):
    final_rows = []
    issue_rows = []

    for result in check_results:
        final_rows.append(
            {
                "File": result["File"],
                "Status": result["Status"],
                "Errors": len(result["Errors"]),
                "Warnings": len(result["Warnings"]),
                "Cartons": result["Cartons"],
                "Gross Weight": result["Gross Weight"],
            }
        )
        for message in result["Errors"]:
            issue_rows.append({"File": result["File"], "Type": "ERROR", "Message": message})
        for message in result["Warnings"]:
            issue_rows.append({"File": result["File"], "Type": "WARNING", "Message": message})

    return pd.DataFrame(final_rows), pd.DataFrame(issue_rows)


def run_processing(invoice_files, t1_pdf, ship_name, thc_text, mrn_override, date_override, allow_errors):
    invoice_entries = build_invoice_entries(invoice_files)
    rename_df = build_rename_frame(invoice_entries)
    thc = parse_decimal_input("THC", thc_text)

    check_results = [final_check_invoice(entry["name"], entry["bytes"]) for entry in invoice_entries]
    final_check_df, issue_df = build_summary_frames(check_results)

    t1_extract = extract_t1_from_pdf(t1_pdf.getvalue())
    t1_df = pd.DataFrame(t1_extract["rows"], columns=["T1 Position", "Cartons"])

    mrn = (mrn_override or "").strip() or t1_extract.get("t1_no", "")
    date_override = (date_override or "").strip()
    date_info = normalize_t1_date(date_override) if date_override else t1_extract.get("t1_date", "")
    analysis = analyze_t1_match(t1_extract["rows"], invoice_entries) if t1_extract["ok"] else {
        "success": False,
        "message": "T1 extraction confidence is too low.",
        "result_df": pd.DataFrame(),
        "failed_df": pd.DataFrame(),
        "unused_df": pd.DataFrame(),
        "suggest_df": pd.DataFrame(),
    }

    return {
        "invoice_entries": invoice_entries,
        "rename_df": rename_df,
        "ship_name": ship_name,
        "thc": thc,
        "final_check_df": final_check_df,
        "issue_df": issue_df,
        "t1_extract": t1_extract,
        "t1_df": t1_df,
        "mrn": mrn,
        "date_info": date_info,
        "analysis": analysis,
        "has_final_errors": any(result["Errors"] for result in check_results),
        "allow_errors": allow_errors,
    }


def render_fixed_values():
    st.sidebar.divider()
    st.sidebar.markdown("### Fixed values")
    st.sidebar.caption(f"Destination: {DESTINATION}")
    st.sidebar.caption(f"VAT mode: {VAT_MODE}")
    st.sidebar.caption("MRN: detected from uploaded T1")
    st.sidebar.caption("Date: detected from uploaded T1")


def main():
    configure_page()
    render_fixed_values()

    st.title(APP_TITLE)
    st.caption("Upload invoice Excel files and one T1 PDF, then generate a ZIP with modified E-origin invoices.")

    invoice_files = st.file_uploader(
        "Upload invoice Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
    )
    t1_pdf = st.file_uploader("Upload T1 PDF", type=["pdf"], accept_multiple_files=False)

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
    ship_name = c1.text_input("Ship name")
    thc_text = c2.text_input("THC", placeholder="Required")
    mrn_override = c3.text_input("MRN override", placeholder="Optional")
    date_override = c4.text_input("Date override", placeholder="Optional DD/MM/YYYY")

    allow_errors = st.checkbox("Allow output even if Final Check has errors", value=False)
    run = st.button("Run full processing", type="primary")

    if not run:
        st.info("Upload the invoice files and the T1 PDF to start.")
        return

    if not invoice_files:
        st.error("Upload at least one invoice Excel file.")
        return
    if t1_pdf is None:
        st.error("Upload one T1 PDF.")
        return
    if not ship_name.strip():
        st.error("Enter the ship name.")
        return
    if not thc_text.strip():
        st.error("Enter the THC amount.")
        return

    try:
        with st.spinner("Processing files..."):
            result = run_processing(
                invoice_files=invoice_files,
                t1_pdf=t1_pdf,
                ship_name=ship_name.strip(),
                thc_text=thc_text,
                mrn_override=mrn_override,
                date_override=date_override,
                allow_errors=allow_errors,
            )
    except Exception as exc:
        st.error(f"Processing error: {exc}")
        return

    final_check_df = result["final_check_df"]
    issue_df = result["issue_df"]
    rename_df = result["rename_df"]
    t1_extract = result["t1_extract"]
    t1_df = result["t1_df"]
    analysis = result["analysis"]

    if not rename_df.empty:
        st.subheader("File Names")
        st.dataframe(rename_df, use_container_width=True, hide_index=True)

    st.subheader("Final Check")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Files checked", len(final_check_df))
    c2.metric("Files with errors", int((final_check_df["Errors"] > 0).sum()))
    c3.metric("Files with warnings", int((final_check_df["Warnings"] > 0).sum()))
    c4.metric("Total cartons", f"{final_check_df['Cartons'].sum():,.0f}")
    st.dataframe(final_check_df, use_container_width=True, hide_index=True)

    if not issue_df.empty:
        st.dataframe(issue_df, use_container_width=True, hide_index=True)
    else:
        st.success("No Final Check issue found.")

    st.subheader("T1 Extract")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("T1 rows", len(t1_df))
    c2.metric("Confidence", f"{t1_extract['score']:.3f}")
    c3.metric("MRN", result["mrn"] or "Not found")
    c4.metric("Date", result["date_info"] or "Not found")
    st.caption(f"Parser: {t1_extract['parser']} | {t1_extract['notes']}")

    if not t1_extract["ok"]:
        st.error("T1 extraction confidence is too low. The app did not modify invoices.")
        report_bytes = build_report_workbook(final_check_df, issue_df, t1_df, analysis, rename_df=rename_df)
        st.download_button(
            "Download report Excel",
            data=report_bytes,
            file_name="EORIGIN_processing_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return

    st.dataframe(t1_df, use_container_width=True, hide_index=True)

    st.subheader("T1 Analyzer")
    c1, c2, c3 = st.columns(3)
    c1.metric("Packing cartons", analysis.get("total_pack", 0))
    c2.metric("T1 cartons", analysis.get("total_t1", 0))
    c3.metric("Match", "YES" if analysis["success"] else "NO")

    if analysis["success"]:
        st.success(analysis["message"])
        st.dataframe(analysis["result_df"], use_container_width=True, hide_index=True)
    else:
        st.error(analysis["message"])
        if not analysis.get("failed_df", pd.DataFrame()).empty:
            st.markdown("Failed T1 lines")
            st.dataframe(analysis["failed_df"], use_container_width=True, hide_index=True)
        if not analysis.get("unused_df", pd.DataFrame()).empty:
            st.markdown("Unused packing lines")
            st.dataframe(analysis["unused_df"], use_container_width=True, hide_index=True)
        if not analysis.get("suggest_df", pd.DataFrame()).empty:
            st.markdown("Suggested files to split")
            st.dataframe(analysis["suggest_df"], use_container_width=True, hide_index=True)

        report_bytes = build_report_workbook(final_check_df, issue_df, t1_df, analysis, rename_df=rename_df)
        st.download_button(
            "Download report Excel",
            data=report_bytes,
            file_name="EORIGIN_processing_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return

    if result["has_final_errors"] and not result["allow_errors"]:
        st.error("Final Check has errors. Fix the files or tick the checkbox to force output.")
        report_bytes = build_report_workbook(final_check_df, issue_df, t1_df, analysis, rename_df=rename_df)
        st.download_button(
            "Download report Excel",
            data=report_bytes,
            file_name="EORIGIN_processing_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return

    if not result["mrn"]:
        st.error("MRN was not detected. Fill the MRN override field and run again.")
        report_bytes = build_report_workbook(final_check_df, issue_df, t1_df, analysis, rename_df=rename_df)
        st.download_button(
            "Download report Excel",
            data=report_bytes,
            file_name="EORIGIN_processing_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return

    if not result["date_info"]:
        st.error("T1 date was not detected. Fill the date override field and run again.")
        report_bytes = build_report_workbook(final_check_df, issue_df, t1_df, analysis, rename_df=rename_df)
        st.download_button(
            "Download report Excel",
            data=report_bytes,
            file_name="EORIGIN_processing_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return

    try:
        with st.spinner("Modifying invoices and building ZIP..."):
            processing = process_modified_invoices(
                invoice_entries=result["invoice_entries"],
                analysis=analysis,
                mrn=result["mrn"],
                date_info=result["date_info"],
                ship_name=result["ship_name"],
                thc=result["thc"],
            )
            report_bytes = build_report_workbook(final_check_df, issue_df, t1_df, analysis, processing, rename_df)
            zip_bytes = build_zip(processing["modified_files"], report_bytes, processing["vat_csv_bytes"])
    except Exception as exc:
        st.error(f"Modification error: {exc}")
        return

    st.subheader("Output")
    st.dataframe(processing["modification_df"], use_container_width=True, hide_index=True)
    st.subheader("VAT Charge")
    st.caption(
        f"Mode: {VAT_MODE} | Base: {VAT_BASE_AMOUNT} + THC {result['thc']} = {processing['vat_base']} | "
        f"Total gross: {processing['vat_total_gross']} | Coefficient: {processing['vat_coefficient']}"
    )
    st.dataframe(processing["vat_df"], use_container_width=True, hide_index=True)

    if processing["modified_files"]:
        c1, c2 = st.columns(2)
        c1.download_button(
            "Download modified invoices ZIP",
            data=zip_bytes,
            file_name="EORIGIN_modified_invoices.zip",
            mime="application/zip",
            type="primary",
        )
        c2.download_button(
            f"Download {VAT_CSV_NAME}",
            data=processing["vat_csv_bytes"],
            file_name=VAT_CSV_NAME,
            mime="text/csv",
        )
        default_batch_name = f"E-Origin {result['mrn'] or time.strftime('%Y%m%d-%H%M%S')}"
        render_eorigin_upload_box(processing, default_batch_name)
    else:
        st.error("No invoice was modified.")


if __name__ == "__main__":
    main()
