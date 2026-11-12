"""
Nexgen Analytics Dashboard
===========================
Streamlit app – fully offline, multi-sheet Excel-formula aware, 
custom dashboard designer, manual column mapping, and modern glassmorphism UI.
"""

from __future__ import annotations

import copy
import difflib
import glob
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ─────────────────────────────────────────────────────────────────────────────
# Streamlit Page Config – MUST be the very first Streamlit call
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Nexgen Analytics Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# Paths & Directories
# ─────────────────────────────────────────────────────────────────────────────
BASE = Path(__file__).parent.resolve()
DATA_DIR = BASE / "data"
TMPL_DIR = BASE / "templates"
INBOX_DIR = BASE / "inbox"
META_FILE = BASE / "nexgen_meta.json"
PRESETS_FILE = BASE / "nexgen_presets.json"
EXCEL_EXE = Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE")

for _d in (DATA_DIR, TMPL_DIR, INBOX_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# Persistent Metadata & Presets Helpers
# ─────────────────────────────────────────────────────────────────────────────
def load_meta() -> dict:
    """Load persistent metadata containing templates, active template, and alert thresholds."""
    if META_FILE.exists():
        try:
            data = json.loads(META_FILE.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
    else:
        data = {}

    data.setdefault("templates", {})
    data.setdefault("thresholds", {})
    data.setdefault("presets", {})
    data.setdefault("active_template", None)

    # Normalize existing template objects
    for tmpl_id, tmpl in data["templates"].items():
        tmpl.setdefault("target_sheets", [tmpl.get("sheet", "Dashboard")])
        tmpl.setdefault("column_mapping", {})
        tmpl.setdefault("formula_cols", {})
        tmpl.setdefault("dashboard_layout", [])

    return data


def save_meta(m: dict) -> None:
    """Save metadata securely to nexgen_meta.json."""
    META_FILE.write_text(json.dumps(m, indent=2, default=str), encoding="utf-8")


def load_presets() -> dict:
    """Load saved view presets."""
    if PRESETS_FILE.exists():
        try:
            return json.loads(PRESETS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    meta = load_meta()
    return meta.get("presets", {})


def save_preset(name: str, config: dict) -> None:
    """Save a dashboard layout / filter preset."""
    meta = load_meta()
    meta.setdefault("presets", {})
    meta["presets"][name] = config
    save_meta(meta)
    PRESETS_FILE.write_text(json.dumps(meta["presets"], indent=2, default=str), encoding="utf-8")


def delete_preset(name: str) -> None:
    """Delete a named preset."""
    meta = load_meta()
    if "presets" in meta and name in meta["presets"]:
        del meta["presets"][name]
        save_meta(meta)
        PRESETS_FILE.write_text(json.dumps(meta["presets"], indent=2, default=str), encoding="utf-8")


def clean_header(v: Any) -> str:
    """Normalise a header: fix NBSP and stray mojibake chars (e.g. 'Meeting\u00c2') and trim."""
    t = str(v).replace("\u00a0", " ").replace("\u00c2", "")
    return re.sub(r"\s+", " ", t).strip()


def scan_formula_errors(path: Path) -> Dict[str, int]:
    """Count Excel error values (#NAME?, #REF!, ...) left in a calculated workbook."""
    errs: Dict[str, int] = {}
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and v.startswith("#") and v.endswith(("!", "?", "A", "0")) and len(v) < 9:
                        errs[v] = errs.get(v, 0) + 1
        wb.close()
    except Exception:
        pass
    return errs


def show_formula_error_warning(path: Path) -> None:
    """Warn in the UI when recalculation left Excel errors (typically XLOOKUP on Excel 2016)."""
    errs = scan_formula_errors(path)
    if not errs:
        return
    summary = ", ".join(f"{k} x{v}" for k, v in sorted(errs.items()))
    msg = f"Excel left formula errors in the calculated workbook: {summary}."
    if "#NAME?" in errs:
        msg += (" #NAME? usually means the installed Excel does not support a function used in the template "
                "(XLOOKUP needs Excel 2021 / Microsoft 365; Excel 2016 does not have it).")
    st.warning(msg)


# ─────────────────────────────────────────────────────────────────────────────
# Excel Recalculation Helpers (COM Automation & Subprocess Fallback)
# ─────────────────────────────────────────────────────────────────────────────
def recalc_via_excel_com(path: Path) -> bool:
    """Recalculate workbook using MS Excel COM automation (Windows)."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client as win32

        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = excel.Workbooks.Open(str(path.resolve()))
        wb.Application.CalculateFull()
        wb.Save()
        wb.Close(False)
        excel.Quit()
        return True
    except Exception:
        try:
            excel.Quit()
        except Exception:
            pass
        return False


def recalc_via_excel_subprocess(path: Path) -> bool:
    """Recalculate workbook by invoking Excel automation via a temporary VBScript."""
    vbs = tempfile.NamedTemporaryFile(suffix=".vbs", delete=False, mode="w", encoding="utf-8")
    vbs.write(f"""
Dim xlApp
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.ScreenUpdating = False
On Error Resume Next
Dim wb
Set wb = xlApp.Workbooks.Open("{str(path.resolve())}")
If Err.Number = 0 Then
    xlApp.CalculateFull
    wb.Save
    wb.Close False
End If
xlApp.Quit
Set xlApp = Nothing
""")
    vbs.close()
    try:
        res = subprocess.run(
            ["cscript", "//NoLogo", vbs.name],
            capture_output=True,
            timeout=75,
        )
        return res.returncode == 0
    except Exception:
        return False
    finally:
        try:
            os.unlink(vbs.name)
        except Exception:
            pass


def recalculate_workbook(path: Path) -> bool:
    """Execute full workbook calculation with automatic fallback."""
    ok = recalc_via_excel_com(path)
    if not ok:
        ok = recalc_via_excel_subprocess(path)
    return ok


def get_excel_data(path: Path) -> Optional[Dict[str, pd.DataFrame]]:
    """Read all sheets from a calculated Excel file into a dict of DataFrames."""
    try:
        wb = load_workbook(path, data_only=True)
        result: Dict[str, pd.DataFrame] = {}
        for name in wb.sheetnames:
            ws = wb[name]
            rows_iter = ws.iter_rows(values_only=True)
            header_row = None
            for row in rows_iter:
                if any(cell is not None for cell in row):
                    header_row = [clean_header(c) if c is not None else f"Col_{i+1}" for i, c in enumerate(row)]
                    break
            if header_row is None:
                continue

            data_rows = []
            for row in rows_iter:
                if any(cell is not None for cell in row):
                    padded_row = list(row) + [None] * (len(header_row) - len(row))
                    data_rows.append(padded_row[:len(header_row)])

            if data_rows:
                df = pd.DataFrame(data_rows, columns=header_row)
            else:
                df = pd.DataFrame(columns=header_row)
            result[name] = df
        return result
    except Exception:
        try:
            # Fallback to pandas read_excel
            return pd.read_excel(path, sheet_name=None)
        except Exception:
            return None


# ─────────────────────────────────────────────────────────────────────────────
# Template Analysis & Multi-Sheet Inspector
# ─────────────────────────────────────────────────────────────────────────────
def analyse_template(path: Path) -> dict:
    """Inspect all sheets of a template Excel file to extract headers, formulas, and data start rows."""
    wb = load_workbook(path, data_only=False)
    schema = {"sheets": {}}

    for name in wb.sheetnames:
        ws = wb[name]
        headers: List[str] = []
        formula_cols: Dict[str, str] = {}
        data_start_row = 2

        # Search for header row in top 10 rows
        for row in ws.iter_rows(max_row=10):
            non_empty = [c for c in row if c.value is not None]
            if non_empty:
                headers = [clean_header(c.value) for c in non_empty]
                header_row_idx = row[0].row
                data_start_row = header_row_idx + 1

                # Inspect next row for formula templates
                if ws.max_row >= data_start_row:
                    for cell in ws.iter_rows(min_row=data_start_row, max_row=data_start_row, max_col=ws.max_column).__next__():
                        if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                            col_header = headers[cell.column - 1] if cell.column <= len(headers) else f"Col_{cell.column}"
                            formula_cols[col_header] = cell.value
                break

        schema["sheets"][name] = {
            "headers": headers,
            "formula_cols": formula_cols,
            "data_start_row": data_start_row,
        }

    return schema


# ─────────────────────────────────────────────────────────────────────────────
# Fuzzy String Matching for Bulk Column Mapping Suggestion
# ─────────────────────────────────────────────────────────────────────────────
def suggest_mappings(upload_cols: List[str], template_cols: List[str]) -> Dict[str, str]:
    """Suggest best matches between raw upload columns and template columns."""
    suggestions: Dict[str, str] = {}
    norm_upload = {re.sub(r"[^a-z0-9]", "", c.lower()): c for c in upload_cols}

    for t_col in template_cols:
        norm_t = re.sub(r"[^a-z0-9]", "", t_col.lower())
        # Exact normalized match
        if norm_t in norm_upload:
            suggestions[t_col] = norm_upload[norm_t]
            continue
        # Substring search
        matched = False
        for nu, original in norm_upload.items():
            if norm_t and (norm_t in nu or nu in norm_t):
                suggestions[t_col] = original
                matched = True
                break
        if matched:
            continue
        # Fuzzy match
        close = difflib.get_close_matches(norm_t, list(norm_upload.keys()), n=1, cutoff=0.55)
        if close:
            suggestions[t_col] = norm_upload[close[0]]

    return suggestions


# ─────────────────────────────────────────────────────────────────────────────
# Multi-Sheet Data Injection Engine
# ─────────────────────────────────────────────────────────────────────────────
def adjust_formula_row(formula: str, row_offset: int, base_row: int, last_row: Optional[int] = None) -> str:
    """Adjust relative cell references in standard Excel formulas when copied downward."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    # If structured reference formula like Table1[[#This Row], ...], leave unmodified
    if "[#This Row]" in formula or "[@" in formula or "Table" in formula:
        return formula

    def _replace_cell(m):
        col_part = m.group(1)
        dollar = m.group(2)
        row_str = m.group(3)
        if dollar == "$":
            return m.group(0)
        orig_row = int(row_str)
        if orig_row <= base_row - 1:
            return m.group(0)  # Keep header or upper range fixed
        return f"{col_part}{dollar}{orig_row + row_offset}"

    # Grow hardcoded absolute ranges like $C$2:$C$44086 to the real last data row
    if last_row:
        def _grow(m):
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            if r1 == base_row and r2 >= base_row:
                return f"${c1}${r1}:${c2}${last_row}"
            return m.group(0)
        formula = re.sub(r"\$([A-Za-z]{1,3})\$(\d+):\$([A-Za-z]{1,3})\$(\d+)", _grow, formula)

    pattern = r"(?<![A-Za-z0-9_])([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])"
    return re.sub(pattern, _replace_cell, formula)


def inject_multi_sheet_data(
    template_path: Path,
    upload_sources: Dict[str, pd.DataFrame],
    template_info: dict,
    dest_path: Path,
) -> bool:
    """
    Inject raw upload data across multiple target sheets according to manual column mappings.
    Preserves template formulas, extends them downward, and triggers Excel full recalculation.
    """
    shutil.copy2(template_path, dest_path)
    wb = load_workbook(dest_path, data_only=False)

    target_sheets = template_info.get("target_sheets", [])
    if not target_sheets:
        target_sheets = [template_info.get("sheet", wb.sheetnames[0])]

    column_mapping = template_info.get("column_mapping", {})
    formula_cols_cfg = template_info.get("formula_cols", {})

    for target_sheet in target_sheets:
        if target_sheet not in wb.sheetnames:
            continue
        ws = wb[target_sheet]

        # Resolve the source dataframe for this sheet
        source_df = upload_sources.get(target_sheet)
        if source_df is None and upload_sources:
            source_df = next(iter(upload_sources.values()))
        if source_df is None or source_df.empty:
            continue

        sheet_mapping = {clean_header(k): v for k, v in column_mapping.get(target_sheet, {}).items()}
        sheet_formula_cols = [clean_header(x) for x in formula_cols_cfg.get(target_sheet, [])]

        # Detect headers
        header_row = 1
        headers_map: Dict[str, int] = {}
        for r in range(1, min(10, ws.max_row + 1)):
            row_cells = [ws.cell(row=r, column=c) for c in range(1, ws.max_column + 1)]
            non_empty = [c for c in row_cells if c.value is not None]
            if len(non_empty) >= 1:
                header_row = r
                headers_map = {clean_header(c.value): c.column for c in non_empty}
                break

        data_start = header_row + 1

        # Collect formula templates from data row
        formula_row: Dict[str, str] = {}
        if ws.max_row >= data_start:
            for header, col_idx in headers_map.items():
                c_val = ws.cell(row=data_start, column=col_idx).value
                if c_val and isinstance(c_val, str) and c_val.startswith("="):
                    formula_row[header] = c_val

        # Clear old data rows
        max_r = max(ws.max_row, data_start)
        for r_idx in range(data_start, max_r + 1):
            for c_idx in range(1, ws.max_column + 1):
                ws.cell(row=r_idx, column=c_idx).value = None

        # Write new rows
        upload_cols = [clean_header(c) for c in source_df.columns]
        source_df = source_df.copy()
        source_df.columns = upload_cols
        last_row = data_start + len(source_df) - 1
        for row_offset, (_, row_data) in enumerate(source_df.iterrows()):
            curr_row = data_start + row_offset
            for header, col_idx in headers_map.items():
                mapped_val = sheet_mapping.get(header)

                # Check if this column should be a formula
                is_formula = (
                    mapped_val == "-- formula column --"
                    or header in sheet_formula_cols
                    or (header in formula_row and mapped_val in (None, "", "-- formula column --"))
                )

                if is_formula:
                    base_formula = formula_row.get(header)
                    if base_formula:
                        adj_formula = adjust_formula_row(base_formula, row_offset, data_start, last_row)
                        ws.cell(row=curr_row, column=col_idx).value = adj_formula
                elif mapped_val == "-- ignore --":
                    continue
                else:
                    # Map to upload column
                    mv = clean_header(mapped_val) if mapped_val else None
                    target_upload_col = (
                        mv
                        if (mv and mv in upload_cols)
                        else (header if header in upload_cols else None)
                    )
                    if target_upload_col is not None:
                        val = row_data.get(clean_header(target_upload_col))
                        if pd.isna(val):
                            val = None
                        elif isinstance(val, (pd.Timestamp, datetime)):
                            val = val.strftime("%Y-%m-%d %H:%M:%S")
                        elif isinstance(val, (np.integer, int)):
                            val = int(val)
                        elif isinstance(val, (np.floating, float)):
                            val = float(val)
                        ws.cell(row=curr_row, column=col_idx).value = val

        # Resize Excel Tables so structured references (Table1[[#This Row],...]) cover the new rows
        for tbl in list(ws.tables.values()):
            try:
                from openpyxl.utils.cell import range_boundaries
                min_c, min_r, max_c, _max_r = range_boundaries(tbl.ref)
                new_last = max(last_row, min_r + 1)
                tbl.ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{new_last}"
                if tbl.autoFilter is not None:
                    tbl.autoFilter.ref = tbl.ref
            except Exception:
                pass

    wb.save(dest_path)
    return recalculate_workbook(dest_path)


# ─────────────────────────────────────────────────────────────────────────────
# Daily Data Storage & Retrieval
# ─────────────────────────────────────────────────────────────────────────────
def get_date_dir(d: date) -> Path:
    """Return storage directory for a specific date, creating if necessary."""
    dd = DATA_DIR / d.strftime("%Y-%m-%d")
    dd.mkdir(parents=True, exist_ok=True)
    return dd


def save_day_data(
    d: date,
    upload_bytes: bytes,
    upload_name: str,
    calc_path: Path,
    template_info: dict,
) -> None:
    """Store calculated workbook, individual sheets, and JSON metadata."""
    dd = get_date_dir(d)
    raw_path = dd / f"raw{Path(upload_name).suffix}"
    raw_path.write_bytes(upload_bytes)

    dest_calc = dd / "calculated.xlsx"
    if calc_path.resolve() != dest_calc.resolve():
        shutil.copy2(calc_path, dest_calc)

    sheets_data = get_excel_data(dest_calc)
    sheets_dir = dd / "sheets"
    sheets_dir.mkdir(parents=True, exist_ok=True)

    target_sheets = template_info.get("target_sheets", [])
    primary_sheet = template_info.get("sheet", "")
    if not primary_sheet and target_sheets:
        primary_sheet = target_sheets[0]

    sheet_counts = {}
    if sheets_data:
        for sname, sdf in sheets_data.items():
            safe_sname = re.sub(r"[^a-zA-Z0-9_\-]", "_", sname)
            sdf.to_csv(sheets_dir / f"{safe_sname}.csv", index=False)
            sdf.to_json(sheets_dir / f"{safe_sname}.json", orient="records", date_format="iso")
            sheet_counts[sname] = len(sdf)

        # Primary dashboard df
        main_df = sheets_data.get(primary_sheet)
        if main_df is None and sheets_data:
            main_df = next(iter(sheets_data.values()))
        if main_df is not None:
            main_df.to_csv(dd / "calculated.csv", index=False)
            main_df.to_json(dd / "dashboard.json", orient="records", date_format="iso")

    meta = {
        "date": str(d),
        "upload": upload_name,
        "template": template_info.get("name", ""),
        "primary_sheet": primary_sheet,
        "sheet_counts": sheet_counts,
        "saved_at": str(datetime.now()),
    }
    (dd / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")


def list_available_dates() -> List[date]:
    """Return all stored daily dates sorted from most recent to oldest."""
    dates = []
    for p in DATA_DIR.glob("*"):
        if p.is_dir():
            try:
                dates.append(date.fromisoformat(p.name))
            except Exception:
                pass
    return sorted(dates, reverse=True)


def load_day_sheets(d: date) -> Dict[str, pd.DataFrame]:
    """Load all calculated sheet DataFrames for a given date."""
    dd = DATA_DIR / d.strftime("%Y-%m-%d")
    sheets_dir = dd / "sheets"
    result: Dict[str, pd.DataFrame] = {}

    if sheets_dir.exists():
        for p in sheets_dir.glob("*.json"):
            try:
                sname = p.stem
                result[sname] = pd.read_json(p, orient="records")
            except Exception:
                pass

    if not result:
        calc_xlsx = dd / "calculated.xlsx"
        if calc_xlsx.exists():
            data = get_excel_data(calc_xlsx)
            if data:
                return data

        dash_json = dd / "dashboard.json"
        if dash_json.exists():
            try:
                result["Dashboard"] = pd.read_json(dash_json, orient="records")
            except Exception:
                pass

    return result


def load_day_main_df(d: date, sheet: Optional[str] = None) -> Optional[pd.DataFrame]:
    """Load the primary DataFrame for a given date."""
    sheets = load_day_sheets(d)
    if not sheets:
        return None
    if sheet and sheet in sheets:
        return sheets[sheet]
    # Check meta for primary sheet
    meta = load_day_meta(d)
    p_sheet = meta.get("primary_sheet")
    if p_sheet and p_sheet in sheets:
        return sheets[p_sheet]
    return next(iter(sheets.values()))


def load_day_meta(d: date) -> dict:
    """Load JSON metadata for a specific date."""
    p = DATA_DIR / d.strftime("%Y-%m-%d") / "meta.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


# ─────────────────────────────────────────────────────────────────────────────
# Smart Column Classifier & Anomaly Detection
# ─────────────────────────────────────────────────────────────────────────────
def classify_columns(df: pd.DataFrame) -> Tuple[List[str], List[str], List[str]]:
    """Partition columns into numeric, categorical text, and datetime collections."""
    numeric, text, date_cols = [], [], []
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            numeric.append(col)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            date_cols.append(col)
        else:
            text.append(col)
    return numeric, text, date_cols


def detect_anomalies(
    df: pd.DataFrame,
    col: str,
    threshold_std: float = 2.0,
) -> Tuple[int, List[int]]:
    """Detect values in a numeric column that exceed N standard deviations from the mean."""
    if col not in df.columns or not pd.api.types.is_numeric_dtype(df[col]):
        return 0, []
    s = df[col].dropna()
    if len(s) < 4:
        return 0, []
    mean = s.mean()
    std = s.std()
    if std == 0 or np.isnan(std):
        return 0, []
    z_scores = np.abs((s - mean) / std)
    anom_indices = s[z_scores > threshold_std].index.tolist()
    return len(anom_indices), anom_indices


def compute_column_stats(df: pd.DataFrame, col: str) -> dict:
    """Compute summary statistics for a column."""
    if col not in df.columns:
        return {}
    s = df[col]
    total = len(s)
    nulls = int(s.isna().sum())
    unique = int(s.nunique())
    stats = {
        "column": col,
        "total": total,
        "nulls": nulls,
        "null_pct": round((nulls / total) * 100, 1) if total > 0 else 0,
        "unique": unique,
    }
    if pd.api.types.is_numeric_dtype(s):
        s_clean = s.dropna()
        if not s_clean.empty:
            stats.update({
                "min": round(float(s_clean.min()), 2),
                "max": round(float(s_clean.max()), 2),
                "mean": round(float(s_clean.mean()), 2),
                "median": round(float(s_clean.median()), 2),
                "std": round(float(s_clean.std()), 2) if len(s_clean) > 1 else 0.0,
            })
    return stats


# ─────────────────────────────────────────────────────────────────────────────
# Visual Charts & Plotly Engine
# ─────────────────────────────────────────────────────────────────────────────
NEXGEN_PALETTE = [
    "#6366f1",  # Indigo
    "#8b5cf6",  # Purple
    "#10b981",  # Emerald Green
    "#06b6d4",  # Cyan
    "#f59e0b",  # Amber
    "#ef4444",  # Red
    "#ec4899",  # Pink
    "#14b8a6",  # Teal
    "#f97316",  # Orange
    "#84cc16",  # Lime
]

CHART_BG = "rgba(0,0,0,0)"
PLOT_BG = "rgba(26, 34, 56, 0.4)"
FONT_COLOR = "#f1f5f9"

BASE_LAYOUT = dict(
    paper_bgcolor=CHART_BG,
    plot_bgcolor=PLOT_BG,
    font=dict(color=FONT_COLOR, family="Inter, system-ui, sans-serif"),
    margin=dict(l=40, r=20, t=45, b=40),
    legend=dict(
        bgcolor="rgba(17, 24, 39, 0.6)",
        bordercolor="rgba(255, 255, 255, 0.1)",
        borderwidth=1,
    ),
    colorway=NEXGEN_PALETTE,
    xaxis=dict(gridcolor="rgba(255, 255, 255, 0.05)", zerolinecolor="rgba(255, 255, 255, 0.1)"),
    yaxis=dict(gridcolor="rgba(255, 255, 255, 0.05)", zerolinecolor="rgba(255, 255, 255, 0.1)"),
)


def build_bar_chart(
    df: pd.DataFrame,
    x: str,
    y_cols: List[str],
    title: str = "",
    barmode: str = "group",
    orientation: str = "v",
) -> go.Figure:
    """Construct grouped, stacked, or horizontal bar charts."""
    if orientation == "h":
        fig = px.bar(
            df,
            x=y_cols,
            y=x,
            title=title,
            barmode=barmode,
            orientation="h",
            color_discrete_sequence=NEXGEN_PALETTE,
        )
    else:
        fig = px.bar(
            df,
            x=x,
            y=y_cols,
            title=title,
            barmode=barmode,
            color_discrete_sequence=NEXGEN_PALETTE,
        )
    fig.update_layout(**BASE_LAYOUT)
    return fig


def build_line_chart(
    df: pd.DataFrame,
    x: str,
    y_cols: List[str],
    title: str = "",
    fill_area: bool = False,
) -> go.Figure:
    """Construct smooth multi-series line chart with optional area fill."""
    fig = go.Figure()
    for i, col in enumerate(y_cols):
        color = NEXGEN_PALETTE[i % len(NEXGEN_PALETTE)]
        fill_mode = "tozeroy" if fill_area else "none"
        fig.add_trace(go.Scatter(
            x=df[x],
            y=df[col],
            mode="lines+markers",
            name=col,
            line=dict(color=color, width=2.5),
            marker=dict(size=6, color=color),
            fill=fill_mode,
            fillcolor=f"rgba({int(color[1:3], 16)}, {int(color[3:5], 16)}, {int(color[5:7], 16)}, 0.12)" if fill_area else None,
        ))
    fig.update_layout(title=title, **BASE_LAYOUT)
    return fig


def build_pie_chart(
    df: pd.DataFrame,
    names: str,
    values: str,
    title: str = "",
    donut: bool = True,
) -> go.Figure:
    """Construct donut or standard pie chart."""
    hole_size = 0.45 if donut else 0.0
    fig = px.pie(
        df,
        names=names,
        values=values,
        title=title,
        hole=hole_size,
        color_discrete_sequence=NEXGEN_PALETTE,
    )
    fig.update_layout(**BASE_LAYOUT)
    return fig


def build_scatter_chart(
    df: pd.DataFrame,
    x: str,
    y: str,
    title: str = "",
    color_col: Optional[str] = None,
) -> go.Figure:
    """Construct scatter correlation chart."""
    fig = px.scatter(
        df,
        x=x,
        y=y,
        title=title,
        color=color_col,
        color_discrete_sequence=NEXGEN_PALETTE,
    )
    fig.update_layout(**BASE_LAYOUT)
    return fig


def build_area_chart(
    df: pd.DataFrame,
    x: str,
    y_cols: List[str],
    title: str = "",
) -> go.Figure:
    """Construct stacked or overlapping area chart."""
    fig = px.area(
        df,
        x=x,
        y=y_cols,
        title=title,
        color_discrete_sequence=NEXGEN_PALETTE,
    )
    fig.update_layout(**BASE_LAYOUT)
    return fig


def build_heatmap_chart(
    df: pd.DataFrame,
    numeric_cols: List[str],
    title: str = "Correlation Matrix",
) -> go.Figure:
    """Construct correlation heatmap for numeric columns."""
    corr = df[numeric_cols].corr().round(2)
    fig = px.imshow(
        corr,
        text_auto=True,
        aspect="auto",
        title=title,
        color_continuous_scale="Viridis",
    )
    fig.update_layout(**BASE_LAYOUT)
    return fig


def build_waterfall_chart(
    categories: List[str],
    values: List[float],
    title: str = "Delta Variance",
) -> go.Figure:
    """Construct waterfall chart for delta changes."""
    fig = go.Figure(go.Waterfall(
        name="Variance",
        orientation="v",
        measure=["relative"] * len(values),
        x=categories,
        textposition="outside",
        text=[f"{v:+,.1f}" for v in values],
        y=values,
        connector={"line": {"color": "rgba(255,255,255,0.2)"}},
        decreasing={"marker": {"color": "#ef4444"}},
        increasing={"marker": {"color": "#10b981"}},
    ))
    fig.update_layout(title=title, **BASE_LAYOUT)
    return fig


def build_trend_chart(
    dates: List[date],
    values: List[float],
    label: str,
) -> go.Figure:
    """Construct historical trend line with glowing gradient fill."""
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dates,
        y=values,
        mode="lines+markers",
        name=label,
        line=dict(color="#6366f1", width=3),
        marker=dict(size=7, color="#8b5cf6", line=dict(color="#ffffff", width=1.5)),
        fill="tozeroy",
        fillcolor="rgba(99, 102, 241, 0.15)",
    ))
    fig.update_layout(title=f"{label} – Historical Trend", **BASE_LAYOUT)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# Modern Glassmorphism UI Renderers
# ─────────────────────────────────────────────────────────────────────────────
def render_kpi_card(
    label: str,
    value: Any,
    delta: Optional[float] = None,
    delta_label: str = "",
    unit: str = "",
    color_theme: str = "indigo",
    is_anomaly: bool = False,
) -> None:
    """Render a modern glassmorphism KPI card with gradient accents and delta badge."""
    theme_colors = {
        "indigo": ("#6366f1", "rgba(99, 102, 241, 0.2)"),
        "purple": ("#8b5cf6", "rgba(139, 92, 246, 0.2)"),
        "green": ("#10b981", "rgba(16, 185, 129, 0.2)"),
        "red": ("#ef4444", "rgba(239, 68, 68, 0.2)"),
        "amber": ("#f59e0b", "rgba(245, 158, 11, 0.2)"),
        "cyan": ("#06b6d4", "rgba(6, 182, 212, 0.2)"),
    }
    accent_color, glow = theme_colors.get(color_theme, theme_colors["indigo"])

    val_str = f"{value:,.2f}" if isinstance(value, float) else f"{value:,}" if isinstance(value, int) else str(value)
    if unit:
        val_str = f"{val_str} {unit}"

    delta_html = ""
    if delta is not None:
        delta_color = "#10b981" if delta >= 0 else "#ef4444"
        arrow = "▲" if delta >= 0 else "▼"
        delta_val_str = f"{abs(delta):,.2f}" if isinstance(delta, float) else f"{abs(delta)}"
        delta_html = f"""
        <div style="display:inline-flex;align-items:center;gap:4px;color:{delta_color};font-size:0.75rem;font-weight:600;margin-top:6px;">
            <span>{arrow} {delta_val_str}</span>
            <span style="color:#94a3b8;font-weight:400">{delta_label}</span>
        </div>
        """

    anomaly_html = ""
    if is_anomaly:
        anomaly_html = """
        <div style="display:inline-block;background:rgba(239,68,68,0.2);border:1px solid rgba(239,68,68,0.5);color:#fca5a5;font-size:0.65rem;font-weight:700;padding:2px 7px;border-radius:20px;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:4px;">
            🚨 Outlier (&gt;2σ)
        </div>
        """

    card_html = f"""
    <div style="background:rgba(26,34,56,0.75);backdrop-filter:blur(12px);border:1px solid rgba(255,255,255,0.08);border-radius:14px;padding:18px 20px;position:relative;overflow:hidden;box-shadow:0 8px 20px rgba(0,0,0,0.25);margin-bottom:12px;">
      <div style="position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg, {accent_color}, transparent);border-radius:14px 14px 0 0"></div>
      {anomaly_html}
      <div style="font-size:0.75rem;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:0.07em;">{label}</div>
      <div style="font-size:2rem;font-weight:800;color:#f1f5f9;margin:8px 0 2px;line-height:1.1;letter-spacing:-0.02em;">{val_str}</div>
      {delta_html}
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)


def render_summary_stats_row(df: pd.DataFrame, numeric_cols: List[str]) -> None:
    """Render a compact summary stats row showing min/avg/max for key metrics."""
    if not numeric_cols:
        return
    cols = numeric_cols[:6]
    st.markdown("##### 📌 Metric Summary Bar")
    cards = st.columns(len(cols))
    for i, col in enumerate(cols):
        s = df[col].dropna()
        if s.empty:
            continue
        c_min = float(s.min())
        c_avg = float(s.mean())
        c_max = float(s.max())
        with cards[i]:
            st.markdown(f"""
            <div style="background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.06);border-radius:10px;padding:10px 14px;font-size:0.75rem;">
              <div style="color:#cbd5e1;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{col}">{col}</div>
              <div style="display:flex;justify-content:space-between;margin-top:6px;color:#94a3b8;">
                <span>Min: <b style="color:#f1f5f9">{c_min:,.1f}</b></span>
                <span>Avg: <b style="color:#6366f1">{c_avg:,.1f}</b></span>
                <span>Max: <b style="color:#10b981">{c_max:,.1f}</b></span>
              </div>
            </div>
            """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PDF Export Generator
# ─────────────────────────────────────────────────────────────────────────────
def generate_pdf(
    df: pd.DataFrame,
    day: date,
    kpi_summary: dict,
    template_name: str = "",
) -> bytes:
    """Generate a high-quality PDF report."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
    )
    styles = getSampleStyleSheet()
    story = []

    # Title & Header
    title_text = f"<b>Nexgen Analytics – {template_name or 'Operational Report'}</b>"
    story.append(Paragraph(title_text, styles["Title"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph(f"Reporting Date: <b>{day.strftime('%d %B %Y')}</b> | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 0.4 * cm))

    # KPIs Section
    if kpi_summary:
        story.append(Paragraph("<b>Key Performance Indicators</b>", styles["Heading2"]))
        kpi_data = [["Metric Name", "Value"]] + [[str(k), f"{v:,.2f}" if isinstance(v, float) else str(v)] for k, v in kpi_summary.items()]
        t_kpi = Table(kpi_data, colWidths=[11 * cm, 6 * cm])
        t_kpi.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#6366f1")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#1a2238"), colors.HexColor("#111827")]),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#374151")),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 0.6 * cm))

    # Data Snapshot
    story.append(Paragraph("<b>Data Records Snapshot (Top 40)</b>", styles["Heading2"]))
    display_df = df.head(40)
    cols_list = list(display_df.columns)[:8]
    table_data = [cols_list] + display_df[cols_list].astype(str).values.tolist()

    col_w = (17 * cm) / max(len(cols_list), 1)
    t_data = Table(table_data, colWidths=[col_w] * len(cols_list), repeatRows=1)
    t_data.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#6366f1")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#1a2238"), colors.HexColor("#111827")]),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#374151")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t_data)

    doc.build(story)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Global Modern CSS Injection
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
  
  html, body, [class*="css"] {
      font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
  }
  
  .stApp {
      background: #0a0f1e;
      color: #f1f5f9;
  }
  
  section[data-testid="stSidebar"] {
      background: #0f172a !important;
      border-right: 1px solid rgba(255, 255, 255, 0.08);
  }
  
  /* Buttons */
  .stButton > button {
      background: linear-gradient(135deg, #6366f1, #8b5cf6);
      color: #ffffff;
      border: 1px solid rgba(255, 255, 255, 0.15);
      border-radius: 9px;
      font-weight: 600;
      padding: 9px 20px;
      transition: all 0.2s ease-in-out;
  }
  .stButton > button:hover {
      transform: translateY(-1.5px);
      box-shadow: 0 6px 20px rgba(99, 102, 241, 0.35);
      border-color: rgba(255, 255, 255, 0.3);
  }
  
  /* Metric containers fallback */
  div[data-testid="metric-container"] {
      background: rgba(26, 34, 56, 0.8);
      backdrop-filter: blur(10px);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 12px;
      padding: 16px;
  }
  
  /* Uploader styling */
  .stFileUploader {
      border: 2px dashed rgba(99, 102, 241, 0.35);
      background: rgba(99, 102, 241, 0.03);
      border-radius: 14px;
      padding: 10px;
  }
  
  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {
      gap: 8px;
  }
  .stTabs [data-baseweb="tab"] {
      background: rgba(255, 255, 255, 0.04);
      border-radius: 8px;
      border: 1px solid rgba(255, 255, 255, 0.06);
      padding: 8px 16px;
      color: #94a3b8;
  }
  .stTabs [aria-selected="true"] {
      background: linear-gradient(135deg, rgba(99, 102, 241, 0.2), rgba(139, 92, 246, 0.2)) !important;
      border-color: #6366f1 !important;
      color: #ffffff !important;
  }
  
  /* Badges */
  .badge {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 3px 10px;
      border-radius: 20px;
      font-size: 0.72rem;
      font-weight: 600;
  }
  .badge-success { background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }
  .badge-warning { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }
  .badge-danger  { background: rgba(239, 68, 68, 0.15);  color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
  .badge-indigo  { background: rgba(99, 102, 241, 0.15); color: #a5b4fc; border: 1px solid rgba(99, 102, 241, 0.3); }
  
  /* Tables */
  .stDataFrame {
      border-radius: 10px;
      overflow: hidden;
      border: 1px solid rgba(255, 255, 255, 0.08);
  }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar: Navigation, Quick Date Stepper & System Status
# ─────────────────────────────────────────────────────────────────────────────
meta = load_meta()
available_dates = list_available_dates()

with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:12px;padding:8px 0 16px 0;">
      <div style="width:42px;height:42px;background:linear-gradient(135deg,#6366f1,#8b5cf6);border-radius:12px;
                  display:flex;align-items:center;justify-content:center;font-size:1.3rem;box-shadow:0 6px 15px rgba(99,102,241,0.4)">⚡</div>
      <div>
        <div style="font-size:1.15rem;font-weight:800;color:#f1f5f9;letter-spacing:-0.01em;">Nexgen</div>
        <div style="font-size:0.72rem;color:#94a3b8;font-weight:500;">Analytics Platform</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    nav_pages = [
        "📊 Dashboard",
        "📁 Template Setup",
        "📤 Upload Data",
        "📅 History",
        "📈 Comparison",
        "📋 Data Validation",
        "⚙️ Settings",
    ]

    if "page_override" in st.session_state:
        override_target = st.session_state.pop("page_override")
        default_nav_idx = nav_pages.index(override_target) if override_target in nav_pages else 0
    else:
        default_nav_idx = 0

    page = st.selectbox(
        "Navigation",
        nav_pages,
        index=default_nav_idx,
    )

    st.markdown("---")

    # Quick Date Navigator
    if available_dates:
        st.markdown("<div style='font-size:0.75rem;font-weight:700;color:#94a3b8;text-transform:uppercase;'>📅 Quick Date Navigator</div>", unsafe_allow_html=True)
        if "selected_date_idx" not in st.session_state:
            st.session_state["selected_date_idx"] = 0

        curr_idx = min(st.session_state["selected_date_idx"], len(available_dates) - 1)
        active_date = available_dates[curr_idx]

        c_prev, c_lbl, c_next = st.columns([1, 3, 1])
        if c_prev.button("◀", help="Previous day"):
            if curr_idx < len(available_dates) - 1:
                st.session_state["selected_date_idx"] = curr_idx + 1
                st.rerun()
        if c_next.button("▶", help="Next day"):
            if curr_idx > 0:
                st.session_state["selected_date_idx"] = curr_idx - 1
                st.rerun()

        with c_lbl:
            st.markdown(
                f"<div style='text-align:center;font-weight:700;font-size:0.85rem;color:#f1f5f9;padding-top:6px;'>"
                f"{active_date.strftime('%d %b %Y')}</div>",
                unsafe_allow_html=True,
            )

    # Active Template Indicator
    st.markdown("---")
    active_tmpl_id = meta.get("active_template")
    if active_tmpl_id and active_tmpl_id in meta.get("templates", {}):
        tmpl_obj = meta["templates"][active_tmpl_id]
        st.markdown(
            f"<div class='badge badge-indigo'>Template: <b>{tmpl_obj.get('name')}</b></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown("<div class='badge badge-warning'>No Active Template</div>", unsafe_allow_html=True)

    # Inbox Watcher
    inbox_files = list(INBOX_DIR.glob("*.xlsx")) + list(INBOX_DIR.glob("*.csv"))
    if inbox_files:
        st.warning(f"📥 {len(inbox_files)} file(s) waiting in inbox")

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption(f"📁 Records: {len(available_dates)} Day(s) stored")

    st.markdown("---")
    with st.expander("🛑 Emergency Kill Switch", expanded=False):
        st.caption("Instantly stop server and release all ports.")
        if st.button("⛔ Kill Server", type="primary", use_container_width=True):
            st.error("Shutting down Nexgen Dashboard...")
            kill_script = BASE / "kill_nexgen.bat"
            if kill_script.exists():
                subprocess.Popen(["cmd.exe", "/c", str(kill_script)], creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
            else:
                os._exit(0)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 1: 📊 DASHBOARD (Uses Custom Designed Layout)
# ─────────────────────────────────────────────────────────────────────────────
if page == "📊 Dashboard":
    st.title("📊 Operational Dashboard")

    if not available_dates:
        st.info("👋 Welcome to Nexgen Analytics! No daily data is loaded yet. Upload your template in **Template Setup** and then upload daily data.")
        st.stop()
    else:
        # Determine current date from session state or dropdown
        if "selected_date_idx" not in st.session_state:
            st.session_state["selected_date_idx"] = 0
        curr_idx = min(st.session_state["selected_date_idx"], len(available_dates) - 1)
        current_date = available_dates[curr_idx]

        # Top Toolbar: Date Picker, Template Info & Global Search
        col_t1, col_t2, col_t3 = st.columns([2, 1, 3])
        with col_t1:
            chosen_date = st.selectbox(
                "Select Reporting Date",
                available_dates,
                index=curr_idx,
                format_func=lambda d: d.strftime("%d %B %Y (%A)"),
            )
            if chosen_date != current_date:
                st.session_state["selected_date_idx"] = available_dates.index(chosen_date)
                st.rerun()

        # Load calculated data
        sheets_data = load_day_sheets(current_date)
        if not sheets_data:
            st.warning(f"No processed data found for {current_date}.")
            st.stop()

        active_tmpl_id = meta.get("active_template")
        tmpl_config = meta.get("templates", {}).get(active_tmpl_id, {})
        layout = tmpl_config.get("dashboard_layout", [])

        with col_t2:
            sheet_options = list(sheets_data.keys())
            active_sheet = st.selectbox("Active Sheet", sheet_options, index=0)
            df_sheet = sheets_data[active_sheet]

        with col_t3:
            global_search = st.text_input("🔍 Quick Search across all columns", placeholder="Filter by agent, branch, status, ID...")

        col_btn, _ = st.columns([2, 4])
        with col_btn:
            if st.button("🎨 Edit Dashboard Layout", help="Go to Template Setup to redesign your dashboard"):
                st.session_state["setup_step"] = 3
                st.session_state["page_override"] = "📁 Template Setup"
                st.rerun()

        # Filter by global search
        if global_search:
            mask = df_sheet.astype(str).apply(
                lambda r: r.str.contains(global_search, case=False, na=False).any(), axis=1
            )
            df_sheet = df_sheet[mask]
            st.caption(f"Filtered to **{len(df_sheet)}** matching rows.")

        numeric_cols, text_cols, date_cols = classify_columns(df_sheet)

        # ── Presets Bar ───────────────────────────────────────────────────────
        with st.expander("⚙️ Dashboard Presets & Filters", expanded=False):
            p_col1, p_col2, p_col3 = st.columns([2, 2, 2])
            presets = load_presets()
            with p_col1:
                p_names = ["(None)"] + list(presets.keys())
                chosen_preset = st.selectbox("Load Saved Preset", p_names)
                if chosen_preset != "(None)":
                    preset_data = presets[chosen_preset]
                    st.info(f"Loaded preset: **{chosen_preset}**")
            with p_col2:
                new_preset_name = st.text_input("New Preset Name", placeholder="e.g. TL Review")
                if st.button("💾 Save View Preset"):
                    if new_preset_name:
                        save_preset(new_preset_name, {"sheet": active_sheet, "search": global_search})
                        st.success(f"Preset '{new_preset_name}' saved!")
                        st.rerun()
            with p_col3:
                if chosen_preset != "(None)" and st.button("🗑️ Delete Preset"):
                    delete_preset(chosen_preset)
                    st.success("Preset deleted.")
                    st.rerun()

        # Check for anomalies across numeric columns
        detected_anomalies_map: Dict[str, int] = {}
        for col in numeric_cols:
            count, _ = detect_anomalies(df_sheet, col)
            if count > 0:
                detected_anomalies_map[col] = count

        if detected_anomalies_map:
            anom_desc = ", ".join([f"**{col}** ({cnt} rows)" for col, cnt in list(detected_anomalies_map.items())[:4]])
            st.markdown(f"""
            <div style="background:rgba(239,68,68,0.12);border:1px solid rgba(239,68,68,0.3);border-radius:10px;padding:12px 18px;margin:10px 0 16px 0;">
                <span style="color:#f87171;font-weight:700;">🚨 Statistical Outlier Alert:</span>
                <span style="color:#fecaca;font-size:0.88rem;"> Significant deviations (&gt;2σ) detected in {anom_desc}.</span>
            </div>
            """, unsafe_allow_html=True)

        # Yesterday comparison data
        prev_date = current_date - timedelta(days=1)
        df_prev = load_day_main_df(prev_date, active_sheet)

        # ── RENDER DASHBOARD LAYOUT ──────────────────────────────────────────
        # Check if dashboard layout is designed - NO AUTO DETECTION FALLBACK
        if not layout:
            st.markdown("""
            <div style="background:rgba(99,102,241,0.1);border:2px dashed rgba(99,102,241,0.4);border-radius:16px;padding:40px;text-align:center;margin:20px 0;">
                <div style="font-size:2.5rem;margin-bottom:12px;">🎨</div>
                <div style="font-size:1.2rem;font-weight:700;color:#f1f5f9;margin-bottom:8px;">No Dashboard Layout Designed Yet</div>
                <div style="color:#8693aa;margin-bottom:20px;">Go to <b>Template Setup → Step 3: Dashboard Designer</b> to choose exactly which KPIs, charts and tables appear here.</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🎨 Go to Dashboard Designer", type="primary"):
                st.session_state["setup_step"] = 3
                st.session_state["page_override"] = "📁 Template Setup"
                st.rerun()
            st.stop()

        kpi_items = [item for item in layout if item.get("type") == "kpi"]
        chart_items = [item for item in layout if item.get("type") == "chart"]
        table_items = [item for item in layout if item.get("type") == "table"]

        kpi_summary_export: Dict[str, Any] = {}

        st.markdown("<br>", unsafe_allow_html=True)

        if kpi_items:
            st.subheader("Key Performance Indicators")
            kpi_cols_grid = st.columns(min(4, max(1, len(kpi_items))))
            for idx, kpi in enumerate(kpi_items):
                target_sname = kpi.get("sheet", active_sheet)
                target_df = sheets_data.get(target_sname, df_sheet)
                col_name = kpi.get("col")
                if col_name not in target_df.columns:
                    continue

                agg = kpi.get("agg", "sum").lower()
                s = target_df[col_name].dropna()
                if s.empty:
                    val = 0
                elif agg in ("sum", "total"):
                    val = float(s.sum())
                elif agg in ("avg", "average", "mean"):
                    val = float(s.mean())
                elif agg == "max":
                    val = float(s.max())
                elif agg == "min":
                    val = float(s.min())
                elif agg == "count":
                    val = int(len(s))
                else:
                    val = float(s.iloc[-1])

                # Delta comparison
                delta = None
                if df_prev is not None and col_name in df_prev.columns:
                    s_p = df_prev[col_name].dropna()
                    if not s_p.empty:
                        if agg in ("sum", "total"):
                            p_val = float(s_p.sum())
                        elif agg in ("avg", "average", "mean"):
                            p_val = float(s_p.mean())
                        elif agg == "max":
                            p_val = float(s_p.max())
                        elif agg == "min":
                            p_val = float(s_p.min())
                        else:
                            p_val = float(len(s_p))
                        delta = round(val - p_val, 2)

                kpi_summary_export[kpi.get("label", col_name)] = val

                with kpi_cols_grid[idx % len(kpi_cols_grid)]:
                    render_kpi_card(
                        label=kpi.get("label", col_name),
                        value=val,
                        delta=delta,
                        delta_label="vs yesterday",
                        unit=kpi.get("unit", ""),
                        color_theme=kpi.get("color", "indigo"),
                        is_anomaly=(col_name in detected_anomalies_map),
                    )

            # Summary stats row below KPIs
            render_summary_stats_row(df_sheet, numeric_cols)

        st.markdown("---")

        # ── Charts Section ────────────────────────────────────────────────────
        if chart_items:
            st.subheader("Visual Analytics")
            for idx in range(0, len(chart_items), 2):
                chart_cols = st.columns(2 if idx + 1 < len(chart_items) else 1)
                for c_offset, c_cfg in enumerate(chart_items[idx : idx + 2]):
                    with chart_cols[c_offset]:
                        c_sheet = sheets_data.get(c_cfg.get("sheet", active_sheet), df_sheet)
                        ctype = c_cfg.get("chart_type", "bar").lower().replace(" ", "_")
                        cx = c_cfg.get("x")
                        cy = c_cfg.get("y", [])
                        if isinstance(cy, str):
                            cy = [cy]
                        ctitle = c_cfg.get("title", f"{ctype.title()} Chart")

                        if cx and cy and cx in c_sheet.columns and all(y in c_sheet.columns for y in cy):
                            if ctype in ("bar", "grouped_bar"):
                                fig = build_bar_chart(c_sheet, cx, cy, ctitle, barmode="group")
                            elif ctype == "stacked_bar":
                                fig = build_bar_chart(c_sheet, cx, cy, ctitle, barmode="stack")
                            elif ctype == "horizontal_bar":
                                fig = build_bar_chart(c_sheet, cx, cy, ctitle, orientation="h")
                            elif ctype == "line":
                                fig = build_line_chart(c_sheet, cx, cy, ctitle, fill_area=False)
                            elif ctype == "area":
                                fig = build_area_chart(c_sheet, cx, cy, ctitle)
                            elif ctype in ("pie", "donut"):
                                fig = build_pie_chart(c_sheet, cx, cy[0], ctitle, donut=(ctype == "donut"))
                            elif ctype == "scatter":
                                fig = build_scatter_chart(c_sheet, cx, cy[0], ctitle)
                            elif ctype == "heatmap":
                                fig = build_heatmap_chart(c_sheet, cy, ctitle)
                            else:
                                fig = build_bar_chart(c_sheet, cx, cy, ctitle)
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning(f"Could not render chart '{ctitle}': missing columns.")
            st.markdown("---")

        # ── Tables & Data Explorer ────────────────────────────────────────────
        if table_items:
            st.subheader("📋 Data Records")
            for t_cfg in table_items:
                t_sheet = sheets_data.get(t_cfg.get("sheet", active_sheet), df_sheet)
                t_cols = t_cfg.get("cols", list(t_sheet.columns))
                t_cols = [c for c in t_cols if c in t_sheet.columns]
                max_r = t_cfg.get("max_rows", 50)
                st.markdown(f"##### {t_cfg.get('title', 'Data Table')}")
                st.dataframe(t_sheet[t_cols].head(max_r), use_container_width=True)
            st.markdown("---")

        # ── Export Hub ────────────────────────────────────────────────────────
        st.markdown("---")
        st.subheader("📥 Export & Report Hub")
        e1, e2, e3, e4 = st.columns(4)

        with e1:
            csv_bytes = df_sheet.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Download CSV",
                csv_bytes,
                f"nexgen_{active_sheet}_{current_date}.csv",
                "text/csv",
                use_container_width=True,
            )

        with e2:
            xlsx_buf = io.BytesIO()
            with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as writer:
                for sname, sdf in sheets_data.items():
                    sdf.to_excel(writer, index=False, sheet_name=sname[:31])
            st.download_button(
                "⬇️ Download Excel (All Sheets)",
                xlsx_buf.getvalue(),
                f"nexgen_all_sheets_{current_date}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with e3:
            tmpl_name = tmpl_config.get("name", active_sheet)
            pdf_bytes = generate_pdf(df_sheet, current_date, kpi_summary_export, tmpl_name)
            st.download_button(
                "⬇️ Download PDF Report",
                pdf_bytes,
                f"nexgen_report_{current_date}.pdf",
                "application/pdf",
                use_container_width=True,
            )

        with e4:
            calc_file = DATA_DIR / current_date.strftime("%Y-%m-%d") / "calculated.xlsx"
            if calc_file.exists():
                st.download_button(
                    "⬇️ Download Recalculated Excel (Formulas)",
                    calc_file.read_bytes(),
                    f"nexgen_calculated_{current_date}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 2: 📁 TEMPLATE SETUP (Wizard: Upload -> Mapping -> Designer -> Save)
# ─────────────────────────────────────────────────────────────────────────────
elif page == "📁 Template Setup":
    st.title("📁 Template Setup & Dashboard Designer")
    st.markdown("Configure your **Excel Template**, map raw columns to multiple template sheets, and design your custom dashboard.")

    # Wizard Steps State
    if "setup_step" not in st.session_state:
        st.session_state["setup_step"] = 1

    # Step Progress Indicator
    steps = ["1. Upload Template", "2. Manual Column Mapping", "3. Dashboard Designer", "4. Preview & Save"]
    step_cols = st.columns(4)
    for i, s_label in enumerate(steps):
        is_active = (st.session_state["setup_step"] == i + 1)
        bg = "linear-gradient(135deg,#6366f1,#8b5cf6)" if is_active else "rgba(255,255,255,0.05)"
        text_color = "#ffffff" if is_active else "#94a3b8"
        step_cols[i].markdown(
            f"""<div style="background:{bg};color:{text_color};padding:8px 12px;border-radius:8px;text-align:center;font-weight:700;font-size:0.8rem;border:1px solid rgba(255,255,255,0.1);">{s_label}</div>""",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── STEP 1: Upload Template ───────────────────────────────────────────
    if st.session_state["setup_step"] == 1:
        st.subheader("Step 1: Upload Your Master Excel Template")
        tmpl_upload = st.file_uploader(
            "Upload Excel template with formulas, sheets, and headers (.xlsx)",
            type=["xlsx"],
            key="wizard_tmpl_upload",
        )
        tmpl_label = st.text_input("Template Name", placeholder="e.g. NPA Daily Operations Report")

        # Or select existing template to edit
        if meta.get("templates"):
            st.markdown("---")
            st.markdown("##### Or Edit an Existing Template:")
            edit_tmpl_id = st.selectbox("Select Template to Edit", list(meta["templates"].keys()), format_func=lambda k: meta["templates"][k].get("name", k))
            if st.button("✏️ Load Selected Template into Wizard"):
                t_obj = meta["templates"][edit_tmpl_id]
                st.session_state["wizard_template_id"] = edit_tmpl_id
                st.session_state["wizard_template_name"] = t_obj.get("name", edit_tmpl_id)
                st.session_state["wizard_template_path"] = t_obj.get("file")
                st.session_state["wizard_schema"] = t_obj.get("schema", analyse_template(Path(t_obj.get("file"))))
                st.session_state["wizard_target_sheets"] = t_obj.get("target_sheets", list(st.session_state["wizard_schema"]["sheets"].keys())[:2])
                st.session_state["wizard_column_mapping"] = t_obj.get("column_mapping", {})
                st.session_state["wizard_formula_cols"] = t_obj.get("formula_cols", {})
                st.session_state["wizard_dashboard_layout"] = t_obj.get("dashboard_layout", [])
                st.session_state["setup_step"] = 2
                st.rerun()

        if tmpl_upload and tmpl_label:
            if st.button("Next: Proceed to Column Mapping ➡️"):
                safe_key = re.sub(r"[^a-z0-9_]", "_", tmpl_label.lower().strip())
                dest = TMPL_DIR / f"{safe_key}.xlsx"
                dest.write_bytes(tmpl_upload.read())
                schema = analyse_template(dest)

                st.session_state["wizard_template_id"] = safe_key
                st.session_state["wizard_template_name"] = tmpl_label
                st.session_state["wizard_template_path"] = str(dest)
                st.session_state["wizard_schema"] = schema
                st.session_state["wizard_target_sheets"] = list(schema["sheets"].keys())[:2]
                st.session_state["wizard_column_mapping"] = {}
                st.session_state["wizard_formula_cols"] = {s: list(sinfo.get("formula_cols", {}).keys()) for s, sinfo in schema["sheets"].items()}
                st.session_state["wizard_dashboard_layout"] = []
                st.session_state["setup_step"] = 2
                st.rerun()

    # ── STEP 2: Manual Column Mapping ─────────────────────────────────────
    elif st.session_state["setup_step"] == 2:
        st.subheader("Step 2: Manual Multi-Sheet Column Mapping")
        schema = st.session_state.get("wizard_schema", {})
        all_sheets = list(schema.get("sheets", {}).keys())

        st.info(f"Configuring template: **{st.session_state.get('wizard_template_name')}** | Sheets in file: **{', '.join(all_sheets)}**")

        target_sheets = st.multiselect(
            "Which sheet(s) receive uploaded data?",
            all_sheets,
            default=st.session_state.get("wizard_target_sheets", all_sheets[:2]),
        )
        st.session_state["wizard_target_sheets"] = target_sheets

        # Sample upload file to detect raw columns
        st.markdown("##### 🧪 Optional: Upload a Sample Raw File for Auto-Suggestions")
        sample_file = st.file_uploader("Upload sample raw xlsx or csv to populate raw column choices", type=["xlsx", "csv"], key="sample_uploader")
        sample_cols: List[str] = []
        if sample_file:
            try:
                if sample_file.name.endswith(".csv"):
                    sample_df = pd.read_csv(sample_file, nrows=5)
                else:
                    sample_df = pd.read_excel(sample_file, nrows=5)
                sample_cols = [str(c).strip() for c in sample_df.columns]
                st.success(f"Detected {len(sample_cols)} raw upload column(s): {', '.join(sample_cols[:8])}...")
            except Exception as e:
                st.error(f"Could not read sample file: {e}")

        # Bulk Suggest Button
        if sample_cols and st.button("✨ Auto-Suggest Mappings (Fuzzy Match)"):
            for s_name in target_sheets:
                s_headers = schema["sheets"].get(s_name, {}).get("headers", [])
                suggestions = suggest_mappings(sample_cols, s_headers)
                st.session_state.setdefault("wizard_column_mapping", {})
                st.session_state["wizard_column_mapping"].setdefault(s_name, {})
                for th, uh in suggestions.items():
                    st.session_state["wizard_column_mapping"][s_name][th] = uh
            st.success("Automated fuzzy match suggestions applied!")
            st.rerun()

        # Mapping grid for each target sheet
        st.session_state.setdefault("wizard_column_mapping", {})
        st.session_state.setdefault("wizard_formula_cols", {})

        for s_name in target_sheets:
            st.markdown(f"#### 📄 Sheet: `{s_name}`")
            sheet_info = schema["sheets"].get(s_name, {})
            headers = sheet_info.get("headers", [])
            formula_cols_detected = sheet_info.get("formula_cols", {})

            st.session_state["wizard_column_mapping"].setdefault(s_name, {})
            current_sheet_map = st.session_state["wizard_column_mapping"][s_name]

            # Form list of dropdown options
            dropdown_options = ["-- select --", "-- formula column --", "-- ignore --"] + sample_cols

            grid_container = st.container()
            with grid_container:
                m_hdr1, m_hdr2, m_hdr3 = st.columns([3, 3, 2])
                m_hdr1.markdown("**Template Column**")
                m_hdr2.markdown("**Map to Upload Column / Action**")
                m_hdr3.markdown("**Status**")

                for t_col in headers:
                    c1, c2, c3 = st.columns([3, 3, 2])
                    c1.write(f"`{t_col}`")

                    # Check current value
                    existing_val = current_sheet_map.get(t_col)
                    if not existing_val and t_col in formula_cols_detected:
                        existing_val = "-- formula column --"
                    elif not existing_val and t_col in sample_cols:
                        existing_val = t_col

                    opts = dropdown_options.copy()
                    if existing_val and existing_val not in opts:
                        opts.append(existing_val)

                    def_idx = opts.index(existing_val) if existing_val in opts else 0
                    chosen_map = c2.selectbox(
                        f"Map {t_col}",
                        opts,
                        index=def_idx,
                        key=f"map_{s_name}_{t_col}",
                        label_visibility="collapsed",
                    )
                    current_sheet_map[t_col] = chosen_map

                    # Status badge
                    if chosen_map == "-- formula column --":
                        c3.markdown("<span class='badge badge-indigo'>Formula (Calculated)</span>", unsafe_allow_html=True)
                    elif chosen_map == "-- ignore --":
                        c3.markdown("<span class='badge badge-warning'>Ignored</span>", unsafe_allow_html=True)
                    elif chosen_map != "-- select --":
                        c3.markdown("<span class='badge badge-success'>Mapped</span>", unsafe_allow_html=True)
                    else:
                        c3.markdown("<span class='badge badge-danger'>Unmapped</span>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        nav1, nav2 = st.columns([1, 1])
        if nav1.button("⬅️ Back to Step 1"):
            st.session_state["setup_step"] = 1
            st.rerun()
        if nav2.button("Next: Dashboard Designer ➡️"):
            # Update formula_cols summary
            for s_name in target_sheets:
                f_list = [k for k, v in st.session_state["wizard_column_mapping"].get(s_name, {}).items() if v == "-- formula column --"]
                st.session_state["wizard_formula_cols"][s_name] = f_list
            st.session_state["setup_step"] = 3
            st.rerun()

    # ── STEP 3: Dashboard Designer ────────────────────────────────────────
    elif st.session_state["setup_step"] == 3:
        st.subheader("Step 3: Dashboard Designer (Manual Output Definition)")
        st.markdown(
            "Define the exact **KPI Cards**, **Charts**, and **Data Tables** that will appear on the live dashboard. "
            "Nothing is auto-generated—you choose the exact source sheets, columns, aggregations, and layout order."
        )

        active_tmpl_id = meta.get("active_template")

        # Hydrate template & schema if Step 3 is accessed directly or session state is uninitialized
        if not st.session_state.get("wizard_schema"):
            tmpl_to_load = st.session_state.get("wizard_template_id") or active_tmpl_id
            if tmpl_to_load and tmpl_to_load in meta.get("templates", {}):
                t_obj = meta["templates"][tmpl_to_load]
                st.session_state["wizard_template_id"] = tmpl_to_load
                st.session_state["wizard_template_name"] = t_obj.get("name", tmpl_to_load)
                st.session_state["wizard_template_path"] = t_obj.get("file")
                schema_data = t_obj.get("schema")
                if not schema_data and t_obj.get("file") and Path(t_obj.get("file")).exists():
                    schema_data = analyse_template(Path(t_obj.get("file")))
                st.session_state["wizard_schema"] = schema_data or {}
                st.session_state["wizard_target_sheets"] = t_obj.get(
                    "target_sheets", list(st.session_state["wizard_schema"].get("sheets", {}).keys())
                )
                st.session_state["wizard_column_mapping"] = t_obj.get("column_mapping", {})
                st.session_state["wizard_formula_cols"] = t_obj.get("formula_cols", {})
                if "wizard_dashboard_layout" not in st.session_state:
                    st.session_state["wizard_dashboard_layout"] = copy.deepcopy(t_obj.get("dashboard_layout", []))

        schema = st.session_state.get("wizard_schema", {})
        all_sheet_names = list(schema.get("sheets", {}).keys())
        target_sheets = st.session_state.get("wizard_target_sheets") or all_sheet_names
        available_sheets = all_sheet_names if all_sheet_names else (target_sheets or ["Sheet1"])

        if not all_sheet_names and not schema.get("sheets"):
            st.warning("⚠️ No template schema detected. Please upload an Excel template in Step 1 first.")
            if st.button("⬅️ Go to Step 1: Upload Template"):
                st.session_state["setup_step"] = 1
                st.rerun()
            st.stop()

        if "wizard_dashboard_layout" not in st.session_state:
            st.session_state["wizard_dashboard_layout"] = []
        layout = st.session_state["wizard_dashboard_layout"]

        # Current Template & Direct Save Action Banner
        tmpl_display = st.session_state.get("wizard_template_name") or active_tmpl_id or "Custom Template"
        total_kpis = len([i for i in layout if i.get("type") == "kpi"])
        total_charts = len([i for i in layout if i.get("type") == "chart"])
        total_tables = len([i for i in layout if i.get("type") == "table"])

        top_col1, top_col2 = st.columns([3, 2])
        with top_col1:
            st.markdown(
                f"""<div style="background:rgba(99,102,241,0.08);border:1px solid rgba(99,102,241,0.25);border-radius:10px;padding:10px 16px;">
                    <span style="font-weight:700;color:#818cf8;">Template:</span> {tmpl_display} &nbsp;|&nbsp;
                    <span style="color:#94a3b8;">Configured:</span> <b>{total_kpis}</b> KPIs, <b>{total_charts}</b> Charts, <b>{total_tables}</b> Tables
                </div>""",
                unsafe_allow_html=True,
            )
        with top_col2:
            if st.button("💾 Save Dashboard Layout to Template", type="primary", use_container_width=True):
                cur_tmpl_id = st.session_state.get("wizard_template_id") or active_tmpl_id
                if cur_tmpl_id and cur_tmpl_id in meta.get("templates", {}):
                    meta["templates"][cur_tmpl_id]["dashboard_layout"] = copy.deepcopy(layout)
                    save_meta(meta)
                    st.success(f"✅ Dashboard layout successfully saved to template **{meta['templates'][cur_tmpl_id].get('name', cur_tmpl_id)}**!")
                else:
                    st.success("✅ Dashboard layout saved in wizard session. Proceed to Step 4 to finalize template.")

        st.markdown("<br>", unsafe_allow_html=True)
        designer_tabs = st.tabs(["📌 KPI Cards", "📊 Charts", "📋 Tables", "👁️ Live Layout Preview & Reorder"])

        # ── TAB 1: KPI Cards ─────────────────────────────────────────────────
        with designer_tabs[0]:
            st.markdown("#### 📌 Configure KPI Metric Cards")
            st.caption("Add prominent metric cards displayed at the top of the operational dashboard.")

            with st.expander("➕ Add New KPI Card", expanded=(total_kpis == 0)):
                k_c1, k_c2, k_c3 = st.columns([2, 2, 2])
                with k_c1:
                    k_sheet = st.selectbox("Source Sheet", available_sheets, key="kpi_src_sheet")
                    k_headers = schema.get("sheets", {}).get(k_sheet, {}).get("headers", [])
                    if not k_headers:
                        st.warning(f"No headers detected in sheet '{k_sheet}'")
                    k_col = st.selectbox("Metric Column", k_headers if k_headers else ["-- none --"], key="kpi_col_picker")
                with k_c2:
                    k_default_lbl = k_col if k_col and k_col != "-- none --" else ""
                    k_label = st.text_input("Card Label", value=k_default_lbl, placeholder="e.g. Total Inflow", key="kpi_lbl_input")
                    k_agg = st.selectbox(
                        "Aggregation Function",
                        ["sum", "average", "count", "max", "min", "latest value"],
                        format_func=lambda x: {
                            "sum": "Sum (Total)",
                            "average": "Average (Mean)",
                            "count": "Count (Rows)",
                            "max": "Maximum",
                            "min": "Minimum",
                            "latest value": "Latest Value",
                        }.get(x, x),
                        key="kpi_agg_picker",
                    )
                with k_c3:
                    k_unit = st.text_input("Unit / Symbol (optional)", placeholder="e.g. ₹, %, Calls, Accounts", key="kpi_unit_input")
                    k_color = st.selectbox("Theme Color", ["indigo", "green", "red", "amber", "purple", "cyan", "blue"], key="kpi_color_picker")

                if st.button("➕ Add KPI Card", key="btn_add_kpi"):
                    if k_col and k_col != "-- none --":
                        new_kpi = {
                            "type": "kpi",
                            "id": f"kpi_{uuid.uuid4().hex[:6]}",
                            "sheet": k_sheet,
                            "col": k_col,
                            "label": k_label.strip() if k_label.strip() else k_col,
                            "agg": k_agg,
                            "unit": k_unit.strip(),
                            "color": k_color,
                        }
                        layout.append(new_kpi)
                        st.success(f"Added KPI Card: **{new_kpi['label']}**")
                        st.rerun()
                    else:
                        st.error("Please choose a valid metric column.")

            st.markdown("---")
            st.markdown("##### Configured KPI Cards")
            kpis = [item for item in layout if item.get("type") == "kpi"]
            if not kpis:
                st.info("ℹ️ No KPI cards configured yet. Click **➕ Add New KPI Card** above to define your first metric.")
            else:
                for idx, item in enumerate(kpis):
                    with st.container():
                        card_c1, card_c2, card_c3, card_c4 = st.columns([5, 1, 1, 1])
                        with card_c1:
                            st.markdown(
                                f"""<div style="background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);border-left:4px solid #6366f1;border-radius:8px;padding:10px 14px;margin-bottom:6px;">
                                    <div style="font-weight:700;font-size:0.95rem;color:#f1f5f9;">KPI Card {idx + 1}: <span style="color:#818cf8;">{item.get('label')}</span></div>
                                    <div style="font-size:0.8rem;color:#94a3b8;margin-top:4px;">
                                        Sheet: <code>{item.get('sheet')}</code> &nbsp;•&nbsp; Column: <code>{item.get('col')}</code> &nbsp;•&nbsp; 
                                        Agg: <b style="color:#e2e8f0;">{item.get('agg')}</b> &nbsp;•&nbsp; 
                                        Unit: <code>{item.get('unit') or 'None'}</code> &nbsp;•&nbsp; 
                                        Color: <span class="badge badge-{item.get('color', 'indigo')}">{item.get('color')}</span>
                                    </div>
                                </div>""",
                                unsafe_allow_html=True,
                            )
                        with card_c2:
                            if idx > 0 and st.button("⬆️", key=f"kpi_up_{idx}", help="Move Up"):
                                l_idx = layout.index(item)
                                prev_l_idx = layout.index(kpis[idx - 1])
                                layout[l_idx], layout[prev_l_idx] = layout[prev_l_idx], layout[l_idx]
                                st.rerun()
                        with card_c3:
                            if idx < len(kpis) - 1 and st.button("⬇️", key=f"kpi_down_{idx}", help="Move Down"):
                                l_idx = layout.index(item)
                                next_l_idx = layout.index(kpis[idx + 1])
                                layout[l_idx], layout[next_l_idx] = layout[next_l_idx], layout[l_idx]
                                st.rerun()
                        with card_c4:
                            if st.button("🗑️", key=f"kpi_del_{idx}", help="Delete KPI"):
                                layout.remove(item)
                                st.rerun()

        # ── TAB 2: Charts ────────────────────────────────────────────────────
        with designer_tabs[1]:
            st.markdown("#### 📊 Configure Visual Charts")
            st.caption("Add analytical charts such as bar, line, pie, donut, or area charts.")

            with st.expander("➕ Add New Chart", expanded=(total_charts == 0)):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    chart_title = st.text_input("Chart Title", value="Performance Overview", key="chart_title_input")
                    chart_sheet = st.selectbox("Source Sheet", available_sheets, key="chart_src_sheet")
                    chart_type = st.selectbox(
                        "Chart Type",
                        [
                            "Bar",
                            "Stacked Bar",
                            "Horizontal Bar",
                            "Line",
                            "Area",
                            "Pie",
                            "Donut",
                            "Scatter",
                            "Heatmap",
                        ],
                        key="chart_type_picker",
                    )
                with c_f2:
                    c_headers = schema.get("sheets", {}).get(chart_sheet, {}).get("headers", [])
                    if not c_headers:
                        st.warning(f"No headers detected in sheet '{chart_sheet}'")
                    chart_x = st.selectbox("X-Axis (Category / Dimension)", c_headers if c_headers else ["-- none --"], key="chart_x_picker")
                    chart_y = st.multiselect(
                        "Y-Axis (Metrics to Plot)",
                        c_headers,
                        default=[c_headers[0]] if c_headers else [],
                        key="chart_y_picker",
                    )

                if st.button("➕ Add Chart to Layout", key="btn_add_chart"):
                    if chart_x and chart_x != "-- none --" and chart_y:
                        new_chart = {
                            "type": "chart",
                            "id": f"chart_{uuid.uuid4().hex[:6]}",
                            "sheet": chart_sheet,
                            "chart_type": chart_type.lower().replace(" ", "_"),
                            "title": chart_title.strip() if chart_title.strip() else f"{chart_type} Chart",
                            "x": chart_x,
                            "y": chart_y,
                        }
                        layout.append(new_chart)
                        st.success(f"Added Chart: **{new_chart['title']}**")
                        st.rerun()
                    else:
                        st.error("Please specify both an X-Axis column and at least one Y-Axis column.")

            st.markdown("---")
            st.markdown("##### Configured Charts")
            charts = [item for item in layout if item.get("type") == "chart"]
            if not charts:
                st.info("ℹ️ No charts configured yet. Click **➕ Add New Chart** above to add one.")
            else:
                for idx, item in enumerate(charts):
                    with st.container():
                        ch_c1, ch_c2, ch_c3, ch_c4 = st.columns([5, 1, 1, 1])
                        with ch_c1:
                            st.markdown(
                                f"""<div style="background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);border-left:4px solid #10b981;border-radius:8px;padding:10px 14px;margin-bottom:6px;">
                                    <div style="font-weight:700;font-size:0.95rem;color:#f1f5f9;">Chart {idx + 1}: <span style="color:#34d399;">{item.get('title')}</span></div>
                                    <div style="font-size:0.8rem;color:#94a3b8;margin-top:4px;">
                                        Type: <b style="color:#e2e8f0;">{item.get('chart_type')}</b> &nbsp;•&nbsp; Sheet: <code>{item.get('sheet')}</code> &nbsp;•&nbsp; 
                                        X-Axis: <code>{item.get('x')}</code> &nbsp;•&nbsp; 
                                        Y-Axis: <code>{', '.join(item.get('y', []))}</code>
                                    </div>
                                </div>""",
                                unsafe_allow_html=True,
                            )
                        with ch_c2:
                            if idx > 0 and st.button("⬆️", key=f"chart_up_{idx}", help="Move Up"):
                                l_idx = layout.index(item)
                                prev_l_idx = layout.index(charts[idx - 1])
                                layout[l_idx], layout[prev_l_idx] = layout[prev_l_idx], layout[l_idx]
                                st.rerun()
                        with ch_c3:
                            if idx < len(charts) - 1 and st.button("⬇️", key=f"chart_down_{idx}", help="Move Down"):
                                l_idx = layout.index(item)
                                next_l_idx = layout.index(charts[idx + 1])
                                layout[l_idx], layout[next_l_idx] = layout[next_l_idx], layout[l_idx]
                                st.rerun()
                        with ch_c4:
                            if st.button("🗑️", key=f"chart_del_{idx}", help="Delete Chart"):
                                layout.remove(item)
                                st.rerun()

        # ── TAB 3: Tables ────────────────────────────────────────────────────
        with designer_tabs[2]:
            st.markdown("#### 📋 Configure Data Tables")
            st.caption("Add granular data grids with selected columns for operational drill-down.")

            with st.expander("➕ Add New Table", expanded=(total_tables == 0)):
                t_f1, t_f2 = st.columns(2)
                with t_f1:
                    tbl_title = st.text_input("Table Title", value="Operational Records", key="tbl_title_input")
                    tbl_sheet = st.selectbox("Source Sheet", available_sheets, key="tbl_src_sheet")
                    tbl_max_r = st.number_input("Max Rows to Display", min_value=5, max_value=1000, value=50, step=10, key="tbl_max_rows_input")
                with t_f2:
                    tbl_headers = schema.get("sheets", {}).get(tbl_sheet, {}).get("headers", [])
                    if not tbl_headers:
                        st.warning(f"No headers detected in sheet '{tbl_sheet}'")
                    tbl_cols = st.multiselect(
                        "Columns to Include",
                        tbl_headers,
                        default=tbl_headers[:min(8, len(tbl_headers))] if tbl_headers else [],
                        key="tbl_cols_picker",
                    )

                if st.button("➕ Add Table to Layout", key="btn_add_tbl"):
                    new_table = {
                        "type": "table",
                        "id": f"table_{uuid.uuid4().hex[:6]}",
                        "sheet": tbl_sheet,
                        "title": tbl_title.strip() if tbl_title.strip() else "Data Table",
                        "cols": tbl_cols if tbl_cols else tbl_headers,
                        "max_rows": int(tbl_max_r),
                    }
                    layout.append(new_table)
                    st.success(f"Added Table: **{new_table['title']}**")
                    st.rerun()

            st.markdown("---")
            st.markdown("##### Configured Tables")
            tables = [item for item in layout if item.get("type") == "table"]
            if not tables:
                st.info("ℹ️ No tables configured yet. Click **➕ Add New Table** above to add one.")
            else:
                for idx, item in enumerate(tables):
                    with st.container():
                        tb_c1, tb_c2, tb_c3, tb_c4 = st.columns([5, 1, 1, 1])
                        with tb_c1:
                            col_preview = ', '.join(item.get('cols', [])[:5])
                            if len(item.get('cols', [])) > 5:
                                col_preview += f" ... (+{len(item.get('cols', [])) - 5} more)"
                            st.markdown(
                                f"""<div style="background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);border-left:4px solid #f59e0b;border-radius:8px;padding:10px 14px;margin-bottom:6px;">
                                    <div style="font-weight:700;font-size:0.95rem;color:#f1f5f9;">Table {idx + 1}: <span style="color:#fbbf24;">{item.get('title')}</span></div>
                                    <div style="font-size:0.8rem;color:#94a3b8;margin-top:4px;">
                                        Sheet: <code>{item.get('sheet')}</code> &nbsp;•&nbsp; 
                                        Columns ({len(item.get('cols', []))}): <code>{col_preview}</code> &nbsp;•&nbsp; 
                                        Max Rows: <b style="color:#e2e8f0;">{item.get('max_rows')}</b>
                                    </div>
                                </div>""",
                                unsafe_allow_html=True,
                            )
                        with tb_c2:
                            if idx > 0 and st.button("⬆️", key=f"tbl_up_{idx}", help="Move Up"):
                                l_idx = layout.index(item)
                                prev_l_idx = layout.index(tables[idx - 1])
                                layout[l_idx], layout[prev_l_idx] = layout[prev_l_idx], layout[l_idx]
                                st.rerun()
                        with tb_c3:
                            if idx < len(tables) - 1 and st.button("⬇️", key=f"tbl_down_{idx}", help="Move Down"):
                                l_idx = layout.index(item)
                                next_l_idx = layout.index(tables[idx + 1])
                                layout[l_idx], layout[next_l_idx] = layout[next_l_idx], layout[l_idx]
                                st.rerun()
                        with tb_c4:
                            if st.button("🗑️", key=f"tbl_del_{idx}", help="Delete Table"):
                                layout.remove(item)
                                st.rerun()

        # ── TAB 4: Live Layout Preview & Reorder ──────────────────────────────
        with designer_tabs[3]:
            st.markdown("#### 👁️ Complete Dashboard Layout Order")
            st.caption("This represents the exact sequence in which components will be rendered on the live Dashboard.")

            if not layout:
                st.info("ℹ️ Dashboard layout is currently empty. Use the tabs above to add KPI cards, charts, and tables.")
            else:
                for idx, item in enumerate(layout):
                    col_p1, col_p2, col_p3, col_p4 = st.columns([5, 1, 1, 1])
                    itype = item.get("type", "").upper()
                    ititle = item.get("title") or item.get("label") or "Untitled"
                    isheet = item.get("sheet", "")

                    badge_style = {
                        "KPI": "background:rgba(99,102,241,0.2);color:#818cf8;border:1px solid rgba(99,102,241,0.4);",
                        "CHART": "background:rgba(16,185,129,0.2);color:#34d399;border:1px solid rgba(16,185,129,0.4);",
                        "TABLE": "background:rgba(245,158,11,0.2);color:#fbbf24;border:1px solid rgba(245,158,11,0.4);",
                    }.get(itype, "")

                    col_p1.markdown(
                        f"""<div style="background:rgba(255,255,255,0.02);border:1px solid rgba(255,255,255,0.06);border-radius:6px;padding:8px 12px;display:flex;align-items:center;gap:10px;">
                            <span style="font-weight:700;color:#94a3b8;min-width:28px;">#{idx + 1}</span>
                            <span style="{badge_style}padding:2px 8px;border-radius:4px;font-size:0.75rem;font-weight:700;">{itype}</span>
                            <span style="font-weight:600;color:#f1f5f9;">{ititle}</span>
                            <span style="color:#64748b;font-size:0.8rem;margin-left:auto;">Sheet: <code>{isheet}</code></span>
                        </div>""",
                        unsafe_allow_html=True,
                    )
                    with col_p2:
                        if idx > 0 and st.button("⬆️", key=f"glob_up_{idx}", help="Move Up in Layout"):
                            layout[idx], layout[idx - 1] = layout[idx - 1], layout[idx]
                            st.rerun()
                    with col_p3:
                        if idx < len(layout) - 1 and st.button("⬇️", key=f"glob_down_{idx}", help="Move Down in Layout"):
                            layout[idx], layout[idx + 1] = layout[idx + 1], layout[idx]
                            st.rerun()
                    with col_p4:
                        if st.button("🗑️", key=f"glob_del_{idx}", help="Delete Item"):
                            layout.remove(item)
                            st.rerun()

                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️ Clear Entire Dashboard Layout", type="secondary"):
                    st.session_state["wizard_dashboard_layout"] = []
                    st.rerun()

        st.markdown("<br><hr>", unsafe_allow_html=True)
        nav1, nav2, nav3 = st.columns([1, 1, 1])
        if nav1.button("⬅️ Back to Step 2"):
            st.session_state["setup_step"] = 2
            st.rerun()
        if nav2.button("💾 Save & Continue to Step 4 ➡️", type="primary"):
            st.session_state["setup_step"] = 4
            st.rerun()
        if nav3.button("📊 Go to Live Dashboard", help="Switch immediately to the Dashboard to see your layout in action"):
            st.session_state["page_override"] = "📊 Dashboard"
            st.rerun()

    # ── STEP 4: Preview & Save ────────────────────────────────────────────
    elif st.session_state["setup_step"] == 4:
        st.subheader("Step 4: Review Configuration & Save Template")

        tmpl_id = st.session_state.get("wizard_template_id")
        tmpl_name = st.session_state.get("wizard_template_name")
        tmpl_path = st.session_state.get("wizard_template_path")
        schema = st.session_state.get("wizard_schema", {})
        target_sheets = st.session_state.get("wizard_target_sheets", [])
        mapping = st.session_state.get("wizard_column_mapping", {})
        formula_cols = st.session_state.get("wizard_formula_cols", {})
        layout = st.session_state.get("wizard_dashboard_layout", [])

        st.markdown(f"""
        <div style="background:rgba(26,34,56,0.8);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px;margin-bottom:16px;">
            <h4 style="margin-top:0;color:#6366f1;">📄 {tmpl_name}</h4>
            <p><b>Target Ingestion Sheets:</b> {', '.join(target_sheets)}</p>
            <p><b>Total Designed Dashboard Items:</b> {len(layout)} ({len([i for i in layout if i.get('type')=='kpi'])} KPIs, {len([i for i in layout if i.get('type')=='chart'])} Charts, {len([i for i in layout if i.get('type')=='table'])} Tables)</p>
        </div>
        """, unsafe_allow_html=True)

        set_active = st.checkbox("Set this template as the ACTIVE default template", value=True)

        c_s1, c_s2 = st.columns([1, 1])
        if c_s1.button("⬅️ Back to Step 3"):
            st.session_state["setup_step"] = 3
            st.rerun()

        if c_s2.button("💾 Finalize & Save Template", type="primary"):
            meta = load_meta()
            meta["templates"][tmpl_id] = {
                "name": tmpl_name,
                "file": tmpl_path,
                "schema": schema,
                "target_sheets": target_sheets,
                "sheet": target_sheets[0] if target_sheets else list(schema.get("sheets", {}).keys())[0],
                "column_mapping": mapping,
                "formula_cols": formula_cols,
                "dashboard_layout": layout,
                "saved_at": str(datetime.now()),
            }
            if set_active or not meta.get("active_template"):
                meta["active_template"] = tmpl_id

            save_meta(meta)
            st.success(f"🎉 Template **{tmpl_name}** successfully saved with full column mappings and custom dashboard layout!")
            st.session_state["setup_step"] = 1
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 3: 📤 UPLOAD DAILY DATA (Multi-Sheet Aware & Validating)
# ─────────────────────────────────────────────────────────────────────────────
elif page == "📤 Upload Data":
    st.title("📤 Upload Daily Operational Data")

    active_key = meta.get("active_template")
    if not active_key or active_key not in meta.get("templates", {}):
        st.warning("⚠️ No active template found. Go to **Template Setup** first to configure one.")
        st.stop()

    tmpl_info = meta["templates"][active_key]
    tmpl_path = Path(tmpl_info["file"])
    target_sheets = tmpl_info.get("target_sheets", [tmpl_info.get("sheet", "Dashboard")])

    st.markdown(f"""
    <div style="background:rgba(99,102,241,0.08);border:1px solid rgba(99,102,241,0.25);border-radius:12px;padding:14px 20px;margin-bottom:16px;">
        <span style="font-weight:700;color:#f1f5f9;">Active Template:</span> <b>{tmpl_info.get('name')}</b> | 
        <span style="font-weight:700;color:#f1f5f9;">Target Ingestion Sheets:</span> <code>{', '.join(target_sheets)}</code>
    </div>
    """, unsafe_allow_html=True)

    col_u1, col_u2 = st.columns([2, 1])

    with col_u1:
        upload_date = st.date_input("Date for Data Records", value=date.today())
        raw_upload = st.file_uploader(
            "Drop daily raw file (.xlsx or .csv)",
            type=["xlsx", "csv"],
            key="daily_raw_uploader",
        )
        overwrite = st.checkbox("Overwrite existing data if record for this date already exists", value=False)

    with col_u2:
        st.markdown("##### Template Expected Mapping:")
        mapping = tmpl_info.get("column_mapping", {})
        for sname in target_sheets:
            s_map = mapping.get(sname, {})
            req_cols = [k for k, v in s_map.items() if v not in ("-- formula column --", "-- ignore --", None, "")]
            f_cols = [k for k, v in s_map.items() if v == "-- formula column --"]
            st.write(f"**Sheet `{sname}`:**")
            st.caption(f"Inputs ({len(req_cols)}): {', '.join(req_cols[:5])}...")
            if f_cols:
                st.caption(f"Formulas ({len(f_cols)}): {', '.join(f_cols[:3])}...")

    if raw_upload:
        raw_bytes = raw_upload.read()
        upload_sources: Dict[str, pd.DataFrame] = {}

        try:
            if raw_upload.name.endswith(".csv"):
                df_csv = pd.read_csv(io.BytesIO(raw_bytes))
                for sname in target_sheets:
                    upload_sources[sname] = df_csv
            else:
                excel_sheets = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None)
                # Map sheets: if exact match exists use it; otherwise assign primary
                for sname in target_sheets:
                    if sname in excel_sheets:
                        upload_sources[sname] = excel_sheets[sname]
                    else:
                        upload_sources[sname] = next(iter(excel_sheets.values()))
        except Exception as e:
            st.error(f"Error reading uploaded file: {e}")
            st.stop()

        # Validation Report
        st.markdown("---")
        st.subheader("📋 Pre-Injection Data Validation")
        val_tabs = st.tabs([f"Sheet: {s}" for s in target_sheets])

        can_proceed = True
        for idx, sname in enumerate(target_sheets):
            with val_tabs[idx]:
                df_src = upload_sources.get(sname, pd.DataFrame())
                s_map = mapping.get(sname, {})
                needed_upload_cols = [v for k, v in s_map.items() if v not in ("-- formula column --", "-- ignore --", "-- select --", None, "")]
                src_cols = [str(c).strip() for c in df_src.columns]

                missing = [c for c in needed_upload_cols if c not in src_cols]
                blanks = int(df_src.isna().sum().sum())
                dups = int(df_src.duplicated().sum())

                v1, v2, v3, v4 = st.columns(4)
                v1.metric("Rows Found", len(df_src))
                v2.metric("Missing Columns", len(missing))
                v3.metric("Blank Values", blanks)
                v4.metric("Duplicates", dups)

                if missing:
                    st.error(f"❌ Missing required columns: {', '.join(missing)}")
                    can_proceed = False
                else:
                    st.success("✅ All mapped columns are present in upload.")

                st.dataframe(df_src.head(5), use_container_width=True)

        # Trigger Calculation Button
        if can_proceed:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("⚡ Inject Data & Run Excel Full Recalculation", type="primary", use_container_width=True):
                dd = get_date_dir(upload_date)
                calc_dest = dd / "calculated.xlsx"

                if calc_dest.exists() and not overwrite:
                    st.error(f"Data for {upload_date} already exists. Check 'Overwrite existing' to update.")
                    st.stop()

                with st.spinner("⏳ Writing records to sheets and invoking Microsoft Excel calculation engine..."):
                    calc_ok = inject_multi_sheet_data(tmpl_path, upload_sources, tmpl_info, calc_dest)

                if calc_dest.exists():
                    show_formula_error_warning(calc_dest)
                if calc_ok:
                    save_day_data(upload_date, raw_bytes, raw_upload.name, calc_dest, tmpl_info)
                    st.success(f"🎉 Recalculation successful! Data persisted for **{upload_date}**.")
                    st.balloons()
                else:
                    # Still save even if COM calculation encountered warnings
                    save_day_data(upload_date, raw_bytes, raw_upload.name, calc_dest, tmpl_info)
                    st.warning("Workbook saved with injected formulas. (Excel background calculation finished with warnings).")

    # Inbox Auto-Import Section
    st.markdown("---")
    st.subheader("📥 Inbox Folder Auto-Import")
    st.markdown(f"Drop daily files directly into `{INBOX_DIR}` for batch processing.")
    inbox_files = list(INBOX_DIR.glob("*.xlsx")) + list(INBOX_DIR.glob("*.csv"))
    if inbox_files:
        for f in inbox_files:
            ci1, ci2 = st.columns([3, 1])
            ci1.write(f"📄 `{f.name}` ({f.stat().st_size // 1024} KB)")
            if ci2.button(f"Process Import", key=f"btn_inbox_{f.name}"):
                f_bytes = f.read_bytes()
                dd = get_date_dir(date.today())
                calc_dest = dd / "calculated.xlsx"
                try:
                    if f.suffix.lower() == ".csv":
                        u_df = pd.read_csv(io.BytesIO(f_bytes))
                        u_sources = {s: u_df for s in target_sheets}
                    else:
                        u_sheets = pd.read_excel(io.BytesIO(f_bytes), sheet_name=None)
                        u_sources = {s: u_sheets.get(s, next(iter(u_sheets.values()))) for s in target_sheets}

                    inject_multi_sheet_data(tmpl_path, u_sources, tmpl_info, calc_dest)
                    save_day_data(date.today(), f_bytes, f.name, calc_dest, tmpl_info)
                    f.unlink()  # Remove processed file from inbox
                    st.success(f"Imported `{f.name}` for today!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error processing inbox file: {e}")
    else:
        st.info(f"Inbox is empty. Drop files into `{INBOX_DIR}` to auto-detect.")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 4: 📅 HISTORY EXPLORER
# ─────────────────────────────────────────────────────────────────────────────
elif page == "📅 History":
    st.title("📅 Historical Records Explorer")

    if not available_dates:
        st.info("No historical data stored yet.")
        st.stop()
    else:
        st.markdown(f"Total **{len(available_dates)}** daily report(s) archived.")

        for d_item in available_dates:
            m = load_day_meta(d_item)
            sheets = load_day_sheets(d_item)
            total_rows = sum(len(df) for df in sheets.values()) if sheets else 0

            with st.expander(f"📅 {d_item.strftime('%d %B %Y (%A)')} — {total_rows} total rows across {len(sheets)} sheet(s)"):
                c_h1, c_h2, c_h3 = st.columns(3)
                c_h1.metric("Ingested File", m.get("upload", "N/A"))
                c_h2.metric("Template Used", m.get("template", "N/A"))
                c_h3.metric("Saved At", m.get("saved_at", "N/A")[:16])

                if sheets:
                    st.markdown("##### Sheet Row Distribution:")
                    for sname, sdf in sheets.items():
                        st.write(f"• **`{sname}`**: {len(sdf)} rows, {len(sdf.columns)} columns")

                # Outlier check on this day
                p_df = load_day_main_df(d_item)
                if p_df is not None:
                    num_cols, _, _ = classify_columns(p_df)
                    if num_cols:
                        st.dataframe(p_df[num_cols[:6]].describe().round(2), use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 5: 📈 DAY-OVER-DAY COMPARISON
# ─────────────────────────────────────────────────────────────────────────────
elif page == "📈 Comparison":
    st.title("📈 Day-over-Day & Period Comparison")

    if len(available_dates) < 2:
        st.info("At least 2 days of data are required to run comparative variance analysis.")
        st.stop()
    else:
        col_cmp1, col_cmp2 = st.columns(2)
        with col_cmp1:
            base_d = st.selectbox("Base Date (Prior)", available_dates[1:], format_func=lambda d: d.strftime("%d %b %Y"))
        with col_cmp2:
            comp_d = st.selectbox("Comparison Date (Recent)", available_dates, format_func=lambda d: d.strftime("%d %b %Y"))

        df_base = load_day_main_df(base_d)
        df_comp = load_day_main_df(comp_d)

        if df_base is None or df_comp is None:
            st.error("Failed to load records for one of the dates.")
            st.stop()

        num1, _, _ = classify_columns(df_base)
        num2, _, _ = classify_columns(df_comp)
        common_nums = [c for c in num1 if c in num2]

        if not common_nums:
            st.warning("No overlapping numeric columns found between the selected dates.")
            st.stop()

        st.markdown("---")
        st.subheader("Metric Variance Overview")
        v_cols = st.columns(min(4, len(common_nums)))
        deltas_for_waterfall = []

        for i, col in enumerate(common_nums[:8]):
            v_base = float(df_base[col].sum())
            v_comp = float(df_comp[col].sum())
            diff = v_comp - v_base
            deltas_for_waterfall.append((col, diff))
            with v_cols[i % len(v_cols)]:
                st.metric(
                    label=col,
                    value=f"{v_comp:,.2f}",
                    delta=f"{diff:+,.2f} vs {base_d.strftime('%d %b')}",
                )

        # Waterfall Variance Chart
        st.markdown("---")
        st.subheader("Waterfall Delta Breakdown")
        wf_cats = [x[0] for x in deltas_for_waterfall[:6]]
        wf_vals = [x[1] for x in deltas_for_waterfall[:6]]
        fig_wf = build_waterfall_chart(wf_cats, wf_vals, "Day-over-Day Variances")
        st.plotly_chart(fig_wf, use_container_width=True)

        # Side-by-side comparison
        st.markdown("---")
        st.subheader("Side-by-Side Category Comparison")
        chosen_metric = st.selectbox("Compare Metric", common_nums)
        text1, _, _ = classify_columns(df_comp)
        cat_col = text1[0] if text1 else None

        if cat_col and cat_col in df_base.columns:
            s1 = df_base.groupby(cat_col)[chosen_metric].sum().reset_index().rename(columns={chosen_metric: base_d.strftime("%d %b")})
            s2 = df_comp.groupby(cat_col)[chosen_metric].sum().reset_index().rename(columns={chosen_metric: comp_d.strftime("%d %b")})
            merged = pd.merge(s1, s2, on=cat_col, how="outer").fillna(0)
            fig_bar = build_bar_chart(merged, cat_col, [base_d.strftime("%d %b"), comp_d.strftime("%d %b")], f"{chosen_metric} Comparison")
            st.plotly_chart(fig_bar, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 6: 📋 DATA VALIDATION AUDIT
# ─────────────────────────────────────────────────────────────────────────────
elif page == "📋 Data Validation":
    st.title("📋 Data Integrity & Anomaly Audit")

    if not available_dates:
        st.info("No data available to audit.")
        st.stop()
    else:
        v_date = st.selectbox("Select Date to Audit", available_dates, format_func=lambda d: d.strftime("%d %B %Y"))
        sheets = load_day_sheets(v_date)

        if not sheets:
            st.error("Could not load data.")
            st.stop()

        audit_sheet = st.selectbox("Select Sheet to Inspect", list(sheets.keys()))
        df_audit = sheets[audit_sheet]

        # Metrics Overview
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("Total Records", len(df_audit))
        a2.metric("Total Columns", len(df_audit.columns))
        a3.metric("Missing Cells", int(df_audit.isna().sum().sum()))
        a4.metric("Duplicate Rows", int(df_audit.duplicated().sum()))

        st.markdown("---")

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("##### Missing Values per Column")
            na_counts = df_audit.isna().sum()
            na_counts = na_counts[na_counts > 0]
            if na_counts.empty:
                st.success("✅ No null values found across any columns.")
            else:
                st.dataframe(pd.DataFrame({"Column": na_counts.index, "Missing Count": na_counts.values}), use_container_width=True)

        with col_r:
            st.markdown("##### Data Types Breakdown")
            type_df = pd.DataFrame({"Column": df_audit.columns, "Inferred Type": [str(t) for t in df_audit.dtypes]})
            st.dataframe(type_df, use_container_width=True)

        # Anomaly Scanner Table (>2 Std Devs)
        st.markdown("---")
        st.markdown("##### 🚨 Statistical Outliers (>2σ from Mean)")
        num_cols, _, _ = classify_columns(df_audit)
        outlier_rows = []

        for c in num_cols:
            count, indices = detect_anomalies(df_audit, c, threshold_std=2.0)
            if count > 0:
                mean = df_audit[c].mean()
                std = df_audit[c].std()
                for idx in indices[:5]:
                    val = df_audit.loc[idx, c]
                    outlier_rows.append({
                        "Row Index": idx,
                        "Column": c,
                        "Value": val,
                        "Mean": round(mean, 2),
                        "Std Dev": round(std, 2),
                        "Z-Score": round(abs((val - mean) / std), 2),
                    })

        if outlier_rows:
            st.dataframe(pd.DataFrame(outlier_rows), use_container_width=True)
        else:
            st.success("✅ No extreme statistical outliers detected in numeric metrics.")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 7: ⚙️ SETTINGS & PRESETS
# ─────────────────────────────────────────────────────────────────────────────
elif page == "⚙️ Settings":
    st.title("⚙️ System Settings & Diagnostics")

    # Alert Thresholds
    st.subheader("🔔 Metric Alert Thresholds")
    st.markdown("Configure operational thresholds. Metrics crossing these thresholds will trigger dashboard alerts.")

    if available_dates:
        sample_df = load_day_main_df(available_dates[0])
        if sample_df is not None:
            numeric_cols, _, _ = classify_columns(sample_df)
            curr_th = meta.get("thresholds", {})
            new_th: Dict[str, dict] = {}

            for c in numeric_cols[:6]:
                with st.expander(f"Threshold for `{c}`"):
                    tc1, tc2 = st.columns(2)
                    w_below = tc1.text_input("Warn if Below", value=str(curr_th.get(c, {}).get("min", "")), key=f"th_min_{c}")
                    w_above = tc2.text_input("Warn if Above", value=str(curr_th.get(c, {}).get("max", "")), key=f"th_max_{c}")
                    t_entry = {}
                    try:
                        if w_below:
                            t_entry["min"] = float(w_below)
                    except ValueError:
                        pass
                    try:
                        if w_above:
                            t_entry["max"] = float(w_above)
                    except ValueError:
                        pass
                    if t_entry:
                        new_th[c] = t_entry

            if st.button("💾 Save Thresholds"):
                meta["thresholds"] = new_th
                save_meta(meta)
                st.success("Alert thresholds saved!")
    else:
        st.info("Upload daily data first to configure column thresholds.")

    # Excel Diagnostics
    st.markdown("---")
    st.subheader("🧪 Microsoft Excel COM Diagnostics")
    st.code(f"Excel Executable: {EXCEL_EXE}\nCOM Available:   win32com.client active\nStorage Path:     {DATA_DIR}\nTemplates Path:   {TMPL_DIR}")

    if st.button("Run Excel COM Engine Test"):
        with st.spinner("Testing Excel COM automation..."):
            test_xlsx = TMPL_DIR / "_test_com.xlsx"
            wb_test = load_workbook(TMPL_DIR / next(iter(meta.get("templates", {}).values()))["file"]) if meta.get("templates") else load_workbook()
            wb_test.save(test_xlsx)
            com_res = recalc_via_excel_com(test_xlsx)
            try:
                test_xlsx.unlink()
            except Exception:
                pass
        if com_res:
            st.success("✅ Microsoft Excel COM automation dispatched and recalculated successfully!")
        else:
            st.warning("⚠️ COM automation returned false; VBS subprocess automation will be used as fallback.")

    # Data Reset
    st.markdown("---")
    st.subheader("🗄️ Storage Management")
    if st.button("🗑️ Clear All Stored Data Records (Irreversible)", type="secondary"):
        shutil.rmtree(DATA_DIR, ignore_errors=True)
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        st.warning("All data records have been wiped clean.")
        st.rerun()
