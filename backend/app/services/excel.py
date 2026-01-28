from __future__ import annotations

import re
from io import BytesIO
from typing import Any, Dict, Iterable, List, Tuple

import math
import numpy as np
import pandas as pd
from scipy import stats


CELL_RE = re.compile(r"^([A-Za-z]+)(\\d+)$")
RC_RE = re.compile(r"^R(\\d+)C(\\d+)$", re.IGNORECASE)
COMMA_RE = re.compile(r"^(\\d+)\\s*,\\s*(\\d+)$")


def load_excel(file_bytes: bytes, filename: str | None = None) -> Dict[str, pd.DataFrame]:
    if filename and filename.lower().endswith(".csv"):
        df = pd.read_csv(BytesIO(file_bytes))
        return {"Sheet1": _normalize_df(df)}
    sheets = pd.read_excel(BytesIO(file_bytes), sheet_name=None, engine="openpyxl")
    normalized = {}
    for name, df in sheets.items():
        normalized[name] = _normalize_df(df)
    return normalized


def create_empty(sheet_name: str) -> Dict[str, pd.DataFrame]:
    return {sheet_name: pd.DataFrame()}


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(col) for col in df.columns]
    return df


def preview(df: pd.DataFrame, limit: int) -> Tuple[List[str], List[Dict[str, Any]]]:
    if df is None or df.empty:
        return [], []
    sample = df.head(limit).copy()
    sample.replace([np.inf, -np.inf], None, inplace=True)
    sample = sample.where(pd.notna(sample), None)
    rows = sample.to_dict(orient="records")
    return list(sample.columns), _sanitize_rows(rows)


def _sanitize_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    for row in rows:
        for key, value in list(row.items()):
            if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                row[key] = None
            elif value is pd.NA:
                row[key] = None
    return rows


def apply_operations(
    sheets: Dict[str, pd.DataFrame],
    operations: Iterable[Dict[str, Any]],
    format_rules: List[dict] | None = None,
) -> Tuple[List[str], List[dict]]:
    if format_rules is None:
        format_rules = []
    changed = set()
    analysis: List[dict] = []
    for op in operations:
        op_type = op.get("type")
        if op_type == "add_sheet":
            name = op.get("to") or op.get("sheet") or "Sheet"
            if name not in sheets:
                sheets[name] = pd.DataFrame()
            changed.add(name)
        elif op_type == "rename_sheet":
            src = op.get("from")
            dst = op.get("to")
            if src and dst and src in sheets:
                sheets[dst] = sheets.pop(src)
                changed.add(dst)
        elif op_type == "add_column":
            sheet = _get_sheet(sheets, op)
            column_name = op.get("column_name") or op.get("column")
            value = op.get("value")
            if column_name:
                sheet[column_name] = value
                changed.add(_sheet_name(sheets, sheet))
        elif op_type == "rename_column":
            sheet = _get_sheet(sheets, op)
            old = op.get("column")
            new = op.get("new_name") or op.get("column_name")
            if old and new and old in sheet.columns:
                sheet.rename(columns={old: new}, inplace=True)
                changed.add(_sheet_name(sheets, sheet))
        elif op_type == "swap_columns":
            sheet_name = op.get("sheet") or (list(sheets.keys())[0] if sheets else "Sheet1")
            sheet = _get_sheet(sheets, op)
            a = op.get("column_a") or op.get("from")
            b = op.get("column_b") or op.get("to")
            idx_a = op.get("column_index_a")
            idx_b = op.get("column_index_b")
            if (idx_a is not None) and (idx_b is not None):
                try:
                    a = sheet.columns[int(idx_a)]
                    b = sheet.columns[int(idx_b)]
                except (ValueError, IndexError):
                    return sorted(changed)
            if a and a not in sheet.columns and isinstance(a, str) and a.isalpha():
                try:
                    a = sheet.columns[_col_letters_to_index(a)]
                except IndexError:
                    return sorted(changed)
            if b and b not in sheet.columns and isinstance(b, str) and b.isalpha():
                try:
                    b = sheet.columns[_col_letters_to_index(b)]
                except IndexError:
                    return sorted(changed)
            if a and b and a in sheet.columns and b in sheet.columns:
                cols = list(sheet.columns)
                ai = cols.index(a)
                bi = cols.index(b)
                cols[ai], cols[bi] = cols[bi], cols[ai]
                sheets[sheet_name] = sheet[cols]
                changed.add(sheet_name)
        elif op_type == "round_column":
            sheet = _get_sheet(sheets, op)
            col = op.get("column")
            decimals = op.get("decimals", 0)
            if col in sheet.columns:
                sheet[col] = pd.to_numeric(sheet[col], errors="coerce").round(int(decimals))
                changed.add(_sheet_name(sheets, sheet))
                format_rules.append(
                    {
                        "type": "number_format",
                        "sheet": _sheet_name(sheets, sheet),
                        "column": col,
                        "format": f"0.{('0' * int(decimals))}" if int(decimals) > 0 else "0",
                    }
                )
        elif op_type == "format_lt":
            sheet_name = op.get("sheet") or (list(sheets.keys())[0] if sheets else "Sheet1")
            col = op.get("column")
            threshold = op.get("threshold")
            color = op.get("color") or "red"
            if col and threshold is not None:
                format_rules.append(
                    {"type": "lt", "sheet": sheet_name, "column": col, "threshold": float(threshold), "color": color}
                )
        elif op_type == "t_test":
            sheet_name = op.get("sheet") or (list(sheets.keys())[0] if sheets else "Sheet1")
            sheet = _get_sheet(sheets, op)
            col_a = op.get("column_a")
            col_b = op.get("column_b")
            equal_var = op.get("equal_var", False)
            output = op.get("output") or "sheet"
            prefix = op.get("output_prefix") or "t_test"
            if col_a in sheet.columns and col_b in sheet.columns:
                a = pd.to_numeric(sheet[col_a], errors="coerce").dropna()
                b = pd.to_numeric(sheet[col_b], errors="coerce").dropna()
                if len(a) > 1 and len(b) > 1:
                    t_stat, p_value = stats.ttest_ind(a, b, equal_var=bool(equal_var), nan_policy="omit")
                    if equal_var:
                        df = float(len(a) + len(b) - 2)
                    else:
                        var_a = a.var(ddof=1)
                        var_b = b.var(ddof=1)
                        df = float(
                            (var_a / len(a) + var_b / len(b)) ** 2
                            / (((var_a / len(a)) ** 2) / (len(a) - 1) + ((var_b / len(b)) ** 2) / (len(b) - 1))
                        )
                    analysis.append(
                        {
                            "type": "t_test",
                            "sheet": sheet_name,
                            "column_a": col_a,
                            "column_b": col_b,
                            "n_a": int(len(a)),
                            "n_b": int(len(b)),
                            "mean_a": float(a.mean()),
                            "mean_b": float(b.mean()),
                            "t_stat": float(t_stat),
                            "p_value": float(p_value),
                            "df": df,
                        }
                    )
                    if output == "column":
                        t_col = f"{prefix}_t"
                        p_col = f"{prefix}_p"
                        df_col = f"{prefix}_df"
                        mean_a_col = f"{prefix}_mean_a"
                        mean_b_col = f"{prefix}_mean_b"
                        sheet[t_col] = float(t_stat)
                        sheet[p_col] = float(p_value)
                        sheet[df_col] = float(df)
                        sheet[mean_a_col] = float(a.mean())
                        sheet[mean_b_col] = float(b.mean())
                        changed.add(sheet_name)
                    else:
                        result_df = pd.DataFrame(
                            [
                                {
                                    "column_a": col_a,
                                    "column_b": col_b,
                                    "n_a": int(len(a)),
                                    "n_b": int(len(b)),
                                    "mean_a": float(a.mean()),
                                    "mean_b": float(b.mean()),
                                    "t_stat": float(t_stat),
                                    "p_value": float(p_value),
                                    "df": df,
                                }
                            ]
                        )
                        sheets["统计结果"] = result_df
                        changed.add("统计结果")
        elif op_type == "set_cell":
            sheet = _get_sheet(sheets, op)
            cell = op.get("cell")
            value = op.get("value")
            if cell:
                row_idx, col_name = _cell_to_indices(sheet, cell)
                _ensure_row(sheet, row_idx)
                if col_name not in sheet.columns:
                    sheet[col_name] = None
                sheet.at[row_idx, col_name] = value
                changed.add(_sheet_name(sheets, sheet))
        elif op_type == "set_range":
            sheet = _get_sheet(sheets, op)
            range_ref = op.get("range")
            value = op.get("value")
            if range_ref:
                for row_idx, col_name in _range_to_indices(sheet, range_ref):
                    _ensure_row(sheet, row_idx)
                    if col_name not in sheet.columns:
                        sheet[col_name] = None
                    sheet.at[row_idx, col_name] = value
                changed.add(_sheet_name(sheets, sheet))
        elif op_type == "delete_rows":
            sheet = _get_sheet(sheets, op)
            rows = op.get("rows") or []
            if rows:
                drop_idx = [r - 1 for r in rows if r > 0]
                sheet.drop(index=drop_idx, inplace=True, errors="ignore")
                sheet.reset_index(drop=True, inplace=True)
                changed.add(_sheet_name(sheets, sheet))
        elif op_type == "update_cells":
            sheet = _get_sheet(sheets, op)
            where = op.get("where") or {}
            set_values = op.get("set") or {}
            col = where.get("column")
            value = where.get("value")
            if col in sheet.columns:
                mask = sheet[col] == value
                for target_col, target_val in set_values.items():
                    if target_col not in sheet.columns:
                        sheet[target_col] = None
                    sheet.loc[mask, target_col] = target_val
                changed.add(_sheet_name(sheets, sheet))
        elif op_type == "sort":
            sheet = _get_sheet(sheets, op)
            by = op.get("by")
            ascending = op.get("ascending", True)
            if by in sheet.columns:
                sheet.sort_values(by=by, ascending=ascending, inplace=True, kind="mergesort")
                sheet.reset_index(drop=True, inplace=True)
                changed.add(_sheet_name(sheets, sheet))
    return sorted(changed), analysis


def _get_sheet(sheets: Dict[str, pd.DataFrame], op: Dict[str, Any]) -> pd.DataFrame:
    name = op.get("sheet") or (list(sheets.keys())[0] if sheets else "Sheet1")
    if name not in sheets:
        sheets[name] = pd.DataFrame()
    return sheets[name]


def _sheet_name(sheets: Dict[str, pd.DataFrame], df: pd.DataFrame) -> str:
    for name, sheet in sheets.items():
        if sheet is df:
            return name
    return list(sheets.keys())[0]


def _cell_to_indices(df: pd.DataFrame, cell: str) -> Tuple[int, str]:
    raw = cell.strip()
    match = CELL_RE.match(raw)
    if match:
        col_letters, row_str = match.groups()
        col_idx = _col_letters_to_index(col_letters)
        col_name = _col_index_to_name(df, col_idx)
        return int(row_str) - 1, col_name
    rc_match = RC_RE.match(raw)
    if rc_match:
        row_str, col_str = rc_match.groups()
        col_idx = int(col_str) - 1
        col_name = _col_index_to_name(df, col_idx)
        return int(row_str) - 1, col_name
    comma_match = COMMA_RE.match(raw)
    if comma_match:
        row_str, col_str = comma_match.groups()
        col_idx = int(col_str) - 1
        col_name = _col_index_to_name(df, col_idx)
        return int(row_str) - 1, col_name
    raise ValueError("invalid_cell")


def _range_to_indices(df: pd.DataFrame, range_ref: str) -> List[Tuple[int, str]]:
    parts = range_ref.split(":")
    if len(parts) != 2:
        raise ValueError("invalid_range")
    start = _cell_to_indices(df, parts[0])
    end = _cell_to_indices(df, parts[1])
    rows = range(min(start[0], end[0]), max(start[0], end[0]) + 1)
    cols = range(
        _col_letters_to_index(CELL_RE.match(parts[0]).group(1)),
        _col_letters_to_index(CELL_RE.match(parts[1]).group(1)) + 1,
    )
    indices = []
    for r in rows:
        for c in cols:
            indices.append((r, _col_index_to_name(df, c)))
    return indices


def _col_letters_to_index(letters: str) -> int:
    result = 0
    for char in letters.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def _col_index_to_name(df: pd.DataFrame, idx: int) -> str:
    if idx < len(df.columns):
        return df.columns[idx]
    return _index_to_letters(idx)


def _index_to_letters(idx: int) -> str:
    idx += 1
    letters = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(rem + ord("A")) + letters
    return letters


def _ensure_row(df: pd.DataFrame, row_idx: int) -> None:
    if row_idx < len(df.index):
        return
    missing = row_idx - len(df.index) + 1
    for _ in range(missing):
        df.loc[len(df)] = [None for _ in range(len(df.columns))]
