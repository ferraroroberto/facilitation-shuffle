"""
Summary building — combine roster with group assignments and serialise to xlsx.

Pure data logic (pandas/openpyxl only); never imports streamlit.
"""
from __future__ import annotations

import io

import pandas as pd
from openpyxl.styles import Alignment, Font

from src.randomize_groups import _assign_ids


def build_summary_df(
    roster_df: pd.DataFrame,
    present_set: set[str],
    groups: dict[str, list[list[str]]],
) -> pd.DataFrame:
    """
    Combine roster with current group assignments into a single DataFrame.
    Non-present participants get empty group columns.
    """
    # Build lookup: name → group number
    name_col = roster_df.columns[0]

    pair_map = _assign_ids(groups["pairs"])
    g4a_map  = _assign_ids(groups["g4a"])
    g4b_map  = _assign_ids(groups["g4b"])

    df = roster_df.copy()
    df["present"]            = df[name_col].apply(lambda n: 1 if n in present_set else 0)
    df["pair (round 1)"]     = df[name_col].apply(lambda n: pair_map.get(n, pd.NA))
    df["group_4A (round 2)"] = df[name_col].apply(lambda n: g4a_map.get(n, pd.NA))
    df["group_4B (round 3)"] = df[name_col].apply(lambda n: g4b_map.get(n, pd.NA))

    # Cast to nullable integer so Arrow serialization works cleanly (absent rows → <NA>)
    for col in ["pair (round 1)", "group_4A (round 2)", "group_4B (round 3)"]:
        df[col] = df[col].astype(pd.Int64Dtype())

    return df


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Serialise a DataFrame to xlsx bytes with light formatting."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Groups")
        ws = writer.sheets["Groups"]
        # Bold headers + auto-width
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for col_cells in ws.columns:
            width = max(len(str(c.value or "")) for c in col_cells) + 4
            ws.column_dimensions[col_cells[0].column_letter].width = min(width, 40)
    buf.seek(0)
    return buf.read()
