"""
Facilitation Shuffle — Streamlit app for 1-2-4-all group generation.

Use this during sessions to:
  1. Check who is present (all ticked by default).
  2. Click Shuffle Groups to generate pairs → groups of 4 (round A) → groups of 4 (round B).
  3. Copy-paste the plain-text output into Zoom breakout rooms.
  4. Download the full summary as an xlsx to review all assignments at once.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
DEFAULT_XLSX = ROOT / "tmp" / "participants.xlsx"

# Make src/ importable so we can reuse the data logic.
sys.path.insert(0, str(ROOT))
from src.randomize_groups import build_groups  # noqa: E402  (after sys.path patch)
from src.roster_io import (  # noqa: E402
    load_participants_from_bytes,
    load_participants_from_path,
)
from src.summary import build_summary_df, df_to_xlsx_bytes  # noqa: E402


# ---------------------------------------------------------------------------
# Helpers — caching wrapper (UI concern: keeps @st.cache_data out of src/)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def _load_from_path(path: str) -> tuple[list[str], pd.DataFrame]:
    """Cached version for the default file path."""
    return load_participants_from_path(path)


# ---------------------------------------------------------------------------
# Helpers — group generation
# ---------------------------------------------------------------------------

def generate_groups(present: list[str]) -> dict[str, list[list[str]]]:
    """Run all three phases and return group assignments."""
    return build_groups(present)


def zoom_text(groups: list[list[str]], label: str = "Room") -> str:
    """Plain-text list ready to paste into Zoom breakout room names."""
    return "\n".join(f"{label} {i}: {', '.join(g)}" for i, g in enumerate(groups, 1))


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Facilitation Shuffle",
    page_icon="🎲",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("🎲 Facilitation Shuffle")
    st.caption("1-2-4-all group randomizer")
    st.divider()

    uploaded = st.file_uploader(
        "Upload participants list",
        type=["xlsx"],
        help="Leave empty to use **tmp/participants.xlsx** in the project folder.",
        key="upload_roster",
    )

    with st.expander("📋 Required file format"):
        st.markdown(
            """
**Excel (.xlsx) with at least column A:**

| Col | Header | Required? |
|-----|--------|-----------|
| A | `name` | ✅ Yes |
| B | `role` | optional |
| C | `company` | optional |
| D | `country` | optional |
| E | `present` | optional |

- **Row 1** must be a header row (any text).
- **Column A** must contain participant names — one per row, no blank names.
- **Column E** (`present`) is ignored by the app — use the checkboxes instead.
- Any pre-existing group columns (F–I) are ignored on upload.
"""
        )

    st.divider()
    st.caption("Generates: pairs → groups of 4 (A) → groups of 4 (B)")

# ---------------------------------------------------------------------------
# Load participant data
# ---------------------------------------------------------------------------
if uploaded is not None:
    raw_bytes = uploaded.read()
    all_names, roster_df = load_participants_from_bytes(raw_bytes)
elif DEFAULT_XLSX.exists():
    all_names, roster_df = _load_from_path(str(DEFAULT_XLSX))
else:
    all_names, roster_df = [], pd.DataFrame()

# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------
st.title("🎲 Facilitation Shuffle")

if not all_names:
    st.warning(
        "No participants found. "
        "Upload an .xlsx file using the sidebar, "
        "or place your file at `tmp/participants.xlsx`."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Participant checklist
# ---------------------------------------------------------------------------
n_total = len(all_names)
st.subheader(f"Participants — {n_total} total")

col_sel, col_desel = st.columns([1, 1], gap="small")
with col_sel:
    if st.button("✅ Select all", width="stretch", key="btn_select_all"):
        for i in range(n_total):
            st.session_state[f"p_{i}"] = True
with col_desel:
    if st.button("☐ Deselect all", width="stretch", key="btn_deselect_all"):
        for i in range(n_total):
            st.session_state[f"p_{i}"] = False

# Initialise checkbox state on first load (all present by default).
# Must happen before rendering so that session_state owns the value —
# avoids the "created with default value but also set via Session State API" warning.
for i in range(n_total):
    if f"p_{i}" not in st.session_state:
        st.session_state[f"p_{i}"] = True

# Three-column checkbox grid — fill each column top-to-bottom so names read A-Z down each column
cols = st.columns(3)
checked: dict[str, bool] = {}
col_size = (n_total + 2) // 3  # ceiling division → equal chunks
for col_idx, col in enumerate(cols):
    with col:
        for i in range(col_idx * col_size, min((col_idx + 1) * col_size, n_total)):
            name = all_names[i]
            checked[name] = st.checkbox(name, key=f"p_{i}")

present = [name for name, ok in checked.items() if ok]
absent_count = n_total - len(present)
status_parts = [f"**{len(present)}** present"]
if absent_count:
    status_parts.append(f"**{absent_count}** absent")
st.caption(" · ".join(status_parts))

st.divider()

# ---------------------------------------------------------------------------
# Shuffle button
# ---------------------------------------------------------------------------
if st.button("🔀 Shuffle Groups", type="primary", width="stretch", key="btn_shuffle"):
    if len(present) < 2:
        st.error("Need at least 2 present participants to form groups.")
    else:
        _new_groups = generate_groups(present)
        st.session_state["groups"] = _new_groups
        st.session_state["present_set"] = set(present)
        st.session_state["n_present"] = len(present)
        st.session_state["ta_pairs"] = zoom_text(_new_groups["pairs"], "Room")
        st.session_state["ta_g4a"]   = zoom_text(_new_groups["g4a"],   "Room")
        st.session_state["ta_g4b"]   = zoom_text(_new_groups["g4b"],   "Room")

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
if "groups" not in st.session_state:
    st.stop()

groups = st.session_state["groups"]
n_present = st.session_state["n_present"]
present_set: set[str] = st.session_state["present_set"]

st.subheader(f"Groups for {n_present} participants")

tab1, tab2, tab3, tab4 = st.tabs([
    "👥  Pairs (Round 1)",
    "👥👥  Groups of 4 — A (Round 2)",
    "👥👥  Groups of 4 — B (Round 3)",
    "📊  Summary",
])


def _render_round_tab(tab_id: str, label: str, phase_groups: list[list[str]], caption: str) -> None:
    st.caption(caption)

    # value= is intentionally omitted: the shuffle handler always writes
    # st.session_state["ta_<id>"] before this widget is rendered, so
    # session_state is the sole owner of the text-area value.  Adding value=
    # here would trigger Streamlit's "created with a default value but also set
    # via Session State API" warning on every shuffle.
    st.text_area(
        "Copy and paste into Zoom:",
        height=max(120, len(phase_groups) * 30),
        key=f"ta_{tab_id}",
    )

    n_cols = min(len(phase_groups), 4)
    cards = st.columns(n_cols) if n_cols > 1 else [st.container()]
    for i, group in enumerate(phase_groups):
        with cards[i % n_cols]:
            members = "\n".join(f"- {p}" for p in group)
            st.markdown(f"**{label} {i + 1}**\n{members}")


with tab1:
    _render_round_tab(
        "pairs", "Room", groups["pairs"],
        "Use for the **pair reflection** phase — 2 people per room.",
    )

with tab2:
    _render_round_tab(
        "g4a", "Room", groups["g4a"],
        "Use for the **first group-of-4** phase (Personal Readme). "
        "Pairs from Round 1 are kept together.",
    )

with tab3:
    _render_round_tab(
        "g4b", "Room", groups["g4b"],
        "Use for the **second group-of-4** phase (Common Enemy). "
        "Optimised to maximise mixing across Round 2 groups.",
    )

with tab4:
    st.caption(
        "Full participant roster with all group assignments. "
        "Absent participants are shown with empty group columns."
    )

    summary_df = build_summary_df(roster_df, present_set, groups)

    st.dataframe(
        summary_df,
        width="stretch",
        hide_index=True,
        column_config={
            "present":           st.column_config.NumberColumn("Present", width="small"),
            "pair (round 1)":    st.column_config.NumberColumn("Pair (R1)", width="small"),
            "group_4A (round 2)": st.column_config.NumberColumn("Group 4A (R2)", width="small"),
            "group_4B (round 3)": st.column_config.NumberColumn("Group 4B (R3)", width="small"),
        },
    )

    xlsx_bytes = df_to_xlsx_bytes(summary_df)
    st.download_button(
        label="⬇️  Download as xlsx",
        data=xlsx_bytes,
        file_name="facilitation_groups.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        key="btn_download_xlsx",
    )
