# Facilitation Shuffle (Streamlit)

## Context
**What**: Streamlit app that randomizes workshop participants into 1-2-4-all breakout groups (pairs → groups of 4A → groups of 4B) for Zoom sessions.
**Why**: The program manager needs to check who is present, shuffle groups, and copy-paste room assignments into Zoom breakout rooms — without relying on the maintainer during live sessions.
**Stack**: Python 3.x, Streamlit, openpyxl, pandas. Windows-first (`launch_app.bat` / `launch_server.bat` + Cloudflare tunnel).

## Codebase Map
- `app/app.py` — Streamlit entry point; all UI logic (sidebar, checklist, shuffle, tabs, summary).
- `app/.streamlit/config.toml` — theme (indigo primary, white background).
- `src/randomize_groups.py` — group algorithm: Phase 1 (pairs/trios), Phase 2 (merge pairs into groups of 4), Phase 3 (optimized re-split minimizing phase-2 overlap).
- `tmp/participants.xlsx` — participant roster (gitignored — contains personal names).
- `requirements.txt` — `streamlit>=1.40.0`, `openpyxl>=3.1.0`.
- `launch_app.bat` — local launch shortcut.
- `launch_server.bat` — Streamlit + Cloudflare tunnel for public HTTPS sharing.

## Commands
```bash
# Run locally
launch_app.bat
# or
.venv\Scripts\python.exe -m streamlit run app\app.py

# Run with public Cloudflare tunnel (share URL with participants)
launch_server.bat
```

## Participant File Format
Excel `.xlsx` with at least column A. Row 1 = headers.

| Col | Header | Required? |
|-----|--------|-----------|
| A | `name` | Yes |
| B | `role` | optional |
| C | `company` | optional |
| D | `country` | optional |
| E | `present` | optional (ignored — UI checkboxes take precedence) |

Default path: `tmp/participants.xlsx`. A different file can be uploaded via the sidebar.

## Group Generation Algorithm (src/randomize_groups.py)

### Phase 1 — pairs
Shuffled present names split into atomic units:
- `n ≡ 0 (mod 4)` — all pairs
- `n ≡ 1 (mod 4)` — one trio first, then pairs
- `n ≡ 2 (mod 4)` — all pairs
- `n ≡ 3 (mod 4)` — pairs, then one trio at the end

### Phase 2 — groups of 4 (A)
Built by merging **whole** phase-1 units (pairs/trios never split):
- `n ≡ 2 (mod 4), n ≥ 4` — pattern `4+4+…+6` (last block is three pairs to avoid a leftover pair of 2).
- Otherwise — adjacent units merged two at a time.

### Phase 3 — groups of 4 (B)
Many random shuffles of names; each sliced into sizes from a greedy 4s-and-3s partition. The layout with the lowest phase-2 cohort overlap penalty is kept.

## Standards
- **Imports order**: stdlib → third-party → local (`src.*`).
- **Naming**: `snake_case` files/functions, `UPPER_CASE` constants.
- **Sort**: always use diacritic-aware sort for participant names (`unicodedata.normalize("NFD", ...)` stripping category "Mn" chars) so accented names sort correctly.

## Streamlit Conventions
- `app/app.py` handles all UI; business logic stays in `src/`.
- `st.session_state` for groups and present set (persist across reruns).
- `@st.cache_data` for the default xlsx load (`_load_from_path`).
- Use `width="stretch"` / `width="content"` — **not** `use_container_width` (deprecated after 2025-12-31).
- Widget keys must be unique across tabs — use `key=f"ta_{tab_id}"` not `key=f"ta_{label}"`.

## Safety
- Never modify `.venv/`.
- Never commit `tmp/*.xlsx` — contains participant personal data (enforced by `.gitignore`).
- The Cloudflare tunnel URL is ephemeral — a new URL is generated each `launch_server.bat` run.
