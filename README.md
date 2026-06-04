# Facilitation Shuffle

Streamlit app that randomizes workshop participants into **1-2-4-all** breakout groups for Zoom sessions.

## What it does

The program manager uses this during live workshop facilitation:

1. **Check presence** — all participants ticked by default; uncheck absentees.
2. **Shuffle Groups** — generates pairs → groups of 4 (round A) → groups of 4 (round B).
3. **Copy-paste** the plain-text room lists into Zoom breakout rooms.
4. **Download** the full summary as `.xlsx` to review all assignments at once.

## Quick start

```bat
launch_app.bat
```

To share the app over the internet (Cloudflare tunnel):

```bat
launch_server.bat
```

The tunnel prints a public `https://` URL — share it with participants as needed. Requires [`cloudflared`](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/) installed.

## Setup

```bat
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

Place the participant roster at `tmp/participants.xlsx` (or upload via the sidebar).

## Participant file format

Excel `.xlsx`, row 1 = headers:

| Col | Header | Required? |
|-----|--------|-----------|
| A | `name` | Yes |
| B | `role` | optional |
| C | `company` | optional |
| D | `country` | optional |
| E | `present` | optional (ignored — use the UI checkboxes) |

The file is gitignored (`tmp/*.xlsx`) — participant names never reach the repository.

## Group generation algorithm

### Phase 1 — pairs (`personal_readme_g2`)

A shuffled list of present names is split into atomic units:

| n mod 4 | Layout |
|---------|--------|
| 0 | all pairs |
| 1 | one trio first, then pairs |
| 2 | all pairs |
| 3 | pairs, then one trio at the end |

### Phase 2 — groups of 4 A (`personal_readme_g4`)

Built by merging **whole** phase-1 units — pairs/trios are never split:

- `n ≡ 2 (mod 4), n ≥ 4` → pattern `4+4+…+6` (last block is three pairs, avoiding a leftover pair of 2).
- Otherwise → adjacent units merged two at a time.

**Invariant:** everyone in the same phase-1 unit shares the same phase-2 group.

### Phase 3 — groups of 4 B (`common_enemy_g4`)

Sizes come from a greedy partition of `n` into 4s and 3s (groups are at least 3 when `n ≥ 3`). Many random shuffles are evaluated; the one with the lowest **phase-2 cohort overlap penalty** is kept — minimizing how often former phase-2 tablemates end up in the same phase-3 group.

## CLI usage

`src/randomize_groups.py` ships a self-contained CLI that writes group ids directly back into the workbook (columns F–I) without launching the Streamlit UI.

```powershell
# Generate a sample workbook (creates tmp/participants.xlsx with 15 placeholder rows)
& .\.venv\Scripts\python.exe -m src.randomize_groups --make-fixture

# Run with the default workbook (tmp/participants.xlsx)
& .\.venv\Scripts\python.exe -m src.randomize_groups

# Run with a custom path
& .\.venv\Scripts\python.exe -m src.randomize_groups path/to/roster.xlsx

# Reproducible run — same seed always produces the same groups
& .\.venv\Scripts\python.exe -m src.randomize_groups --seed 42
```

**Arguments:**

| Argument | Default | Description |
|---|---|---|
| `xlsx` (positional) | `tmp/participants.xlsx` | Path to the participant workbook |
| `--seed N` | none (random) | Integer RNG seed for reproducible runs |
| `--make-fixture` | — | Create `tmp/participants.xlsx` with 15 placeholder names if the file is missing, then exit |

**Output columns written (F–I):**

| Column | Header | Content |
|---|---|---|
| F | `personal_readme_g2` | Phase-1 pair group id (1-based) |
| G | `personal_readme_g4` | Phase-2 group-of-4 id (round A) |
| H | `common_enemy_g4` | Phase-3 group-of-4 id (round B) |
| I | `group_repeat` | Count of tablemates shared across both phase-2 and phase-3 groups |

**CLI vs. UI:** the CLI writes ids back into the workbook and exits — useful for scripting, CI fixtures, or offline use. The Streamlit UI instead emits copy-paste Zoom room lists and a downloadable summary xlsx without modifying the original roster file.

## Project structure

```
app/
  app.py                  Streamlit entry point
  .streamlit/config.toml  Theme (indigo)
src/
  randomize_groups.py     Group algorithm + CLI entry point (phases 1–3)
tmp/
  participants.xlsx       Participant roster (gitignored)
launch_app.bat            Local launch
launch_server.bat         Streamlit + Cloudflare tunnel
requirements.txt
```

## License

[MIT](LICENSE) — Copyright (c) 2026 Roberto Ferraro.
