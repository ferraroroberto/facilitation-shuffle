# Facilitation Shuffle

Streamlit app that randomizes workshop participants into **1-2-4-all** breakout groups for Zoom sessions.

## What it does

Julia (program manager) uses this during ESADE facilitation sessions:

1. **Check presence** — all participants ticked by default; uncheck absentees.
2. **Shuffle Groups** — generates pairs → groups of 4 (round A) → groups of 4 (round B).
3. **Copy-paste** the plain-text room lists into Zoom breakout rooms.
4. **Download** the full summary as `.xlsx` to review all assignments at once.

## Quick start

```bat
launch_app.bat
```

To share with Julia over the internet (Cloudflare tunnel):

```bat
launch_server.bat
```

The tunnel prints a public `https://` URL — share it with Julia. Requires [`cloudflared`](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/) installed.

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

## Project structure

```
app/
  app.py                  Streamlit entry point
  .streamlit/config.toml  Theme (indigo)
src/
  randomize_groups.py     Group algorithm (phases 1–3)
tmp/
  participants.xlsx       Participant roster (gitignored)
launch_app.bat            Local launch
launch_server.bat         Streamlit + Cloudflare tunnel
requirements.txt
```
