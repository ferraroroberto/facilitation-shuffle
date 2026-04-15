# Facilitation Shuffle — Claude Code Instructions

## Project
Streamlit app for 1-2-4-all group randomization. Julia Garcia Marina (program manager, ESADE) is the end-user; Roberto Ferraro is the developer/facilitator. See `AGENTS.md` for full context, codebase map, and algorithm docs.

## Key rules

### Streamlit
- Use `width="stretch"` / `width="content"` — **never** `use_container_width` (deprecated after 2025-12-31).
- Widget `key=` values must be unique across all tabs. Use `key=f"ta_{tab_id}"` patterns, not label-based keys.
- `@st.cache_data` on the default xlsx load; skip cache for uploaded bytes.

### Participant names
- Always sort with diacritic-aware key: `unicodedata.normalize("NFD", s.lower())` stripping category "Mn" chars.
- Checkbox grid is **column-major** (A-Z reads top-to-bottom per column, not left-to-right across columns). Formula: `col_size = (n_total + 2) // 3`.

### Privacy
- `tmp/*.xlsx` is gitignored — participant names must never reach the repository.
- Never hardcode participant file paths in `src/` or `app/`.

### Code style
- Don't add comments, docstrings, or type annotations to code you didn't change.
- Don't add error handling for scenarios that can't happen.
- Business logic stays in `src/`; UI stays in `app/app.py`.
