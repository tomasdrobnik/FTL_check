# FTL_check

Flight Time Limitation (FDP/rest) violation tracker, rebuilt from the daily
"last 3 days" xlsx exports in `~/Downloads/FTL_violations`.

**Live page:** https://tomasdrobnik.github.io/FTL_check/

## Update with new data

Drop new `.xlsx` exports anywhere under `~/Downloads/FTL_violations` (any
subfolder is fine), then run:

```bash
./rebuild.sh
```

This merges everything, dedupes, regenerates `index.html`, and commits +
pushes so the live page updates. "Handled" marks are stored in each viewer's
browser (localStorage), keyed by a stable per-duty ID, so they survive
rebuilds as long as the pilot/date/route/times don't change.

## How it works

- `build.py` reads every `.xlsx`, keeps the newest version of any
  re-exported duty, computes why each row was flagged, links duty pairs
  flown by two pilots (so marking one handled clears both), and writes
  `data/final_records.json`.
- `template.html` + `data/final_records.json` → `index.html` (self-contained,
  no build step needed to view it — just open the file or the live page).
- `data/seed_handled.json` (optional, not normally present) is a one-time
  seed used only once, the first time this repo replaced an earlier tracker.
