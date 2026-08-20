# How this project is tested

Four layers, from fast to manual. Run layers 1-3 before any merge; layer 4
before any release milestone.

## 1. Unit and module tests (seconds, run constantly)

```powershell
.venv\Scripts\python -m pytest
```

48 tests. Each audit module has its own test file that builds tiny decks
with planted violations and asserts the exact issue_type, severity, and
Arabic-flag behavior, plus a clean control deck that must produce zero
findings (the false-positive guard). `tests/test_engine.py` covers the
pipeline: preflight, module selection, manifest schema, Arabic index.

## 2. Spike experiments (the technical-risk regression rig)

```powershell
.venv\Scripts\python -m spike.fixtures    # regenerate corpus if needed
.venv\Scripts\python -m spike.run_all     # U1-U4 checks, writes out/spike-results.json
```

Keeps the load-bearing claims true as code evolves: inheritance resolution,
color math, Arabic detection, round-trip byte preservation, 200-slide speed.

## 3. PowerPoint desktop validation (semi-automated, ~1 min)

```powershell
.\tools\Validate-InPowerPoint.ps1
```

Opens every fixture and generated output in PowerPoint via COM (invisible,
read-only) and fails on any file PowerPoint cannot open cleanly. This is the
"no repair prompts" criterion from PRD 11.1. Requires desktop PowerPoint;
never run server-side.

## 4. Manual / external checks (before milestones)

- Open a sample of `out/*.pptx` in PowerPoint **on the web** (M365): the web
  renderer is stricter than desktop in different ways.
- Real-deck corpus: run `python -m qc.cli <deck> --profile prezlab_en --verbose`
  over 5-10 real (anonymized) Prezlab decks, including one 150-plus-slide
  deck, Arabic client decks, and SmartArt-heavy decks (python-pptx cannot
  generate SmartArt, so only real decks exercise that preflight path).
  Review findings with a senior designer: every disputed flag feeds the
  per-fix-type false-positive metric that gates any future auto-apply.
- Color-math hand-check (U4 caveat): compare a handful of resolved
  tint/lumMod colors against PowerPoint's rendered values.
- U5 (external): Microsoft Graph PDF conversion test + Operations/IT-security
  data-residency ruling. Blocked on tenant access, not on code.

## Testing the fix flow (v1.5 core)

In the web UI, audit a deck, tick fixes in the Apply column (deterministic
fixes are pre-ticked, suggestions need an explicit tick, AR rows are never
fixable), click Apply, then download the cleaned `.pptx` and open it in
PowerPoint. The re-audit path: run the cleaned file back through an audit
and confirm the applied issue types are gone. Known pilot limitation: the
PDF report renders Arabic as unshaped glyphs (reportlab has no bidi); the
CSV and manifest JSON are the lossless channels for Arabic content.

## Auditing a deck from the command line

```powershell
.venv\Scripts\python -m qc.cli path\to\deck.pptx --profile prezlab_bilingual --verbose --json manifest.json
```

`--modules font,color_palette` runs a targeted subset. Profiles live in
`qc/profiles/`; pass a name or a path to a custom profile JSON.
