# PowerPoint Formatting QC Tool

Internal Prezlab tool that audits (and later fixes) formatting consistency on `.pptx` decks against per-client profiles. See [PRD.md](PRD.md) for the full specification.

## Current phase: Technical Spike (Section 11.1 of the PRD)

The spike retires five unknowns (U1-U5) before the v1 timeline is committed:

| # | Unknown | Where |
|---|---------|-------|
| U1 | Master enforcement without corruption | `spike/u1_master.py` |
| U2 | Safe write fidelity (round-trip preservation) | `spike/u2_roundtrip.py` |
| U3 | 200-slide processing time | `spike/u3_perf.py` |
| U4 | Inheritance/color resolution + Arabic detection | `spike/resolver.py`, `spike/color_resolver.py`, `spike/arabic.py` |
| U5 | Microsoft Graph renderer viability + data-residency ruling | External: needs M365 tenant access and an Operations/IT-security decision |

## Setup

```powershell
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## Run the tool

```powershell
.venv\Scripts\python -m qc.web        # http://127.0.0.1:8000
```

Four destinations, in the header of every page.

**Prepare a deck** is the main flow and it is three steps. Drop a finished
master to save it as a profile; drop a messy client deck against that profile;
then **choose the layouts** for the slides the file could not place with
confidence, and only then is the master applied and the rebuilt deck audited.
Nothing is rewritten until that second press.

**Audit only** reads a deck against a profile and changes nothing.

**Profiles** lists every saved profile and edits one: type, palette, margins and
tolerances, in inches and points rather than EMU. Use it to change one number
instead of re-reading a master.

**Team** is the pilot roster; saving a profile needs a lead or admin.

Applying a master is pure code - it asks no model and produces the same file
from the same inputs every time. The passes that DO use a model are the visual
ones: what a designer would adjust on the rebuilt slides, what the things on a
slide are, the triage questions, the ask box, and proposing a layout for a gap.
All of them go to Gemini (`GEMINI_API_KEY` in a gitignored `.env`), all of them
are optional, and `QC_AI=0` turns the lot off at the source.

## Run the spike

```powershell
.venv\Scripts\python -m spike.fixtures     # generate the synthetic test corpus into fixtures/
.venv\Scripts\python -m spike.run_all      # run U1-U4 experiments, print the findings memo
```

Outputs land in `out/`. Files written there (round-trip and enforcement outputs) must ALSO be opened manually in PowerPoint desktop and PowerPoint on the web; "opens without repair prompt" is a spike success criterion that cannot be automated here.

## Notes

- The synthetic corpus covers the mechanics. Before the spike is signed off, re-run against 5-10 real (anonymized) Prezlab decks, including one 150-plus-slide deck and real Arabic client decks (PRD 11.1).
- Engine decision (python-pptx vs Aspose.Slides) is made at spike close per PRD 11.1.
