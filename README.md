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

## Run the spike

```powershell
.venv\Scripts\python -m spike.fixtures     # generate the synthetic test corpus into fixtures/
.venv\Scripts\python -m spike.run_all      # run U1-U4 experiments, print the findings memo
```

Outputs land in `out/`. Files written there (round-trip and enforcement outputs) must ALSO be opened manually in PowerPoint desktop and PowerPoint on the web; "opens without repair prompt" is a spike success criterion that cannot be automated here.

## Notes

- The synthetic corpus covers the mechanics. Before the spike is signed off, re-run against 5-10 real (anonymized) Prezlab decks, including one 150-plus-slide deck and real Arabic client decks (PRD 11.1).
- Engine decision (python-pptx vs Aspose.Slides) is made at spike close per PRD 11.1.

123
