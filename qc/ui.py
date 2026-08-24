"""Page rendering for the local pilot UI, in the Prezlab brand system.

Source of truth: Prezlab Brand Guidelines v1.0 (April 2026). Corporate
palette as the working surface (Dark Teal anchors, Light Teal and Sand
breathe, Slate for functional details, Off-White as the neutral ground);
light mode only, per the guidelines' own philosophy for working contexts.
Typography: Figtree for everything functional (bundled locally as a
variable font; no runtime egress per the PRD), PP Editorial New for the
single display moment when locally installed, Georgia as its fallback.

Semantic status colors are mapped onto APPROVED brand pairings (guidelines
pp. 25-26): error = Burgundy with Off-White text, warning = Orange with
Dark Teal text, info = Slate with Off-White text, Arabic review = Lavender
with Dark Teal text, success = Light Teal with Dark Teal text.
"""

import html
from pathlib import Path

from .config import DEMO_BANNER

MODULE_LABELS = {
    "master_slide": "Master slide",
    "font": "Fonts",
    "margin_alignment": "Margins & alignment",
    "color_palette": "Color palette",
    "shape_size": "Shape sizes",
    "header_footer": "Headers & footers",
    "preflight": "Preflight",
}

# Human labels for the issue-type summary band. Falls back to the prettified
# last segment of the issue_type for anything not listed.
ISSUE_LABELS = {
    "margin_alignment.edge_misaligned": "Edge misaligned",
    "margin_alignment.uneven_spacing": "Uneven spacing",
    "margin_alignment.outside_safe_zone": "Outside safe zone",
    "margin_alignment.heading_past_margin": "Heading past the margin (ask the client)",
    "margin_alignment.body_band_intrusion": "Content in the reserved header band",
    "margin_alignment.text_overlap": "Text overlapping text",
    "margin_alignment.squeezed_text": "Squeezed text box",
    "margin_alignment.text_anchor_mismatch": "Text anchored differently in an aligned row",
    "color_palette.off_palette_rgb": "Off-palette color",
    "font.family_out_of_set": "Font not in set",
    "font.size_off_role": "Font size off",
    "font.title_autofit_shrunk": "Title auto-shrunk",
    "font.theme_ref_disallowed": "Disallowed theme font",
    "font.mixed_weight": "Mixed bold weight",
    "font.cs_typeface_missing": "Arabic font missing",
    "shape_size.size_mismatch": "Shape size mismatch",
    "master_slide.foreign_master": "Foreign master",
    "master_slide.layout_outlier": "Layout outlier",
    "master_slide.placeholder_geometry_off": "Placeholder moved",
    "master_slide.no_usable_master": "Broken master chain",
    "header_footer.text_mismatch": "Footer text wrong",
    "header_footer.missing": "Header/footer missing",
    "preflight.unmodifiable_content": "Preserved as-is",
    "margin_alignment.body_band_intrusion": "Body in the header's clear strip",
    "margin_alignment.body_below_band": "Body starts below its guide",
    "margin_alignment.overlap_check_capped": "Slide too crowded to check fully",
    "margin_alignment.space_edge_misaligned": "Off the presentation space edge",
}


def issue_label(issue_type: str) -> str:
    if issue_type in ISSUE_LABELS:
        return ISSUE_LABELS[issue_type]
    tail = issue_type.split(".")[-1].replace("_", " ")
    return tail[:1].upper() + tail[1:]


def esc(value) -> str:
    return html.escape(str(value), quote=True)


_CSS = """
@font-face {
  font-family: 'Figtree';
  src: url('/static/fonts/figtree-var.woff2') format('woff2');
  font-weight: 300 900; font-style: normal; font-display: swap;
}
:root {
  --teal: #002528; --teal-light: #b6edf3; --sand: #e3d8cc;
  --slate: #62848c; --offwhite: #f4f6f6;
  --burgundy: #40182d; --orange: #ff7c4a; --lavender: #cabefc;
  --slate-text: #4a666e;
  --line: rgba(0, 37, 40, 0.12); --line-soft: rgba(0, 37, 40, 0.07);
  --hover: rgba(182, 237, 243, 0.18);
}
* { box-sizing: border-box; }
input, button { font-family: inherit; }
body { background: var(--offwhite); color: var(--teal); margin: 0;
  font: 500 15px/1.5 Figtree, 'Segoe UI', system-ui, sans-serif; }
.wrap { max-width: 72rem; margin: 0 auto; padding: 1.6rem 1.5rem 4rem; }
a { color: var(--teal); text-decoration: underline; text-underline-offset: 3px;
  text-decoration-color: var(--slate); }
a:hover { text-decoration-color: var(--teal); }

.brandrow { display: flex; align-items: baseline; gap: 0.9rem; margin-bottom: 1.6rem; }
.wordmark { font-weight: 900; font-size: 1.25rem; letter-spacing: -0.02em;
  color: var(--teal); }
.kicker { color: var(--slate); font-size: 11px; font-weight: 600;
  letter-spacing: 0.16em; text-transform: uppercase; }
h1 { font-family: 'PP Editorial New', Georgia, 'Times New Roman', serif;
  font-weight: 400; font-size: 2.3rem; letter-spacing: -0.01em;
  margin: 0 0 0.35rem; text-wrap: balance; }
h1.file { font-family: Figtree, 'Segoe UI', sans-serif; font-weight: 700;
  letter-spacing: -0.02em; font-size: 2rem; overflow-wrap: anywhere; }
.sub { color: var(--slate-text); margin: 0 0 1.6rem; font-weight: 500; }
.sub b { color: var(--teal); font-weight: 700; }

.card { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 1.4rem 1.6rem; margin: 1rem 0; }
.banner { border-radius: 10px; padding: 0.8rem 1.1rem; margin: 0.9rem 0;
  font-weight: 600; }
.banner.ok { background: var(--teal-light); color: var(--teal);
  border-left: 4px solid var(--teal); }
.banner.warn { background: var(--sand); color: var(--burgundy); }
.demostrip { background: var(--burgundy); color: var(--offwhite);
  text-align: center; font-size: 0.8rem; font-weight: 700;
  letter-spacing: 0.04em; padding: 0.4rem 1rem; }

/* buttons */
.btn { display: inline-flex; align-items: center; gap: 0.5rem; border: 0;
  border-radius: 10px; padding: 0.55rem 1.3rem; font-weight: 700;
  font-size: 0.92rem; cursor: pointer; text-decoration: none; }
.btn.primary { background: var(--teal); color: var(--offwhite); }
.btn.primary:hover { background: #0b3d42; }
.btn.primary:disabled { background: var(--offwhite); color: var(--slate-text);
  border: 1px solid var(--line); cursor: not-allowed; }
.btn.success { background: var(--teal-light); color: var(--teal); }
.btn.success:hover { background: #a3e4ec; }
.btn.ghost { background: none; border: 1px solid var(--line); color: var(--teal); }
.btn.ghost:hover { background: var(--hover); }

/* sticky action bar (report) */
.actionbar { position: sticky; top: 0; z-index: 10; background: #fff;
  border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 0.65rem 1rem; margin: 0.4rem 0 1rem; display: flex; gap: 0.9rem;
  align-items: center; flex-wrap: wrap; box-shadow: 0 2px 12px rgba(0, 37, 40, 0.05); }
.actionbar .grow { flex: 1; }
.minicounts { display: flex; gap: 0.9rem; font-size: 0.85rem; color: var(--slate-text);
  font-weight: 600; font-variant-numeric: tabular-nums; }
.minicounts .dot { display: inline-block; width: 9px; height: 9px;
  border-radius: 50%; margin-right: 0.3rem; }
.exports { font-size: 0.85rem; color: var(--slate-text); }

/* upload */
.drop { border: 1.5px dashed var(--slate); border-radius: 12px; padding: 3rem 1.5rem;
  text-align: center; cursor: pointer; background: #fff;
  transition: border-color 0.15s, background 0.15s; }
.drop.armed, .drop:hover { border-color: var(--teal); background: var(--hover); }
.drop strong { font-size: 1.05rem; font-weight: 700; }
.drop .hint { color: var(--slate); font-size: 0.85rem; margin-top: 0.5rem; }
.drop .file { color: var(--teal); font-weight: 700; margin-top: 0.6rem; }
.config { display: flex; gap: 1.1rem; flex-wrap: wrap; margin-top: 1.1rem;
  align-items: flex-start; }
.config > fieldset { flex: 1 1 20rem; border: 1px solid var(--line-soft);
  border-radius: 12px; padding: 0.9rem 1.1rem 1.1rem; margin: 0; background: #fff; }
legend { color: var(--slate); font-size: 11px; font-weight: 600;
  letter-spacing: 0.14em; text-transform: uppercase; padding: 0 0.4rem; }
.radio-card { display: flex; gap: 0.6rem; align-items: baseline;
  padding: 0.6rem 0.75rem; border: 1px solid var(--line-soft); border-radius: 10px;
  margin: 0.45rem 0; cursor: pointer; }
.radio-card:has(input:checked) { border-color: var(--teal); background: var(--hover); }
.radio-card b { font-weight: 700; }
.radio-card small { color: var(--slate); display: block; }
input[type="radio"], input[type="checkbox"] { accent-color: var(--teal); }
select, input[type="text"], input[type="password"], input[type="number"] {
  border: 1px solid var(--line); border-radius: 10px; padding: 0.42rem 0.85rem;
  background: #fff; color: var(--teal); font-size: 0.88rem; font-weight: 600; }
select { appearance: none; -webkit-appearance: none; padding-right: 2rem;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%23002528' stroke-width='1.6' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat: no-repeat; background-position: right 0.8rem center; cursor: pointer; }
select:focus, input:focus { outline: 2px solid var(--teal-light); border-color: var(--teal); }
.chips { display: flex; flex-wrap: wrap; gap: 0.45rem; margin-top: 0.45rem; }
.chip { display: inline-flex; align-items: center; gap: 0.4rem;
  border: 1px solid var(--line); border-radius: 999px; padding: 0.3rem 0.85rem;
  cursor: pointer; font-size: 0.85rem; font-weight: 600; background: #fff; }
.chip:has(input:checked) { border-color: var(--teal); background: var(--teal-light); }
.actions { margin-top: 1.4rem; }
.note { color: var(--slate-text); font-size: 0.8rem; }

/* KPI band */
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(10rem, 1fr));
  gap: 0.8rem; margin: 1rem 0; }
.kpi { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 0.85rem 1rem 0.9rem; border-left: 4px solid var(--line); }
.kpi .n { font-size: 1.8rem; font-weight: 700; line-height: 1.1;
  font-variant-numeric: tabular-nums; }
.kpi .l { color: var(--slate-text); font-size: 0.75rem; letter-spacing: 0.08em;
  text-transform: uppercase; font-weight: 600; margin-top: 0.15rem; }
.kpi.err { border-left-color: var(--burgundy); }
.kpi.warn { border-left-color: var(--orange); }
.kpi.info { border-left-color: var(--slate); }
.kpi.ar { border-left-color: var(--lavender); }

/* filters + search */
.controls { display: flex; gap: 0.45rem; flex-wrap: wrap; align-items: center;
  margin: 1rem 0 0.8rem; }
.fchip { border: 1px solid var(--line); background: #fff; color: var(--teal);
  border-radius: 999px; padding: 0.3rem 0.95rem; font-size: 0.85rem;
  font-weight: 600; cursor: pointer; }
.fchip[aria-pressed="true"] { background: var(--teal); color: var(--offwhite);
  border-color: var(--teal); }
.fchip:disabled { opacity: 0.4; cursor: default; }
.search { flex: 1; min-width: 12rem; border: 1px solid var(--line);
  border-radius: 999px; padding: 0.38rem 1rem; font-size: 0.88rem;
  background: #fff; color: var(--teal); }
.search::placeholder { color: var(--slate); }
.vdiv { width: 1px; height: 1.5rem; background: var(--line); }
.shown { color: var(--slate-text); font-size: 0.82rem; font-weight: 600;
  font-variant-numeric: tabular-nums; }

/* slide groups */
details.grp { background: #fff; border: 1px solid var(--line-soft);
  border-radius: 12px; margin: 0.55rem 0; overflow: hidden; }
details.grp > summary { list-style: none; cursor: pointer; display: flex;
  align-items: center; gap: 0.8rem; padding: 0.7rem 1.1rem; font-weight: 700;
  font-variant-numeric: tabular-nums; }
details.grp > summary::-webkit-details-marker { display: none; }
details.grp > summary::before { content: '+'; color: var(--slate);
  font-weight: 700; width: 1rem; }
details.grp[open] > summary::before { content: '\\2212'; }
details.grp > summary:hover { background: var(--hover); }
.grpcounts { color: var(--slate-text); font-size: 0.82rem; font-weight: 600; }
table { border-collapse: collapse; width: 100%; font-size: 0.88rem; }
td { padding: 0.55rem 0.8rem; border-top: 1px solid var(--line-soft);
  vertical-align: top; }
tr:hover td { background: var(--hover); }
td.fix { width: 2.6rem; text-align: center; }
td.fix input { width: 17px; height: 17px; cursor: pointer; }
td.sev { width: 7.5rem; white-space: nowrap; }
td.code { font-family: 'Cascadia Mono', Consolas, monospace; font-size: 0.78rem;
  white-space: nowrap; color: var(--teal); width: 19rem; }
td.msg { color: #1d3f44; }

/* status pills: approved brand pairings only */
.pill { display: inline-block; border-radius: 999px; padding: 0.14rem 0.65rem;
  font-size: 0.72rem; font-weight: 700; letter-spacing: 0.04em; }
.pill.error { background: var(--burgundy); color: var(--offwhite); }
.pill.warning { background: var(--orange); color: var(--teal); }
.pill.info { background: var(--slate); color: var(--offwhite); }
.pill.ar { background: var(--lavender); color: var(--teal); margin-left: 0.35rem; }
.pill.changed { background: var(--teal-light); color: var(--teal); }

.clean { text-align: center; padding: 3.4rem 1rem; }
.clean .mark { width: 64px; height: 64px; border-radius: 50%; margin: 0 auto 1rem;
  background: var(--teal-light); color: var(--teal); font-size: 2rem;
  line-height: 64px; font-weight: 700; }
.clean h2 { font-family: 'PP Editorial New', Georgia, serif; font-weight: 400;
  margin: 0 0 0.3rem; }
.clean p { color: var(--slate); }

/* before/after diff */
.diffslide { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 1rem 1.2rem 1.2rem; margin: 1rem 0; }
.diffslide h3 { margin: 0 0 0.2rem; font-size: 1.05rem; }
.difflabels { color: var(--slate-text); font-size: 0.8rem; margin-bottom: 0.8rem;
  font-family: 'Cascadia Mono', Consolas, monospace; }
.panes { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; }
@media (max-width: 60rem) { .panes { grid-template-columns: 1fr; } }
.pane .tag { font-size: 11px; font-weight: 600; letter-spacing: 0.14em;
  text-transform: uppercase; color: var(--slate-text); margin-bottom: 0.35rem; }
.shot { position: relative; border: 1px solid var(--line); border-radius: 8px;
  overflow: hidden; }
.shot img { display: block; width: 100%; height: auto; }
.hl { position: absolute; border-radius: 3px; pointer-events: none; }
.hl.before { border: 2px dashed var(--orange); background: rgba(255, 124, 74, 0.12); }
.hl.after { border: 2px solid #0e7c66; background: rgba(14, 124, 102, 0.10); }
.legend { display: flex; gap: 1.4rem; color: var(--slate-text); font-size: 0.85rem;
  margin: 0.4rem 0 0.6rem; align-items: center; }
.legend .swatch { display: inline-block; width: 14px; height: 14px;
  border-radius: 3px; margin-right: 0.4rem; vertical-align: -2px; }
/* triage */
.tri { display: inline-flex; gap: 0.25rem; }
.tri button { border: 1px solid var(--line); background: #fff; color: var(--slate-text);
  border-radius: 7px; width: 1.7rem; height: 1.7rem; cursor: pointer;
  font-size: 0.85rem; line-height: 1; padding: 0; }
.tri button:hover { background: var(--hover); }
tr.t-confirmed .tri .ok { background: var(--teal-light); color: var(--teal);
  border-color: var(--teal); }
tr.t-false_positive .tri .fp { background: var(--sand); color: var(--burgundy);
  border-color: var(--burgundy); }
tr.t-false_positive td { opacity: 0.5; }
tr.t-false_positive td.trit { opacity: 1; }

/* slide preview panel: sticky beside the findings so the slide stays in
   view while the list scrolls */
.prevbtn { margin-left: auto; }
.prev { padding: 0 1.1rem 1rem; }
.prev .shot { max-width: 56rem; }
.prev .loading { color: var(--slate-text); font-size: 0.85rem; padding: 0.6rem 0; }
.grp .split { display: flex; gap: 1.1rem; align-items: flex-start;
  padding: 0 1.1rem 1rem; }
.grp .split .prev { flex: 0 0 46%; min-width: 0; padding: 0;
  position: sticky; top: 0.75rem; }
.grp .split table { flex: 1; min-width: 0; }
@media (max-width: 980px) {
  .grp .split { display: block; }
  .grp .split .prev { position: static; margin-bottom: 0.8rem; }
}
.hl { cursor: pointer; }
.hl.hlhover { outline: 3px solid var(--teal); outline-offset: 2px; z-index: 2; }
.hl.error { border: 2px solid var(--burgundy); background: rgba(64, 24, 45, 0.10); }
.hl.warning { border: 2px solid var(--orange); background: rgba(255, 124, 74, 0.12); }
.hl.info { border: 2px solid var(--slate); background: rgba(98, 132, 140, 0.10); }
.hl.arflag { border-style: dashed; }

/* numbered pins: badge N in the findings list = box N on the preview */
.pinno { display: inline-flex; align-items: center; justify-content: center;
  min-width: 1.2rem; height: 1.2rem; padding: 0 0.2rem; border: 0;
  border-radius: 50%; color: #fff; background: var(--slate);
  font: 700 0.68rem/1 "Figtree", sans-serif; cursor: pointer; }
.pinno.error { background: var(--burgundy); }
.pinno.warning { background: var(--orange); }
.pinno.whole { background: transparent; color: var(--slate-text);
  border: 1px solid var(--line); border-radius: 8px; cursor: default; }
td.pinc { width: 2rem; text-align: center; }
.hl .pinno { position: absolute; top: -0.55rem; left: -0.55rem;
  box-shadow: 0 0 0 2px #fff; pointer-events: none; }
.pinhot { animation: pinhot 1.6s ease-out; }
@keyframes pinhot {
  0%, 55% { outline: 3px solid var(--teal); outline-offset: 2px; }
  100% { outline: 3px solid transparent; outline-offset: 2px; }
}

/* issue-type summary band */
.summary { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 0.7rem 0.5rem; margin: 0.4rem 0 0.9rem; }
.summary .sumhead { font-weight: 700; font-size: 0.9rem; padding: 0.1rem 0.6rem 0.5rem;
  color: var(--teal); }
.summary .sumhead .note { font-weight: 400; margin-left: 0.4rem; }
.sumrow { display: grid; grid-template-columns: 0.9rem 2.4rem 1fr auto auto;
  align-items: center; gap: 0.7rem; width: 100%; text-align: left;
  background: none; border: 0; border-radius: 8px; cursor: pointer;
  padding: 0.4rem 0.6rem; font-size: 0.9rem; color: var(--teal); }
.sumrow:hover { background: var(--hover); }
.sumrow[aria-pressed="true"] { background: var(--teal-light); }
.sumrow .dot { width: 0.62rem; height: 0.62rem; border-radius: 50%; }
.sumrow .dot.error { background: var(--burgundy); }
.sumrow .dot.warning { background: var(--orange); }
.sumrow .dot.info { background: var(--slate); }
.sumrow .sn { font-weight: 700; font-variant-numeric: tabular-nums; text-align: right; }
.sumrow .sl { min-width: 0; }
.sumrow .ssev { font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.06em;
  color: var(--slate-text); white-space: nowrap; }
.sumrow .ssev .error { color: var(--burgundy); }
.sumrow .ssev .warning { color: var(--orange); }
.sumrow .sfix { font-size: 0.76rem; color: var(--slate-text); min-width: 5rem;
  text-align: right; }

/* assistant triage */
.assist .assistq { display: flex; gap: 0.6rem; align-items: flex-start;
  padding: 0.55rem 0; border-top: 1px solid var(--line-soft); cursor: pointer; }
.assist .assistq input { margin-top: 0.25rem; }
.assist .assistrow { display: flex; gap: 0.8rem; align-items: center;
  padding-top: 0.7rem; }
.assist .copilotrow { display: flex; gap: 0.8rem; align-items: center;
  border-top: 1px solid var(--line-soft); margin-top: 0.8rem;
  padding-top: 0.8rem; }
.assist .copilotrow form { margin: 0; }

/* stats */
table.stats { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  border-collapse: separate; border-spacing: 0; overflow: hidden; width: 100%; }
table.stats th { background: var(--offwhite); color: var(--slate-text); font-size: 0.72rem;
  text-transform: uppercase; letter-spacing: 0.1em; text-align: left;
  padding: 0.6rem 0.9rem; }
table.stats td { padding: 0.55rem 0.9rem; border-top: 1px solid var(--line-soft);
  font-variant-numeric: tabular-nums; }
.fpbar { display: inline-block; height: 8px; border-radius: 4px;
  background: var(--burgundy); vertical-align: middle; margin-right: 0.5rem; }

/* busy overlay */
.busy { position: fixed; inset: 0; z-index: 100; display: flex; align-items: center;
  justify-content: center; background: rgba(244, 246, 246, 0.82);
  backdrop-filter: blur(3px); }
.busy[hidden] { display: none; }
.busycard { background: #fff; border: 1px solid var(--line-soft); border-radius: 14px;
  padding: 2rem 3rem; text-align: center; box-shadow: 0 8px 40px rgba(0, 37, 40, 0.12);
  max-width: min(32rem, calc(100vw - 3rem)); }
.busycard .wordmark { font-size: 1.6rem; }
.busycard .colon { display: inline-block; animation: pulse 1.1s ease-in-out infinite; }
@keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.15; } }
/* long filenames wrap instead of stretching the card off-centre */
.busymsg { font-weight: 700; margin-top: 0.7rem; overflow-wrap: anywhere; }
.busysub { color: var(--slate-text); font-size: 0.85rem; margin-top: 0.25rem;
  max-width: 26rem; margin-inline: auto; }
@media (prefers-reduced-motion: reduce) { .busycard .colon { animation: none; } }
:focus-visible { outline: 2px solid var(--teal); outline-offset: 2px; }
@media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
@media print { .actionbar, .controls, .drop, .no-print { display: none !important; } }
"""


def _shell(title: str, body: str) -> str:
    demo_strip = (f'<div class="demostrip">{esc(DEMO_BANNER)}</div>'
                  if DEMO_BANNER else "")
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)}</title><style>{_CSS}</style></head>
<body>{demo_strip}<div class="wrap">
<div class="brandrow"><span class="wordmark">prezlab:</span>
<span class="kicker">Pre-delivery QC</span>
<span style="flex:1"></span><span class="note"><a href="/">Run an audit</a> &middot; <a href="/master">Read a master</a> &middot; <a href="/format">Apply a master</a> &middot; <a href="/profiles">Profiles</a> &middot; <a href="/team">Team</a> &middot; <a href="/stats">Stats</a></span><span id="who" class="note"></span></div>
{body}
</div>
<div class="busy" id="busy" hidden role="status" aria-live="polite">
 <div class="busycard"><span class="wordmark">prezlab<span class="colon">:</span></span>
  <div class="busymsg" id="busymsg"></div>
  <div class="busysub" id="busysub"></div></div></div>
<script>
function showBusy(msg, sub) {{
  document.getElementById('busymsg').textContent = msg;
  document.getElementById('busysub').textContent = sub || '';
  document.getElementById('busy').hidden = false;
}}
// back/forward cache restores must never resurrect the overlay
window.addEventListener('pageshow', () => {{
  document.getElementById('busy').hidden = true;
}});
// signed-in state in the header; signing in lives on its own page
fetch('/me').then(r => r.json()).then(d => {{
  const who = document.getElementById('who');
  if (d.user) {{
    who.innerHTML = `Signed in as <b>${{d.user.name}}</b> (${{d.user.role}})` +
      ` &nbsp;<a href="#" id="signout">sign out</a>`;
    document.getElementById('signout').addEventListener('click', async e => {{
      e.preventDefault();
      await fetch('/signout', {{method: 'POST'}});
      location.reload();
    }});
  }} else {{
    who.innerHTML = `<a href="/signin"><b>Sign in</b></a>`;
  }}
}}).catch(() => {{}});
</script>
</body></html>"""


def render_index(profiles: list[dict], modules: tuple, message: str = "") -> str:
    banner = f'<div class="banner warn">{esc(message)}</div>' if message else ""
    default_id = ("prezlab_bilingual"
                  if any(p["id"] == "prezlab_bilingual" for p in profiles)
                  else (profiles[0]["id"] if profiles else ""))
    cards = "".join(
        f"""<label class="radio-card"><input type="radio" name="profile"
        value="{esc(p['id'])}"{' checked' if p['id'] == default_id else ''}>
        <span><b>{esc(p.get('label', p['id']))}</b><small>{esc(p['name'])}</small></span></label>"""
        for p in profiles)
    module_chips = "".join(
        f"""<label class="chip"><input type="checkbox" name="modules" value="{m}" checked>
        {esc(MODULE_LABELS.get(m, m))}</label>""" for m in modules)

    body = f"""
<h1>Run a formatting audit.</h1>
<p class="sub">Upload a deck and pick what to hold it to: a saved profile, the
master it was built from, or the deck's own conventions. You get a slide-by-slide
read on what matches and what needs a designer's eye. The deck is read, never
changed.</p>
{banner}
<form action="/audit" method="post" enctype="multipart/form-data" id="f">
  <div class="drop" id="drop" tabindex="0" role="button"
       aria-label="Drop a .pptx file here or press Enter to browse">
    <strong>Drop a .pptx here</strong> or click to browse
    <div class="hint">Audited locally. The file is deleted after the audit;
    findings stay in memory on this machine only.</div>
    <div class="file" id="fname" aria-live="polite"></div>
    <input type="file" name="deck" id="deck" accept=".pptx" required hidden>
  </div>
  <div class="config">
    <fieldset><legend>Rules</legend>{cards}</fieldset>
    <fieldset><legend>Checks to run</legend><div class="chips">{module_chips}</div></fieldset>
  </div>
  <div class="drop" id="mdrop" tabindex="0" role="button" hidden
       aria-label="Drop the master .pptx here or press Enter to browse">
    <strong>Drop the master .pptx here</strong> or click to browse
    <div class="hint">Read for its theme, layouts, and grid. The master is not
    changed and not kept; only the rules it declares carry into this audit.
    <a href="/master">See what gets read</a>.</div>
    <div class="file" id="mname" aria-live="polite"></div>
    <input type="file" name="master" id="master" accept=".pptx" hidden>
  </div>
  <div class="actions">
    <button class="btn primary" id="go" type="submit" disabled>Run audit</button>
  </div>
</form>
<script>
const drop = document.getElementById('drop'), input = document.getElementById('deck'),
      fname = document.getElementById('fname'), go = document.getElementById('go');
const mdrop = document.getElementById('mdrop'), minput = document.getElementById('master'),
      mname = document.getElementById('mname');

// The master drop zone only exists for the upload-a-master option, and its
// input only becomes required while it is visible: a hidden required input
// blocks submission with a validation message the user cannot see.
function syncRuleSource() {{
  const picked = document.querySelector('input[name=profile]:checked');
  const wantsMaster = picked && picked.value === '__master__';
  mdrop.hidden = !wantsMaster;
  minput.required = !!wantsMaster;
  if (!wantsMaster) {{ minput.value = ''; mname.textContent = ''; }}
  arm();
}}
document.querySelectorAll('input[name=profile]').forEach(
  r => r.addEventListener('change', syncRuleSource));
mdrop.addEventListener('click', () => minput.click());
mdrop.addEventListener('keydown', e => {{
  if (e.key === 'Enter' || e.key === ' ') minput.click(); }});
minput.addEventListener('change', () => {{
  mname.textContent = minput.files.length ? minput.files[0].name : '';
  arm();
}});
['dragover', 'dragenter'].forEach(ev => mdrop.addEventListener(ev, e => {{
  e.preventDefault(); mdrop.classList.add('armed'); }}));
['dragleave', 'drop'].forEach(ev => mdrop.addEventListener(ev, e => {{
  e.preventDefault(); mdrop.classList.remove('armed'); }}));
mdrop.addEventListener('drop', e => {{
  if (e.dataTransfer.files.length) {{
    minput.files = e.dataTransfer.files;
    mname.textContent = minput.files[0].name;
    arm();
  }} }});

function arm() {{
  const needMaster = minput.required && !minput.files.length;
  if (input.files.length) {{ fname.textContent = input.files[0].name; }}
  go.disabled = !input.files.length || needMaster;
}}
drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', e => {{ if (e.key === 'Enter' || e.key === ' ') input.click(); }});
input.addEventListener('change', arm);
document.getElementById('f').addEventListener('submit', () => {{
  go.disabled = true;
  showBusy('Auditing ' + (input.files[0] ? input.files[0].name : 'deck'),
           'Reading every slide and checking it against the profile. Larger decks take longer.');
}});
['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => {{
  e.preventDefault(); drop.classList.add('armed'); }}));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {{
  e.preventDefault(); drop.classList.remove('armed'); }}));
drop.addEventListener('drop', e => {{
  if (e.dataTransfer.files.length) {{ input.files = e.dataTransfer.files; arm(); }} }});
syncRuleSource();
</script>"""
    return _shell("Prezlab PPT QC", body)


_SEV_RANK = {"error": 0, "warning": 1, "info": 2}


def render_report(manifest: dict, job_id: str, can_fix: bool = False,
                  banner: str = "", has_cleaned: bool = False,
                  diff_href: str | None = None,
                  triage: dict | None = None,
                  promoted: set | None = None,
                  comments: dict | None = None,
                  archived: bool = False,
                  assist: bool = False,
                  design: int | None = None) -> str:
    from .fixer import is_fixable, tick_reason

    s = manifest["summary"]
    sev = s.get("by_severity", {})
    deck = esc(Path(manifest["deck"]).name)
    records = sorted(manifest["records"],
                     key=lambda r: (r["slide_index"], _SEV_RANK.get(r["severity"], 3)))
    any_fixable = can_fix and any(is_fixable(r) for r in records)
    triage = triage or {}
    promoted = promoted or set()
    comments = comments or {}
    n_slides = manifest["slides"]

    # sticky action bar: counts left, every action right, always in view
    apply_btn = ('<button class="btn ghost" id="selectallbtn" type="button" '
                 'title="Tick every fixable finding, including the ones the '
                 'tool left unticked for your judgment">Select all fixable'
                 '</button>'
                 '<button class="btn primary" id="applybtn" type="submit" '
                 'form="applyform">Apply selected fixes</button>') if any_fixable else ""
    dl_btn = (f'<a class="btn primary" href="/download/{esc(job_id)}" '
              f'download>Download cleaned .pptx</a>') if has_cleaned else ""
    diff_btn = (f'<a class="btn success" href="{esc(diff_href)}">Review changes</a>'
                if diff_href else "")
    # The design pass is a separate door on purpose. Its findings are decisions
    # with several right answers, not rule violations with one fix, and mixing a
    # "pick one of three" control into a list whose every other row is a single
    # tick teaches the tick to mean something it does not.
    design_btn = "" if archived else (
        f'<a class="btn ghost" href="/design/{esc(job_id)}" '
        f'title="Palette conflicts, unreadable text, overlapping shapes and '
        f'content outside the frame - each with the ways to fix it">Design QC'
        + (f' <b>{design}</b>' if design else "") + '</a>')
    exports_span = "" if archived else (
        f'<span class="exports">Export: '
        f'<a href="/report/{esc(job_id)}.pdf">PDF</a> &middot; '
        f'<a href="/report/{esc(job_id)}.csv">CSV</a> &middot; '
        f'<a href="/manifest/{esc(job_id)}">JSON</a> &middot; '
        f'<a href="/annotated/{esc(job_id)}" title="Deck copy with findings '
        f'and comments in speaker notes">Annotated .pptx</a></span>')
    actionbar = f"""
<div class="actionbar no-print">
  <span class="minicounts">
    <span><span class="dot" style="background:var(--burgundy)"></span>{sev.get('error', 0)} errors</span>
    <span><span class="dot" style="background:var(--orange)"></span>{sev.get('warning', 0)} warnings</span>
    <span><span class="dot" style="background:var(--slate)"></span>{sev.get('info', 0)} info</span>
    <span><span class="dot" style="background:var(--lavender)"></span>{s.get('arabic_flagged', 0)} Arabic review</span>
  </span>
  <span class="grow"></span>
  {apply_btn}{diff_btn}{dl_btn}{design_btn}
  {exports_span}
  <a class="btn ghost" href="/">New audit</a>
</div>"""

    banner_html = (f'<div class="banner ok" role="status">&#10003;&nbsp; '
                   f'{esc(banner)}</div>') if banner else ""

    if not records:
        findings = """<div class="card clean"><div class="mark">&#10003;</div>
<h2>No findings</h2><p>This deck matches the profile. Ready for the next pass.</p></div>"""
        controls = groups_js = summary_band = ""
    else:
        # pin badges: same numbering rule as the preview overlays, so badge
        # N in this list is box N on the rendered slide
        from .render import pin_numbers

        pin_of = pin_numbers(records)

        # group records by slide
        by_slide: dict[int, list] = {}
        for r in records:
            by_slide.setdefault(r["slide_index"], []).append(r)

        modules_present = sorted({r["module"] for r in records})
        groups = []
        for slide_idx in sorted(by_slide):
            rows = []
            counts = {"error": 0, "warning": 0, "info": 0}
            ar_count = 0
            for r in by_slide[slide_idx]:
                counts[r["severity"]] = counts.get(r["severity"], 0) + 1
                ar_count += 1 if r["arabic_flag"] else 0
                ar = ' <span class="pill ar">AR</span>' if r["arabic_flag"] else ""
                if r["action"] == "changed":
                    fix_cell = '<span class="pill changed">fixed</span>'
                elif can_fix and is_fixable(r):
                    # Pre-ticked = the tool is confident it's wrong AND has a
                    # safe fix: deterministic changes, triage-promoted types,
                    # and errors (by taxonomy, error means confidently wrong).
                    # Ticking is reserved for the genuinely ambiguous, and
                    # Arabic font substitutions are NEVER pre-selected:
                    # shaping changes with the font, the tick is the approval.
                    hold = tick_reason(r)
                    pre = (hold is None
                           and (r["confidence"] == "deterministic"
                                or r["issue_type"] in promoted
                                or r["severity"] == "error"))
                    checked = " checked" if pre else ""
                    why = (hold if hold
                           else "deterministic fix"
                           if r["confidence"] == "deterministic"
                           else "confidently wrong (error) with a safe fix"
                           if r["severity"] == "error"
                           else "validated by designer triage" if pre
                           else "suggestion: ticking it is your approval")
                    fix_cell = (f'<input type="checkbox" name="record_ids" '
                                f'value="{esc(r["record_id"])}" form="applyform"'
                                f'{checked} title="{esc(why)}" aria-label="Apply this fix">')
                else:
                    fix_cell = ""
                fixable_row = "1" if (can_fix and is_fixable(r)) else "0"
                tri_state = triage.get(r["record_id"], "")
                tri_cls = f" class=\"t-{tri_state}\"" if tri_state else ""
                tri_cell = ("" if (r["module"] == "preflight" or archived) else
                            f'<span class="tri">'
                            f'<button type="button" class="ok" title="Agree: real issue"'
                            f' data-t="confirmed">&#10003;</button>'
                            f'<button type="button" class="fp" title="False alarm"'
                            f' data-t="false_positive">&#10005;</button></span>')
                pin = pin_of.get(r["record_id"])
                if pin is not None:
                    pin_cell = (f'<button type="button" class="pinno {esc(r["severity"])}"'
                                f' data-pin="{pin}" title="Box {pin} on the slide '
                                f'preview; click to highlight">{pin}</button>')
                elif str(r.get("shape_id") or "-") == "-" and r["module"] != "preflight":
                    pin_cell = ('<span class="pinno whole" '
                                'title="Applies to the whole slide">S</span>')
                else:
                    pin_cell = ""
                data_pin = f' data-pin="{pin}"' if pin is not None else ""
                rows.append(
                    f"""<tr{tri_cls} data-sev="{esc(r['severity'])}" data-ar="{'1' if r['arabic_flag'] else '0'}"
 data-mod="{esc(r['module'])}" data-type="{esc(r['issue_type'])}" data-fix="{fixable_row}" data-record="{esc(r['record_id'])}"{data_pin}>
<td class="fix">{fix_cell}</td>
<td class="pinc">{pin_cell}</td>
<td class="sev"><span class="pill {esc(r['severity'])}">{esc(r['severity'])}</span>{ar}</td>
<td class="code">{esc(r['issue_type'])}</td>
<td class="msg" dir="auto">{esc(r['message'])}</td>
<td class="trit">{tri_cell}</td></tr>""")
            summary_bits = []
            if counts["error"]:
                n = counts["error"]
                summary_bits.append(f"{n} error{'s' if n != 1 else ''}")
            if counts["warning"]:
                n = counts["warning"]
                summary_bits.append(f"{n} warning{'s' if n != 1 else ''}")
            if counts["info"]:
                summary_bits.append(f"{counts['info']} info")
            if ar_count:
                summary_bits.append(f"{ar_count} AR")
            # Groups start collapsed: the summary band up top is the entry
            # point; a filter click or expand-all opens what matters.
            open_attr = ""
            prev_btn = "" if archived else (
                f'<button type="button" class="fchip prevbtn" '
                f'data-slide="{slide_idx}">Preview slide</button>')
            groups.append(f"""
<details class="grp"{open_attr}><summary>Slide {slide_idx + 1}
 <span class="grpcounts">{esc(' · '.join(summary_bits))}</span>
 {prev_btn}
 <button type="button" class="fchip cmtbtn" data-slide="{slide_idx}">Comments
 ({comments.get(slide_idx, 0)})</button>
</summary>
<table><tbody>{''.join(rows)}</tbody></table></details>""")

        def _chip(label, count, key):
            dis = " disabled" if count == 0 else ""
            pressed = "true" if key == "all" else "false"
            return (f'<button class="fchip" aria-pressed="{pressed}" data-f="{key}"'
                    f'{dis}>{label} {count}</button>')

        mod_chips = "".join(
            f'<button class="fchip" aria-pressed="false" data-m="{m}">'
            f'{esc(MODULE_LABELS.get(m, m))}</button>' for m in modules_present)

        # Issue-type summary: one row per type, worst severity wins its dot,
        # sorted by severity then count. Clicking a row filters the list to
        # that type. Preflight (info-only disclosure) is folded to the end.
        by_type: dict[str, dict] = {}
        for r in records:
            slot = by_type.setdefault(
                r["issue_type"],
                {"n": 0, "sev": "info", "fix": 0,
                 "counts": {"error": 0, "warning": 0, "info": 0}})
            slot["n"] += 1
            slot["counts"][r["severity"]] += 1
            if _SEV_RANK[r["severity"]] < _SEV_RANK[slot["sev"]]:
                slot["sev"] = r["severity"]
            if can_fix and is_fixable(r):
                slot["fix"] += 1
        ordered = sorted(by_type.items(),
                         key=lambda kv: (_SEV_RANK[kv[1]["sev"]], -kv[1]["n"]))

        def _sev_label(slot: dict) -> str:
            # uniform severity -> one word; mixed -> the split, so a row of
            # "60 edge misaligned" reads honestly as 26 error / 34 warning
            present = [(sv, slot["counts"][sv]) for sv in
                       ("error", "warning", "info") if slot["counts"][sv]]
            if len(present) == 1:
                return f'<span class="{present[0][0]}">{present[0][0]}</span>'
            return " &middot; ".join(
                f'<span class="{sv}">{n} {sv}</span>' for sv, n in present)

        summary_rows = "".join(
            f'<button type="button" class="sumrow" data-type="{esc(it)}">'
            f'<span class="dot {esc(slot["sev"])}"></span>'
            f'<span class="sn">{slot["n"]}</span>'
            f'<span class="sl">{esc(issue_label(it))}</span>'
            f'<span class="ssev">{_sev_label(slot)}</span>'
            f'<span class="sfix">{("&#10003; " + str(slot["fix"]) + " fixable") if slot["fix"] else ""}</span>'
            '</button>'
            for it, slot in ordered)
        summary_band = f"""
<div class="summary no-print">
 <div class="sumhead">What was found <span class="note">click a row to filter;
  slides start collapsed</span></div>
 {summary_rows}
</div>"""
        controls = f"""
<div class="controls no-print" role="group" aria-label="Filter findings">
  {_chip("All", s.get('total', 0), "all")}
  {_chip("Errors", sev.get('error', 0), "error")}
  {_chip("Warnings", sev.get('warning', 0), "warning")}
  {_chip("Info", sev.get('info', 0), "info")}
  {_chip("Arabic", s.get('arabic_flagged', 0), "ar")}
  {_chip("Fixable", sum(1 for r in records if can_fix and is_fixable(r)), "fix")}
  <span class="vdiv"></span>
  {mod_chips}
  <input class="search" id="search" type="search" placeholder="Search findings"
   aria-label="Search findings">
  <span class="shown" id="shown"></span>
  <span class="shown" id="trisum"></span>
  <button class="fchip" id="expandall" type="button">Expand all</button>
</div>"""
        findings = "".join(groups)
        groups_js = """
<script>
let sevFilter = 'all', modFilter = null, typeFilter = null, q = '';
const rows = [...document.querySelectorAll('details.grp tbody tr')];
const grps = [...document.querySelectorAll('details.grp')];
const shown = document.getElementById('shown');
function refresh() {
  let visible = 0;
  rows.forEach(tr => {
    const okSev = sevFilter === 'all' || tr.dataset.sev === sevFilter ||
                  (sevFilter === 'ar' && tr.dataset.ar === '1') ||
                  (sevFilter === 'fix' && tr.dataset.fix === '1');
    const okMod = !modFilter || tr.dataset.mod === modFilter;
    const okType = !typeFilter || tr.dataset.type === typeFilter;
    const okQ = !q || tr.textContent.toLowerCase().includes(q);
    const ok = okSev && okMod && okType && okQ;
    tr.style.display = ok ? '' : 'none';
    if (ok) visible++;
  });
  grps.forEach(g => {
    const any = [...g.querySelectorAll('tbody tr')].some(tr => tr.style.display !== 'none');
    g.style.display = any ? '' : 'none';
    if (sevFilter !== 'all' || modFilter || typeFilter || q) g.open = true;
  });
  shown.textContent = visible + ' shown';
}
function clearTypeFilter() {
  typeFilter = null;
  document.querySelectorAll('.sumrow').forEach(x => x.setAttribute('aria-pressed', 'false'));
}
document.querySelectorAll('.fchip[data-f]').forEach(c => c.addEventListener('click', () => {
  document.querySelectorAll('.fchip[data-f]').forEach(x => x.setAttribute('aria-pressed', 'false'));
  c.setAttribute('aria-pressed', 'true'); sevFilter = c.dataset.f; clearTypeFilter(); refresh();
}));
// issue-type summary rows: click to filter the list to that type; click the
// active row again to clear
document.querySelectorAll('.sumrow').forEach(rw => rw.addEventListener('click', () => {
  const on = rw.getAttribute('aria-pressed') === 'true';
  document.querySelectorAll('.sumrow').forEach(x => x.setAttribute('aria-pressed', 'false'));
  typeFilter = on ? null : rw.dataset.type;
  if (!on) rw.setAttribute('aria-pressed', 'true');
  // reset the severity chips so the two filter surfaces don't fight
  sevFilter = 'all';
  document.querySelectorAll('.fchip[data-f]').forEach(
    x => x.setAttribute('aria-pressed', x.dataset.f === 'all' ? 'true' : 'false'));
  refresh();
  if (!on) {
    const c = document.querySelector('.controls');
    if (c) c.scrollIntoView({behavior: 'smooth', block: 'start'});
  }
}));
document.querySelectorAll('.fchip[data-m]').forEach(c => c.addEventListener('click', () => {
  const on = c.getAttribute('aria-pressed') === 'true';
  document.querySelectorAll('.fchip[data-m]').forEach(x => x.setAttribute('aria-pressed', 'false'));
  modFilter = on ? null : c.dataset.m;
  if (!on) c.setAttribute('aria-pressed', 'true');
  clearTypeFilter(); refresh();
}));
document.getElementById('search').addEventListener('input', e => {
  q = e.target.value.trim().toLowerCase(); refresh();
});
const ex = document.getElementById('expandall');
ex.addEventListener('click', () => {
  const anyClosed = grps.some(g => !g.open && g.style.display !== 'none');
  grps.forEach(g => g.open = anyClosed);
  ex.textContent = anyClosed ? 'Collapse all' : 'Expand all';
});
function jobIdFromLinks() {
  const m = document.querySelector('a[href^="/manifest/"]');
  return m ? m.getAttribute('href').split('/').pop() : '';
}
const JOB = jobIdFromLinks();
// triage: capture designer judgment per finding
document.querySelectorAll('.tri button').forEach(b => b.addEventListener('click', async e => {
  e.preventDefault();
  const tr = b.closest('tr');
  const want = 't-' + b.dataset.t;
  const state = tr.classList.contains(want) ? 'cleared' : b.dataset.t;
  const res = await fetch('/triage', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({job_id: JOB, record_id: tr.dataset.record, state}),
  });
  if (!res.ok) return;
  tr.classList.remove('t-confirmed', 't-false_positive');
  if (state !== 'cleared') tr.classList.add('t-' + state);
  const c = (await res.json()).counts;
  document.getElementById('trisum').textContent =
    `Triage: ${c.confirmed} agreed · ${c.false_positive} false alarm${c.false_positive === 1 ? '' : 's'}`;
}));
// pin sync: badge in the list <-> numbered box on the preview
function flash(el) {
  el.classList.remove('pinhot');
  void el.offsetWidth;  // restart the animation
  el.classList.add('pinhot');
}
document.querySelectorAll('tr .pinno[data-pin]').forEach(p => p.addEventListener('click', async e => {
  e.preventDefault(); e.stopPropagation();
  const grp = p.closest('details');
  if (!grp.querySelector('.prev')) {
    const btn = grp.querySelector('.prevbtn');
    if (!btn) return;
    btn.click();
    // first render drives PowerPoint and can take a while
    for (let i = 0; i < 80 && !grp.querySelector('.hl'); i++)
      await new Promise(res => setTimeout(res, 250));
  }
  const hl = grp.querySelector(`.hl[data-pin="${p.dataset.pin}"]`);
  if (hl) { hl.scrollIntoView({behavior: 'smooth', block: 'center'}); flash(hl); }
}));
// slide previews: rendered by PowerPoint on first request, cached after
document.querySelectorAll('.prevbtn').forEach(b => b.addEventListener('click', async e => {
  e.preventDefault(); e.stopPropagation();
  const grp = b.closest('details');
  const existing = grp.querySelector('.prev:not(.cmts)');
  if (existing) {
    // unwrap: put the findings table back where the split container sits
    const split = grp.querySelector('.split');
    if (split) {
      const table = split.querySelector('table');
      if (table) grp.insertBefore(table, split);
      split.remove();
    } else {
      existing.remove();
    }
    b.textContent = 'Preview slide';
    return;
  }
  grp.open = true;
  b.textContent = 'Rendering…';
  b.disabled = true;
  try {
    const res = await fetch(`/slide/${JOB}/${b.dataset.slide}`);
    if (!res.ok) {
      const err = await res.json();
      alert(err.error || 'Preview unavailable'); return;
    }
    const data = await res.json();
    const panel = document.createElement('div');
    panel.className = 'prev';
    const overlays = data.rects.map(r =>
      `<span class="hl ${r.severity}${r.arabic ? ' arflag' : ''}" title="${r.label}"` +
      ` data-pin="${r.pin}"` +
      ` style="left:${(r.x * 100).toFixed(2)}%;top:${(r.y * 100).toFixed(2)}%;` +
      `width:${(r.w * 100).toFixed(2)}%;height:${(r.h * 100).toFixed(2)}%">` +
      `<i class="pinno ${r.severity}">${r.pin}</i></span>`).join('');
    panel.innerHTML = `<div class="shot"><img src="${data.png}" alt="Slide preview">${overlays}</div>`;
    // slide stays visible beside the findings while they scroll: preview and
    // table share a split container, preview sticky
    const table = grp.querySelector('table');
    const split = document.createElement('div');
    split.className = 'split';
    grp.insertBefore(split, table);
    split.appendChild(panel);
    split.appendChild(table);
    // box -> findings list: clicking a pin flashes its rows
    panel.querySelectorAll('.hl').forEach(hl => hl.addEventListener('click', () => {
      grp.querySelectorAll(`tbody tr[data-pin="${hl.dataset.pin}"]`).forEach((tr, i) => {
        if (i === 0) tr.scrollIntoView({behavior: 'smooth', block: 'center'});
        flash(tr);
      });
    }));
    // findings list -> box: hovering a row lights up its box on the slide
    grp.querySelectorAll('tbody tr[data-pin]').forEach(tr => {
      if (tr.dataset.hoverwired) return;
      tr.dataset.hoverwired = '1';
      tr.addEventListener('mouseenter', () => {
        const hl = grp.querySelector(`.hl[data-pin="${tr.dataset.pin}"]`);
        if (hl) hl.classList.add('hlhover');
      });
      tr.addEventListener('mouseleave', () => {
        grp.querySelectorAll('.hl.hlhover').forEach(h => h.classList.remove('hlhover'));
      });
    });
    b.textContent = 'Hide preview';
  } finally {
    b.disabled = false;
    if (b.textContent === 'Rendering…') b.textContent = 'Preview slide';
  }
}));
// per-slide comments: attribution via the pilot identity cookie
const DECK = (document.getElementById('deckname') || {textContent: ''}).textContent;
document.querySelectorAll('.cmtbtn').forEach(b => b.addEventListener('click', async e => {
  e.preventDefault(); e.stopPropagation();
  const grp = b.closest('details');
  const existing = grp.querySelector('.cmts');
  if (existing) { existing.remove(); return; }
  grp.open = true;
  const res = await fetch(`/comments?deck=${encodeURIComponent(DECK)}&slide=${b.dataset.slide}`);
  const data = await res.json();
  const panel = document.createElement('div');
  panel.className = 'cmts prev cmtspanel';
  const items = data.comments.map(c =>
    `<p class="note"><b>${c.author}</b> · ${c.created_at.slice(0, 16).replace('T', ' ')}<br>` +
    `${c.text.replace(/</g, '&lt;')}</p>`).join('') || '<p class="note">No comments yet.</p>';
  panel.innerHTML = `${items}
    <textarea class="search" rows="2" placeholder="Add a comment" style="border-radius:10px;width:100%"></textarea>
    <button type="button" class="btn ghost" style="margin-top:0.4rem">Save comment</button>`;
  // the table may live inside the preview's split container; anchor on
  // whichever of the two is a direct child of the group
  const anchor = grp.querySelector(':scope > .split') || grp.querySelector('table');
  grp.insertBefore(panel, anchor);
  panel.querySelector('button').addEventListener('click', async () => {
    const text = panel.querySelector('textarea').value.trim();
    if (!text) return;
    const r = await fetch('/comments', {method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({deck: DECK, slide_index: parseInt(b.dataset.slide), text})});
    if (r.status === 401) { alert('Pick your name in the top right first.'); return; }
    if (!r.ok) return;
    b.textContent = `Comments (${data.comments.length + 1})`;
    panel.remove(); b.click();
  });
}));
const btn = document.getElementById('applybtn');
const applyForm = document.getElementById('applyform');
if (applyForm) applyForm.addEventListener('submit', () => {
  const n = [...document.querySelectorAll('input[name="record_ids"]')].filter(b => b.checked).length;
  showBusy('Applying ' + n + ' fix' + (n === 1 ? '' : 'es'),
           'Writing the changes to a copy of the deck, then re-auditing it to verify.');
});
document.querySelectorAll('a[href^="/diff/"]').forEach(a => a.addEventListener('click', () => {
  showBusy('Rendering before and after',
           'PowerPoint is exporting the changed slides. The first open can take up to a minute; after that it is instant.');
}));
if (btn) {
  const boxes = () => document.querySelectorAll('input[name="record_ids"]');
  const sync = () => {
    const n = [...boxes()].filter(b => b.checked).length;
    btn.textContent = n ? `Apply ${n} selected fix${n === 1 ? '' : 'es'}` : 'Apply selected fixes';
    btn.disabled = n === 0;
  };
  boxes().forEach(b => b.addEventListener('change', sync));
  sync();
  const selAll = document.getElementById('selectallbtn');
  if (selAll) selAll.addEventListener('click', () => {
    const all = [...boxes()];
    const everyOn = all.every(b => b.checked);
    all.forEach(b => { b.checked = !everyOn; });  // second click deselects
    selAll.textContent = everyOn ? 'Select all fixable' : 'Deselect all';
    sync();
  });
}
// assistant triage: questions come from the server with actions held there;
// only accepted question ids go back, never action payloads
const askBtn = document.getElementById('askassist');
if (askBtn) askBtn.addEventListener('click', async () => {
  const out = document.getElementById('assistout');
  const escq = s => String(s).replace(/[&<>"]/g,
    c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
  askBtn.disabled = true;
  askBtn.textContent = 'Reading the findings…';
  try {
    const res = await fetch(`/assist/${JOB}`, {method: 'POST'});
    const data = await res.json();
    if (!res.ok) { out.innerHTML = `<p class="note">${escq(data.error)}</p>`; return; }
    if (!data.questions.length) {
      out.innerHTML = '<p class="note">Nothing worth asking: no recurring ambiguous findings on this deck.</p>';
      return;
    }
    out.innerHTML = data.questions.map(q => `
      <label class="assistq"><input type="checkbox" value="${escq(q.id)}" checked>
       <span><b>${escq(q.question)}</b><br>
       <span class="note">${escq(q.rationale)} ${escq(q.impact)}</span></span></label>`).join('') +
      `<div class="assistrow">
        <button type="button" class="btn primary" id="assistapply">Apply accepted answers</button>
        <span class="note">${data.source.startsWith('fallback') ? 'Offline phrasing' : 'Phrased by Claude'}
         · updates profile <b>${escq(data.profile)}</b></span></div>`;
    document.getElementById('assistapply').addEventListener('click', async () => {
      const ids = [...out.querySelectorAll('.assistq input:checked')].map(b => b.value);
      if (!ids.length) return;
      const form = new URLSearchParams();
      ids.forEach(id => form.append('accepted', id));
      const r = await fetch(`/assist/${JOB}/apply`, {method: 'POST', body: form});
      const d = await r.json();
      if (!r.ok) { out.insertAdjacentHTML('beforeend', `<p class="note">${escq(d.error)}</p>`); return; }
      out.innerHTML = `<p class="note">&#10003; Profile <b>${escq(d.profile)}</b> updated to
        v${escq(d.version)}: ${d.applied.map(escq).join('; ')}.</p>
        <form method="post" action="/reaudit/${JOB}"
         onsubmit="showBusy('Re-auditing with the updated profile',
          'Running the same deck against the new rules.')">
         <button class="btn primary" type="submit">Re-audit this deck now</button></form>`;
    });
  } finally {
    askBtn.disabled = false;
    askBtn.textContent = 'Ask the assistant';
  }
});
refresh();
</script>"""

    apply_form = (f'<form id="applyform" method="post" action="/apply" class="no-print">'
                  f'<input type="hidden" name="job_id" value="{esc(job_id)}"></form>'
                  if any_fixable else "")

    assist_panel = ""
    if assist and can_fix and records and not archived:
        assist_panel = f"""
<div class="card assist no-print">
 <b>Assistant</b>
 <p class="note">Turns recurring findings into quick yes/no questions;
  accepted answers update the profile so future audits stop flagging them.
  Only finding metadata (colors, fonts, counts) is used, never slide
  content. Applying answers needs a lead or admin.</p>
 <button type="button" class="btn ghost" id="askassist">Ask the assistant</button>
 <div id="assistout"></div>
 <div class="copilotrow">
  <form method="post" action="/copilot/{esc(job_id)}"
    onsubmit="showBusy('Design copilot is reviewing the slides',
     'Rendering each slide, then asking Claude what a designer would adjust. A full deck takes a minute or two.')">
   <button class="btn ghost" type="submit">Layout review with Claude vision</button>
  </form>
  <span class="note">Sends slide <b>images</b> to the Anthropic API
   (unlike the assistant above). Use only on decks approved for cloud
   processing. Suggestions arrive tickable, never pre-selected.</span>
 </div>
 <div class="copilotrow">
  <form method="post" action="/components/{esc(job_id)}"
    onsubmit="showBusy('Working out what the components are',
     'Claude names the things on each slide and which line they belong on; the geometry is measured here. A full deck takes a minute or two.')">
   <button class="btn ghost" type="submit">Component review</button>
  </form>
  <span class="note">Answers the two questions geometry cannot: which
   shapes are <b>one thing</b> (a card with its icon and label), and which
   line they were meant to share &mdash; the master's frame, or another
   component. Claude never supplies a coordinate; every target is measured
   here. Slide <b>images</b> are sent, as above.</span>
 </div>
</div>"""
    fix_note = ("""<p class="note no-print">Confident fixes (errors and
 deterministic changes) come pre-selected; ambiguous suggestions need your
 tick. Rows marked AR are never auto-fixed. The original file is never
 overwritten.</p>""" if any_fixable else "")

    body = f"""
<span class="kicker">Audit report</span>
<h1 class="file">{deck}</h1>
<p class="sub">Profile <b>{esc(manifest['profile_id'])}</b>
 v{esc(manifest['profile_version'])} &middot; <b>{n_slides}</b>
 slide{'s' if n_slides != 1 else ''}</p>
{actionbar}
{banner_html}
<div class="kpis">
 <div class="kpi err"><div class="n">{sev.get('error', 0)}</div><div class="l">Errors</div></div>
 <div class="kpi warn"><div class="n">{sev.get('warning', 0)}</div><div class="l">Warnings</div></div>
 <div class="kpi info"><div class="n">{sev.get('info', 0)}</div><div class="l">Info</div></div>
 <div class="kpi ar"><div class="n">{s.get('arabic_flagged', 0)}</div><div class="l">Arabic review</div></div>
</div>
{assist_panel}
{summary_band}
{controls}
{fix_note}
{apply_form}
<span id="deckname" hidden>{deck}</span>
{findings}
{groups_js}"""
    return _shell(f"Audit: {Path(manifest['deck']).name}", body)


def render_diff(deck_name: str, job_id: str, diff: dict | None,
                error: str = "") -> str:
    deck = esc(Path(deck_name).name)
    top = f"""
<span class="kicker">Before / after review</span>
<h1 class="file">{deck}</h1>
<div class="actionbar no-print">
  <span class="grow"></span>
  <a class="btn primary" href="/download/{esc(job_id)}" download>Download cleaned .pptx</a>
  <a class="btn ghost" href="/diff/{esc(job_id)}.pdf">Export review PDF</a>
  <a class="btn ghost" href="javascript:history.back()">Back to report</a>
  <a class="btn ghost" href="/">New audit</a>
</div>"""

    if error:
        return _shell(f"Review: {deck_name}",
                      top + f'<div class="banner warn">{esc(error)}</div>')

    slides_html = []
    for sl in diff["slides"]:
        idx = sl["index"]

        def _overlays(rects, cls):
            return "".join(
                f'<span class="hl {cls}" title="{esc(r["label"])}" '
                f'style="left:{r["x"] * 100:.2f}%;top:{r["y"] * 100:.2f}%;'
                f'width:{r["w"] * 100:.2f}%;height:{r["h"] * 100:.2f}%"></span>'
                for r in rects)

        n = sl["changes"]
        slides_html.append(f"""
<div class="diffslide">
 <h3>Slide {idx + 1} <span class="grpcounts">{n} change{'s' if n != 1 else ''}</span></h3>
 <div class="difflabels">{esc(' · '.join(sl['labels']))}</div>
 <div class="panes">
  <div class="pane"><div class="tag">Before</div>
   <div class="shot"><img src="/render/{esc(job_id)}/before-{idx}.png"
        alt="Slide {idx + 1} before fixes">{_overlays(sl['before_rects'], 'before')}</div></div>
  <div class="pane"><div class="tag">After</div>
   <div class="shot"><img src="/render/{esc(job_id)}/after-{idx}.png"
        alt="Slide {idx + 1} after fixes">{_overlays(sl['after_rects'], 'after')}</div></div>
 </div>
</div>""")

    n_slides = len(diff["slides"])
    legend = f"""
<p class="sub">{n_slides} slide{'s' if n_slides != 1 else ''} changed. Rendered by
 PowerPoint on this machine; images stay in memory and are never written to disk.</p>
<div class="legend">
 <span><span class="swatch" style="border:2px dashed var(--orange);
  background:rgba(255,124,74,0.12)"></span>changed element, before</span>
 <span><span class="swatch" style="border:2px solid #0e7c66;
  background:rgba(14,124,102,0.10)"></span>same element, after the fix</span>
</div>"""
    return _shell(f"Review: {deck_name}", top + legend + "".join(slides_html))


def render_stats(rows: list[dict]) -> str:
    if not rows:
        body = """
<span class="kicker">Detection quality</span>
<h1>Triage stats</h1>
<p class="sub">No triage data yet. Open an audit report and use the
 &#10003; / &#10005; buttons on findings; every judgment lands here and
 tunes which fixes graduate to one-click.</p>
<p><a class="btn ghost" href="/">Back to audits</a></p>"""
        return _shell("Triage stats", body)

    trs = []
    for r in rows:
        pct = round(r["fp_rate"] * 100)
        bar = f'<span class="fpbar" style="width:{max(2, pct)}px"></span>' if pct else ""
        trs.append(
            f"""<tr><td class="code">{esc(r['issue_type'])}</td>
<td>{esc(MODULE_LABELS.get(r['module'], r['module']))}</td>
<td>{r['reviewed']}</td><td>{r['confirmed']}</td>
<td>{r['false_alarms']}</td><td>{bar}{pct}%</td></tr>""")
    body = f"""
<span class="kicker">Detection quality</span>
<h1>Triage stats</h1>
<p class="sub">Latest designer judgment per finding, aggregated per check.
 High false-alarm rates point at rules to tune; low ones qualify fixes for
 one-click. <a href="/">Back to audits</a></p>
<table class="stats"><thead><tr><th>Check</th><th>Module</th><th>Reviewed</th>
<th>Agreed</th><th>False alarms</th><th>False-alarm rate</th></tr></thead>
<tbody>{''.join(trs)}</tbody></table>"""
    return _shell("Triage stats", body)


def render_signin(users: list[dict], message: str = "") -> str:
    banner = f'<div class="banner warn">{esc(message)}</div>' if message else ""
    cards = "".join(
        f"""<label class="radio-card"><input type="radio" name="name"
        value="{esc(u['name'])}" required>
        <span><b>{esc(u['name'])}</b><small>{esc(u['role'])}
        {'&middot; PIN set' if u.get('pin_hash') else '&middot; first sign-in sets your PIN'}</small></span></label>"""
        for u in users) or ('<p class="note">Nobody on the team yet. '
                            '<a href="/team">Add the first person</a>, then sign in.</p>')
    body = f"""
<span class="kicker">Sign in</span>
<h1>Who is working?</h1>
<p class="sub">Sign-in attributes triage judgments, comments, audits, and
 profile edits to you. Your first sign-in sets your PIN.</p>
{banner}
<form method="post" action="/signin">
 <div class="config">
  <fieldset><legend>Name</legend>{cards}</fieldset>
  <fieldset><legend>PIN</legend>
   <input type="password" name="pin" minlength="4" required
    placeholder="At least 4 characters" autocomplete="current-password">
   <p class="note">Forgot it? Any lead or admin can reset your PIN from
    the team page.</p>
  </fieldset>
 </div>
 <div class="actions">
  <button class="btn primary" type="submit">Sign in</button>
  <a class="btn ghost" href="/team">Manage team</a>
 </div>
</form>"""
    return _shell("Sign in", body)


def render_history(rows: list[dict]) -> str:
    trs = []
    for r in rows:
        when = r["created_at"].replace("T", " ")[:16]
        kind = ('<span class="pill changed">fix pass</span>' if r["kind"] == "fix"
                else '<span class="pill info">audit</span>')
        trs.append(f"""<tr>
<td>{esc(when)}</td><td class="code">{esc(r['deck'])}</td>
<td>{esc(r['profile_id'])} v{esc(r['profile_version'])}</td>
<td>{esc(r['user_name'])}</td><td>{r['slides']}</td>
<td><span class="pill error">{r['errors']}</span>
 <span class="pill warning">{r['warnings']}</span>
 <span class="pill ar">{r['arabic']}</span></td>
<td>{kind}</td>
<td><a href="/history/{r['id']}">Open</a></td></tr>""")
    table = ("""<table class="stats"><thead><tr><th>When (UTC)</th><th>Deck</th>
<th>Profile</th><th>By</th><th>Slides</th><th>Findings</th><th>Type</th><th></th></tr></thead>
<tbody>""" + "".join(trs) + "</tbody></table>") if trs else (
        '<p class="note">No audits recorded yet. Run one and it lands here.</p>')
    body = f"""
<span class="kicker">Audit history</span>
<h1>History.</h1>
<p class="sub">Every audit and fix pass, kept with its full findings record.
 Reports stay readable here even after the working copy leaves memory.
 Decks themselves are never stored. <a href="/">Back to audits</a></p>
{table}"""
    return _shell("Audit history", body)
