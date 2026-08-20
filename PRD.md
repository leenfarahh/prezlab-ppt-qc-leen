# PowerPoint Formatting QC Tool: Product Requirements Document

**Status:** Draft v0.1 (for review)  
**Date:** 30/06/2026  
**Prepared for:** Sanad Zaqtan, Prezlab  
**Confidentiality:** Internal. Contains no client data.

---

## Document control and key decisions

This PRD specifies an in-house, web-based QC assistant that audits and (in later phases) cleans formatting on PowerPoint decks against per-client profiles. It was produced from a structured feasibility analysis (technical, architecture, product/UX, rendering/Arabic, delivery, and prior-art lenses) whose riskiest technical claims were independently verified against the python-pptx source, the OOXML/ECMA-376 specification, and Microsoft guidance.

Three decisions shape everything that follows:

1. **Audit-first, not auto-fix-first.** The original plan shipped silent bulk auto-fix in the MVP and deferred audit and preview to Phase 2. That inverts the risk curve. This PRD ships audit, flagging, reporting, and preview in v1 (the same detection engine auto-fix would use), adds safe deterministic fixes with per-change accept/reject in v1.5, and defers the hard, judgment-heavy fixes to v2.
2. **A 2-week technical spike precedes any timeline commitment.** Two capabilities (enforcing a Master Slide on populated slides, and trustworthy alignment/shape auto-fix) are open-ended editing problems, not bounded ones. The spike retires those unknowns and settles the engine choice before an estimate is promised. The originally bundled auto-fix MVP is realistically 10 to 14 weeks; the audit-first v1 fits the ~8-week target.
3. **Build in-house, with the engine as an open library decision.** python-pptx is the baseline; Aspose.Slides (commercial) is the contingency where its better RTL and layout-reassignment support is needed. The spike decides.

Two corrections from the verification pass are load-bearing in the spec below:

- **Master Slide enforcement is constrained.** python-pptx has no API to re-apply a layout to an existing populated slide, so v1 enforces an explicitly chosen existing master and flags outliers; it never synthesizes a master.
- **Arabic safety is a v1 requirement, not a Phase 2 feature.** Every v1 module touches the OOXML properties that carry Arabic, and an RTL-unaware font swap silently destroys Arabic glyphs. v1 detects Arabic and guards direction-sensitive and font operations; full RTL rules follow in v2.

The canonical schemas, enumerations, and endpoint names are defined once in **Appendix A** and govern wherever a section appears to differ.

---

## Table of contents

1. Overview, Goals & Success Metrics
2. Personas, Roles & User Flows
3. Functional Requirements: Modules 1-3 (Master Slide, Font, Margin & Alignment)
4. Functional Requirements: Modules 4-6 (Color Palette, Shape & Size, Header & Footer)
5. Formatting Profiles (Data Model)
6. System Architecture & Stack
7. Change Manifest, Audit Report & Preview
8. Arabic / RTL Safety
9. Non-Functional Requirements (Performance, Security, Reliability)
10. API Specification
11. Delivery Plan, Risks & Validation
- Appendix A. Canonical Data Contracts

---



---

## 1. Overview, Goals & Success Metrics

### 1.1 Product summary

A web-based, in-house QC assistant that helps Prezlab designers audit and clean last-mile formatting on PowerPoint (.pptx) decks against a per-client or per-project profile (fonts, colors, margins, master layout). A designer uploads a deck, selects a profile, and receives a slide-by-slide issue log with severity, a structured change manifest, and exportable reports. The tool ships v1 (audit-first): detect, flag, report, and preview, then adds v1.5 (safe deterministic fix) for accept/reject auto-fixes, with v2 (hard fix + platform) extending to inference-heavy fixes and broader platform capabilities. It runs in the browser with Microsoft 365 sign-in, role-based access, and a REST API as its primary surface so Odoo and other M365 systems can call it. It enforces formatting consistency only; it does not make design judgments, render decks as the deliverable, or touch copy.

### 1.2 Problem statement

Designers lose meaningful, high-value time to manual formatting QC on the final stretch before delivery, and that QC is error-prone exactly where it matters most:

- **High slide counts.** Decks routinely run to 100 to 200 slides. Checking margins, fonts, palette, alignment, headers, and footers by eye across that volume is slow and inconsistent across reviewers.
- **Bilingual EN/AR complexity.** Arabic runs use a separate complex-script typeface and RTL direction. A careless font or alignment change silently breaks Arabic shaping, and manual review rarely catches it reliably.
- **High-stakes clients.** Recipients are consulting firms, investment banks, and government bodies across UAE, KSA, Jordan, and Qatar. A single formatting slip in a delivered deck carries reputational cost out of proportion to the effort that produced it.
- **No shared baseline.** Brand rules live in people's heads and in reference files, not in an enforceable, auditable profile, so consistency depends on who happens to do the final pass.

The result: senior design time spent on tedious, mechanical checking instead of design, with residual defects still reaching clients.

### 1.3 Goals

| # | Goal | Why it matters |
|---|------|----------------|
| G1 | Cut time spent on last-mile formatting QC per deck | Frees senior design capacity for design work |
| G2 | Catch more formatting defects than a manual pass, especially on 100+ slide bilingual decks | Reduces defects reaching high-stakes clients |
| G3 | Make brand/formatting rules explicit and enforceable via per-client/per-project profiles | Removes reliance on individual reviewers' memory |
| G4 | Produce a clear, exportable audit record (issue log, manifest, PDF) per deck | Supports pre-delivery review and accountability |
| G5 | Protect Arabic content by default during detection and any fix | Prevents silent glyph/shaping damage in bilingual decks |
| G6 | Earn designer trust through precision and non-destructive operation | First-release precision drives adoption in a 140-person shop |
| G7 | Integrate cleanly with the existing M365 + Odoo stack (SSO, RBAC, API) | Fits how Prezlab already works, no new silo |

### 1.4 Non-Goals

The tool enforces formatting consistency. It explicitly does not do the following.

**Mirrored from out-of-scope:**

- Redesigning layouts or deck structure.
- Creating new visual elements or icons.
- Writing or editing copy.
- Changing narrative flow or slide order.
- Custom infographic design.
- Motion or animation.

**Additional non-goals:**

- **Not a design-judgment engine.** It does not decide whether a layout is good, whether a color is the right creative choice, or whether a slide communicates well. It checks conformance to a defined profile, nothing more.
- **Not a deck renderer.** It is not a presentation viewer or a deck-to-PDF production tool. Any rendering is in service of QC artifacts (issue reports and before/after diffs), not a deliverable.
- **Not a copy author.** It never generates, rewrites, translates, or proofreads text content.

### 1.5 Target users

| Role | Who | Primary use |
|------|-----|-------------|
| Designer | In-house designers preparing client decks | Run audits and (later) apply fixes, review manifests, download cleaned .pptx and reports |
| Reviewer | Senior designers, QC leads, project leads | View manifests and reports, add comments; no profile edits |
| Admin | Design ops / brand owners, tool admins | Manage per-client/per-project profiles and user access |
| Machine caller | Odoo and other M365 systems | Trigger jobs and pull results via the REST API under an application identity (not a human role) |

Initial pilot: 2 to 3 senior designers as co-owners, positioning the tool as a pre-delivery QC assistant rather than an enforcement or policing system.

### 1.6 Success metrics

Targets are set during the pilot baseline and refined per phase. The metrics, not the targets, are fixed here.

| Metric | Definition | Why we track it |
|--------|------------|-----------------|
| Audit time per 100-slide deck | Median end-to-end audit time (upload to reviewable issue log) for a normalized 100-slide deck | Direct read on G1; processing must stay well within the 200-slide no-timeout requirement |
| Issues caught vs manual baseline | Defects flagged by the tool against defects found by an expert manual pass on the same decks | Direct read on G2; proves added detection value |
| False-positive rate per fix type | Share of flags or proposed fixes a reviewer rejects, tracked separately for each module/operation (e.g. font unify, palette, margin, alignment, shape, header/footer) | Precision gate; an auto-fix is enabled only when its false-positive rate clears threshold on the regression corpus |
| Designer adoption | Active designers running jobs per period as a share of eligible designers, plus repeat usage | Read on G6; adoption, not feature count, defines success |
| % decks passed through pre-delivery | Share of client-facing decks that go through the tool before delivery | Read on whether the QC step becomes standard practice |
| Defect escape rate (secondary) | Formatting defects found after delivery on decks that passed through the tool | Trailing quality signal; should fall as adoption and precision rise |
| Arabic-run incident count (secondary) | Confirmed cases of Arabic glyph/shaping damage attributable to a tool action | Hard guardrail on G5; target is zero |

Precision and false-positive rate are measured per fix type on a real-deck regression corpus and must be gated before any auto-apply is enabled for that fix type. Auto-apply is eligible only when a finding's `confidence` is `deterministic`, in line with the v1.5 (safe deterministic fix) phase.

### 1.7 Guiding principles

- **Audit-first.** Detection, flagging, reporting, and preview ship in v1 (audit-first) before any bulk auto-fix. We earn the right to change a file by first proving we read it correctly.
- **Original never overwritten.** The uploaded .pptx is immutable. Every output is a new artifact (cleaned .pptx, manifest, reports), so a bad run is never a lost deck.
- **Human-in-the-loop for judgment calls.** Only `deterministic`-confidence fixes are eligible for auto-apply in v1.5 (safe deterministic fix). Inference-heavy operations (nearest-color replacement, font-size hierarchy inference, master synthesis, alignment intent) stay flag-and-suggest behind a per-change accept/reject review, with a per-slide before/after diff as the real review surface, and graduate under v2 (hard fix + platform).
- **Arabic-safe by default.** Arabic content is detected, and direction-sensitive operations (alignment, margins, footer position) and font substitution on Arabic runs are skipped or hard-flagged as "Arabic content, manual review." The default for ambiguous Arabic cases is to do nothing and flag.
- **Confidentiality by default.** Client decks (including investment-bank and government work) are treated as highly sensitive from v1: access is gated behind M365 sign-in, uploads are auto-deleted after processing, and data is encrypted in transit and at rest. Data-handling, retention, and data-residency (KSA/UAE government rules) are confirmed with Operations and IT-security before launch.

---

## 2. Personas, Roles & User Flows

This product is audit-first. v1 ships detection, flagging, reporting, and preview only (Flow B). v1.5 adds safe deterministic auto-fix with per-change accept/reject and a visual before/after diff (Flow A). Roles and permissions below are designed for both phases, so v1.5 does not require a re-model of access control.

Access-control phasing: in v1, all authenticated users are treated as Designers (single-role behind the Entra ID sign-in gate). The per-role grants in Appendix A.7 and A.8 define the contract from day one but are enforced starting in v1.5, when `POST /v1/jobs/{job_id}/apply` introduces the first role-restricted operation. Role administration (assigning Entra app roles) is an Admin/IT task and gets its in-product surface with the v2 admin capabilities.

### 2.1 Personas

| Persona | Who they are | Primary goal | Why they care |
|---|---|---|---|
| Designer | In-house presentation designer producing client decks. Bilingual EN/AR. | Catch formatting inconsistencies before a deck goes to a client, then (v1.5) apply safe fixes without hand-editing each slide. | Speed and confidence pre-delivery. One QC pass instead of manual slide-by-slide checking. |
| Reviewer | Senior designer, design lead, or QA acting as a second set of eyes. May also be an account-side reviewer. | Read the audit, judge severity, leave comments, sign off. | Quality gate before client hand-off. No need to run or change anything, just assess. |
| Admin | Design ops / studio lead owning brand and client standards. | Maintain per-client and per-project formatting profiles; manage who can do what. | Profiles are the source of truth for what "correct" means per account. |

Adoption framing for all personas: this is a pre-delivery QC assistant, not an enforcement tool. The original .pptx is never overwritten, and (in v1.5) no change applies without an explicit per-change accept.

### 2.2 Roles and Permissions

Roles are driven by Microsoft Entra ID app roles delivered in the token `roles` claim (`Designer`, `Reviewer`, `Admin`). Machine callers (Odoo, Teams) authenticate via OAuth2 client-credentials as an application identity, authorized per scope, and are never mapped to a human role. The REST API is the primary surface and enforces these permissions server-side; the web UI is one client of it and must not be the only place a permission is checked.

The matrix below is the projection of the canonical role permission matrix (Appendix A.8) onto the user-facing capabilities, with the canonical endpoints from Appendix A.7. Where this section and the appendix differ, the appendix governs.

| Capability | Endpoint (canonical) | Designer | Reviewer | Admin | Machine |
|---|---|---|---|---|---|
| Upload .pptx | `POST /v1/uploads` | Yes (own) | No | Yes | Yes |
| Run audit (Flow B) | `POST /v1/jobs` with `mode=audit` | Yes (own) | No | Yes | Yes |
| Run fix pass (Flow A, v1.5) | `POST /v1/jobs` with `mode=fix` | Yes (own) | No | Yes | Yes |
| Cancel a job | `POST /v1/jobs/{job_id}/cancel` | Yes (own) | No | Yes | Yes |
| View job status | `GET /v1/jobs/{job_id}` | Yes | Yes | Yes | Yes |
| View findings / issue log | `GET /v1/jobs/{job_id}/findings` | Yes | Yes | Yes | Yes |
| View summary roll-up | `GET /v1/jobs/{job_id}/summary` | Yes | Yes | Yes | Yes |
| View outputs (report, diff, cleaned deck) | `GET /v1/jobs/{job_id}/outputs` | Yes | Yes | Yes | Yes |
| Apply selected fixes (v1.5) | `POST /v1/jobs/{job_id}/apply` | Yes | No | Yes | No |
| Add inline comment | `POST /v1/jobs/{job_id}/comments` | Yes | Yes | Yes | No |
| Sign off | `POST /v1/jobs/{job_id}/signoff` | No | Yes | Yes | No |
| Create/edit formatting profiles | `POST` / `PUT /v1/profiles/{id}` | No | No | Yes | No |
| View profiles (to select at submit) | `GET /v1/profiles` | Yes | Yes | Yes | No |
| Manage users / role assignment | Delegated to Entra app roles | No | No | Yes | No |

Permission notes:

- Reviewer is read-plus-comment-plus-signoff. A Reviewer can read every job's findings, summary, and outputs, can comment, and can sign off, but cannot create or cancel jobs, apply fixes, or edit profiles. This keeps the Reviewer a pure quality gate.
- Designer owns the working loop: upload, create and cancel own jobs, review, comment, and (v1.5) apply fixes. A Designer can comment but cannot sign off, and cannot edit profiles or manage users.
- Admin is a superset of Designer for job operations plus exclusive ownership of profiles and user management. Admin can also comment and sign off.
- Machine callers (the Odoo application identity) can upload, create, cancel, and read jobs and findings on behalf of an automated flow, but cannot comment, sign off, apply fixes, or manage profiles. They are authorized explicitly as a registered service principal and are not mapped to a human role.
- Confidentiality is a v1 control, not a later phase. Client decks (investment-bank and government) are highly sensitive, so every endpoint above sits behind the M365 sign-in gate; uploads auto-delete after processing per the configured retention window; data is encrypted at rest and in transit. An unauthenticated "internal-only" endpoint is not an acceptable control.

### 2.3 Flow B: Audit Only (v1 primary)

Flow B is the v1 primary flow. Audit reuses the full module pipeline with write disabled, so audit findings and (later) fix actions cannot drift apart. No cleaned .pptx is produced in this flow.

1. Upload. Designer drags a .pptx onto the drop zone or uses the file picker. Client-side guards: extension `.pptx`, size cap, slide-count cap (200). The file is streamed to object storage via `POST /v1/uploads`, which returns `{ upload_id, expires_at }`; the UI shows the file name and slide count read from the package.
2. Pre-flight. The tool inspects the package and surfaces blocking and advisory notices before any work: a "no usable master" guard (non-conformant Canva/Gamma/Google Slides/Keynote exports), and a mandatory "contains SmartArt / advanced charts / embedded media we will not modify" notice. Arabic content is detected here (Unicode ranges plus `a:pPr@rtl`) and runs are tagged so direction-sensitive checks self-guard later, surfacing as records with `arabic_flag = true`.
3. Select profile. Designer picks a per-client or per-project formatting profile (fonts, colors, margins, header/footer rules) from `GET /v1/profiles`. If none fits, the deck can be audited against the studio default profile.
4. Choose scope. Full audit (all six modules) or targeted module(s) by `ModuleKey`: `master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer`.
5. Create job and enqueue. `POST /v1/jobs` with `{ upload_id, profile_id, modules[], mode: "audit", preview }` returns a `job_id` in state `queued`. The job is placed on the async queue; the upload connection is not held open.
6. Process. A worker pulls the job (`processing`) and runs the enabled modules through the shared inheritance resolver (font size/color, theme refs, tint/shade math, schemeClr resolution). A per-slide progress counter streams to the client via SSE or polling (`GET /v1/jobs/{job_id}`, reading `slides_processed` / `slides_total`). Each module emits structured `FindingRecord` entries: `{ record_id, job_id, slide_index, shape_id, shape_path, module, issue_type, property, old_value, new_value, severity, action: "flagged", confidence, arabic_flag, profile_rule_id, message }`.
7. Job states. The client renders the live `JobStatus` for the job:

   | State | Meaning | UI affordance |
   |---|---|---|
   | `queued` | Accepted, awaiting a worker | Spinner, position if available |
   | `processing` | Worker running modules | Per-slide progress (e.g. 87/200) |
   | `ready` | Audit complete, findings available | Open issue log / preview |
   | `failed` | Pipeline error or pre-flight block | Reason + retry / re-upload |
   | `cancelled` | Job cancelled before completion | Re-create or re-upload |

8. Review flagged issues. On `ready`, the Designer opens the issue log: slide-by-slide findings with `severity` (`error` / `warning` / `info`) and a summary roll-up per `issue_type` and per module (`GET /v1/jobs/{job_id}/summary`). Filters: by slide, by module, by severity, and by "Arabic content, manual review" (records carrying `arabic_flag = true`, where direction-sensitive operations and font substitution on Arabic runs are flagged, never silently changed). The preview step here is read-only: the deck is shown with flagged shapes highlighted in place; no changes exist to diff yet.
9. Comment (optional inline-comment mode, ships in v2). A Designer, Reviewer, or Admin can attach inline comments anchored to a `slide_index` and optional `record_id` (`POST /v1/jobs/{job_id}/comments`). Reviewers use this to direct the Designer; Designers use it to note intentional deviations. In v1 and v1.5 this step is out-of-band (review meeting or Teams thread).
10. Sign-off (optional, ships in v2). A Reviewer or Admin can mark the audit resolved (`POST /v1/jobs/{job_id}/signoff`) once findings are accepted or dispositioned. Designers cannot sign off. In v1 and v1.5, sign-off is out-of-band.
11. Export. Designer or Reviewer retrieves the report via `GET /v1/jobs/{job_id}/outputs`, which returns signed URLs for `report_pdf`, `report_csv`, and `manifest_json`. The v1 report is report-only (issue logs, severity, counts, summaries, no slide thumbnails) and is rendered from the stored `FindingRecord` set. The CSV export is rendered from the same record set.
12. Cleanup. The uploaded .pptx auto-deletes per the retention setting; the findings and report are retained per policy.

### 2.4 Flow A: Format & Deliver (v1.5)

Flow A is the v1.5 flow. It is Flow B plus a write phase. It runs the identical pipeline with writing enabled, applies only high-confidence deterministic fixes by default, and gates every applied change behind a per-change accept/reject and a visual before/after diff. The original is never overwritten; the cleaned deck is a new output.

1. Upload, pre-flight, select profile, choose scope. Identical to Flow B steps 1 to 4. Pre-flight blocks and Arabic guards carry over unchanged.
2. Create job and enqueue. `POST /v1/jobs` with `{ upload_id, profile_id, modules[], mode: "fix", preview }` returns a `job_id` in `queued`.
3. Process (audit + propose). The worker runs the modules (`processing`) and, in addition to flagging, computes proposed changes. Each module classifies every operation as deterministic remediation (safe to auto-apply: e.g. exact-match shape sizing where `prstGeom` prst, ext within threshold, and placeholder idx/role all match; theme-color refs already valid by name) versus inference-heavy remediation (nearest-color replacement, font-size hierarchy inference, master synthesis, alignment intent) which is proposed as suggest-only. Records carry `action: "changed"` (applied, deterministic) or `action: "flagged"` (suggested, requires human decision) plus `confidence`. Auto-apply is eligible only when `confidence = deterministic`.
4. Job states. Same `JobStatus` values as Flow B (`queued`, `processing`, `ready`, `failed`, `cancelled`). `failed` additionally covers a write/round-trip validation failure: the worker validates that the rewritten package opens cleanly and preserves unmapped parts (SmartArt, media, OLE) before marking `ready`.
5. Review findings with visual diff. On `ready`, the Designer opens the change review surface. The primary surface is the per-slide visual before/after diff, not the text record list. For each proposed change the Designer sees: `old_value`, `new_value`, module, `severity`, `confidence`, and the affected shape highlighted in both states.
6. Per-change accept/reject. The Designer accepts or rejects each change individually, then applies the accepted set via `POST /v1/jobs/{job_id}/apply` with `{ record_ids[] }`, which returns a new job/output. Deterministic changes default to accepted but remain reversible; suggest-only changes default to rejected and require an explicit accept. Bulk accept/reject is available per module and per severity, but the unit of record is always the individual `FindingRecord`. Arabic-tagged runs (`arabic_flag = true`) cannot be auto-accepted for direction-sensitive or font-substitution changes; they stay manual-review and surface with `action: "skipped"`.
7. Comment (optional inline-comment mode, ships in v2). Same as Flow B: comments anchor to `slide_index` and optional `record_id` via `POST /v1/jobs/{job_id}/comments`. Out-of-band until v2.
8. Finalize and export. Once `POST /v1/jobs/{job_id}/apply` completes, the worker writes the cleaned .pptx surgically (touching only the XML for accepted changes, with brand and Arabic fonts embedded). Outputs are retrieved via `GET /v1/jobs/{job_id}/outputs`: signed URLs for `cleaned_pptx`, `report_pdf`, `report_csv`, and `manifest_json`. The records reflect final per-change decisions.
9. Download. Designer downloads the cleaned .pptx plus PDF from the signed `/outputs` URLs. The original remains intact in storage until retention deletes it.
10. Cleanup. Upload and intermediate artifacts auto-delete per the retention setting; findings and final report are retained per policy.

### 2.5 Cross-flow rules

- Single pipeline. Audit (Flow B, `mode=audit`) and fix (Flow A, `mode=fix`) are the same code path with writing toggled. This guarantees that what was flagged in audit is exactly what is proposed for fix.
- FindingRecord is authoritative. Reports (PDF/CSV), the `manifest_json`, and the diff UI are all projections of the stored `FindingRecord` set, never reconstructed by diffing two .pptx zips (a full open/save round-trip rewrites untouched XML and produces noisy false diffs).
- Original is immutable. No flow overwrites the source file. Flow A always emits a new output (`cleaned_pptx`).
- Precision before breadth. Auto-apply is gated per fix type on a real-deck regression corpus and enabled only when its false-positive rate clears threshold and `confidence = deterministic`. A single mangled client deck is an adoption risk in a 140-person studio, so first-release precision outranks feature count.

---

## 3. Functional Requirements: Modules 1-3

This section specifies the first three audit/remediation modules. All three depend on two shared services defined once and reused: the **Inheritance Resolver** (run rPr to paragraph/list level to txBody list style to placeholder-on-layout to placeholder-on-master to master text styles to theme/presentation defaults) and the **Transform Resolver** (applies rotation, flips, text insets, and group `chOff`/`chExt` to `off`/`ext` mapping before any cross-shape or cross-slide comparison). Each module emits one `FindingRecord` per detected issue or applied change to the shared record set (the system of record defined in Appendix A.2). The canonical record carries, among other fields, `record_id`, `slide_index` (zero-based), `shape_id` (string), `shape_path`, `module` (a `ModuleKey`), `issue_type` (a dotted code from A.3), `property`, `old_value`, `new_value`, `severity` (a `Severity`), `action` (an `Action`), `confidence` (a `Confidence`), `arabic_flag`, `profile_rule_id`, and `message`. Here `action ∈ {flagged, changed, skipped}` and `severity ∈ {error, warning, info}`. The change manifest, report, and UI surfaces are all projections of this record set, never reconstructed by diffing `.pptx` files. Audit mode is the same pipeline with writes disabled, so audit and fix cannot drift.

A guiding constraint across all three modules: **detection is mostly objective and ships first; remediation is often a judgment call and is gated behind preview and per-change accept/reject.** Auto-apply (v1.5) is reserved for fixes whose record carries `confidence = deterministic`; the `Confidence` enum is `deterministic | high | medium | low`.

---

### 3.1 Module 1: Master Slide Engine

ModuleKey: `master_slide`.

#### Purpose
Enforce structural consistency by detecting the dominant slide layout in a deck, flagging slides that deviate, and (later) re-binding outliers to an explicitly chosen existing master/layout. The module does not invent new masters or redesign layouts.

#### Detection rules

| Rule | `issue_type` | What it flags | OOXML / python-pptx specifics |
|---|---|---|---|
| Dominant-layout detection | `master_slide.layout_outlier` | Computes the modal `slide_layout` across the deck; slides not on the dominant (or profile-pinned) layout are flagged as outliers. | `Slide.slide_layout` is read-only and resolved via the slide-to-layout relationship. Group by layout `rId` / layout name. |
| Placeholder conformance | `master_slide.placeholder_geometry_off` | Slides whose populated placeholders do not match the chosen layout's placeholder set (missing expected idx, extra orphan shapes, or content in a non-placeholder where a placeholder exists). | Placeholders bind by `ph idx` in `spPr`/`nvSpPr`. Compare slide `ph idx` set against layout `ph idx` set. |
| Off-layout geometry | `master_slide.placeholder_geometry_off` | Title/body/content shapes whose position or size diverge materially from the layout's placeholder geometry. | Resolve placeholder `off`/`ext` from layout via inheritance; apply Transform Resolver before comparing. |
| No usable master (defensive guard) | `master_slide.no_usable_master` | Decks with absent or non-conformant master/layout structure (common in Canva, Gamma, Google Slides, Keynote exports). | Every conformant .pptx has a master; third-party exports may not. Detect missing/broken master relationships and hard-flag. |

#### Remediation behavior

| Operation | Classification | v1 | v1.5 | v2 |
|---|---|---|---|---|
| Detect dominant layout + flag outliers | Detection (objective) | Flag | Flag | Flag |
| Enforce an explicitly chosen existing master/layout | Remediation (inference-heavy: intent + content remap) | Flag only | Flag only | Auto-fix behind preview + per-change accept/reject |
| Auto-invent a master when none exists | Out of scope | Cut | Cut | Cut |

**v1 scoping (enforce-existing-only):** the module never synthesizes a master in v1 or v1.5; it only enforces an existing, explicitly chosen master/layout. This is pinned in the profile as `master_slide.enforce_existing_only = true`.

**Master re-apply limitation (verified):** python-pptx has **no API to re-apply or change the layout of an existing populated slide**. `Slide.slide_layout` is read-only and a raw relationship swap does not cleanly remap placeholder content (placeholders bind by idx, and a swap leaves content orphaned). Therefore master enforcement is **flag-first in v1**. The v2 hard auto-fix requires either a verified surgical XML remap (placeholder idx re-mapping per shape) or the Aspose.Slides engine path, and must be retired in the technical spike (retire item: "master enforcement without corruption") before the timeline commits to it.

#### Configuration inputs (from profile, `config.master_slide`)
- `master_slide.enforce_existing_only` (bool): v1/v1.5 always true; never synthesize a master.
- `master_slide.pinned_layout_id` (string | null): explicit master/layout to enforce, overriding modal detection as the conformance target.
- `master_slide.layout_allowlist` (string[]): permitted layout ids; others are flagged as outliers.
- `master_slide.geometry_tolerance_emu`: tolerance for off-layout placeholder geometry flags.

#### Severity assignment
- **Error:** no usable master / non-conformant deck (`master_slide.no_usable_master`); slide on a layout outside `layout_allowlist`; missing expected title/body placeholder.
- **Warning:** off-layout placeholder geometry within a soft band; orphan shapes; minority-layout outlier where no layout is pinned.

#### Edge cases & guards
- **No usable master:** never attempt enforcement; emit `issue_type = master_slide.no_usable_master`, `action = skipped`, `severity = error`, `message = "No usable master, manual review."`
- **Grouped/rotated shapes:** apply Transform Resolver before comparing placeholder geometry; do not flag rotation-induced bounding-box changes as off-layout.
- **Inheritance:** resolve placeholder geometry through layout/master before comparison; a slide inheriting position from the layout is conformant even when `off`/`ext` are absent on the slide shape.
- **Arabic:** master enforcement is structural, not direction-sensitive, so layout flagging proceeds; do not infer footer/margin direction here (that belongs to Modules 3 and 6).
- **Master re-apply limitation:** in v1/v1.5 the module must not perform relationship swaps; enforcement records carry `action = flagged` only.

#### Phase
- **v1 (audit):** dominant-layout detection, outlier flagging, placeholder conformance, no-usable-master guard.
- **v1.5 (safe fix):** none for hard enforcement; deterministic geometry nudges of placeholders back to layout values may ship behind preview if the spike confirms safety, and only when the record carries `confidence = deterministic`.
- **v2 (hard fix):** master/layout enforcement on existing slides via verified surgical remap or Aspose engine.

---

### 3.2 Module 2: Font Auditor & Fixer

ModuleKey: `font`.

#### Purpose
Unify font family, size, and weight against the profile's typographic system, and flag every run that deviates. This is the module most likely to silently destroy content (Arabic shaping, theme-font references), so detection ships broadly and remediation is tightly gated.

#### Detection rules

| Rule | `issue_type` | What it flags | OOXML / python-pptx specifics |
|---|---|---|---|
| Off-brand family | `font.family_out_of_set` | Latin runs whose effective typeface is not in the profile font set. | Resolve via Inheritance Resolver; `run.font.name` touches only `a:latin`. Theme fonts are `+mn-lt` / `+mj-lt` references and must be resolved to the theme, not treated as literal mismatches. |
| Off-system size | `font.size_off_role` | Runs whose effective point size is outside the role's allowed size for its role. | `run.font.size` returns `None` when inherited; never read it raw. Resolve effective size through the full chain. |
| Weight / style mismatch | `font.mixed_weight` | Bold/italic that contradicts the profile's role styling. | `b`, `i` attributes on `rPr`, resolved through inheritance. |
| Arabic typeface (complex script) | `font.cs_typeface_missing` | Arabic runs whose complex-script font differs from the profile's complex-script font. | Arabic uses `a:cs` (complex script), **distinct from `a:latin`**. python-pptx `Font.name` does not touch `cs`; detection and any future fix require raw XML on `a:cs`. |
| Disallowed theme reference | `font.theme_ref_disallowed` | `+mn-lt`/`+mj-lt` references when theme references are not permitted by the profile. | Read `font.theme_font_refs_allowed`; if false, a theme reference is reported rather than resolved as on-brand. |
| Mixed fonts within a frame | `font.mixed_weight` | Text frames containing multiple effective families/sizes where the role expects one. | Iterate runs; compare resolved values per run. |
| Overflow (advisory) | `font.size_off_role` | Text that likely overflows its frame at the resolved font/size. | `fit_text()` / TextFitter approximate overflow for uniform font/size frames; requires the font file installed, uses integer point sizes, and PIL metrics diverge from PowerPoint. In-process flag only; renderer validates high fidelity. |

#### Remediation behavior

| Operation | Classification | v1 | v1.5 | v2 |
|---|---|---|---|---|
| Flag off-brand family / size / weight | Detection (objective) | Flag | Flag | Flag |
| Exact family swap to brand font (Latin runs only, deterministic 1:1) | Remediation (deterministic) | Flag only | Auto-fix with per-change accept/reject | Auto-fix |
| Size-hierarchy inference (infer intended title/subtitle/body and resize) | Remediation (inference-heavy) | Flag only | Flag / suggest | Suggest behind preview |
| Arabic (`a:cs`) font substitution | Remediation (high-risk) | Flag only | Flag only | Behind preview + Arabic regression gate |
| Embed brand + complex-script fonts in cleaned .pptx | Output hygiene | n/a | On write | On write |

**Guard on substitution:** an RTL-unaware font swap to a Latin-only brand font destroys Arabic glyphs and shaping silently. Latin-only exact-match swaps are the only auto-apply candidate; they carry `confidence = deterministic` and ship only after the precision/false-positive gate on the real-deck regression corpus passes for this fix type.

#### Configuration inputs (from profile, `config.font`)
- `font.roles.{role}.latin[]` (role ∈ {title, subtitle, body, caption}): allowed Latin families per role, primary plus alternates.
- `font.roles.{role}.complex_script[]`: allowed complex-script (`a:cs`) families per role.
- `font.roles.{role}.size_pt`: target point size per role.
- `font.roles.{role}.allowed_weights[]`: permitted weights per role.
- `font.theme_font_refs_allowed` (bool): treat `+mn-lt`/`+mj-lt` as on-brand when they resolve to a profile font.
- `font.size_tolerance_pt`: allowed deviation from the role's `size_pt` before flagging.

#### Severity assignment (per A.4)
- **Error:** off-brand family on a non-Arabic run (`font.family_out_of_set`).
- **Warning:** size off role (`font.size_off_role`) beyond `font.size_tolerance_pt`; weight / style mismatch (`font.mixed_weight`); likely overflow.
- **Arabic-guarded (skipped):** every Arabic-run finding (never auto-fixed in v1/v1.5) is emitted with `severity = warning`, `action = skipped`, and `arabic_flag = true`, message "Arabic content, manual review." These are not `info`.

#### Edge cases & guards
- **Arabic:** detect via Unicode ranges (U+0600-06FF, U+0750-077F, U+08A0-08FF, presentation forms U+FB50-FDFF / U+FE70-FEFF) and/or `rtl='1'` on `a:pPr`. Mark all Arabic-run findings with `arabic_flag = true`, `severity = warning`, `action = skipped`, message "Arabic content, manual review," and **never substitute fonts on Arabic runs** before v2. A run mixing Latin and Arabic is treated as Arabic-bearing for guard purposes.
- **Inheritance:** all family/size/weight reads go through the Inheritance Resolver; `None` from python-pptx means inherited, not absent. Do not flag a run as off-brand when its resolved value is on-brand via theme/master.
- **Theme fonts:** `+mn-lt`/`+mj-lt` references are resolved before comparison when `font.theme_font_refs_allowed = true`; do not report a theme reference as a mismatch when it resolves to a profile font. When theme references are disallowed, emit `font.theme_ref_disallowed`.
- **Overflow:** treated as advisory only; never auto-resizes text in v1/v1.5.
- **Grouped/rotated shapes:** font reads are run-level and unaffected by transform, but overflow approximation should account for the resolved frame extent after Transform Resolver.

#### Phase
- **v1 (audit):** family/size/weight detection, Arabic detection + guard, theme-font resolution, advisory overflow flag.
- **v1.5 (safe fix):** Latin-only exact-match family swap with per-change accept/reject (`confidence = deterministic`); brand + complex-script font embedding on write.
- **v2 (hard fix):** size-hierarchy inference, Arabic `a:cs` substitution behind preview and the Arabic regression gate.

---

### 3.3 Module 3: Margin & Alignment Engine

ModuleKey: `margin_alignment`.

#### Purpose
Detect content that violates the safe-zone margins and flag misalignment between shapes, then (later, behind preview) snap to guides and align similar groups. Detection of "outside safe-zone" is reliable; alignment auto-fix is heuristic and high-risk, so it is flag-first throughout.

#### Detection rules

| Rule | `issue_type` | What it flags | OOXML / python-pptx specifics |
|---|---|---|---|
| Outside safe-zone margin | `margin_alignment.outside_safe_zone` | Shapes whose bounding BOX crosses the profile safe zone, on any of the four sides. | `left`/`top`/`width`/`height` are exact EMU. Compare box edges against `geometry.safe_zone_margins_emu`: `left`/`top` against the near margin, `left + width` / `top + height` against the far one. Text insets (`bodyPr` `lIns`/`tIns`/`rIns`/`bIns`) are deliberately NOT applied — see the guard below. Rotation excludes the shape (stored `off`/`ext` do not describe a rotated bbox). Reliable detection. |
| Heading past the margin | `margin_alignment.heading_past_margin` | A title or subtitle whose BOX crosses a safe-zone margin by more than 1 mm. | Flag-only and permanently outside `FIXABLE_ISSUES`: whether a heading may break the frame is the client's house style, not a defect (design lead, 19/08/2026). Emitted instead of `outside_safe_zone` for the same shape, and the shape is excluded from `content_overflow`'s rescale selection and from the cross-slide `recurring_off_position` pin, so no approved fix can reposition it. Heading identity comes from the title/subtitle placeholder role, with a bounded largest-type fallback for placeholder-free export decks (`qc.util.heading_ids`). |
| Content in the reserved header band | `margin_alignment.body_band_intrusion` | Any element — picture or text — standing in the strip a master reserves under its subtitle, between the floor the subtitle may not cross and the ceiling the body may not cross (`geometry.body_band_emu`, read from the master's two interior horizontal guides: `qc.stylespec.read_content_band`). | One record per SLIDE, cleared by ONE vertical translate of the whole body block, so the arrangement inside the body never changes. Deliberately not text-only: the strip is a statement about the page, so a photo in it is the same defect as a paragraph. Excluded from the block: headings, placeholders, full-bleed and page-deep panels (`qc.util.full_height_panel`), the bottom furniture strip, and header-zone boxes that merely overhang the floor on descender slack. Both guides must be stated; with only a ceiling there is no way to tell an eyebrow from body content that crept up. Severity `error` from 3 mm deep, `warning` below; never pre-ticked (`qc.fixer.tick_reason`) because moving every element on a slide is the designer's call. A block that cannot come down without leaving the canvas is reported flag-only. |
| Edge / center misalignment | `margin_alignment.edge_misaligned` | Shapes that nearly share an edge or center axis but are off by a small EMU delta. Judged on COLLECTIONS: a shape that rides a larger one (a corner rule across a photo, a caption under it, a chip stacked on its box) is compared to how peers sit in THEIR anchors, never against the absolute clusters, and any fix translates the whole collection so the composition inside it cannot change (`qc.util.rides_with`, `qc.fixer._carried_contents`). Membership is judged ACROSS the axis of the move: shapes along it are the spacing peers being corrected, not collection members. | Compare resolved edges/centers across shapes after transforms; threshold-banded by `geometry.alignment.edge_tolerance_emu` / `center_tolerance_emu`. |
| Inconsistent gutters / spacing | `margin_alignment.uneven_spacing` | Repeated elements (e.g., card rows) with uneven inter-shape spacing. | Compute pairwise gaps from resolved geometry; band by `geometry.alignment.spacing_tolerance_emu`. |
| Group coordinate space | `margin_alignment.edge_misaligned` | Child shapes inside a `groupShape` compared without transform normalization. | Children live in a transformed space (`chOff`/`chExt` mapped to `off`/`ext`); transforms must be applied before any cross-slide or cross-shape comparison, or alignment readings are wrong. |

#### Remediation behavior

| Operation | Classification | v1 | v1.5 | v2 |
|---|---|---|---|---|
| Flag outside-safe-zone | Detection (objective, reliable) | Flag | Flag | Flag |
| Flag edge/center misalignment | Detection (objective within threshold) | Flag | Flag | Flag |
| Snap to guides | Remediation (heuristic, risky) | Flag only | Suggest behind preview | Auto-fix behind preview + accept/reject |
| Align similar groups | Remediation (heuristic: requires alignment-intent inference) | Flag only | Suggest behind preview | Auto-fix behind preview + accept/reject |

No alignment operation auto-applies without preview and per-change accept/reject. Alignment intent is inference-heavy, never reaches `confidence = deterministic`, and stays flag/preview.

#### Configuration inputs (from profile, `config.geometry`)
- `geometry.safe_zone_margins_emu` (left/right/top/bottom): the safe-zone margins compared against effective shape edges. Projected from the best statement the master makes, in this order (`style_spec_source.grid_source` records which): the **presentation space**, a rectangle the designer draws on the master and names "Presentation space"; then the drawing **guides**; then the master's content **placeholders**. The rectangle outranks the guides because it needs no interpreting - a master can carry an outer page margin, a column grid, a header band and a bleed line, and choosing among them is a guess (design lead, 21/08/2026: "some cases have multiple margins, so presentation space is safer"). The reserved header band is still read from the guides: one rectangle cannot state a pair of lines inside itself.
- `geometry.body_band_emu` (`subtitle_floor`, `body_top`): the reserved strip under the header. Carried only when the master DREW it — there is no default, because a guessed ceiling is a line the client never drew. Absent from a profile, the module falls back to the guides on the deck's own master, which is the same source the profile is projected from.
- `geometry.alignment.edge_tolerance_emu` / `geometry.alignment.center_tolerance_emu`: bands for edge/center misalignment flags.
- `geometry.alignment.spacing_tolerance_emu`: gutter consistency band.

#### Severity assignment (per A.4)
- **Warning:** element outside the safe zone (`margin_alignment.outside_safe_zone`); edge/center misalignment within the flag band (`margin_alignment.edge_misaligned`); inconsistent gutters (`margin_alignment.uneven_spacing`). All margin, alignment, and spacing deviations default to `warning`.

#### Edge cases & guards
- **The box, not the text (design lead, 19/08/2026):** a margin is measured against the shape's stored `off`/`ext` box, on all four sides, and text insets (`bodyPr lIns/tIns/rIns/bIns`) are NOT normalized into it. Insets, autofit scale and glyph extent are rendering, and rendering varies with the font actually installed; the box is the only edge the deck itself states. A line of text spilling out of a box that sits correctly is a copy-length conversation for the designer and the client, not a margin breach, and this module must not estimate glyph widths to invent one. (This supersedes the earlier requirement to account for insets.)
- **Rotation / flips:** never compare raw `off`/`ext` on a rotated or flipped shape. Apply Transform Resolver (rotation, `flipH`/`flipV`); where rotation cannot be resolved confidently, emit an informational skip rather than a mis-normalized flag.
- **Grouped shapes:** normalize child coordinates from group space (`chOff`/`chExt` to `off`/`ext`) before measurement; report at the child `shape_id` and preserve the group relationship via `shape_path` in the record.
- **Arabic (direction-sensitive guard):** margins, alignment, and footer position are direction-sensitive. On slides/runs detected as Arabic (same detection signals as Module 2, including `rtl='1'` on `a:pPr`), guard margin/alignment remediation: emit the record with `severity = warning`, `action = skipped`, and `arabic_flag = true`, message "Arabic content, manual review." LTR safe-zone math does not transfer to RTL content. Detection-only flags may still be emitted with `arabic_flag = true`.
- **Inheritance:** position can be inherited from the placeholder/layout; resolve effective geometry through the chain before flagging, consistent with Module 1.

#### Phase
- **v1 (audit):** outside-safe-zone detection, edge/center misalignment flags, gutter consistency, group-space normalization, Arabic direction guard.
- **v1.5 (safe fix):** none auto-applied; snap-to-guides and align-similar-groups available as preview suggestions only.
- **v2 (hard fix):** snap-to-guides and align-similar-groups auto-fix behind preview with per-change accept/reject.

---

## 4. Functional Requirements: Modules 4-6

This section specifies the Color Palette Checker, Shape & Size Normalizer, and Header & Footer Consistency modules, keyed as `color_palette`, `shape_size`, and `header_footer`. All three share the inheritance/color resolver and the per-shape geometry resolver defined in Section 3, emit one `FindingRecord` per detected issue or applied change (canonical fields: `record_id`, `job_id`, `slide_index` zero-based, `shape_id` as string, `shape_path`, `module`, `issue_type`, `property`, `old_value`, `new_value`, `severity`, `action`, `confidence`, `arabic_flag`, `profile_rule_id`, `message`, `created_at`), and run the same pipeline in `audit` mode (write disabled) and `fix` mode (write enabled) so audit and remediation cannot drift.

A guiding rule across all three: **detection is mostly objective; remediation is frequently a judgment call.** We auto-apply only fixes whose `confidence = deterministic` (eligible in v1.5) and route inference-heavy operations (nearest-color replacement, fuzzy shape alignment, footer text rewrites) to suggest-only behind the per-change accept/reject preview. Inference-heavy suggestions carry `confidence` of `high`, `medium`, or `low`, never `deterministic`.

### 4.1 Module 4: Color Palette Checker

#### Purpose

Audit every color-bearing property against the active profile's palette and flag off-palette colors. The goal is to catch colors that are not part of the approved brand set (stray RGB values pasted from other decks, near-miss tints, default Office colors) and, where confident, propose the nearest on-palette color. The hard constraint is correct color resolution: most brand colors in a real deck are theme references with modifiers, not literal hex, so a naive RGB read produces false positives on perfectly compliant slides.

#### Detection rules

Color resolution is the core of this module. python-pptx does **not** resolve theme references or tint/shade math to final RGB, so we resolve in our own layer.

| Step | Rule |
|---|---|
| Enumerate sources | For each shape, resolve fill (`a:solidFill`), line (`a:ln`), and text run color (`a:solidFill` inside `rPr`). Group, picture, gradient, and pattern fills are handled as guards (below). |
| Resolve scheme refs | A color expressed as `a:schemeClr val="accent1"` (etc.) is resolved against `theme1.xml` `a:clrScheme` to a base RGB. Map the theme slots in `color_palette.theme_color_slots` (`dk1`, `lt1`, `dk2`, `lt2`, `accent1` through `accent6`), including the `tx1`/`tx2`/`bg1`/`bg2` aliases. |
| Apply modifiers | Apply child transforms in document order: `a:lumMod`, `a:lumOff`, `a:tint`, `a:shade`, `a:satMod`, `a:alpha`. Tint/shade math is computed to produce the effective RGB; this is what the eye sees and what we compare against `named_colors[].allowed_tints` and `named_colors[].allowed_shades`. |
| On-palette test (two-mode) | A color is on-palette if EITHER (a) **by-name** (`on_palette_mode = "by_name"`): it is a valid theme-color reference (`schemeClr`) whose slot is in `theme_color_slots`, regardless of tint/shade (recommended default, lowest false-positive rate), OR (b) **by-resolved-RGB** (`on_palette_mode = "by_resolved_rgb"`): its resolved RGB is within `match_tolerance_deltaE` of an allowed `named_colors[]` entry (matched on `hex`). The profile selects the mode. |
| Off-palette flag | A literal `a:srgbClr` (or `a:prstClr`) not matching any allowed entry under the active mode raises `color_palette.off_palette_rgb`. A `schemeClr` whose slot is not in `theme_color_slots` raises `color_palette.disallowed_theme_slot`; a `schemeClr` in an allowed slot is never flagged regardless of modifier. A resolved tint/shade outside `named_colors[].allowed_tints`/`allowed_shades` raises `color_palette.tint_out_of_range`. |
| Nearest-match suggestion | Distance is computed in **CIEDE2000 (Lab)**, never RGB Euclidean. The suggested replacement is the allowed `named_colors[]` entry with the smallest ΔE00 to the resolved RGB. The ΔE00 value is recorded on the record and drives `confidence`. A nearest-color match is always a suggestion, carrying `confidence` of `high` or `medium`, never `deterministic`. |

#### Remediation

| Operation | Mode | Rationale |
|---|---|---|
| Off-palette literal RGB within `auto_replace_max_deltaE` of exactly one allowed `named_colors[]` entry | **Auto (v1.5)** | Eligible for auto-apply only when `confidence = deterministic` and ambiguity is near zero (a clear near-miss tint). Written as a `schemeClr` ref when the matched entry has a `theme_ref`, preserving downstream theme behavior. |
| Off-palette color where the nearest match exceeds `auto_replace_max_deltaE` (within `ambiguity_band_deltaE`), or two allowed entries fall within the ambiguity band of each other | **Suggest only** | Nearest-color replacement is inference; the designer chooses. Top 1-3 candidates are presented with ΔE00, at `confidence` `high`/`medium`. |
| `schemeClr` in a disallowed slot (`color_palette.disallowed_theme_slot`) | **Suggest only** | Re-mapping a theme slot can cascade across the deck; it requires intent. |
| Any color on a gradient/pattern/picture fill, or any color on an Arabic run | **Flag only, never auto** | See guards. |

No remediation runs in v1 (audit-only). Color resolution, on-palette testing, nearest-match computation, and the report ship in v1; auto-replace of deterministic near-misses ships in v1.5 behind per-change accept/reject.

#### Configuration inputs (per profile)

Read from `color_palette` in the profile config:

- `theme_color_slots`: list of permitted theme slots (`dk1`, `lt1`, `dk2`, `lt2`, `accent1` through `accent6`).
- `named_colors[]`: list of `{name, hex, theme_ref, allowed_tints, allowed_shades}` brand colors for by-resolved-RGB mode and for nearest-match candidates. `hex` is `#RRGGBB`; `theme_ref` ties a named color to a theme slot when applicable; `allowed_tints`/`allowed_shades` bound the permitted modifier range.
- `on_palette_mode`: `by_name` or `by_resolved_rgb`.
- `match_tolerance_deltaE`: ΔE00 within which a resolved color counts as on-palette (CIEDE2000).
- `auto_replace_max_deltaE`: ΔE00 ceiling at or below which a near-match yields a safe suggestion (and, in v1.5, a deterministic auto-replace).
- `ambiguity_band_deltaE`: beyond `auto_replace_max_deltaE` and within this band, flag but do not suggest a replacement.

#### Severity

Per Appendix A.4:

| Condition | Severity |
|---|---|
| Literal RGB off-palette with no near match (beyond the ambiguity band) | `error` |
| Off-palette but within the ambiguity band | `warning` |
| `schemeClr` in a disallowed theme slot | `warning` |
| Resolved tint/shade out of allowed range | `warning` |
| Color on gradient/picture/pattern fill (unresolvable) | `warning` (informational, manual review) |

#### Edge cases & guards

- **Gradient / pattern / picture fills:** no single color exists. Do not attempt resolution or replacement. Emit one informational record per shape (`property: "fill.type"`, value = `gradient`/`pattern`/`blipFill`) tagged "manual review," never an auto-fix candidate.
- **Inherited color:** when a run's color is `None` it is inherited; resolve through the shared inheritance chain before testing, otherwise inherited-but-compliant text is falsely flagged.
- **Theme-aliased slots:** treat `tx1`/`dk1`, `bg1`/`lt1` (and dark/light pairs) consistently so identical effective colors are not double-counted.
- **Arabic runs:** the color audit itself is direction-agnostic and may run, but any color change on an Arabic run is suppressed from auto-apply, recorded with `arabic_flag = true` and `action = skipped`, and marked "Arabic content, manual review" to keep this module aligned with the global Arabic guard. Color is never the trigger for a font change.

#### Phase

| Capability | Phase |
|---|---|
| Resolution, on-palette test, nearest-match report | v1 (audit) |
| Deterministic near-miss auto-replace, per-change accept/reject | v1.5 |

### 4.2 Module 5: Shape & Size Normalizer

#### Purpose

Detect shapes that should be identical in geometry but are not, and unify their size (and, where deterministic, position) so repeated elements (cards, KPI tiles, agenda rows, icon frames) line up. The core distinction, drawn directly from the geometry findings, is **exact-match similarity (deterministic, safe to auto-apply) versus fuzzy "looks similar" (heuristic, suggest-only).** We are conservative here because a single mangled client deck kills adoption.

#### Detection rules

All geometry is read as exact EMU but resolved through the shared geometry resolver before any cross-shape comparison: rotation, h/v flips, text insets (`bodyPr lIns/tIns/rIns/bIns`), and group transforms (`chOff/chExt` to `off/ext`) are applied first. Comparing raw `off`/`ext` without this resolution is invalid for grouped or rotated shapes.

**Exact-match cohort (deterministic).** Two shapes are exact-match peers when ALL hold:

1. Same `prstGeom@prst` (same preset geometry), or both custom geometry with identical path.
2. Same placeholder idx/role, OR both non-placeholder of the same shape type.
3. Resolved `ext` (width and height) differ by no more than `shape_size.size_tolerance_emu`.

Within an exact-match cohort, shapes whose size deviates from the cohort's dominant size (per `dominant_size_strategy`) beyond tolerance are flagged as size outliers (`shape_size.size_mismatch`). This is the safe, auto-applicable case, recorded with `confidence = deterministic`.

**Fuzzy cohort (heuristic, suggest-only).** Shapes grouped by approximate visual similarity (similar aspect ratio, similar role inferred from position/text, near sizes outside the exact tolerance) are flagged for review only. "Looks similar" is an inference and is never auto-applied; these records carry `confidence` of `high`, `medium`, or `low`, never `deterministic`. Off-grid placement is recorded as `shape_size.off_grid`.

#### Remediation

| Operation | Mode |
|---|---|
| Resize a size-outlier within an **exact-match cohort** to the cohort dominant `ext` | **Auto (v1.5)**, `confidence = deterministic` |
| Normalize size across a **fuzzy cohort** | **Suggest only** |
| Reposition shapes ("align similar groups", snap to guides) | **Suggest only**, always behind preview (alignment intent is heuristic) |
| Any resize that would change the aspect ratio of a picture or a shape containing a picture fill | **Suggest only** |

Auto-resize writes only `ext` (and `off` only when the deterministic rule fully specifies it); we operate surgically on the transform XML and touch nothing else. v1 reports cohorts and outliers; v1.5 enables exact-match auto-resize with accept/reject; broader alignment normalization is v2.

#### Configuration inputs (per profile)

Read from `shape_size` in the profile config:

- `size_tolerance_emu`: tolerance for "same size" within an exact cohort.
- `min_cohort_size`: minimum peer count before outlier logic runs (default 3) to avoid normalizing a pair toward an arbitrary value.
- `preserve_picture_aspect`: when true, never auto-change aspect on picture-bearing shapes.
- `dominant_size_strategy`: `median` (default) | `mode` | `largest`.

#### Severity

Per Appendix A.4, a shape size mismatch within a cohort defaults to `warning`:

| Condition | Severity |
|---|---|
| Size outlier within exact-match cohort | `warning` |
| Member of fuzzy cohort with size variance | `warning` |
| Rotated/flipped shape excluded from comparison | `warning` (informational) |

#### Edge cases & guards

- **Rotation, flips, text insets:** must be normalized before comparison; raw `off/ext` ignore all three. Shapes with non-zero rotation are compared on resolved geometry and, if the resolver cannot confidently account for rotation, excluded with an informational flag rather than mis-normalized.
- **Group shapes:** children live in a transformed coordinate space; apply the `chOff/chExt` to `off/ext` mapping before comparing a grouped child to an ungrouped peer. Do not auto-resize a grouped child in a way that fights the parent group transform.
- **SmartArt / charts / media:** flagged by the mandatory pre-flight as "will not modify." This module skips them entirely.
- **Pictures:** aspect-ratio-preserving by default (`preserve_picture_aspect`); never auto-stretch.
- **Arabic content:** size normalization is geometry-only and generally direction-safe, but repositioning is direction-sensitive; any reposition suggestion on a slide with detected Arabic is recorded with `arabic_flag = true` and `action = skipped`, marked "Arabic content, manual review," and excluded from auto-apply, consistent with the global RTL guard.

#### Phase

| Capability | Phase |
|---|---|
| Cohort detection, outlier report (exact + fuzzy) | v1 (audit) |
| Exact-match auto-resize, accept/reject | v1.5 |
| Alignment/reposition normalization | v2 |

### 4.3 Module 6: Header & Footer Consistency

#### Purpose

Check that footers, slide numbers, and dates are present, consistent, and correctly sourced across the deck, and flag "fake footers" (free-floating text boxes manually placed where a footer belongs instead of the real footer placeholder). The challenge is that footer state is governed by placeholders and layout-level non-propagation rules, so naive text scanning both misses real footers and mis-classifies decorative text.

#### Detection rules

| Check | Rule |
|---|---|
| Real footer placeholders | Inspect each slide for footer (`type="ftr"`), slide-number (`type="sldNum"`), and date (`type="dt"`) placeholders, plus the slide's `p:hf` settings and the show-on-title-slide flags. Read the resolved visible text via the inheritance chain. Compare slide number presence against `header_footer.template.slide_number` and date presence/format against `header_footer.template.date`. |
| Consistency | Compare footer text, presence of slide number, and date format across the cohort of content slides. Outliers (different footer string, missing slide number, divergent date format) raise `header_footer.text_mismatch` or `header_footer.position_mismatch` as appropriate. |
| Expected value match | If `header_footer.template.footer_text` is set, compare each slide's resolved footer against it (`header_footer.text_mismatch` on divergence). A required footer that is absent raises `header_footer.missing`. |
| Fake-footer fuzzy match | Detect non-placeholder text boxes positioned in the footer safe-zone (bottom band, configurable EMU) whose text **fuzzily matches** footer-like content (the expected footer string, a page-number pattern `\d+`/`Page \d+`, or a date pattern). Fuzzy string match (normalized Levenshtein ratio at or above the configured similarity threshold) catches near-duplicates of the real footer placed as manual text. |

Placeholder presence is exact/deterministic; fake-footer detection is fuzzy and therefore suggest-only, recorded with `confidence` of `high`, `medium`, or `low`, never `deterministic`.

#### Remediation

| Operation | Mode |
|---|---|
| Standardize footer **text** on slides that already use the real footer placeholder, to `header_footer.template.footer_text` | **Suggest only** in v1.5 (text is content-adjacent; the designer confirms) |
| Enable/show a missing slide-number or footer placeholder defined by the layout | **Suggest only** (toggling visibility can collide with intentional title-slide suppression) |
| Convert a detected fake footer into the real placeholder | **Suggest only**, always; never auto. Re-homing content into a placeholder is structurally risky and is presented as a recommendation with the offending shape highlighted. |
| Reposition a footer placeholder toward `header_footer.template.position_emu` | **Suggest only** (direction- and layout-sensitive) |

No footer auto-fix in v1 (audit-only). Footer remediation is inherently judgment-heavy, so even in v1.5 this module stays predominantly suggest-only; nothing here meets the `confidence = deterministic` bar for unattended auto-apply.

#### Configuration inputs (per profile)

Read from `header_footer.template` in the profile config:

- `footer_text`: canonical footer string (supports tokens, e.g. client name, confidentiality marking); `null` when no expected text is enforced.
- `slide_number`: boolean, whether a slide number is required.
- `date`: `{enabled, format}` for expected date presence and format (format defaults to `DD/MM/YYYY`).
- `position_emu`: `{x, y}` expected footer placeholder position, used for position-mismatch checks and reposition suggestions.
- `font_role`: the `FontRole` the footer text should use (default `caption`), used for font-mismatch checks.

The footer safe-zone band used for fake-footer scanning and the fuzzy similarity ratio are module-level scanning parameters applied during detection.

#### Severity

Per Appendix A.4, header/footer missing defaults to `error`; text, position, or font mismatches default to `warning`:

| Condition | Severity |
|---|---|
| Missing required slide number / footer where expected | `error` |
| Footer text inconsistent across content slides | `warning` |
| Footer position or font mismatch | `warning` |
| Date format inconsistent | `warning` |
| Suspected fake footer (fuzzy match) | `warning` (manual review; suggest-only) |

#### Edge cases & guards

- **Layout-level footer non-propagation:** a footer placeholder defined or hidden at the layout/master level does **not** automatically propagate to a populated slide. Resolve footer state through layout and master before declaring a slide "missing a footer," and never assume a layout-level footer is rendered on the slide. This is the primary false-positive source and must be guarded explicitly.
- **Title/section layouts:** footers are frequently and intentionally suppressed; honor the configured title-layout suppression so these are not flagged.
- **Gradient/picture decorative bars:** a colored bar or image at the bottom is not a footer; only text-bearing shapes enter fake-footer scanning.
- **Inherited footer text:** read via the inheritance resolver; placeholder text shown from the layout/master is not "empty."
- **Arabic footers:** footer text and position are direction-sensitive. On slides with detected Arabic (Unicode ranges per the global rule, or `rtl="1"` on `a:pPr`), suppress any position change and any font substitution, record findings with `arabic_flag = true` and `action = skipped`, and mark them "Arabic content, manual review." An RTL-unaware footer edit can silently break Arabic shaping, which is exactly the failure mode the global guard exists to prevent.

#### Phase

| Capability | Phase |
|---|---|
| Placeholder/consistency audit, fake-footer detection, report | v1 (audit) |
| Suggest-only standardization with accept/reject | v1.5 |
| Any structural footer auto-fix (re-homing, placeholder enable), full RTL footer handling | v2 |

---

## 5. Formatting Profiles (Data Model)

A formatting profile is the single source of truth for what "correct" means on a given deck. Every module (Master Slide Engine, Font Auditor, Margin & Alignment, Color Palette Checker, Shape & Size Normalizer, Header & Footer) reads its rules and thresholds from the resolved profile. Profiles are scoped per client and optionally per project, so a government deck in KSA and a consulting deck in the UAE can enforce different fonts, palettes, and margins without code changes.

The design principle: a profile is **declarative configuration**, not logic. The audit engine interprets it. This keeps the inheritance resolver, color math, and geometry checks (see Sections on those modules) profile-agnostic and lets non-engineers (Admins, in v2) tune rules without a deploy.

The canonical `Profile` shape, including the `config` object and every per-module key, is defined in Appendix A.6 and governs this section. The module config groups are exactly `master_slide`, `font`, `color_palette`, `geometry`, `shape_size`, and `header_footer`, keyed by `ModuleKey` (Appendix A.1).

### 5.1 Resolution and scoping

Profiles resolve from most specific to least specific. The first match wins per field, so a project profile can override a subset of a client profile while inheriting the rest. Scope is carried by two top-level fields, `client_scope` and `project_scope` (both `string | null`, per Appendix A.6).

| Order | Scope | Selected when |
|-------|-------|---------------|
| 1 | Project profile | Job's resolved profile sets `project_scope` to the job's project |
| 2 | Client profile | Job's resolved profile sets `client_scope` (client-level, `project_scope = null`) |
| 3 | Seeded default (EN or AR/bilingual) | No client/project profile, or user explicitly picks a default (`is_default = true`) |

The resolver merges parent fields into child where the child leaves a field null (shallow merge at the top-level `config` group, e.g. `color_palette`, `font`). Each Job records both the **selected** profile reference (`profile_id`, `profile_version`) and the **fully resolved** profile snapshot (`profile_snapshot`, see 5.10).

### 5.2 Top-level schema

All identifiers are strings. All timestamps are ISO 8601 UTC (`2026-06-30T14:05:00Z`). Sizes inside the profile are stored canonically in **EMU** (English Metric Units, 914400 per inch, 360000 per cm) to match OOXML and avoid rounding drift; the Admin UI accepts cm and converts. Font sizes are in points.

The top-level object matches Appendix A.6 exactly:

```jsonc
Profile {
  "id":            "string",
  "name":          "string",                // human label, e.g. "ACME Bank - Board Deck"
  "client_scope":  "string | null",         // client this profile belongs to (Odoo partner id when linked)
  "project_scope": "string | null",         // optional narrower project scope; null => client-level profile
  "is_default":    false,                    // true only for the two seeded profiles
  "version":       1,                        // incremented on each save; jobs stamp the version used
  "owner":         "string",                 // Entra object id of the owner
  "created_at":    "string (ISO 8601)",
  "updated_at":    "string (ISO 8601)",
  "config": {
    "master_slide":  { /* 5.3 */ },
    "font":          { /* 5.4 (roles + family/size) */ },
    "color_palette": { /* 5.5 */ },
    "geometry":      { /* 5.6 (safe zone, grid, alignment) */ },
    "shape_size":    { /* 5.7 */ },
    "header_footer": { /* 5.8 */ }
  }
}
```

Direction and region context (LTR/RTL/bilingual) is not a top-level field in the canonical contract; it is expressed through the font role sets (`complex_script`) and the header/footer template, and is captured by the seeded profile choice (5.11). Versioning is the single monotonic `version` integer; jobs stamp the version they ran against.

### 5.3 Master slide config (`config.master_slide`)

The master slide module enforces existing masters/layouts only (v1/v1.5 never synthesize a master). The profile pins or allowlists layouts and sets the geometry tolerance used to flag placeholder drift.

```jsonc
"master_slide": {
  "enforce_existing_only": true,            // v1/v1.5: never synthesize a master
  "pinned_layout_id":      "string | null", // explicit master/layout to enforce
  "layout_allowlist":      ["string"],      // permitted layout ids; others flagged as outliers
  "geometry_tolerance_emu": 9525
}
```

Notes:
- When `pinned_layout_id` is set, layouts other than the pinned one (and any in `layout_allowlist`) raise `master_slide.layout_outlier`.
- A deck with no usable master raises `master_slide.no_usable_master`; placeholder drift beyond `geometry_tolerance_emu` raises `master_slide.placeholder_geometry_off`.

### 5.4 Allowed font sets and size rules (`config.font`)

Fonts are specified per **text role** (`title`, `subtitle`, `body`, `caption`, the `FontRole` enum in Appendix A.1) and per **script**. The separation of Latin and complex-script (Arabic) typefaces is mandatory: OOXML carries the Arabic face in `a:cs` and the Latin face in `a:latin`, and a swap that ignores `a:cs` silently destroys Arabic shaping. The Font Auditor treats a run as conformant if its resolved face is in the allowed set for that role and script, and its size is within `size_tolerance_pt` of the role's `size_pt`.

The font module reads `font.roles.{role}.{latin[], complex_script[], size_pt, allowed_weights[]}`, per Appendix A.6:

```jsonc
"font": {
  "roles": {
    "title":    { "latin": ["Aptos Display"], "complex_script": ["Dubai", "Dubai Medium"], "size_pt": 36, "allowed_weights": ["regular", "bold"] },
    "subtitle": { "latin": ["Aptos"],         "complex_script": ["Dubai"],                 "size_pt": 24, "allowed_weights": ["regular", "bold"] },
    "body":     { "latin": ["Aptos"],         "complex_script": ["Dubai"],                 "size_pt": 16, "allowed_weights": ["regular", "bold"] },
    "caption":  { "latin": ["Aptos"],         "complex_script": ["Dubai Light"],           "size_pt": 11, "allowed_weights": ["regular"] }
  },
  "theme_font_refs_allowed": true,           // permit +mj-lt / +mn-lt references
  "size_tolerance_pt": 0.5
}
```

Notes:
- Each role lists ordered `latin` and `complex_script` arrays; the first entry is the **preferred** face used for any v1.5+ auto-fix suggestion, the rest are accepted-as-conformant alternates.
- `complex_script` arrays drive the Arabic detection guard: when a run contains Arabic Unicode (ranges per the Arabic detection rule) and its `a:cs` face is outside this set, the module flags `font.cs_typeface_missing` with `arabic_flag = true` and `action = skipped` rather than substituting.
- `theme_font_refs_allowed` permits `+mj-lt` / `+mn-lt` references. When false, theme-font references raise `font.theme_ref_disallowed`.
- A family outside the allowed set raises `font.family_out_of_set` (default severity `error`); a size outside the role's `size_pt` ± `size_tolerance_pt` raises `font.size_off_role`, and weight conflicts raise `font.mixed_weight` (both default `warning`, per Appendix A.4).
- Font names are matched case-insensitively against the resolved effective face, never against the inherited `None` returned by `run.font.name`.

### 5.5 Color palette (`config.color_palette`)

Every named color carries **both** an absolute hex value and a theme-color mapping, because most brand colors live as `schemeClr` references with `lumMod`/`lumOff`/`tint`/`shade`, not literal RGB. The Color Palette Checker accepts a fill as on-palette by either path: (a) it is a valid theme-color reference (`schemeClr` matching a slot in `theme_color_slots`, within an allowed tint/shade), or (b) its resolved RGB is within the CIEDE2000 tolerance of a named hex. Nearest-match uses Lab/CIEDE2000, not RGB Euclidean.

The block uses `theme_color_slots` and `named_colors[]` with `hex`, per Appendix A.6:

```jsonc
"color_palette": {
  "theme_color_slots": ["dk1","lt1","dk2","lt2","accent1","accent2","accent3","accent4","accent5","accent6"],
  "named_colors": [
    {
      "name": "Prezlab Blue",
      "hex": "#1F3A5F",
      "theme_ref": "accent1",                       // null if literal-only color
      "allowed_tints":  [0.0, 0.2, 0.4, 0.6, 0.8],  // tint multipliers permitted
      "allowed_shades": [0.0, 0.25, 0.5]
    },
    {
      "name": "Ink",
      "hex": "#111418",
      "theme_ref": "dk1",
      "allowed_tints":  [0.0, 0.15, 0.35],
      "allowed_shades": [0.0]
    }
  ],
  "on_palette_mode":        "by_name",       // "by_name" or "by_resolved_rgb"
  "match_tolerance_deltaE":  2.0,            // within = on-palette (CIEDE2000)
  "auto_replace_max_deltaE": 5.0,            // <= this => safe nearest-match suggestion
  "ambiguity_band_deltaE":   10.0           // beyond auto_replace and within this => flag, do not suggest
}
```

Severity follows Appendix A.4: a literal RGB off-palette with no near match (outside the ambiguity band) raises `color_palette.off_palette_rgb` at `error`; off-palette but within the ambiguity band, a `color_palette.disallowed_theme_slot`, or a `color_palette.tint_out_of_range` defaults to `warning`. Gradient, pattern, and picture fills have no single resolvable color and are explicitly skipped with a manifest note rather than flagged as violations.

### 5.6 Margins, grid, and alignment (`config.geometry`)

Stored in EMU. Detection of "outside safe-zone" is reliable and deterministic; the alignment/snap **fixes** are heuristic and remain behind preview (see the Margin & Alignment module). Per Appendix A.6, the safe zone is `geometry.safe_zone_margins_emu`, the grid lives under `geometry.grid`, and alignment tolerances live under `geometry.alignment.*`.

```jsonc
"geometry": {
  "safe_zone_margins_emu": { "left": 457200, "right": 457200, "top": 365760, "bottom": 365760 },  // 0.5 in / 0.4 in
  "body_band_emu": { "subtitle_floor": 1509712, "body_top": 1739900 },   // reserved strip; null unless the master draws it
  "grid": { "columns": 12, "gutter_emu": 182880, "enabled": false },
  "alignment": {
    "edge_tolerance_emu":    91440,          // 0.1 in; within => edges considered aligned
    "center_tolerance_emu":  91440,
    "spacing_tolerance_emu": 91440
  }
}
```

Notes:
- An element outside `safe_zone_margins_emu` raises `margin_alignment.outside_safe_zone` (default `warning`).
- Edge misalignment beyond `alignment.edge_tolerance_emu` raises `margin_alignment.edge_misaligned`; uneven spacing beyond `alignment.spacing_tolerance_emu` raises `margin_alignment.uneven_spacing` (both default `warning`).
- The master's own vertical text anchor is followed, never overridden. A master that hangs its title at the bottom of the box is pairing the heading with the subtitle under it; hoisting the text to the top of the box put the heading on the top margin and moved the empty space to the other side of it, which reads worse (design lead, 21/08/2026, reversing the 20/08 decision after seeing both). A title BOX that sits somewhere other than the master says is geometry, and `master_slide.placeholder_geometry_off` owns it.
- A title or subtitle box past a margin raises `margin_alignment.heading_past_margin` (default `warning`) and is never fixable: the designer asks the client whether the heading should be held to the frame. The same record covers a header that has grown through the reserved strip and past `body_band_emu.body_top` (side reported as "body ceiling"); the subtitle floor itself is not measured against, because a text box's descender slack overhangs it by design — the client master's own subtitle placeholder does so by 2.4 mm.
- An element standing in `body_band_emu` raises `margin_alignment.body_band_intrusion` (`error` from 3 mm, else `warning`), one record per slide, fixed by one translate of the whole body block.
- **A STATED frame's top binds** in the migration pass (`qc.migrate`), whatever states it: a presentation-space rectangle, or guides, with or without a header band. The block is seated on that line even when it is too tall for the region, and the overflow past the bottom margin is reported as an alert. Where nothing is stated — a frame inferred from placeholder extents — the old clamp holds: the block moves as far as it fits and no further, and the report says the body was not seated on a stated frame and what to draw in the master to make it exact.

  This rule was first written as "the body CEILING binds", which tied it to a guide PAIR, and that was a bug with a deck-wide symptom (design lead, 21/08/2026): a master stating its frame with a rectangle has no band, so the clamp applied, so every slide whose content was too tall kept the position it arrived with and the deck read as "not following the presentation space" from end to end. Alignment must not be conditional on how tall a slide's content happens to be — that produces a deck where some slides sit on the frame and some do not, which is the complaint both times. The per-slide report names the frame source and where the body ACTUALLY landed, never the target it aimed at: a message that stated the target while the clamp held the block elsewhere is what kept this invisible for two rounds.
- Text insets (`bodyPr lIns/tIns/rIns/bIns`) are NOT folded into the margin check: the box is what is measured (see "The box, not the text" above). Shape rotation/flips are not captured by the raw `off/ext` either, and a rotated shape is skipped rather than compared.
- **Collections are the unit of judgment and of movement.** A shape riding a larger one (contained in it, welded to it, or within 10 mm across the axis of the move) belongs to it: it is compared against how peers sit in THEIR anchors rather than against the absolute clusters, and every positional fix translates the whole collection. Backdrops (≥ 35% of the canvas, or page-deep) and recurring bottom-strip furniture ride nothing, and satellites never chain, so a slide of touching decorative shapes cannot drag half the canvas. Shared with the migration through `qc.util.rides_with`, because a fix that answered "what travels with this" differently from the check that raised it is how a composition comes apart (design lead, 20/08/2026).

### 5.6a The master a profile applies, and replacing it

A profile carries two things: the rules (JSON, reviewed and edited by a design lead) and the **master file** they were read from, stored under `data/templates/`. Applying a profile hands PowerPoint that stored file, because restyling a slide needs real `slideLayout` parts and no amount of extracted numbers substitutes for them.

The file was stored once, when the profile was created, and that made every later change to the designer's master invisible: a presentation-space rectangle, a moved guide, a renamed layout reached no deck at all (design lead, 21/08/2026 — the presentation-space box was missing from every formatted deck because the stored copy predated it). `POST /profiles/{pid}/master` replaces it, and the profile editor shows what is stored (size, digest, date) and which source the frame was read from, so a stale copy is visible rather than inferred.

Replacing re-reads only what the master **states**: margins (presentation space, else guides, else placeholders), the reserved header band, the column grid, and the layout allow-list. Fonts, palette and tolerances are left exactly as edited, because those are decisions about the client rather than readings of the file, and reverting them to a fresh projection would punish every edit made in the editor. The version bumps, and the change note names each field that moved.

**The master that stays behind.** A slide that cannot be rebuilt keeps the deck's ORIGINAL design alive to serve it, so the output carries two slide masters and PowerPoint's master view lists the original **first**. A designer opening that view sees a master with none of the new guides, furniture or presentation-space rectangle and reads it as "the master was not copied", when it was — onto the other one (design lead, 21/08/2026). `ApplyResult` therefore reports `masters` and `stragglers` (the slide indexes still on another design), the result page explains the consequence and names the slides, and the audit's `master_slide.foreign_master` is the remedy: it flags exactly those slides and its fix moves them onto the applied master. The per-slide move note also names which frame it was seated on — presentation space, guides, or placeholder extents — because a stale stored master otherwise looks identical to a bug.

### 5.6b Content the migration removes, and putting it back

The migration removes header TEXT the master has no placeholder for, and reports each removal as an alert carrying the full text. It keeps the removed element's own XML with the change record, so the result page offers each piece with a tick and `POST /format/{job_id}/restore` puts the selected ones back. Listing the words alone left a designer retyping them and placing them by eye; the tool decides what the master has no room for, the designer decides whether that decision was right (design lead, 20/08/2026).

A restore is an insert with four constraints, three of which were real damage found on 20/08/2026 ("bringing back selected pieces screwed up the whole presentation content"):

- **Parsed with python-pptx's parser**, never lxml's. A generic element spliced into a python-pptx tree is a shape the library can no longer read.
- **Inserted before `p:extLst`**, which the schema requires to be the last child of a shape tree. Appending after it produces a deck PowerPoint offers to repair.
- **Renumbered**, so restoring the same piece twice cannot collide with itself. The route also skips ids already restored, since a browser can resubmit the POST.
- **Put back at its ORIGINAL position**, because a restore is an undo, not a re-layout (design lead, 21/08/2026: "ticked pieces should be put back where they were originally, not into the body area"). Where the master has since filled that spot the piece prints over it; the piece is named `RESTORED …` so it is findable in PowerPoint's selection pane, and the report names exactly what it covers. An interim version hunted for the nearest empty space instead and landed an eyebrow three inches down the slide, which is not where an eyebrow belongs however empty that spot was.

Only text-bearing shapes are ever swept. A corner rule, bracket or mark drawn in the header band carries no text and belongs to a composition; removing one and reporting it as "unplaced text" was a defect, however it was phrased.

### 5.7 Shape and size normalization (`config.shape_size`)

The Shape & Size Normalizer unifies sizes within a cohort of like shapes. The profile sets the size tolerance, the minimum cohort size before unifying, picture-aspect handling, and the strategy used to pick the dominant size. Field names match Appendix A.6 (the shape tolerance is `shape_size.size_tolerance_emu`).

```jsonc
"shape_size": {
  "size_tolerance_emu":      9525,
  "min_cohort_size":         3,             // minimum shapes to treat as a cohort before unifying
  "preserve_picture_aspect": true,
  "dominant_size_strategy":  "median"       // "median" | "mode" | "largest"
}
```

Notes:
- A shape outside the cohort's dominant size beyond `size_tolerance_emu` raises `shape_size.size_mismatch`; an off-grid placement raises `shape_size.off_grid` (both default `warning`).
- Exact matches (same preset, extent within tolerance, same index/role) are eligible for auto-apply in v1.5 only when `confidence = deterministic` (Appendix A.1); fuzzy similarity stays suggest-only.

### 5.8 Header/footer template (`config.header_footer`)

Header and footer checks are direction-sensitive. When a slide carries Arabic content, footer **position** enforcement is guarded (flag, do not move) per the Arabic guard. The template matches Appendix A.6: a single `template` object with `footer_text`, `slide_number`, `date`, `position_emu`, and `font_role`.

```jsonc
"header_footer": {
  "template": {
    "footer_text":  "ACME Bank  |  Confidential",   // null to disable footer text
    "slide_number": true,
    "date":         { "enabled": true, "format": "DD/MM/YYYY" },
    "position_emu": { "x": 457200, "y": 6492240 },    // bottom-left for LTR; mirrored for RTL detection
    "font_role":    "caption"
  }
}
```

Notes:
- A missing required footer raises `header_footer.missing` (default `error`).
- Text, position, or font mismatches raise `header_footer.text_mismatch`, `header_footer.position_mismatch`, or `header_footer.font_mismatch` (default `warning`).
- For RTL/bilingual slides, footer position is mirrored in detection only; the move is guarded (`arabic.guarded_operation`, `action = skipped`, `arabic_flag = true`) rather than auto-applied.

### 5.9 Module enablement and severity overrides

Each of the six modules (keyed by `ModuleKey`: `master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer`) runs independently. A job selects which modules to run via the API `modules[]` parameter (empty array = run all). Per-module thresholds live inside the matching `config.{module}` block above (5.3 to 5.8); there is no separate top-level `modules` config object in the canonical contract.

Severity defaults follow Appendix A.4 and are profile-overridable through the relevant config block. Auto-apply eligibility in v1.5 is gated on `confidence = deterministic` (Appendix A.1); inference-heavy findings stay suggest-only regardless of configuration. The `Confidence` enum is `deterministic`, `high`, `medium`, `low`, and `Severity` is `error`, `warning`, `info`.

### 5.10 Versioning and per-Job snapshot (reproducibility requirement)

Profiles are mutable over time, but audits must be reproducible. Two mechanisms enforce this:

1. **Monotonic version on every save.** Any persisted change increments `version` and updates `updated_at`. Prior versions are retained (immutable rows keyed by `(id, version)`), never hard-deleted; archiving hides a profile from selection without breaking historical references.

2. **Job-stamped resolved snapshot.** When a Job starts, the engine resolves the profile (project over client over default, per 5.1), serializes the **fully merged** `config`, and stores it on the Job as `profile_snapshot` together with `profile_id` and `profile_version` (Appendix A.5). The audit and any later re-run read from the snapshot, not the live profile. Editing the profile afterward never changes a completed Job's findings. The manifest references the snapshot so a reviewer can confirm which exact rules produced each finding.

```jsonc
// fields stored on the Job record (see Appendix A.5)
{
  "profile_id":       "string",
  "profile_version":  7,
  "profile_snapshot": { /* full merged Profile.config as of run time */ }
}
```

Because preview/audit and fix share one pipeline with writes disabled (per the manifest architecture), the snapshot guarantees audit and fix cannot drift on the same Job.

### 5.11 Seeded default profiles

Two profiles ship as seed data (`is_default: true`) and cannot be deleted, only cloned. They give v1 immediate utility before any client-specific profile exists.

| | Prezlab EN (default) | Prezlab AR / Bilingual |
|---|---|---|
| Primary direction | LTR (Latin-first font roles) | Bilingual (complex-script roles first-class) |
| Fonts (Latin) | Prezlab Latin brand set per role | Same Latin set |
| Fonts (complex script) | Arabic set present but secondary | Arabic brand set per role, first-class |
| Color palette | Full Prezlab palette (theme slots + named tints) | Same palette |
| Header/footer | Left-aligned LTR footer | RTL-mirrored footer, Arabic-aware position guard |
| Arabic guard | Detect + flag Arabic runs | Detect + flag; direction-sensitive ops suppressed, not auto-moved |
| Intended use | English-primary consulting/IB decks | KSA/UAE government and bilingual decks |

Both share the same palette and geometry; they differ in direction handling, the priority of the complex-script font set in `config.font.roles`, and footer mirroring in `config.header_footer`. Designers select one when no client profile is mapped. Both ship in v1 (CRUD UI is v2, see 5.12).

### 5.12 Admin UI and API (v2 CRUD; seeded in v1)

In **v1, profiles are seeded** (the two defaults above), loaded via migration/seed scripts. There is no in-app create or edit; this keeps v1 scope at audit-only and avoids shipping profile management before role enforcement lands (sign-in gate is v1, app-role RBAC is v1.5, profile administration is v2). Designers may **select** a profile, which is captured on the Job as `profile_snapshot`, but cannot persist new profiles. **Profile CRUD UI is a v2 capability.**

**Phase v2** adds full create/update behind the Admin role (Entra app role `Admin` in the token `roles` claim). Reviewers and Designers get read-only access to profiles; only Admins manage them, per the role permission matrix in Appendix A.8.

The canonical REST surface is Appendix A.7. All paths are under `/v1`. Roles: D = Designer, R = Reviewer, A = Admin.

| Method & path | Purpose | Roles | Phase |
|---|---|---|---|
| `GET /v1/profiles` | List profiles | D, R, A | v1 |
| `GET /v1/profiles/{id}` | Get a profile | D, R, A | v1 |
| `POST /v1/profiles` | Create a profile | A | v2 (seeded in v1) |
| `PUT /v1/profiles/{id}` | Update a profile (increments `version`) | A | v2 |

There are no separate clone, archive, or versioned-read endpoints in the canonical contract; cloning and soft-archive in the v2 UI are implemented over `POST /v1/profiles` and `PUT /v1/profiles/{id}` respectively. Historical versions are retained server-side and surfaced through the same `GET /v1/profiles/{id}` resource.

Admin UI requirements (v2):
- **Form-driven editor** grouped by the `config` sections (master slide, font, color palette, geometry, shape size, header/footer), with cm input that converts to EMU on save and pt for type sizes.
- **Color editor** that accepts hex and a theme-slot picker (mapping to `color_palette.theme_color_slots`), previews resolved RGB after tint/shade, and validates each named color resolves cleanly.
- **Font role matrix** with separate Latin and complex-script columns per role (`config.font.roles.{role}.{latin,complex_script}`), validating that complex-script faces are present for bilingual/RTL profiles.
- **Live validation** against the JSON schema before save; reject saves that leave a required role empty or set inconsistent tolerances.
- **Version history view** with diff between any two versions and a read-only render of the `profile_snapshot` attached to a given Job.
- **Scope binding** to an Odoo client/project (via the client-credentials service identity) so `client_scope`/`project_scope` can be linked to the partner record without manual id entry.
- All writes audit-logged with `owner`/updater identity (Entra OID) and timestamp.

---

## 6. System Architecture & Stack

The tool is an asynchronous, queue-backed web application. A 200-slide deck plus optional rendering exceeds what a synchronous HTTP request should hold open, so we decouple submission from processing: the browser uploads, gets a `job_id` back immediately, then tracks progress while a worker does the heavy lifting out of band. This buys us per-slide progress UX, no held connections, and free retries. The REST API is the primary surface; the React SPA and Odoo are both clients of it.

### 6.1 End-to-End Flow

1. Client streams the `.pptx` to `POST /v1/uploads`, which returns `{ upload_id, expires_at }`. The upload is the single contract for getting deck bytes into the system; no internal storage key is ever exposed to the client.
2. Client `POST`s to `/v1/jobs` with `{ upload_id, profile_id, modules[], mode, preview }`. The API creates a `Job` (`status = queued`) and enqueues a task referencing the resolved upload and profile snapshot.
3. Worker pulls the task, downloads the deck, runs the selected pipeline (full audit or targeted modules), emits one `FindingRecord` per detected issue or applied change, writes outputs (manifest JSON, report PDF/CSV, and the cleaned `.pptx` in fix mode) back to object storage, and updates `Job.status` and `Job.slides_processed` as it advances.
4. Client tracks progress via SSE (primary) or poll (fallback) against `GET /v1/jobs/{job_id}`, showing a per-slide counter driven by `slides_processed` out of `slides_total`.
5. On `status = ready`, the client requests short-TTL signed download URLs for each output via `GET /v1/jobs/{job_id}/outputs`.

Audit mode (`mode = audit`) and fix mode (`mode = fix`) run the **same pipeline with writes disabled**, so a flagged finding and the change that would remediate it cannot drift apart. The `preview` flag forces flag-only behavior even in `fix` mode.

### 6.2 API Surface (primary endpoints)

The canonical endpoint contract is Appendix A.7; the table below summarizes the endpoints this section depends on.

| Method | Path | Purpose |
|---|---|---|
| `POST` | `/v1/uploads` | Stream a `.pptx` to storage; returns `{ upload_id, expires_at }` |
| `POST` | `/v1/jobs` | Create a job: `{ upload_id, profile_id, modules[], mode, preview }` |
| `GET` | `/v1/jobs/{job_id}` | Job status, progress, error: `{ status, slides_total, slides_processed, error }` |
| `POST` | `/v1/jobs/{job_id}/cancel` | Cancel a queued or processing job |
| `GET` | `/v1/jobs/{job_id}/findings` | Full `FindingRecord[]` (filterable by module, severity, slide) |
| `GET` | `/v1/jobs/{job_id}/summary` | Aggregate counts per `issue_type` and `severity` |
| `GET` | `/v1/jobs/{job_id}/outputs` | Signed URLs for `cleaned_pptx`, `report_pdf`, `report_csv`, `manifest_json` |
| `GET`/`POST`/`PUT` | `/v1/profiles`, `/v1/profiles/{id}` | Per-client/per-project formatting profiles (Admin writes) |

Progress is delivered as an SSE stream layered over the job-status surface, with polling of `GET /v1/jobs/{job_id}` as the fallback. There is no separate `events` endpoint, no `manifest` endpoint, and no `artifacts/{kind}` endpoint: structured findings come from `/findings`, and all downloadable outputs (including the manifest JSON) are signed URLs from `/outputs`.

`JobStatus` lifecycle: `queued → processing → ready | failed | cancelled`. A deck that processed but hit a guard on some slides (for example, no usable master or SmartArt present) still resolves to `ready`; the guarded slides surface as `FindingRecord`s with `action = skipped` rather than as a distinct job status.

### 6.3 Component Description

| Component | Responsibility | Notes |
|---|---|---|
| **Web/API tier** | Auth (Entra OIDC), request validation, job CRUD, signed-URL minting, SSE fan-out | Stateless; horizontally scalable. No deck bytes pass through it beyond the upload stream. |
| **Worker pool** | Runs the six-module pipeline, the inheritance/color resolver, overflow checks, `FindingRecord` emission, output writing | CPU-bound; scale by deck volume. Idempotent on `job_id` for safe retries. |
| **Queue / broker** | Durable task hand-off, retry, dead-letter | Decouples API from workers. |
| **Object storage** | Uploaded decks, cleaned `.pptx`, manifests, PDFs | Encrypted at rest. Auto-deletion of uploads after processing (v1 confidentiality requirement). |
| **Relational DB** | `Job`, `Profile`, `User`/role, `FindingRecord` set (JSONB) | Source of truth for status and findings. |
| **Renderer service** | Optional slide thumbnails / deck-to-PDF | Out-of-process, called only when visual artifacts are needed (see 6.5). |

The original upload is never overwritten; the cleaned deck is always a distinct artifact.

### 6.4 Recommended Stack

| Layer | Choice | Justification |
|---|---|---|
| Language / API | **Python + FastAPI** | python-pptx and the OOXML manipulation libraries are Python-native; async FastAPI suits SSE and signed-URL workflows. One language across API and worker. |
| Queue / workers | **Redis + RQ** (escalate to **Celery** if routing/retry needs grow) | RQ is the right weight for a single in-house tool. Celery only if we later need complex chains/canvas. |
| Database | **Postgres with JSONB records** | Relational for `Job`/`Profile`/RBAC; JSONB stores the per-deck `FindingRecord` set without a rigid schema and is directly queryable for the regression corpus and per-fix precision tracking. |
| Object storage | **Azure Blob** (prod), **S3-compatible MinIO** (dev/on-prem option) | Azure fits the existing M365/Entra estate and Graph rendering. MinIO keeps an on-prem path open if data-residency (KSA/UAE government) forces it. Same S3-style signed-URL API both ways. |
| Frontend | **React SPA** | Rich client for the before/after diff review surface, profile admin, and live progress. Pure API client; installs nothing on designer machines. |

Hosting (Azure vs on-prem) is decided by the data-residency and retention review with Operations/IT-security, not by the stack. Both targets are supported by the choices above.

### 6.5 Engine Decision (settled by the spike)

Documented choice, not yet final. The engine that performs the OOXML read/audit and surgical edits is the single highest-leverage decision because formatting quality **is** the product.

| Option | Strengths | Gaps |
|---|---|---|
| **python-pptx + shared resolver** (baseline) | Full control, surgical XML edits, no license cost, preserves unmapped parts (SmartArt/media/OLE) verbatim | No API to re-apply a layout on a populated slide; `Font.name` touches Latin only (Arabic complex-script `a:cs` needs raw XML); does not resolve inherited or theme/`schemeClr` values |
| **Aspose.Slides** (commercial) | More complete RTL handling and layout reassignment; higher-level operations | License cost; less surgical (more rewrite churn risk); third-party fidelity to validate |

**Decision rule:** ship the baseline (python-pptx) backed by a **shared inheritance resolver** (run `rPr` → paragraph/list level → `txBody` list style → placeholder/layout → placeholder/master → master text styles → theme/presentation defaults) and a **color resolver** (`schemeClr` → `theme1.xml` RGB → apply `lumMod`/`lumOff`/`tint`/`shade`, compare in CIEDE2000/Lab). Budget for an Aspose license and adopt it **if the 2-week spike shows python-pptx cannot meet quality on master enforcement, RTL, or layout reassignment**. The resolver is shared across all six modules (`master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer`) regardless of engine, so the modules never read a `None` (inherited) value as if it were the effective value.

### 6.6 Rendering Approach

| Need | v1 approach | Rationale |
|---|---|---|
| Report PDF (issue logs, counts, summaries; no slide images) | **ReportLab / WeasyPrint, no renderer** | python-pptx cannot render. A text/table report needs no renderer at all; generate it directly from the JSONB `FindingRecord` set. |
| Slide thumbnails / deck-to-PDF | **Microsoft Graph** (`driveItem` content `?format=pdf`) | Server-safe, fits the M365/SSO/Azure estate, no local Office. Primary renderer when visual artifacts are required. |
| Renderer fallback | **LibreOffice headless** (internal artifact only) | Font-substitutes and has documented Arabic RTL rendering bugs, so its PDF is fidelity-acceptable for internal use, not pixel-faithful. |
| Explicitly avoided | **PowerPoint COM / Office automation** | Microsoft does not support unattended server automation: single-instance, hangs, not licensable for the pattern. |

Install canonical brand and Arabic fonts on any renderer host; embed fonts in the cleaned `.pptx`.

### 6.7 Findings as Source of Truth

Each module emits one `FindingRecord` per detected issue or applied change. The full schema is canonical in Appendix A.2; the load-bearing fields for this architecture are:

```jsonc
FindingRecord {
  "record_id":       "string (uuid)",   // canonical id, not "id"
  "job_id":          "string",
  "slide_index":     0,                   // integer, ZERO-BASED
  "shape_id":        "string",            // OOXML p:cNvPr @id, as a string
  "shape_path":      "string | null",     // group ancestry, null at top level
  "module":          "ModuleKey",         // master_slide | font | margin_alignment | color_palette | shape_size | header_footer
  "issue_type":      "string",            // stable dotted code, e.g. "font.family_out_of_set"
  "severity":        "Severity",          // error | warning | info
  "action":          "Action",            // flagged | changed | skipped
  "confidence":      "Confidence",        // deterministic | high | medium | low
  "arabic_flag":     false,               // true if the record involves Arabic/RTL content
  "profile_rule_id": "string | null",     // the profile rule that triggered the record
  "message":         "string"             // human-readable, not "note"
}
```

Records are persisted as JSONB and are the single source for the report PDF and CSV exports and for the manifest JSON returned via `/outputs`. We do **not** reconstruct changes by diffing the two `.pptx` zips: a full open/save round-trip rewrites untouched XML and re-zips, producing noisy diffs driven by rewrite churn. Workers operate surgically, touching only the XML they change, and emit each `FindingRecord` as they go.

### 6.8 Words-Based Component Diagram

```
                         ┌──────────────────────────────┐
                         │        React SPA (browser)    │
                         │  upload · progress · diff · UI │
                         └───┬───────────────▲───────────┘
        (1) POST /uploads    │               │ (4) SSE / poll GET /jobs/{job_id}
        deck bytes           │               │ (5) GET /jobs/{job_id}/outputs (signed)
   ┌─────────────────────────▼──┐            │
   │      Object Storage         │           │
   │ Azure Blob / MinIO          │◄──────────┼────── writes outputs (3)
   │ uploads · cleaned.pptx ·    │           │
   │ manifest · report.pdf/csv   │           │
   └───────────▲─────────────────┘           │
               │ download deck (3)           │
               │                  ┌──────────┴───────────┐
   (2) POST /jobs ───────────────►│   Web / API tier      │
                                  │   FastAPI · Entra OIDC │
                                  │   jobs · profiles ·SSE │
                                  └───┬───────────▲────────┘
                                      │ enqueue   │ status / findings
                                      ▼           │
                              ┌───────────────┐   │
                              │ Queue (Redis/RQ)│ │
                              └───────┬─────────┘ │
                                      │ pull      │
                          ┌───────────▼──────────┐│
                          │     Worker pool       ││
                          │  6-module pipeline ·  ││
                          │  inheritance/color    ││
                          │  resolver · findings  ││
                          └───┬──────────────┬────┘│
            renderer call (opt)│             │ persist Job + FindingRecord (JSONB)
                               ▼             ▼     │
                   ┌───────────────────┐  ┌────────┴────────┐
                   │ Renderer service  │  │   Postgres       │
                   │ MS Graph (primary)│  │ Job · Profile ·  │
                   │ LibreOffice (fb)  │  │ findings (JSONB) │
                   └───────────────────┘  └──────────────────┘

External caller: Odoo ──OAuth2 client-credentials──► Web / API tier
```

### 6.9 Single Biggest Architectural Risk

**Risk: semantic correctness of OOXML transforms on messy real-world decks, compounded by Arabic.** The format is forgiving of structural variety, third-party exports (Canva, Gamma, Google Slides, Keynote) are frequently non-conformant, and inheritance/theme resolution plus Arabic complex-script handling create many silent-failure paths. An RTL-unaware font swap to a Latin-only brand font destroys Arabic glyphs with no error. In a 140-person shop, one mangled client deck spreads by word of mouth and kills adoption, so first-release precision outweighs feature count.

**Mitigation:**

- **Audit-first phasing.** v1 detects, flags, and reports only (`mode = audit`, every record carries `action = flagged`). Detection is the objective half of every operation; remediation (the judgment-call half) ships later in `fix` mode behind explicit accept/reject.
- **Preview as the review surface.** The per-slide visual before/after diff, not the text manifest, is what designers approve. No bulk write without it; `preview = true` keeps writes disabled during review.
- **Regression corpus + precision gate.** Maintain a corpus of real client decks (including Arabic fixtures). Track precision and false-positive rate per fix type; gate auto-apply (eligible only when `confidence = deterministic`) on that metric before enabling it.
- **Defensive guards.** Pre-flight flags SmartArt/charts/media as "will not modify" (`issue_type = preflight.unmodifiable_content`, `severity = info`); "no usable master" halts master enforcement; Arabic runs (Unicode `U+0600–06FF`, `U+0750–077F`, `U+08A0–08FF`, presentation forms `FB50–FDFF`/`FE70–FEFF`, or `rtl='1'` on `a:pPr`) are excluded from direction-sensitive operations and font substitution and emitted as `arabic.guarded_operation` with `arabic_flag = true` and `action = skipped`.
- **Spike validation.** The 2-week spike validates safe-write (cleaned deck opens cleanly in desktop PowerPoint and M365 web), round-trip preservation on real decks, 200-slide processing time, and resolver + Arabic-detection reliability before the timeline is committed.

---

## 7. Change Manifest, Audit Report & Preview

This section specifies how the tool records what it found and what it changed, how it reports that to designers and reviewers, and how preview keeps audit and fix honest. The set of `FindingRecord` entries is the system of record. Every report (PDF, CSV) and every UI surface (issue log, before/after diff, accept/reject) is a projection of it. Get the record schema right and the rest is rendering.

### 7.1 Finding Record Schema

Each module emits one `FindingRecord` per finding. A record describes a single observation about a single shape (or slide-level element) and, where applicable, the change applied to it. Records are immutable once emitted within a job: a v1.5 accept/reject decision produces a new applied output, it does not mutate the original finding. The schema below is the canonical `FindingRecord` from Appendix A.2.

| Field | Type | Required | Description |
|---|---|---|---|
| `record_id` | string (UUID) | yes | Canonical identifier for this finding within the job (the id field is `record_id`, not `id`). Referenced by accept/reject and inline-comment payloads. |
| `job_id` | string | yes | The owning job. |
| `slide_index` | integer | yes | Zero-based slide position in the deck. Slide-level findings (header/footer) still carry the slide they apply to. |
| `shape_id` | string | yes | OOXML shape identifier (`p:cNvPr@id`), carried as a string, scoped to the slide. For grouped shapes, the leaf shape id; the group path is carried in `shape_path`. Null only for whole-slide findings. |
| `shape_path` | string \| null | no | Group ancestry (e.g. `12/3` for shape 3 inside group 12) so a shape inside nested `groupShape` elements is locatable after transform resolution. Null at top level. |
| `module` | `ModuleKey` | yes | Emitting module, one of the canonical keys: `master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer`. |
| `issue_type` | string | yes | Stable, dotted machine code for the finding category, used for summary counts and per-fix-type precision tracking. Controlled vocabulary, see 7.1.1. |
| `property` | string \| null | no | The specific OOXML property observed or changed (e.g. `rPr.latin.typeface`, `rPr.cs.typeface`, `fill.color`, `position.left`, `ext.width`, `footer.text`). |
| `old_value` | string \| null | yes | Resolved current value at the property. For inherited values this is the EFFECTIVE value from the inheritance resolver, with the source recorded in `value_source`. Null when the property is absent. |
| `new_value` | string \| null | yes | Proposed (preview) or applied value. Null for pure detections with no remediation (`action: flagged`). |
| `value_source` | enum | no | Where `old_value` was resolved from: `run`, `paragraph`, `txbody_liststyle`, `placeholder_layout`, `placeholder_master`, `master_txstyles`, `theme_default`, `literal`. Required for font and color records so reviewers see whether a value was explicit or inherited. |
| `severity` | `Severity` | yes | `error`, `warning`, `info`. See 7.1.2. |
| `action` | `Action` | yes | `flagged` (detected, no change), `changed` (a fix was applied or proposed), `skipped` (a candidate fix was deliberately not applied; reason in `message`). |
| `confidence` | `Confidence` | yes | `deterministic`, `high`, `medium`, `low`. Gates eligibility for auto-apply, see 7.1.3. |
| `arabic_flag` | boolean | yes | True when the record involves Arabic/RTL content (the target shape/run contains Arabic script or carries `rtl="1"`). Direction-sensitive and font-substitution operations on these records are forced to `skipped` in v1, see the Arabic guard in 7.1.4. |
| `profile_rule_id` | string \| null | no | The formatting-profile rule that triggered the finding (e.g. `palette.primary`, `font.body.family`), linking each finding back to the client/project profile. |
| `message` | string | yes | Human-readable, designer-facing explanation (the human-readable field is `message`, not `note`). For `skipped`, must state the reason (e.g. "Arabic content, manual review"). |
| `created_at` | string (ISO 8601) | yes | Emission timestamp. |

#### 7.1.1 `issue_type` vocabulary (initial set)

`issue_type` is a closed, dotted list scoped by module (per Appendix A.3) so counts roll up cleanly and precision can be tracked per type. Initial codes:

- master_slide: `master_slide.layout_outlier`, `master_slide.no_usable_master`, `master_slide.placeholder_geometry_off`
- font: `font.family_out_of_set`, `font.size_off_role`, `font.mixed_weight`, `font.theme_ref_disallowed`, `font.cs_typeface_missing`
- margin_alignment: `margin_alignment.outside_safe_zone`, `margin_alignment.edge_misaligned`, `margin_alignment.uneven_spacing`
- color_palette: `color_palette.off_palette_rgb`, `color_palette.disallowed_theme_slot`, `color_palette.tint_out_of_range`
- shape_size: `shape_size.size_mismatch`, `shape_size.off_grid`
- header_footer: `header_footer.missing`, `header_footer.text_mismatch`, `header_footer.position_mismatch`, `header_footer.font_mismatch`
- cross-cutting: `preflight.unmodifiable_content`, `arabic.guarded_operation`

New codes require a schema change and a migration of stored records, so the list is deliberately conservative for v1.

#### 7.1.2 Severity semantics

| Severity | Meaning | Example |
|---|---|---|
| `error` | Clear violation of the active profile that should block delivery and will read as a defect. | Body text in a non-brand font; shape positioned outside the safe-zone margin; header/footer missing. |
| `warning` | Likely issue, or a guarded/skipped action needing a human; legitimate exceptions exist. | Color is a near match to a palette entry (CIEDE2000 within the ambiguity band); fuzzy shape-size mismatch; Arabic-guarded skipped operation. |
| `info` | Advisory observation, no action implied. | Pre-flight unmodifiable content (SmartArt, charts, media). |

Severity is assigned by the emitting module from the profile (per the Appendix A.4 defaults, which are profile-overridable), not inferred at report time.

#### 7.1.3 Confidence and auto-apply gating

`confidence` separates objective detection from judgment-call remediation, per the detection-vs-remediation classification. It is per-record certainty that the finding is correct and its proposed fix is safe.

- `deterministic`: the fix is mechanically correct and reversible (e.g. exact-match shape size: same `prstGeom` prst, same placeholder idx/role, `ext` within threshold). Eligible for auto-apply once the regression corpus clears the gate.
- `high`: strong rule match, low ambiguity (e.g. explicit run-level font family differs from the single profile body font).
- `medium`: inference involved (e.g. nearest-color replacement, font-size hierarchy inference). Preview-only, never auto-applied in v1.5.
- `low`: heuristic (e.g. "looks similar" group alignment). Suggest-only.

Auto-apply in v1.5 is restricted to `confidence: deterministic` AND `arabic_flag: false` AND a per-`issue_type` precision gate met on the real-deck regression corpus. Everything else routes through preview and explicit acceptance. Numeric precision is tracked separately as an aggregate metric, not on the record.

#### 7.1.4 Arabic guard at the record level

When `arabic_flag` is true, any record whose `property` is direction-sensitive (`position.*`, `margin.*`, `footer.position`, paragraph alignment) or is a Latin font substitution (`rPr.latin.typeface`) is emitted with `action: skipped` and `message: "Arabic content, manual review"`, carrying `severity: warning` per the Appendix A.4 default for Arabic-guarded skipped operations. The module still records the finding (`issue_type: arabic.guarded_operation`) so it appears in the report and count; it simply does not propose a change. This prevents an RTL-unaware font swap from silently destroying Arabic shaping.

### 7.2 Records Are Emitted by Modules, Not Reconstructed by Zip Diff

The record set is built from `FindingRecord` entries each module emits as it runs. The tool does NOT compute changes by diffing the original .pptx zip against the cleaned .pptx zip. Reasons:

- **Round-trip churn is noisy.** A full open/save through the engine re-serializes and re-zips untouched parts. A zip diff would surface large volumes of rewrite churn (reordered attributes, rewritten untouched XML) that are not semantic changes, drowning the real edits. The problem is rewrite churn, not serialization nondeterminism.
- **Records carry intent the diff cannot recover.** A diff shows that bytes changed; it cannot tell you the value was inherited, what profile rule fired, the confidence level, or that a candidate fix was deliberately skipped. `value_source`, `confidence`, `action: skipped`, and `profile_rule_id` only exist because the module that made the decision wrote them down.
- **Surgical edits stay traceable.** Modules touch only the XML they change and emit one record per change, so the record set is a precise, one-to-one log of decisions rather than an after-the-fact inference.

Storage: records are persisted as JSON/JSONB associated with the Job (records as a JSON array, or a normalized child table keyed by `record_id` for query). The change manifest (`manifest_json`), PDF, and CSV are all projections rendered from this stored record set, never recomputed from the files.

### 7.3 Audit Report

The audit report is the human deliverable for Flow B (Audit Only) and the review artifact for Flow A. All views derive from the stored record set.

#### 7.3.1 Slide-by-slide log

Grouped by `slide_index`, then by `module`. Each entry shows the `severity` badge, `issue_type`, `property`, `old_value` to `new_value` (with `value_source`), `action`, `confidence`, and `message`. Arabic-guarded entries are visually marked. This is the canonical reading order for a reviewer walking the deck.

#### 7.3.2 Summary counts per issue type

A table of every `issue_type` present with its count and a breakdown by `action` (flagged / changed / skipped). This is the at-a-glance picture of deck health and the basis for the per-fix-type precision metric.

#### 7.3.3 Severity rollup

Top-line counts: total `error`, `warning`, `info`, plus a per-module severity matrix (module x severity). Lets a reviewer triage by impact before reading slide by slide.

#### 7.3.4 PDF export

Report-only PDF generated from the stored record set using ReportLab or WeasyPrint. No slide rendering and no renderer dependency in v1: the PDF contains the severity rollup, the summary-counts table, and the slide-by-slide log, not slide thumbnails. This keeps v1 PDF export free of the COM/LibreOffice/Graph rendering stack. Slide-thumbnail or deck-to-PDF output is a later enhancement that would route through Microsoft Graph (`driveItem content?format=pdf`) as the primary renderer; it is out of scope for the v1 report. The PDF is retrieved as `report_pdf` via a signed URL from `GET /v1/jobs/{job_id}/outputs`.

#### 7.3.5 CSV export

Flat CSV, one row per record, columns mirroring the 7.1 schema (`record_id`, `slide_index`, `shape_id`, `module`, `issue_type`, `property`, `old_value`, `new_value`, `severity`, `action`, `confidence`, `arabic_flag`, `message`, `profile_rule_id`). Intended for spreadsheet triage, bulk review, and downstream analysis (e.g. precision tracking in Excel). UTF-8 with BOM so Arabic text in `old_value` / `new_value` / `message` renders correctly in Excel. The CSV is retrieved as `report_csv` via a signed URL from `GET /v1/jobs/{job_id}/outputs`.

#### 7.3.6 Report API surface

Reports and machine-readable findings are projections of the record set, exposed on the canonical REST API under `/v1` (the primary surface; the web UI is one client of it). Two distinct read surfaces exist: a detail endpoint and an aggregate summary endpoint.

| Endpoint | Returns | Roles |
|---|---|---|
| `GET /v1/jobs/{job_id}/findings` | Full `FindingRecord[]` (the record array), filterable by module, severity, and slide. | D, R, A, M |
| `GET /v1/jobs/{job_id}/summary` | Aggregate counts per `issue_type` and `severity` (JSON, for dashboards and the severity rollup). | D, R, A, M |
| `GET /v1/jobs/{job_id}/outputs` | Signed URLs for `report_pdf`, `report_csv`, `manifest_json` (and `cleaned_pptx` in v1.5). | D, R, A, M |

The report PDF, the CSV, and the change manifest (`manifest_json`) are retrieved only as signed URLs via `/outputs`. There are no direct `report.pdf`, `report.csv`, `/manifest`, or `artifacts/{kind}` endpoints. All report endpoints are behind the same Microsoft Entra ID sign-in gate and RBAC as the job: Designer, Reviewer, Admin, and the machine (Odoo) identity can view findings, summary, and outputs, with no profile edits implied (per the Appendix A.8 permission matrix).

### 7.4 Preview Mode

Preview is the pipeline run with writes disabled. It is a property of the Job (`preview: true`), not a separate mode: when `preview` is true, there are no writes even in `fix` mode. Audit and fix execute the SAME module code over the SAME inheritance resolver and the SAME detection logic; preview simply suppresses the XML write step and the cleaned-.pptx output. Consequences:

- **Audit and fix cannot drift.** There is one code path. A finding shown in preview is the exact finding the fix run would act on, with the same `old_value`, `new_value`, `confidence`, and Arabic guard. This is a correctness guarantee, not a convenience.
- **Flow B is preview with the report rendered and no output file.** Flow A's review step is the same preview surfaced before the write is committed.
- Preview is the default first stop for any `confidence` below `deterministic`, and for all `arabic_flag` records.

API: `POST /v1/jobs` accepts `{ upload_id, profile_id, modules[], mode, preview }`. The canonical `JobMode` is `audit` or `fix` (the term `format` is not used, and there is no `preview` mode value). A job submitted with `mode: "audit"`, or with `mode: "fix"` and `preview: true`, runs the pipeline and produces records with `action` values of `flagged` or `skipped` only (no `changed`), and no cleaned .pptx. A job with `mode: "fix"` and `preview: false` runs the same pipeline with writes enabled.

### 7.5 Per-Change Visual Before/After Diff (v1.5 Trust Surface)

The per-slide, per-change visual before/after diff is the real review surface, not the text manifest. It is the primary trust mechanism for v1.5 and the gate for enabling any auto-apply. In a 140-person shop, one mangled client deck spreads by word of mouth, so reviewers must SEE the change, not read a property delta.

Specification:

- For each record with `action` in (`changed`, candidate change), render the affected slide region before and after the proposed edit, side by side, with the changed shape highlighted.
- v1.5 rendering of the diff uses Microsoft Graph PDF/image conversion (M365-native, server-safe) for the before/after slide images; LibreOffice headless is fallback only and its output is labeled a fidelity-acceptable internal artifact, not pixel-faithful, given its Arabic RTL rendering limitations.
- Each diff is tied to its `record_id` so the visual, the record, and the accept/reject decision are the same object.

### 7.6 Granular Accept/Reject and "Apply Selected" Output

v1.5 gives the designer per-change control. The unit of decision is the `record_id`.

- The UI presents each candidate change with its before/after diff, severity, and confidence. The designer marks each accept or reject. Default selection: `deterministic` + non-Arabic pre-checked; everything else unchecked.
- `POST /v1/jobs/{job_id}/apply` with a body of `{ "record_ids": [ ... ] }` (the selected, accepted records) produces a new job and cleaned .pptx applying ONLY the selected records. The original is never overwritten and the preview output is never silently committed. This endpoint is available to Designer and Admin.
- The applied output carries its own resulting record set: accepted records become `action: changed`, rejected candidates are recorded as `action: skipped` with `message` noting designer rejection. This preserves a full audit trail of what the human decided.
- Per-`issue_type` precision and false-positive rate are computed from accept/reject decisions on the real-deck regression corpus and must clear the gate before that `issue_type` becomes eligible for auto-apply.

### 7.7 Inline-Comment Mode (Later Phase)

Optional inline-comment mode places notes directly on the slide (PowerPoint comments anchored to the relevant shape) rather than only in the external report. Scope and phasing:

- Deferred to a later phase (post v1.5), after the diff and accept/reject surfaces are proven. It is an output convenience, not a trust mechanism. Comments are exposed through `GET /v1/jobs/{job_id}/comments` and `POST /v1/jobs/{job_id}/comments` (which accepts `{ slide_index, record_id?, text }`), available to Designer, Reviewer, and Admin.
- Each comment is generated from a record (`message`, anchored via `shape_id` / `slide_index`) so the on-slide note and the record stay consistent.
- Comments are additive annotations only; this mode never alters slide formatting. It is compatible with Audit Only (Flow B), letting a reviewer hand back an annotated deck without any formatting change.
- Arabic-guarded findings appear as comments like any other, carrying the "Arabic content, manual review" message so the designer handles them by hand.

---

## 8. Arabic / RTL Safety

Arabic safety is a **hard requirement for v1**, not a Phase 2 enhancement. The MVP does not need full bidirectional layout logic, but it must reliably **detect** Arabic content and **guard** against direction-sensitive operations and font substitution that silently corrupt Arabic glyphs and shaping. An RTL-unaware font swap to a Latin-only brand font destroys Arabic text without raising any error, so the tool must err toward flagging over fixing on any Arabic hit. Given our client base (IB and government in KSA, UAE, Jordan, Qatar), a mangled Arabic deck is a credible adoption-killer, and the guard layer is the control that prevents it.

### 8.1 Scope (v1)

| Capability | v1 status |
|---|---|
| Detect Arabic at run / paragraph / text-frame level | In scope (hard requirement) |
| Guard direction-sensitive ops on Arabic content | In scope (hard requirement) |
| Guard font substitution on Arabic runs | In scope (hard requirement) |
| Resolve and report `a:cs` complex-script typeface | In scope (read/report) |
| Full RTL remediation (mirrored margins, RTL alignment, EN/AR pairing) | Out of scope, see v2 roadmap (§8.6) |

### 8.2 Detection

Detection runs inside the **existing single-pass traversal** (the same walk used by the Font, Margin/Alignment, Color, Shape, and Header/Footer modules, keyed `font`, `margin_alignment`, `color_palette`, `shape_size`, and `header_footer`). Each run, paragraph, and text-frame is evaluated as it is visited; there is no separate Arabic pass. A hit at any level propagates upward: an Arabic run marks its paragraph and text-frame as Arabic-bearing for guard purposes.

**Signal 1: Unicode block ranges.** A run or text node is Arabic-bearing if any character falls in these blocks:

| Block | Range | Notes |
|---|---|---|
| Arabic | `U+0600`–`U+06FF` | Core letters, diacritics, digits |
| Arabic Supplement | `U+0750`–`U+077F` | Extended letters for non-Arabic languages |
| Arabic Extended-A | `U+08A0`–`U+08FF` | Additional letters and marks |
| Arabic Presentation Forms-A | `U+FB50`–`U+FDFF` | Ligatures, contextual forms |
| Arabic Presentation Forms-B | `U+FE70`–`U+FEFF` | Contextual/joined glyph forms |

**Signal 2: paragraph RTL attribute.** A paragraph is RTL when `a:pPr/@rtl="1"` (or `rtl="true"`). This catches paragraphs declared RTL even when a specific run holds neutral characters (digits, punctuation, Latin), and is read directly from the XML because python-pptx does not expose `rtl` as a typed property.

A unit is treated as Arabic-bearing if **either** signal is present. Detection is intentionally conservative: false positives cost a manual-review flag, false negatives risk a corrupted deliverable.

### 8.3 Guard rules

On any Arabic hit, the affected operations are **skipped or flagged** rather than applied. The two protected categories:

**Direction-sensitive geometry / layout operations**
- Alignment normalization (paragraph alignment, shape-to-guide alignment, cross-shape alignment)
- Margin and safe-zone enforcement (left/right inset semantics invert under RTL)
- Footer repositioning and header/footer alignment

**Font operations**
- Font-family substitution and unification on Arabic runs (a Latin-only brand font has no Arabic glyphs and breaks shaping)
- Weight/style remapping that would route an Arabic run to a Latin-only face

Size-only changes that do not alter the typeface may proceed, but only behind preview and never as part of a substitution that changes the font family on an Arabic run.

**Record tagging.** Every guarded operation emits a standard `FindingRecord` (per the canonical schema in Appendix A.2) with the cross-cutting `issue_type` `arabic.guarded_operation` and the following field values, aligned to the default severity for an Arabic-guarded operation in Appendix A.4:

| Field | Value |
|---|---|
| `action` | `skipped` |
| `arabic_flag` | `true` |
| `severity` | `warning` |
| `message` | `Arabic content, manual review` |
| `confidence` | unchanged from module default; the flag, not confidence, drives the skip |

The guarded record is always tagged `severity = warning` with `action = skipped` and `arabic_flag = true`; it is never emitted as `info`. `arabic_flag` is a first-class boolean on the `FindingRecord` so the report and the `GET /v1/jobs/{job_id}/summary` view (aggregate counts per `issue_type` and `severity`) can filter and count Arabic-guarded items directly. The slide-by-slide issue log surfaces these as a distinct "Arabic / RTL, manual review" group so a designer can triage them in one place.

### 8.4 Complex-script typeface handling (font module)

Arabic does not use the Latin typeface slot. In OOXML, a run's font is declared across separate child elements of `a:rPr`:

- `a:latin` (Latin / `+mn-lt` / `+mj-lt`), what python-pptx `Font.name` reads and writes
- `a:cs` (complex script), the typeface that actually renders Arabic
- `a:ea` (East Asian), not in our scope

Because `Font.name` only touches `a:latin`, the font module (`font`) MUST read and write `a:cs` via **raw XML**, using the shared inheritance resolver (run `rPr` → paragraph/list level → `txBody` list style → placeholder layout → placeholder master → master text styles → theme defaults) to determine the **effective** complex-script face when `a:cs` is absent on the run.

Module behavior:

| Situation | Behavior |
|---|---|
| Run is Arabic-bearing | Do not substitute the family. Report effective `a:cs` and effective `a:latin` separately; set `arabic_flag = true` if any substitution would have applied, with `severity = warning` and `action = skipped` per §8.3. |
| Audit, mixed EN/AR run group | Report `a:latin` and `a:cs` as distinct `FindingRecord` items so the designer sees the real Arabic face, not the Latin slot. |
| `a:cs` resolves to a Latin-only family on Arabic content | Hard-flag with `severity = error` (issue_type `font.cs_typeface_missing`): this is an existing corruption-in-waiting in the source deck. |

The font compliance check therefore evaluates the **complex-script** typeface against the profile's Arabic font (`font.roles.{role}.complex_script[]`) separately from the Latin typeface (`font.roles.{role}.latin[]`), even in v1 where remediation is held back.

### 8.5 Required test fixtures

Ship these as committed fixtures in the regression corpus before the Arabic guard is considered done. Each must assert that guarded operations produce records with `arabic_flag = true`, `severity = warning`, and `action = skipped`, and that no Arabic run's typeface is altered.

| Fixture | Contents | Asserts |
|---|---|---|
| `mixed_en_ar_title_body.pptx` | Slide with EN+AR in the same title and the same body paragraph (mixed-direction runs) | Per-run detection; `a:cs` reported separately from `a:latin`; no substitution on AR runs |
| `rtl_mirrored_layout.pptx` | Slide with `a:pPr/@rtl="1"` paragraphs and an RTL-mirrored placeholder arrangement | RTL attribute detection; alignment and margin/safe-zone ops skipped/flagged |
| `ar_footer.pptx` | Deck with an Arabic footer across slides | Footer repositioning and header/footer alignment guarded; `arabic_flag` on footer records |

Add at least one **neutral-character edge case** (an RTL-declared paragraph containing only digits/punctuation) to confirm Signal 2 catches what Signal 1 alone would miss.

### 8.6 v2 roadmap: full RTL rules

v2 replaces guards with real bidirectional handling, gated on precision metrics from the real-deck regression corpus before any auto-apply:

- **Mirrored margins / safe zones:** interpret left/right insets and safe-zone edges (`geometry.safe_zone_margins_emu`) under RTL so "outer" and "inner" margins map correctly to the reading direction. **Shipped for the BINDING margin (20/08/2026):** wherever a pass seats content ON a margin rather than merely inside one, the margin it binds to follows the script, judged per slide (`qc.util.slide_is_rtl` — an explicit `rtl` paragraph, else the script carrying most of the slide's letters). That covers the migration's content-block move (`qc.migrate`) and the audit's select-all rescale (`margin_alignment.content_overflow`). The four-sided breach checks were always symmetric and need no mirroring.
- **RTL alignment semantics:** treat `start`/`end` per paragraph direction (RTL "start" = right edge), and apply alignment normalization with direction awareness instead of skipping it. **Shipped for text the migration MOVES (21/08/2026):** moving text into a placeholder carries the words and drops everything else by design — the placeholder must style it from the master — but DIRECTION is not styling. The source paragraph's `rtl` and explicit `algn` are carried across, and where the source states neither, Arabic text is marked RTL and started on the right. Without this, Arabic headings inherited an English master's left-to-right paragraph and every Arabic deck came back reading from the wrong edge; which way a language reads is a fact about the language, not a house style the master gets to overrule.
- **EN/AR font pairing per run:** unify the Latin face and the Arabic (`a:cs`) face independently per run from a paired profile entry (e.g. brand Latin + brand Arabic via `font.roles.{role}.latin[]` and `font.roles.{role}.complex_script[]`), so mixed runs render both scripts correctly.
- **Admin UI support:** the `Profile` schema gains a paired Arabic typeface and RTL margin convention per client/project, managed through `POST /v1/profiles` and `PUT /v1/profiles/{id}` (Admin only).

Engine note: if the technical spike shows python-pptx cannot perform RTL-aware layout reassignment cleanly, this is the area most likely to justify Aspose.Slides, which handles RTL and layout reassignment more completely.

### 8.7 Rendering caveat

LibreOffice headless **mis-renders Arabic** (documented RTL shaping/order bugs plus font substitution), so its PDF is not trustworthy for Arabic decks even as an internal artifact. For any deck flagged Arabic-bearing:

- **Prefer Microsoft Graph** (`driveItem` `content?format=pdf`) for thumbnails and deck-to-PDF, consistent with the M365/SSO stack.
- If Graph is unavailable, **skip slide thumbnails** for Arabic decks and fall back to the report-only PDF (issue logs and counts, no rendered slides) rather than ship a mis-rendered preview.
- Ensure canonical brand **Arabic fonts are installed and embedded** on any renderer path.

The visual before/after diff is the primary review surface for designers; for Arabic content it must be sourced from Graph (or omitted), never from LibreOffice output.

---

## 9. Non-Functional Requirements (Performance, Security, Reliability)

These NFRs are binding acceptance criteria, not aspirations. Two of them are hard guarantees that gate release: the original file is never modified, and client uploads are deleted after processing. Everything else is measured and reported.

A note on hosting before the detail below: data residency (KSA/UAE government client constraints), not technology, decides on-prem versus Azure. That decision is owned by Operations and IT-security and must be confirmed before infrastructure is provisioned. The requirements here are written to hold under either deployment.

### 9.1 Performance

The system must process decks up to 200 slides without timing out. The mechanism is the async job queue (upload to object storage via `POST /v1/uploads`, create a `Job` via `POST /v1/jobs`, enqueue, worker runs the pipeline, client polls `GET /v1/jobs/{job_id}` or subscribes via SSE, then downloads via signed URLs from `GET /v1/jobs/{job_id}/outputs`). No request holds an HTTP connection for the duration of a deck. This sidesteps idle-read timeouts (for example the nginx 60s default) by design rather than by tuning.

The slow part is rendering, not XML work. python-pptx parse and surgical edits on a 200-slide deck are fast and in-process. High-fidelity render-to-PDF and per-slide thumbnail generation (Microsoft Graph `driveItem` `content?format=pdf`, or LibreOffice headless as the internal-only fallback) dominate wall-clock time and are the only operations that warrant concurrency control. Size targets accordingly.

**Target processing times (p95, single deck, warm worker):**

| Operation | 50 slides | 200 slides |
|---|---|---|
| Parse + full audit (all six modules, write disabled, `mode = audit`) | < 8 s | < 30 s |
| Parse + safe deterministic fix pass (`mode = fix`, v1.5+) | < 12 s | < 45 s |
| Report-only PDF (issue log, no thumbnails) | < 3 s | < 6 s |
| High-fidelity render to PDF / thumbnails (per slide) | < 1.5 s/slide | < 1.5 s/slide |

- These are targets to validate, not measured fixtures. The 2-week technical spike must retire the 200-slide processing-time risk on real client decks and replace these numbers with observed values before the timeline is committed.
- Time-to-first-feedback matters more than total time for the audit flow. The `Job` exposes `slides_total` and `slides_processed`, so the UI shows movement from the first slide, not a spinner until completion.
- Audit and fix run the same pipeline (write toggled off when `mode = audit`), so audit timing is a reliable lower bound for the fix pass.

**Concurrency limits on any renderer:**

- The renderer is the scarce resource and must sit behind an explicit concurrency cap (configurable, default tuned to host vCPUs, suggested starting point 2 to 4 concurrent render jobs per worker).
- Microsoft Graph PDF conversion is rate-limited upstream; respect `Retry-After` and back off rather than retrying tight.
- LibreOffice headless is single-document-per-process in practice; serialize through a worker pool, never invoke ad hoc per request.
- PowerPoint COM / Office server-side automation is prohibited (Microsoft does not support unattended server automation: single-instance, hangs). It must not appear in any code path.

### 9.2 Security and Confidentiality (v1 requirement)

Client decks include investment-bank and government material and are treated as highly sensitive. Confidentiality controls below are v1 scope, not a later phase. An unauthenticated "internal-only" endpoint is reachable on the network and is not an adequate control; there is no version of this tool without an auth gate.

**Authentication and authorization**

- **Sign-in gate:** Microsoft Entra ID (OIDC) on the M365 tenant. All human access requires interactive sign-in. No anonymous endpoints, including health-adjacent or "preview" routes that touch deck data.
- **RBAC:** driven by Entra app roles delivered in the token `roles` claim. The three human roles match the canonical permission matrix in Appendix A.8:
  - `Designer`: upload, create, and cancel own jobs; view findings, summary, and outputs; comment; apply selected fixes (v1.5).
  - `Reviewer`: view findings, summary, and outputs; comment; sign off. No job execution, no profile edits.
  - `Admin`: all Designer capabilities plus sign-off, manage profiles, and manage users/roles.
- **Machine callers (Odoo and other services):** OAuth2 client-credentials as an application identity, not a human role. Each service principal is authorized explicitly; a machine token never inherits Designer/Reviewer/Admin and is limited to upload, create, cancel, and read access per Appendix A.8 (no comment, no sign-off, no fix apply, no profile or user management). The REST API is the primary surface and enforces RBAC identically for UI and API clients.

**Encryption**

- **In transit:** TLS 1.2 minimum (prefer 1.3) on every hop, including worker-to-object-storage and worker-to-renderer. No plaintext internal traffic, even on a private subnet.
- **At rest:** uploads, outputs, manifests, and audit logs encrypted at rest (platform-managed keys acceptable for v1; customer-managed keys to be revisited if Operations/IT-security require it for government work).

**Upload lifecycle and retention**

- Uploaded source decks auto-delete after processing completes. This is a v1 requirement.
- Retention is configurable: default retention for source uploads is 0 (delete immediately on successful job completion, that is when the `Job` reaches `status = ready`). Generated outputs (`cleaned_pptx`, `report_pdf`, `report_csv`) and the `manifest_json` follow a separate configurable retention window (proposed default 7 days) so designers can re-download via `GET /v1/jobs/{job_id}/outputs`, after which they are purged.
- Failed jobs (`status = failed`) delete the source on the same policy; a failed job must not leave client material lingering in object storage.
- Retention values are confirmed with Operations/IT-security and may be overridden downward (never silently upward) by data-residency policy.

**Audit logging**

- Record who processed what: actor identity (Entra `oid`/`upn` or service principal id), `job_id`, action (`upload`, `audit`, `fix`, `download`, `delete`, `profile_change`), source file name and hash, selected `profile_id` and resolved `profile_version`, timestamp (UTC), and outcome.
- Audit logs are append-only, access-restricted to Admin, and retained independently of the deck retention window.

**No third-party data egress**

- Deck content stays inside the chosen boundary (on-prem or the Prezlab Azure tenant). No client data is sent to any third-party SaaS for processing.
- The single permitted external dependency is Microsoft Graph for PDF rendering, which is inside the M365 tenant and the same trust boundary, not a third party. If data residency rules forbid even Graph round-trips for a given client, the LibreOffice headless fallback (fully in-boundary) is used and the deck is flagged accordingly.
- Aspose.Slides, if adopted as the engine, must run as an in-process / in-boundary library with no telemetry or cloud callout enabled.

**Data residency (deciding constraint, to confirm)**

- KSA and UAE government client rules may require that data never leaves the country or never leaves on-prem infrastructure. This constraint, not the technology stack, decides on-prem versus Azure hosting and may also restrict Microsoft Graph usage.
- Action: Operations and IT-security confirm data-handling, retention, and residency obligations per client tier before hosting is finalized. The build must support both deployment targets so this decision is not blocked by code.

### 9.3 Reliability

The non-negotiable: **the original file is never overwritten.** This is a hard guarantee enforced architecturally, not a convention. The source upload is immutable in object storage; every output is written to a new key. There is no code path that writes back to the source object. Verified by automated test on every build.

| Property | Requirement |
|---|---|
| Idempotency | Jobs are idempotent. Re-running a `job_id` (retry or duplicate enqueue) produces the same outputs and does not double-apply changes or duplicate `FindingRecord` entries. Operate surgically: touch only the XML being changed. |
| Retries | Transient failures (renderer timeout, storage blip, Graph throttling) retry with bounded exponential backoff and a max attempt count. Deterministic failures (corrupt .pptx, no usable master) fail fast, no retry, and set the `Job` to `status = failed` with a populated `error`. |
| Original immutability | Source object is immutable; outputs go to new keys. Enforced and tested. |
| Graceful failure | A module failure does not abort the whole job. The pipeline continues and emits a partial record set covering modules that completed, with failed modules recorded at module level plus the slide/shape scope reached. The user always gets a result they can act on. |
| Defensive guards | The "no usable master" guard (third-party exports from Canva/Gamma/Google Slides/Keynote can be non-conformant) and the mandatory pre-flight (SmartArt/charts/media) fail the affected operation cleanly with a flag, not a crash. |

The `FindingRecord` set is the source of truth for what happened. Each module emits structured records carrying, among the full contract in Appendix A.2, `slide_index` (zero-based), `shape_id` (string), `shape_path`, `module`, `issue_type`, `property`, `old_value`, `new_value`, `severity` (three-valued: `error`, `warning`, `info`), `action`, `confidence`, `arabic_flag`, and `profile_rule_id`, stored as JSON/JSONB. The change manifest, audit report, summary, and all UI surfaces are projections of this record set. It is never reconstructed by diffing the two .pptx zips, since an open/save round-trip rewrites untouched XML and produces noisy false diffs.

### 9.4 Observability

The tool's credibility in a 140-person shop depends on measured precision, not claimed precision. One mangled client deck kills adoption by word of mouth, so false-positive and false-negative rates are first-class metrics, tracked per module and gated before any fix type is allowed to auto-apply.

**Job-level metrics**

- Job throughput, queue depth, p50/p95/p99 processing time by deck size band.
- Per-stage timing (parse, audit, fix, render) to keep the render-is-the-bottleneck assumption honest.
- Renderer concurrency utilization and Graph throttle/`Retry-After` rate.

**Error metrics**

- Job error rate by failure class (corrupt file, no master, pre-flight block, renderer failure, internal error).
- Partial-record-set rate (jobs that completed with at least one module in error).

**Per-module quality metrics (the adoption gate)**

- Per-module precision and false-positive rate, measured against a real-deck regression corpus with known-good labels. Numeric precision is tracked as an aggregate metric per module, separate from the per-record `confidence` field.
- Tracked separately for `flagged` (detection) and `changed` (remediation) actions, because detection is mostly objective and remediation is often a judgment call. These two values are the canonical `Action` enum from Appendix A.1.
- A given fix type is enabled for auto-apply only after it clears a precision threshold on the corpus (threshold owned by the pilot designers as co-owners) and only for records where `confidence = deterministic`. Inference-heavy operations (nearest-color replace, font-size hierarchy inference, master synthesis, alignment intent) stay flag/preview regardless of score until proven.
- Arabic-content items are tracked as a distinct slice via `arabic_flag = true`; any direction-sensitive operation or font swap on Arabic runs must show as guarded (`action = skipped`, severity `warning`), never silently changed.

### 9.5 Compatibility

Outputs must open cleanly in the tools designers and clients actually use. "Opens cleanly" means no repair prompt, no missing-content warning, no visible corruption, in both:

- **Microsoft PowerPoint** (desktop, current supported M365 channel).
- **Microsoft 365 web** (PowerPoint for the web).

Requirements:

- Every `cleaned_pptx` output is validated to open without a repair dialog in both targets. This is part of the spike exit criteria (safe write that opens cleanly in PowerPoint and M365 web) and a standing regression check thereafter, run on real client decks.
- Fonts: install canonical brand and Arabic fonts on any renderer, and embed fonts in the cleaned .pptx so the deck renders correctly on machines that lack them.
- **Pre-flight flag (mandatory):** before processing, detect and flag decks containing **SmartArt, advanced charts, or embedded media**. These produce `FindingRecord` entries with `issue_type = preflight.unmodifiable_content` and severity `info` per Appendix A.4. The parts are preserved verbatim on round-trip (a plain open/save through python-pptx does not routinely flatten SmartArt to an image) but are not editable by this tool. The pre-flight tells the designer up front "this deck contains elements we will not modify," so nothing is silently skipped and no one expects a fix that cannot be delivered.
- Round-trip preservation of unmapped parts (SmartArt, embedded media, customXml, OLE) is validated on real client decks during the spike and kept as a regression check.

---

## 10. API Specification

The REST API is the primary product surface. The web UI is one client of it, and Odoo is a second, machine-to-machine client. Every capability available in the UI is reachable through the API, so we never build behavior that only the browser can invoke. All endpoints are versioned under the base path `/v1`, accept and return `application/json` unless noted, and use ISO 8601 timestamps in UTC.

### 10.1 Design principles

- **API-first.** Build and document endpoints before UI work; the UI consumes the same contracts external callers do.
- **Async by default.** Job creation returns immediately with a Job resource. Long-running audit/fix work happens in the worker (see Architecture). Clients poll `GET /v1/jobs/{job_id}` or subscribe to SSE; no endpoint holds a connection open for the duration of processing.
- **Audit-first phasing carried into the API.** v1 exposes job creation in `audit` mode, status, findings, summary, and report outputs (signed URLs). `POST /v1/jobs/{job_id}/apply` (per-change accept/reject) ships in v1.5. The Odoo integration pattern ships in v2.
- **Same pipeline, write toggle.** `mode: audit` and `preview: true` run the identical detection pipeline with writes disabled, so audit results and fix proposals cannot drift from each other.

### 10.2 Authentication and authorization

Two distinct identity types. Human users authenticate with Microsoft Entra ID (OIDC); machine callers (Odoo) use OAuth2 client-credentials as an application identity. Every request carries `Authorization: Bearer <token>`. Tokens are validated against Entra (issuer, audience, signature, expiry) on every call.

| Caller | Mechanism | Identity model | Authorization source |
|---|---|---|---|
| Human users (web UI) | Entra ID OIDC (authorization code + PKCE) | A signed-in person | Entra **app roles** delivered in the token `roles` claim |
| Machine callers (Odoo, future services) | OAuth2 **client-credentials** | An application (service principal), not a person | Service principal explicitly granted as a registered application identity; no human role assumed |

**App roles (the `roles` claim) and endpoint authorization** (aligned to Appendix A.8; D = Designer, R = Reviewer, A = Admin, M = machine/application identity):

| Role | Capabilities | Endpoints authorized |
|---|---|---|
| **Designer** | Upload, create, and cancel own jobs; view findings, summary, and outputs; comment; apply selected fixes (v1.5) | `POST /v1/uploads`, `POST /v1/jobs`, `GET /v1/jobs/{job_id}`, `POST /v1/jobs/{job_id}/cancel` (own), `GET /v1/jobs/{job_id}/findings`, `GET /v1/jobs/{job_id}/summary`, `GET /v1/jobs/{job_id}/outputs`, `POST /v1/jobs/{job_id}/apply`, `GET /v1/jobs/{job_id}/comments`, `POST /v1/jobs/{job_id}/comments`, `GET /v1/profiles`, `GET /v1/profiles/{id}` |
| **Reviewer** | View findings, summary, and outputs; comment; sign off; no job execution, no profile edits | `GET /v1/jobs/{job_id}`, `GET /v1/jobs/{job_id}/findings`, `GET /v1/jobs/{job_id}/summary`, `GET /v1/jobs/{job_id}/outputs`, `GET /v1/jobs/{job_id}/comments`, `POST /v1/jobs/{job_id}/comments`, `POST /v1/jobs/{job_id}/signoff`, `GET /v1/profiles`, `GET /v1/profiles/{id}` |
| **Admin** | Everything above plus cancel any job, sign off, and profile/user management | All endpoints, including `POST /v1/jobs/{job_id}/cancel` (any), `POST /v1/jobs/{job_id}/signoff`, `POST /v1/profiles`, `PUT /v1/profiles/{id}` |
| **Machine (Odoo)** | Upload, create jobs, and read results under an application identity; no comment, no sign-off, no apply, no profile management | `POST /v1/uploads`, `POST /v1/jobs`, `GET /v1/jobs/{job_id}`, `POST /v1/jobs/{job_id}/cancel` (own), `GET /v1/jobs/{job_id}/findings`, `GET /v1/jobs/{job_id}/summary`, `GET /v1/jobs/{job_id}/outputs` |

Authorization is enforced server-side per endpoint from the validated token claims, never from a client-supplied role field. A missing or insufficient role returns `403`. A missing or invalid token returns `401`. Machine callers are authorized explicitly as a registered service principal and are not mapped to a human role.

**Confidentiality controls apply at the API layer (v1, not deferred):** all transport over TLS; uploads encrypted at rest; uploaded source `.pptx` auto-deleted after processing per the retention policy; output download is only via short-lived signed URLs scoped to the job. There is no unauthenticated "internal" endpoint.

### 10.3 Endpoints

The canonical endpoint set is defined in Appendix A.7 and reproduced here. All paths are under `/v1`. Roles in brackets follow A.8 (D = Designer, R = Reviewer, A = Admin, M = machine/application identity).

| Method & path | Purpose | Roles | Phase |
|---|---|---|---|
| `POST /v1/uploads` | Stream a `.pptx` to storage; returns `{ upload_id, expires_at }` | D, A, M | v1 |
| `POST /v1/jobs` | Create a job `{ upload_id, profile_id, modules[], mode, preview }` | D, A, M | v1 |
| `GET /v1/jobs/{job_id}` | Job status, progress, error | D, R, A, M | v1 |
| `POST /v1/jobs/{job_id}/cancel` | Cancel a queued or processing job | D (own), A | v1 |
| `GET /v1/jobs/{job_id}/findings` | Full `FindingRecord[]` (filterable by module, severity, slide) | D, R, A, M | v1 |
| `GET /v1/jobs/{job_id}/summary` | Aggregate counts per `issue_type` and `severity` | D, R, A, M | v1 |
| `GET /v1/jobs/{job_id}/outputs` | Signed URLs for `cleaned_pptx`, `report_pdf`, `report_csv`, `manifest_json` | D, R, A, M | v1 (report_pdf/csv, manifest_json); `cleaned_pptx` v1.5 |
| `POST /v1/jobs/{job_id}/apply` | Apply selected fixes `{ record_ids[] }`, returns a new job/output | D, A | v1.5 |
| `GET /v1/jobs/{job_id}/comments` | List inline comments | D, R, A | v2 |
| `POST /v1/jobs/{job_id}/comments` | Add a comment `{ slide_index, record_id?, text }` | D, R, A | v2 |
| `POST /v1/jobs/{job_id}/signoff` | Reviewer/Admin sign-off on a job | R, A | v2 |
| `GET /v1/profiles` | List profiles | D, R, A | v1 |
| `GET /v1/profiles/{id}` | Get a profile | D, R, A | v1 |
| `POST /v1/profiles` | Create a profile | A | v2 (seeded in v1) |
| `PUT /v1/profiles/{id}` | Update a profile (increments `version`) | A | v2 |

Reports and the change manifest are retrieved only as signed URLs via `/outputs`. There are no direct `report.pdf`, `report.csv`, or `artifacts/{kind}` endpoints. The single upload contract is `POST /v1/uploads` returning `upload_id`; `storage_key` is internal and never returned.

#### POST /v1/uploads

Stream a source `.pptx` to object storage. Kept separate from job creation so `POST /v1/jobs` stays JSON and idempotent. The original is never overwritten.

**Auth:** Designer, Admin, or machine (Odoo).

**Response `201 Created`:**

```json
{
  "upload_id": "upl_19b7e0",
  "expires_at": "2026-06-30T10:12:04Z"
}
```

**Errors:** `403` role not permitted; `413` upload exceeds size/slide limit; `415` not a `.pptx`.

#### POST /v1/jobs

Create an audit or fix job. Returns immediately with a `queued` Job.

**Auth:** Designer, Admin, or machine (Odoo).

**Request body:**

```json
{
  "upload_id": "upl_19b7e0",
  "profile_id": "prof_8f3a2c",
  "modules": ["master_slide", "font", "margin_alignment", "color_palette", "shape_size", "header_footer"],
  "mode": "audit",
  "preview": true
}
```

| Field | Type | Required | Notes |
|---|---|---|---|
| `upload_id` | string | yes | Identifier from `POST /v1/uploads` (source `.pptx` already in object storage). The original is never overwritten. |
| `profile_id` | string | yes | Per-client/per-project profile to evaluate against. The resolved version and snapshot are stamped onto the Job. |
| `modules` | ModuleKey[] | no | Subset of `ModuleKey` values (`master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer`) for targeted runs. An empty array (or omission) runs all modules. |
| `mode` | JobMode | yes | `audit` (v1) or `fix` (v1.5+). |
| `preview` | boolean | no (default `true` for `fix`) | When `true`, the pipeline runs with writes disabled and emits proposed changes only. |

A mandatory pre-flight runs at job start and records a `FindingRecord` (`issue_type: "preflight.unmodifiable_content"`, severity `info`) when the deck contains SmartArt, advanced charts, or embedded media we will not modify.

**Response `202 Accepted`** (the canonical `Job` shape, Appendix A.5):

```json
{
  "job_id": "job_4d2f91",
  "status": "queued",
  "mode": "audit",
  "preview": true,
  "profile_id": "prof_8f3a2c",
  "profile_version": 7,
  "profile_snapshot": { },
  "modules": ["master_slide", "font", "margin_alignment", "color_palette", "shape_size", "header_footer"],
  "slides_total": 0,
  "slides_processed": 0,
  "created_at": "2026-06-30T09:12:04Z",
  "started_at": null,
  "finished_at": null,
  "error": null
}
```

**Errors:** `400` invalid `ModuleKey` or mode/preview combination; `403` role not permitted; `404` `profile_id` or `upload_id` not found; `413` upload exceeds size/slide limit.

#### GET /v1/jobs/{job_id}

Poll job status and per-slide progress.

**Auth:** Designer/Reviewer/Admin (Designer limited to own jobs); machine (Odoo) for jobs it created.

**Response `200`** (canonical `Job`):

```json
{
  "job_id": "job_4d2f91",
  "status": "processing",
  "mode": "audit",
  "preview": true,
  "profile_id": "prof_8f3a2c",
  "profile_version": 7,
  "profile_snapshot": { },
  "modules": ["master_slide", "font", "margin_alignment", "color_palette", "shape_size", "header_footer"],
  "slides_total": 200,
  "slides_processed": 84,
  "created_at": "2026-06-30T09:12:04Z",
  "started_at": "2026-06-30T09:12:07Z",
  "finished_at": null,
  "error": null
}
```

`status` is one of `queued`, `processing`, `ready`, `failed`, `cancelled` (`JobStatus`). The progress bar is driven by `slides_processed` over `slides_total`. An SSE variant (`GET /v1/jobs/{job_id}/events`) streams the same status and progress transitions for live UI without polling.

#### POST /v1/jobs/{job_id}/cancel

Cancel a job that is still `queued` or `processing`. Idempotent: cancelling an already-terminal job returns the current Job unchanged.

**Auth:** Designer (own jobs) or Admin (any job); machine (Odoo) for its own jobs.

**Response `200`:** the canonical `Job` with `status: "cancelled"` and `finished_at` set.

**Errors:** `403` role not permitted; `404` job not found; `409` job already in a terminal state when cancellation is not applicable.

#### GET /v1/jobs/{job_id}/findings

Return the full `FindingRecord[]` for the job: the system of record from which the manifest, report, and summary are projected (Appendix A.2). One `FindingRecord` per detected issue or applied change. Records are emitted directly by the modules and never reconstructed by diffing `.pptx` files.

**Auth:** Designer/Reviewer/Admin; machine (Odoo).

**Query parameters:** `slide_index`, `module`, `severity`, `action`, `confidence`, `page`, `page_size`.

**Response `200`:** paginated `findings` using the canonical `FindingRecord` shape:

```json
{
  "findings": [
    {
      "record_id": "5f1c0c0e-1f3a-4b2c-9c3a-1a2b3c4d5e6f",
      "job_id": "job_4d2f91",
      "slide_index": 7,
      "shape_id": "12",
      "shape_path": null,
      "module": "font",
      "issue_type": "font.size_off_role",
      "property": "rPr.sz",
      "old_value": "16pt",
      "new_value": "18pt",
      "severity": "warning",
      "action": "flagged",
      "confidence": "high",
      "arabic_flag": false,
      "profile_rule_id": "font.body.size_pt",
      "message": "Body text 16pt is below the role minimum of 18pt.",
      "created_at": "2026-06-30T09:13:02Z"
    },
    {
      "record_id": "7a2d1b2c-3e4f-5061-7182-93a4b5c6d7e8",
      "job_id": "job_4d2f91",
      "slide_index": 11,
      "shape_id": "5",
      "shape_path": "12/3",
      "module": "color_palette",
      "issue_type": "color_palette.off_palette_rgb",
      "property": "solidFill.srgbClr.val",
      "old_value": "#1A1A1A",
      "new_value": null,
      "severity": "error",
      "action": "skipped",
      "confidence": "medium",
      "arabic_flag": true,
      "profile_rule_id": "color_palette.named.dk1",
      "message": "Off-palette fill on Arabic content; guarded for manual review.",
      "created_at": "2026-06-30T09:13:02Z"
    }
  ],
  "page": 1,
  "page_size": 100,
  "total": 119
}
```

The record shape is fixed across all six modules. `slide_index` is zero-based; `shape_id` is the OOXML `p:cNvPr @id` as a string; the identifier is `record_id` (not `id`); the human-readable field is `message` (not `note`). `severity` is three-valued (`error`, `warning`, `info`); `confidence` is the enum `deterministic | high | medium | low`. Direction-sensitive operations on Arabic runs are emitted with `arabic_flag: true` and `action: skipped` (`issue_type: "arabic.guarded_operation"` where applicable).

#### GET /v1/jobs/{job_id}/summary

Aggregate counts over the `FindingRecord` set per `issue_type` and `severity`, for the issue log and summary screens.

**Auth:** Designer/Reviewer/Admin; machine (Odoo).

**Response `200`:**

```json
{
  "job_id": "job_4d2f91",
  "by_severity": { "error": 6, "warning": 113, "info": 4 },
  "by_module": { "font": 71, "color_palette": 22, "margin_alignment": 18, "shape_size": 8 },
  "by_issue_type": {
    "font.size_off_role": 41,
    "font.family_out_of_set": 30,
    "color_palette.off_palette_rgb": 22
  },
  "total": 123
}
```

#### GET /v1/jobs/{job_id}/outputs

List downloadable artifacts as short-lived signed URLs. This is the only way to retrieve reports and the manifest; there are no direct report endpoints.

**Auth:** Designer/Reviewer/Admin; machine (Odoo).

**Response `200`:**

```json
{
  "job_id": "job_4d2f91",
  "outputs": [
    { "type": "report_pdf", "url": "https://...signed...", "expires_at": "2026-06-30T09:28:02Z", "bytes": 184221 },
    { "type": "report_csv", "url": "https://...signed...", "expires_at": "2026-06-30T09:28:02Z", "bytes": 20144 },
    { "type": "manifest_json", "url": "https://...signed...", "expires_at": "2026-06-30T09:28:02Z", "bytes": 33180 },
    { "type": "cleaned_pptx", "url": "https://...signed...", "expires_at": "2026-06-30T09:28:02Z", "bytes": 5532110 }
  ]
}
```

Output types depend on phase and mode: `report_pdf`, `report_csv`, and `manifest_json` exist for audit runs (v1); `cleaned_pptx` exists once changes are applied (v1.5+). URLs are time-limited and scoped to the requesting principal; the API never returns a permanent object-storage path.

#### POST /v1/jobs/{job_id}/apply  (v1.5)

Apply a designer-selected subset of proposed changes from a `fix` job and produce a cleaned `.pptx`. This is the granular accept/reject surface; only the records the designer selects are written.

**Auth:** Designer (own jobs) or Admin.

**Request body:**

```json
{
  "record_ids": [
    "5f1c0c0e-1f3a-4b2c-9c3a-1a2b3c4d5e6f",
    "9b8c7d6e-5f40-3122-1304-a5b6c7d8e9f0"
  ]
}
```

| Field | Type | Required | Notes |
|---|---|---|---|
| `record_ids` | string[] | yes | `FindingRecord.record_id` values to write. Auto-apply eligibility is limited to records with `confidence = deterministic`; inference-heavy records (nearest-color replace, font-hierarchy inference, alignment intent, master synthesis) are suggest-only and rejected if submitted. |

**Response `202 Accepted`:** a new apply Job (canonical `Job` shape) that produces a `cleaned_pptx` output on completion. The original source is untouched; edits are surgical (only the XML for accepted records is modified). Accepting a record tied to Arabic content that was guarded for manual review (`arabic_flag: true`, `action: skipped`) returns `409` unless an explicit override flag is supplied.

**Errors:** `400` unknown `record_id` or non-eligible (non-`deterministic`) record in `record_ids`; `403` role not permitted; `409` job not in `fix`/`preview` state or Arabic-guard conflict.

#### GET /v1/jobs/{job_id}/comments and POST /v1/jobs/{job_id}/comments

Inline collaboration on a job's findings. `GET` lists comments; `POST` adds one.

**Auth:** Designer, Reviewer, or Admin. Machine (Odoo) callers cannot comment.

**POST request body:**

```json
{
  "slide_index": 7,
  "record_id": "5f1c0c0e-1f3a-4b2c-9c3a-1a2b3c4d5e6f",
  "text": "Confirmed with the client; keep the larger body size."
}
```

`slide_index` is zero-based; `record_id` is optional (a comment can be slide-scoped rather than tied to a specific `FindingRecord`).

**Errors:** `403` role not permitted; `404` job, slide, or `record_id` not found.

#### POST /v1/jobs/{job_id}/signoff

Reviewer or Admin sign-off recording formal approval of a job's findings before delivery.

**Auth:** Reviewer or Admin only. Designers and machine (Odoo) callers cannot sign off.

**Response `200`:** the job's sign-off state with the signing principal and timestamp.

**Errors:** `403` role not permitted; `404` job not found; `409` job not in a `ready` state.

#### Profiles

Per-client/per-project profiles defining canonical fonts (Latin and complex-script/Arabic), the color palette (theme color slots plus named colors and tint/shade tolerances), and geometry/safe-zones. The full schema is the canonical `Profile` (Appendix A.6).

| Method | Path | Auth | Purpose |
|---|---|---|---|
| `GET` | `/v1/profiles` | Designer/Reviewer/Admin | List profiles (filter by client/project). |
| `GET` | `/v1/profiles/{id}` | Designer/Reviewer/Admin | Retrieve one profile. |
| `POST` | `/v1/profiles` | Admin | Create a profile (v2; seeded in v1). |
| `PUT` | `/v1/profiles/{id}` | Admin | Update a profile; increments `version` (v2). |

**Profile resource (canonical `Profile`, abbreviated):**

```json
{
  "id": "prof_8f3a2c",
  "name": "Gov KSA Deck Standard",
  "client_scope": "client_gov_ksa",
  "project_scope": null,
  "is_default": false,
  "version": 7,
  "owner": "admin:rana.k@prezlab.com",
  "created_at": "2026-05-02T10:00:00Z",
  "updated_at": "2026-06-12T14:30:00Z",
  "config": {
    "font": {
      "roles": {
        "title":    { "latin": ["Arial"], "complex_script": ["Dubai"], "size_pt": 36, "allowed_weights": ["bold"] },
        "subtitle": { "latin": ["Arial"], "complex_script": ["Dubai"], "size_pt": 24, "allowed_weights": ["regular"] },
        "body":     { "latin": ["Arial"], "complex_script": ["Dubai"], "size_pt": 18, "allowed_weights": ["regular"] },
        "caption":  { "latin": ["Arial"], "complex_script": ["Dubai"], "size_pt": 12, "allowed_weights": ["regular"] }
      },
      "theme_font_refs_allowed": true,
      "size_tolerance_pt": 0.5
    },
    "color_palette": {
      "theme_color_slots": ["dk1","lt1","dk2","lt2","accent1","accent2"],
      "named_colors": [
        { "name": "Brand Blue", "hex": "#0B5CAB", "theme_ref": "accent1", "allowed_tints": [0.0], "allowed_shades": [0.0] }
      ],
      "on_palette_mode": "by_name",
      "match_tolerance_deltaE": 2.0,
      "auto_replace_max_deltaE": 5.0,
      "ambiguity_band_deltaE": 10.0
    },
    "geometry": {
      "safe_zone_margins_emu": { "left": 457200, "right": 457200, "top": 457200, "bottom": 457200 }
    }
  }
}
```

Profiles carry both the complex-script (Arabic) typeface and the Latin typeface under `config.font.roles.{role}.{latin[],complex_script[]}`, so the font module never substitutes an Arabic run with a Latin-only brand font. Color matching uses CIEDE2000 nearest-match (`color_palette.match_tolerance_deltaE`), and valid theme-color references in `theme_color_slots` are treated as on-palette by name.

### 10.4 Odoo integration pattern (v2)

Odoo authenticates as an application identity via OAuth2 client-credentials and drives the QC tool through the same public API. The pattern, triggered from an Odoo client/project record:

1. **Pull active client/project.** Odoo resolves the relevant profile (`GET /v1/profiles`, filtered by client/project), or the `profile_id` is stored on the Odoo record.
2. **Upload and submit.** Odoo uploads the deck (`POST /v1/uploads`, capturing `upload_id`) and creates an audit job (`POST /v1/jobs`) with the resolved `upload_id` and `profile_id`.
3. **Await completion.** Odoo polls `GET /v1/jobs/{job_id}` until `status` is `ready` (or handles `failed`/`cancelled`).
4. **Attach results to the Odoo record.** Odoo retrieves the audit `report_pdf` and `manifest_json` via `GET /v1/jobs/{job_id}/outputs` (signed URLs) and attaches them to the originating client/project record.

No human role is assumed for these calls; the service principal is authorized explicitly as a registered application identity and is limited to upload, job submission, cancellation of its own jobs, and result retrieval. It cannot comment, sign off, apply fixes, or manage profiles.

### 10.5 Conventions and OpenAPI

- **Errors:** consistent envelope `{ "error": { "code": "string", "message": "string", "details": [] } }` with standard HTTP status codes (`400/401/403/404/409/413/415/422/429/500`). The `error` object matches the `Job.error` shape (`{ code, message }`) on failed jobs.
- **Idempotency:** `POST /v1/uploads`, `POST /v1/jobs`, and `POST /v1/jobs/{job_id}/apply` accept an `Idempotency-Key` header to make retries safe.
- **Pagination:** cursor or page/page_size on list endpoints (`/v1/profiles`, `/v1/jobs/{job_id}/findings`).
- **OpenAPI auto-generation:** the OpenAPI 3.1 spec is generated from the service code (request/response models and route definitions), published at `/v1/openapi.json` with interactive docs at `/v1/docs`. The contract is generated from the implementation rather than hand-maintained, which keeps the documentation, the UI client, and the Odoo client in sync as the API evolves across v1, v1.5, and v2.

---

## 11. Delivery Plan, Risks & Validation

The estimate is deliberately staged. We run a 2-week spike to retire the five highest-uncertainty unknowns before committing a timeline, then ship audit-first so designers get value (and we earn trust) before any bulk auto-fix touches a client deck. Effort is sized S/M/L/XL rather than in false-precision weeks, because several items (Arabic RTL, master enforcement) carry estimation risk that the spike is designed to resolve.

### 11.1 Technical Spike (2 weeks, before any timeline commitment)

The spike is a time-boxed engineering investigation, not a feature deliverable. Its only job is to convert five open unknowns into facts and to make the engine decision (python-pptx baseline vs Aspose.Slides) on evidence rather than assumption. Output is a written findings memo plus working proof-of-concept code, not production code.

**Five unknowns to retire**

| # | Unknown | What we must prove | Success criterion |
|---|---------|--------------------|-------------------|
| U1 | Master enforcement without corruption | Detect dominant layout, flag outliers, and enforce an explicitly chosen existing master/layout on a populated slide via surgical XML, without orphaning placeholders or breaking content binding (placeholders bind by `idx`). | On the test corpus, enforcing a chosen layout produces a file that opens cleanly with placeholder content intact in PowerPoint desktop and PowerPoint on the web. Any case we cannot do safely is reclassified to flag-only (`action = flagged`), with the boundary documented. |
| U2 | Safe write fidelity | A surgical open/edit/save (touch only the XML we change) preserves unmapped parts (SmartArt, embedded media, customXml, OLE, charts) verbatim and opens without repair prompts. | Zero "PowerPoint found a problem / wants to repair" prompts across the corpus in PowerPoint desktop and M365 web. SmartArt/charts/media survive byte-stable or visually identical. |
| U3 | 200-slide processing time | Full audit pass (all six modules, write disabled) plus a fix pass completes inside our async-job budget on the 150-plus-slide deck. | Audit pass and fix pass each complete without timeout under the async worker pattern; we record actual wall-clock per pass to size worker resourcing. Target working assumption: full audit of 200 slides in a few minutes, not tens of minutes. |
| U4 | Resolution + Arabic detection reliability | The shared inheritance resolver returns correct effective font family/size/weight and effective color (resolving `schemeClr` + `lumMod`/`lumOff`/`tint`/`shade` against `theme1.xml`), and Arabic detection fires correctly so we can guard direction-sensitive ops and complex-script (`a:cs`) typefaces. | Resolver matches PowerPoint's displayed values on a hand-checked sample with no `None`-leaks for inherited properties. Arabic runs are detected with no false negatives on the Arabic fixtures (a missed Arabic run is the dangerous failure), and detected records carry `arabic_flag = true`. |
| U5 | Renderer viability (Microsoft Graph) | Convert representative decks (at minimum the bilingual EN/AR deck and the 150-plus-slide deck) to PDF via Microsoft Graph (`driveItem` `content?format=pdf`); measure conversion time and throttling behavior; verify Arabic rendering fidelity; obtain an Operations/IT-security ruling on client decks transiting OneDrive/SharePoint under KSA/UAE data-residency constraints. The v1.5 visual before/after diff depends on this renderer. | Graph converts the corpus within the async-job budget with acceptable Arabic fidelity, and the residency ruling permits Graph for the deck categories we serve. If fidelity or residency fails, the v1.5 diff approach is re-planned before v1.5 is sized (LibreOffice is not an acceptable Arabic renderer per Section 8.7). |

**Test decks needed for the spike**

- Native PowerPoint deck built on a clean corporate master (the "happy path" baseline).
- One 150-plus-slide deck (real Prezlab deck, anonymized if needed) to exercise U3 and to surface scale-only failures.
- Bilingual EN/AR deck and an Arabic-only deck, including mixed-direction paragraphs and Arabic in placeholders, footers, and grouped shapes, to exercise U4 and the RTL guards.
- Brand-themed deck using `schemeClr` references with `lumMod`/`tint`/`shade` (not literal hex), to prove color resolution.
- Third-party export deck (Canva / Gamma / Google Slides / Keynote) to exercise the "no usable master" defensive guard (`issue_type = master_slide.no_usable_master`).
- Heavy-asset deck containing SmartArt, native charts, embedded media, and OLE objects, to exercise U2 round-trip preservation.

**How the spike gates the estimate and the engine decision**

- The estimate in 11.2 is provisional until the spike memo is signed off. If U1 or U3 fail or partially fail, scope moves down a phase (for example, master enforcement narrows to flag-only in v1, deferring enforcement to v2) and the affected sizings are re-rated.
- If U5 fails (Graph is ruled out on residency grounds, or its Arabic fidelity is unacceptable), the v1.5 visual before/after diff is re-planned before v1.5 commits. The diff is v1.5's long pole and its only viable renderers are Graph and LibreOffice, and LibreOffice is disqualified for Arabic decks; alternatives to evaluate would include shape-outline overlay diffs (no rendering) or restricting rendered diffs to non-Arabic decks.
- Engine decision rule: python-pptx (read/audit + surgical XML edits) is the baseline. We adopt Aspose.Slides (commercial license) only if the spike shows concrete python-pptx gaps that formatting quality depends on, specifically RTL handling and layout reassignment. Since formatting quality is the product, budget contingency for the Aspose license now, and confirm or release it at spike close. Decision criteria: (1) can python-pptx pass U1 and U2 safely; (2) does Aspose materially reduce Arabic/RTL corruption risk in U4; (3) license cost vs the engineering cost of building the same fidelity on raw XML.

### 11.2 Phasing and effort sizing

Audit-first. We ship detection, flagging, reporting, and preview before any bulk auto-fix, so designers trust the tool on read-only output before it ever writes. Sizes are relative effort (S < M < L < XL). For reference, the originally bundled auto-fix MVP was realistically ~10-14 weeks; splitting it into audit-first phases de-risks that estimate and delivers value earlier.

Staffing assumption (to confirm before the timeline is committed): the sizing and the ~8-week v1 target assume a dedicated team of two engineers (one senior backend engineer owning the OOXML engine and resolvers, one full-stack engineer owning the pipeline, API, and web UI) plus a part-time designer for the review surfaces. If actual staffing differs, re-rate the v1 target at spike close; the S/M/L/XL sizes stand independently of staffing.

**v1: Audit-only (target ~8 weeks, subject to spike)**

Detect, flag, report, preview. No bulk writes to deck content. Jobs run with `mode = audit`.

| Item | Size | Notes |
|------|------|-------|
| Async job pipeline (upload to object storage via `POST /v1/uploads`, Job lifecycle, worker, polling/SSE on `GET /v1/jobs/{job_id}`, signed-URL download via `GET /v1/jobs/{job_id}/outputs`, per-slide progress from `slides_processed`) | L | Foundational; everything else runs on it. |
| Shared inheritance + color resolver | L | Used by all six modules; correctness here is load-bearing (run `rPr` to paragraph/list level to `txBody` list style to placeholder(layout) to placeholder(master) to master text styles to theme/presentation defaults). CIEDE2000/Lab nearest-match for color. |
| Six audit modules (detect/flag only): `master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer` | XL | Detection is mostly objective across all six; this is the bulk of v1. |
| Arabic detection + direction-sensitive guards | M | Detect Unicode ranges (U+0600-06FF, U+0750-077F, U+08A0-08FF, presentation forms FB50-FDFF / FE70-FEFF) and/or `rtl='1'` on `a:pPr`; hard-flag and skip (`action = skipped`, `arabic_flag = true`) direction-sensitive ops and font substitution on Arabic runs. |
| Pre-flight scan (flags SmartArt / charts / media / OLE we will not modify, `issue_type = preflight.unmodifiable_content`) | S | Mandatory gate before any pass. |
| Structured `FindingRecord` set (JSONB records carrying `record_id`, `slide_index`, `shape_id`, `shape_path`, `module`, `issue_type`, `property`, `old_value`, `new_value`, `severity`, `action`, `confidence`, `arabic_flag`, `profile_rule_id`, `message`) | M | Emitted by modules, never reconstructed by zip-diffing. |
| Slide-by-slide issue log + per-`issue_type` summary via `GET /v1/jobs/{job_id}/summary`, severity (`error`/`warning`/`info`) | M | |
| Report-only PDF/CSV export (`report_pdf`, `report_csv` from `/outputs`; issue logs, counts, no thumbnails) | S | ReportLab/WeasyPrint, no renderer required. Rendered thumbnails deferred. |
| Baseline confidentiality controls: Microsoft Entra ID sign-in gate, auto-deletion of uploads after processing, encryption at rest/in transit | M | v1 requirement, not deferred. An unauthenticated "internal" endpoint is not an adequate control for IB/government decks. |
| Per-client/per-project Profiles (read/apply in audit via `GET /v1/profiles` and `GET /v1/profiles/{id}`) + minimal profile config | M | Full admin UI deferred to v2; profiles seeded in v1. |
| Web UI for Flow B (Audit Only): upload, select profile, audit, review flagged issues, export PDF | M | |

**v1.5: Safe deterministic auto-fix with preview/diff**

Only `confidence = deterministic` fixes, behind per-change accept/reject and a visual before/after diff. Jobs run with `mode = fix`.

| Item | Size | Notes |
|------|------|-------|
| Clean-write engine (surgical edits, original never overwritten); produces `cleaned_pptx` via `/outputs` | L | Same pipeline as audit with write enabled, so audit and fix cannot drift. |
| Per-slide visual before/after diff (the real review surface) | L | Requires the renderer (see v2 note); this is the trust mechanism designers actually use, not the text record set. |
| Granular per-change accept/reject UI on `POST /v1/jobs/{job_id}/apply` (`{ record_ids[] }`) | M | |
| Deterministic fixes only (`confidence = deterministic`): font family/size unification where effective values resolve unambiguously; color replace to a named valid theme-color ref; exact-match shape sizing (same `prstGeom` `prst` + same `ext` within threshold + same placeholder `idx`/role); header/footer text/position normalization on non-RTL content | L | Excludes nearest-color replace, font-size hierarchy inference, master synthesis, alignment intent (those stay flag/preview). |
| Embed canonical brand + Arabic fonts in cleaned `.pptx` | S | |
| Web UI for Flow A (Format & Deliver): upload, select profile, fix pass, review records, download cleaned `.pptx` + PDF | M | |
| Renderer integration (Microsoft Graph `driveItem` `content?format=pdf` primary; LibreOffice headless fallback for internal artifacts only) | M | Avoid PowerPoint COM/Office automation (unsupported for unattended server use). Gated on spike U5. |
| App-role RBAC enforcement: Entra token `roles` claim drives Designer/Reviewer/Admin authorization per Appendix A.8 | M | Required here, not v2: `POST /v1/jobs/{job_id}/apply` is the first role-restricted operation (Designer and Admin only). Builds on the v1 OIDC sign-in gate; v1 runs single-role. |

**v2: Hard auto-fix + full Arabic RTL + admin + integrations**

| Item | Size | Notes |
|------|------|-------|
| Master enforcement / layout reassignment auto-fix (only if spike U1 passes) | XL | Highest blast radius. "Auto-invent a master when none exists" remains cut. |
| Heuristic remediation behind preview: snap-to-guides, align similar groups, nearest-color replace, font-size hierarchy inference | L | Suggest/preview, never silent auto-apply. |
| Full Arabic RTL support (complex-script `a:cs` typeface handling, RTL alignment/margins/footer logic) | XL | Lifts the v1 guards into real fixes; gated on spike findings and possibly Aspose. |
| Admin UI for Profiles (fonts/colors/margins, per-client/per-project) + user management; `POST /v1/profiles`, `PUT /v1/profiles/{id}` | L | Role administration builds on the v1.5 app-role RBAC (sign-in gate is v1; role enforcement is v1.5). Per the A.8 matrix: Reviewer = view findings/summary/outputs + comment + sign off, no profile edits; Designer = run own jobs + view + apply fixes + comment; Admin = all capabilities. |
| Odoo integration via OAuth2 client-credentials (machine/application identity, explicitly authorized service principal) over the REST API | M | REST API is the primary surface; web UI is one client of it. Machine callers upload, create, cancel jobs and read findings/summary/outputs; they do not comment or sign off. |
| Teams integration / notifications | M | |
| Optional inline comment mode (`GET`/`POST /v1/jobs/{job_id}/comments`, `POST /v1/jobs/{job_id}/signoff`) | M | |
| Deck-to-PDF with slide thumbnails | M | Builds on the v1.5 renderer. |

### 11.3 Risk register

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| Master re-apply corruption: enforcing a layout on a populated slide orphans placeholders or breaks content binding (no python-pptx API; `idx`-bound placeholders; raw relationship swap does not remap content). | High | High | Spike U1 proves the boundary first. v1 limits scope to detect dominant layout + flag outliers (`master_slide.layout_outlier`) + enforce an explicitly chosen existing master. Cut master synthesis. Defensive `master_slide.no_usable_master` guard for non-conformant third-party exports. Enforcement auto-fix only in v2 if U1 passes. |
| Arabic silent corruption: an RTL-unaware font swap to a Latin-only brand font destroys Arabic glyphs/shaping with no visible error at fix time. | High | High | v1 detects Arabic (Unicode ranges + `rtl='1'`) and hard-flags/skips (`action = skipped`, `arabic_flag = true`, `issue_type = arabic.guarded_operation`) font substitution and direction-sensitive ops on Arabic runs ("Arabic content, manual review"). Handle complex-script `a:cs` separately from `a:latin`. Arabic fixtures in the regression corpus. Full RTL deferred to v2 and possibly to the Aspose engine. |
| Renderer fidelity: the diff/PDF misrepresents the deck, eroding trust or hiding a real defect. | Medium | Medium | Microsoft Graph PDF as primary renderer (M365-native, server-safe). LibreOffice headless is fallback for internal artifacts only and is labeled as such (font substitution, documented Arabic RTL bugs). Install canonical brand + Arabic fonts on every renderer. Never use unattended Office automation. |
| Round-trip data loss: open/save churn rewrites untouched XML or damages SmartArt/charts/media/OLE. | Medium | High | Operate surgically (touch only changed XML). Mandatory pre-flight flags SmartArt/charts/media we will not modify (`preflight.unmodifiable_content`; preserved verbatim but not editable). Spike U2 validates round-trip on real client decks. Original never overwritten. |
| Designer trust / adoption: one mangled client deck spreads by word of mouth in a 140-person shop and kills the tool. | Medium | High | Audit-first release earns trust before any write. Position as a pre-delivery QC assistant, not enforcement/policing. Per-slide visual before/after diff + granular accept/reject. Aggregate precision and false-positive rate tracked per fix type and gated before auto-apply is enabled (separate from the per-record `confidence` field). Pilot with 2-3 senior designers as co-owners. |
| Confidentiality breach: IB/government decks exposed via a reachable "internal" endpoint or undeleted uploads. | Medium | High | v1 baseline: Microsoft Entra ID sign-in gate, auto-deletion of uploads after processing, encryption at rest/in transit. Confirm retention and data-residency (KSA/UAE government rules) with Operations/IT-security before hosting choice; that constraint, not technology, decides on-prem vs Azure. |
| Estimation: hidden complexity in master enforcement or Arabic RTL blows the timeline. | High | Medium | The 2-week spike retires U1-U5 before any timeline is committed. S/M/L/XL sizing avoids false precision. Audit-first phasing means a slip in v1.5/v2 fixes does not block v1 value delivery. Re-rate affected sizings at spike close. |

### 11.4 Validation strategy

**Regression corpus.** Build and maintain a versioned corpus of real (anonymized) Prezlab decks spanning the conditions that break formatting tools: native masters, third-party exports, bilingual EN/AR and Arabic-only, `schemeClr`-themed brand decks, heavy-asset decks (SmartArt/charts/media/OLE), and at least one 150-plus-slide deck. Every fixed defect adds a deck (or slide) to the corpus so it cannot regress. The corpus runs in CI on the audit pipeline.

**Per-fix precision and false-positive targets (gate auto-apply).** Auto-apply is disabled by default per fix type and only enabled once that type clears its gate on the corpus. Precision here is an aggregate accuracy metric measured over the corpus, distinct from the per-record `Confidence` enum (`deterministic`, `high`, `medium`, `low`) stamped on each `FindingRecord`.

| Fix type | Class | Auto-apply gate |
|----------|-------|-----------------|
| Exact-match shape sizing; named theme-color compliance; deterministic font family/size unification; header/footer normalization (non-RTL) | Deterministic (`confidence = deterministic`) | Aggregate precision >= 99%, aggregate false-positive rate <= 1% on the corpus. Eligible for auto-apply once met. |
| Nearest-color replace, font-size hierarchy inference, alignment/snap intent, master enforcement | Inference-heavy | Never auto-applied. Suggest/preview only, regardless of measured precision. |

Aggregate precision and false-positive rate are tracked per fix type and reviewed before any change moves from preview-only to auto-apply. These corpus-level metrics are independent of the per-record `confidence` value, which governs only whether an individual finding is eligible for auto-apply in v1.5.

**Acceptance criteria for v1 (audit-only).**

- All six modules emit structured `FindingRecord` entries (`record_id`, `slide_index`, `shape_id`, `shape_path`, `module`, `issue_type`, `property`, `old_value`, `new_value`, `severity`, `action`, `confidence`, `arabic_flag`, `profile_rule_id`, `message`) for every finding.
- Inheritance + color resolver returns effective values matching PowerPoint's displayed values on the hand-checked sample, with no `None`-leaks for inherited font size/color.
- Arabic detection: zero false negatives on the Arabic fixtures; direction-sensitive ops and Arabic-run font substitution are flagged (`arabic_flag = true`, `action = skipped`), never silently acted on.
- Full audit of the 200-slide deck completes inside the async-job budget without timeout, with a working per-slide progress counter driven by `slides_processed`.
- Pre-flight correctly flags SmartArt/charts/media/OLE on the heavy-asset deck (`preflight.unmodifiable_content`).
- Report PDF/CSV generates from the `FindingRecord` set with correct counts and per-`issue_type` summary.
- Baseline confidentiality controls live: Microsoft Entra ID sign-in gate enforced, uploads auto-deleted after processing, encryption at rest and in transit verified.
- Original `.pptx` is never overwritten in any flow.

**Pilot.** Run v1 (and later v1.5) with 2-3 senior designers as co-owners, not test subjects, on their own live client decks. Success signals: designers run the audit voluntarily on real pre-delivery work; flagged issues are judged accurate and useful (low false-positive complaints); the before/after diff (in v1.5) is trusted enough that designers accept fixes without re-checking in PowerPoint. Exit criterion to widen rollout: no client deck damaged during the pilot, and per-fix gates met for the fix types proposed for auto-apply.

---

## Appendix A. Canonical Data Contracts

This appendix is the single source of truth for shared schemas, enumerations, and endpoint names. Where any other section of this PRD differs, this appendix governs. All field names, enum values, and paths below are normative.

### A.1 Enumerations

| Enum | Values | Notes |
|------|--------|-------|
| `Severity` | `error`, `warning`, `info` | Three-valued everywhere. `error` = profile violation that should block delivery; `warning` = likely issue or a guarded/skipped action needing a human; `info` = advisory. |
| `Confidence` | `deterministic`, `high`, `medium`, `low` | Per-record certainty that the finding is correct and its proposed fix is safe. Auto-apply (v1.5) is eligible only when `confidence = deterministic`. Numeric precision is tracked separately as an aggregate metric, not on the record. |
| `Action` | `flagged`, `changed`, `skipped` | `flagged` = detected, no write (audit/preview). `changed` = a fix was applied or proposed. `skipped` = intentionally not acted on (for example, an Arabic-guarded operation). |
| `ModuleKey` | `master_slide`, `font`, `margin_alignment`, `color_palette`, `shape_size`, `header_footer` | One canonical key per module. Used in records, the API `modules[]` parameter, and profile config keys. |
| `JobStatus` | `queued`, `processing`, `ready`, `failed`, `cancelled` | |
| `JobMode` | `audit`, `fix` | `fix` is the write pass (v1.5+). The term `format` is not used. |
| `FontRole` | `title`, `subtitle`, `body`, `caption` | Text hierarchy roles used by the font module and profiles. |

### A.2 Finding record (system of record)

One `FindingRecord` is emitted per detected issue or applied change. Records are emitted by the modules as they run; they are never reconstructed by diffing `.pptx` files. The change manifest, audit report, summary, and all UI surfaces are projections of this record set.

```jsonc
FindingRecord {
  "record_id":      "string (uuid)",        // canonical id; not "id"
  "job_id":         "string",
  "slide_index":    0,                        // integer, ZERO-BASED
  "shape_id":       "string",                 // OOXML p:cNvPr @id, as a string
  "shape_path":     "string | null",          // group ancestry, e.g. "12/3" for shape 3 inside group 12; null at top level
  "module":         "ModuleKey",
  "issue_type":     "string",                 // stable code, e.g. "font.family_out_of_set" (see A.3)
  "property":       "string | null",          // OOXML property inspected/changed, e.g. "rPr.latin.typeface"
  "old_value":      "string | null",
  "new_value":      "string | null",          // proposed (preview) or applied value; null for flag-only findings
  "severity":       "Severity",
  "action":         "Action",
  "confidence":     "Confidence",
  "arabic_flag":    false,                     // true if the record involves Arabic/RTL content
  "profile_rule_id":"string | null",           // the profile rule that triggered the record
  "message":        "string",                  // human-readable; not "note"
  "created_at":     "string (ISO 8601)"
}
```

Field decisions that supersede the section drafts: the id field is `record_id`; the human-readable field is `message`; `shape_id` is a string; `slide_index` is zero-based; `shape_path`, `issue_type`, `arabic_flag`, `profile_rule_id`, and `created_at` are part of the contract.

### A.3 `issue_type` namespace

Stable, dotted codes scoped by module. Non-exhaustive, extended per module section:

`master_slide.layout_outlier`, `master_slide.no_usable_master`, `master_slide.placeholder_geometry_off`; `font.family_out_of_set`, `font.size_off_role`, `font.mixed_weight`, `font.theme_ref_disallowed`, `font.cs_typeface_missing`; `margin_alignment.outside_safe_zone`, `margin_alignment.edge_misaligned`, `margin_alignment.uneven_spacing`; `color_palette.off_palette_rgb`, `color_palette.disallowed_theme_slot`, `color_palette.tint_out_of_range`; `shape_size.size_mismatch`, `shape_size.off_grid`; `header_footer.missing`, `header_footer.text_mismatch`, `header_footer.position_mismatch`, `header_footer.font_mismatch`; cross-cutting `preflight.unmodifiable_content` and `arabic.guarded_operation`.

### A.4 Default severity by issue class

Defaults are profile-overridable. The module sections must match these.

| Issue class | Default severity |
|-------------|------------------|
| Font family out of allowed set | `error` |
| Font size off role / mixed weight | `warning` |
| Color: literal RGB off-palette with no near match (outside ambiguity band) | `error` |
| Color: off-palette but within ambiguity band, or disallowed theme slot, or tint out of range | `warning` |
| Margin: element outside safe zone | `warning` |
| Alignment / spacing deviation | `warning` |
| Shape size mismatch within a cohort | `warning` |
| Header/footer missing | `error` |
| Header/footer text, position, or font mismatch | `warning` |
| Arabic-guarded (skipped) operation | `warning` with `arabic_flag = true`, `action = skipped` |
| Pre-flight unmodifiable content (SmartArt, charts, media) | `info` |

### A.5 Job model

```jsonc
Job {
  "job_id":          "string",
  "status":          "JobStatus",
  "mode":            "JobMode",
  "preview":         true,                     // if true, no writes even in fix mode
  "profile_id":      "string",
  "profile_version": 0,                        // version resolved at submit
  "profile_snapshot":{ /* resolved Profile.config, frozen for reproducibility */ },
  "modules":         ["ModuleKey"],            // empty array = run all modules
  "slides_total":    0,
  "slides_processed":0,                         // drives the progress bar
  "created_at":      "string (ISO 8601)",
  "started_at":      "string | null",
  "finished_at":     "string | null",
  "error":           { "code": "string", "message": "string" }       // null unless status=failed
}
```

Outputs are not embedded on the Job; they are retrieved via `GET /v1/jobs/{job_id}/outputs` (A.7).

### A.6 Profile model

```jsonc
Profile {
  "id":            "string",
  "name":          "string",
  "client_scope":  "string | null",     // client this profile belongs to
  "project_scope": "string | null",     // optional narrower project scope
  "is_default":    false,
  "version":       1,                     // incremented on each save; jobs stamp the version used
  "owner":         "string",
  "created_at":    "string (ISO 8601)",
  "updated_at":    "string (ISO 8601)",
  "config": {
    "master_slide": {
      "enforce_existing_only": true,            // v1/v1.5: never synthesize a master
      "pinned_layout_id":      "string | null", // explicit master/layout to enforce
      "layout_allowlist":      ["string"],       // permitted layout ids; others flagged as outliers
      "geometry_tolerance_emu": 9525
    },
    "font": {
      "roles": {
        "title":    { "latin": ["string"], "complex_script": ["string"], "size_pt": 0, "allowed_weights": ["string"] },
        "subtitle": { "latin": ["string"], "complex_script": ["string"], "size_pt": 0, "allowed_weights": ["string"] },
        "body":     { "latin": ["string"], "complex_script": ["string"], "size_pt": 0, "allowed_weights": ["string"] },
        "caption":  { "latin": ["string"], "complex_script": ["string"], "size_pt": 0, "allowed_weights": ["string"] }
      },
      "theme_font_refs_allowed": true,           // permit +mj-lt / +mn-lt references
      "size_tolerance_pt": 0.5
    },
    "color_palette": {
      "theme_color_slots": ["dk1","lt1","dk2","lt2","accent1","accent2","accent3","accent4","accent5","accent6"],
      "named_colors": [
        { "name": "string", "hex": "#RRGGBB", "theme_ref": "string | null", "allowed_tints": [0.0], "allowed_shades": [0.0] }
      ],
      "on_palette_mode":        "by_name",       // "by_name" or "by_resolved_rgb"
      "match_tolerance_deltaE":  2.0,            // within = on-palette (CIEDE2000)
      "auto_replace_max_deltaE": 5.0,            // <= this => safe nearest-match suggestion
      "ambiguity_band_deltaE":   10.0            // beyond auto_replace and within this => flag, do not suggest
    },
    "geometry": {
      "safe_zone_margins_emu": { "left": 0, "right": 0, "top": 0, "bottom": 0 },
      "grid": { "columns": 12, "gutter_emu": 0, "enabled": false },
      "alignment": { "edge_tolerance_emu": 0, "center_tolerance_emu": 0, "spacing_tolerance_emu": 0 }
    },
    "shape_size": {
      "size_tolerance_emu":     9525,
      "min_cohort_size":        3,               // minimum shapes to treat as a cohort before unifying
      "preserve_picture_aspect":true,
      "dominant_size_strategy": "median"          // "median" | "mode" | "largest"
    },
    "header_footer": {
      "template": {
        "footer_text":  "string | null",
        "slide_number": true,
        "date":         { "enabled": false, "format": "DD/MM/YYYY" },
        "position_emu": { "x": 0, "y": 0 },
        "font_role":    "caption"
      }
    }
  }
}
```

Canonical field names (these supersede the module and API drafts): safe zone is `geometry.safe_zone_margins_emu`; alignment tolerances live under `geometry.alignment.*`; the shape tolerance is `shape_size.size_tolerance_emu`; color uses `theme_color_slots` and `named_colors[]` with `hex`; the font module reads `font.roles.{role}.{latin[],complex_script[],size_pt,allowed_weights[]}`.

### A.7 REST endpoints (canonical)

All paths are under `/v1`. User calls authenticate with Microsoft Entra ID (OIDC); machine callers (Odoo) use OAuth2 client-credentials as an application identity. Roles in brackets (D = Designer, R = Reviewer, A = Admin; M = machine/application identity).

| Method & path | Purpose | Roles | Phase |
|---|---|---|---|
| `POST /v1/uploads` | Stream a `.pptx` to storage; returns `{ upload_id, expires_at }` | D, A, M | v1 |
| `POST /v1/jobs` | Create a job `{ upload_id, profile_id, modules[], mode, preview }` | D, A, M | v1 |
| `GET /v1/jobs/{job_id}` | Job status, progress, error | D, R, A, M | v1 |
| `POST /v1/jobs/{job_id}/cancel` | Cancel a queued or processing job | D (own), A | v1 |
| `GET /v1/jobs/{job_id}/findings` | Full `FindingRecord[]` (filterable by module, severity, slide) | D, R, A, M | v1 |
| `GET /v1/jobs/{job_id}/summary` | Aggregate counts per `issue_type` and `severity` | D, R, A, M | v1 |
| `GET /v1/jobs/{job_id}/outputs` | Signed URLs for `cleaned_pptx`, `report_pdf`, `report_csv`, `manifest_json` | D, R, A, M | v1 (report_pdf/csv, manifest_json); `cleaned_pptx` v1.5 |
| `POST /v1/jobs/{job_id}/apply` | Apply selected fixes `{ record_ids[] }`, returns a new job/output | D, A | v1.5 |
| `GET /v1/jobs/{job_id}/comments` | List inline comments | D, R, A | v2 |
| `POST /v1/jobs/{job_id}/comments` | Add a comment `{ slide_index, record_id?, text }` | D, R, A | v2 |
| `POST /v1/jobs/{job_id}/signoff` | Reviewer/Admin sign-off on a job | R, A | v2 |
| `GET /v1/profiles` | List profiles | D, R, A | v1 |
| `GET /v1/profiles/{id}` | Get a profile | D, R, A | v1 |
| `POST /v1/profiles` | Create a profile | A | v2 (seeded in v1) |
| `PUT /v1/profiles/{id}` | Update a profile (increments `version`) | A | v2 |

Reports are retrieved only as signed URLs via `/outputs`. There are no direct `report.pdf` or `artifacts/{kind}` endpoints. The single upload contract is `POST /v1/uploads` returning `upload_id`; `storage_key` is internal and never returned.

### A.8 Role permission matrix

| Capability | Designer | Reviewer | Admin | Machine (Odoo) |
|---|---|---|---|---|
| Upload, create, cancel jobs | Yes (own) | No | Yes | Yes |
| View findings, summary, outputs | Yes | Yes | Yes | Yes |
| Apply selected fixes (v1.5) | Yes | No | Yes | No |
| Comment | Yes | Yes | Yes | No |
| Sign off | No | Yes | Yes | No |
| Manage profiles | No | No | Yes | No |
| Manage users / roles | No | No | Yes | No |

Machine callers are authorized explicitly as a registered service principal; they are not mapped to a human role.
