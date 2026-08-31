"""Master unification: detect slides living on foreign masters and move them
to the dominant master.

Two proven paths (validated 14/07/2026 with live COM and package-level
experiments against desktop PowerPoint; see tests/test_unify.py):

1. dedup() - pure package-level repoint for CLONE masters (the copy-paste
   pollution case). When a foreign layout is structurally identical to one on
   the dominant master, re-targeting the slide's layout relationship changes
   nothing visually: the slide's own XML is untouched and every inherited
   property resolves to identical values. Works everywhere (no PowerPoint),
   so it runs on the cloud tier too. Emptied masters are then removed.

2. com_unify() - drives desktop PowerPoint for genuinely DIFFERENT masters.
   Assigning Slide.CustomLayout runs PowerPoint's own placeholder-matching
   engine (verified: content preserved, unmatched placeholders orphaned in
   place with content, never swapped or deleted). Windows + PowerPoint only;
   the visual result can legitimately differ, so these fixes stay behind the
   designer's per-change approval and the before/after review.

Structural identity is deliberately strict: canonicalized master XML (minus
its layout-id list) + the full set of its layouts' canonicalized XML + the
theme, all with relationship ids normalized. Anything less risks calling two
visually different masters "identical".
"""

import hashlib
import io
import os
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_SLIDE_LAYOUT = f"{R}/slideLayout"
REL_SLIDE_MASTER = f"{R}/slideMaster"
REL_THEME = f"{R}/theme"


def com_available() -> bool:
    """Desktop PowerPoint reachable for COM automation on this host.

    A registry check, and deliberately no more: starting PowerPoint to find out
    would make every page load launch Office. It answers "is PowerPoint
    installed", NOT "will automation work right now" - a machine can pass this
    and still refuse the dispatch (see com_failure_advice), which is why the
    pages that need it also handle the later failure."""
    if os.name != "nt":
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application").Close()
        return True
    except OSError:
        return False


# COM failures restated as something a designer can do. The raw tuple - what
# the rebuild used to print, verbatim: "(-2147417856, 'System call
# failed.', None, None)" - names the HRESULT and nothing else, so the person
# reading it has no next step (design lead, 23/08/2026).
_COM_ADVICE = {
    # 0x80010100 RPC_E_SYS_CALL_FAILED. Seen after a run was interrupted: the
    # headless instance PowerPoint was driving stays alive, and every later
    # dispatch fails against it. Not retryable in-process, which is why it is
    # not in _TRANSIENT_HRESULTS - the leftover has to go first.
    -2147417856: (
        "PowerPoint automation is in a bad state on this machine, which "
        "usually means a leftover instance from a run that was interrupted. It "
        "was started with /AUTOMATION and has no window, so closing PowerPoint "
        "normally does not touch it: end POWERPNT.EXE in Task Manager, then try "
        "again."),
    # 0x80040154 REGDB_E_CLASSNOTREG
    -2147221164: (
        "Windows has no PowerPoint registered for automation. Applying a "
        "master runs PowerPoint's own placeholder matching, so it needs "
        "desktop PowerPoint installed on this machine."),
    # 0x80080005 CO_E_SERVER_EXEC_FAILURE. The COM server was launched and did
    # not register its class factory in time. Two causes seen on real hosts, and
    # the message named only the second one for a release - which sent a
    # designer looking for a repair prompt that was not there while a leftover
    # instance from 10:05 that morning was the actual cause (24/08/2026).
    # Leftovers are now swept before every run (sweep_automation), so the prompt
    # is what is left, but the leftover is named first because it is the one a
    # person can check in ten seconds.
    -2146959355: (
        "PowerPoint would not start. Check Task Manager for a POWERPNT.EXE "
        "with no window - a leftover from an interrupted run wedges the next "
        "one, and although these are now cleared automatically before each run, "
        "one that refuses to close has to go by hand. Otherwise this is an "
        "update or repair prompt waiting to be answered: open PowerPoint "
        "yourself once, clear whatever it asks for, then try again."),
}


# Every headless PowerPoint, and ONLY the headless ones. COM launches PowerPoint
# with /AUTOMATION -Embedding; a person opening a deck never does, so filtering on
# it is what makes a forced teardown safe. Matching the image name alone could
# kill a designer's open presentation with unsaved work in it, which would be far
# worse than the leak this exists to stop.
#
# PowerPoint's Application object offers no route to its own process id - unlike
# Word and Excel it has no HWND member ("Member not found", checked 23/08/2026) -
# so the pid has to come from the OS.
_PS_AUTOMATION_PIDS = (
    "Get-CimInstance Win32_Process -Filter \"Name='POWERPNT.EXE'\" | "
    "Where-Object { $_.CommandLine -like '*/AUTOMATION*' } | "
    "ForEach-Object { $_.ProcessId }")


_WQL_AUTOMATION = ("SELECT ProcessId, CommandLine FROM Win32_Process "
                   "WHERE Name = 'POWERPNT.EXE'")


def _automation_pids_wmi() -> set | None:
    """The same answer through in-process WMI, or None when WMI is not
    reachable and the caller should fall back.

    Worth having its own path because this question is asked several times per
    render and it is not cheap either way: reading a process's command line
    means opening the process. Spawning PowerShell to ask costs ~510ms;
    asking WMI directly costs ~330ms (measured 24/08/2026)."""
    try:
        import win32com.client
    except ImportError:
        return None
    try:
        rows = win32com.client.GetObject("winmgmts:").ExecQuery(_WQL_AUTOMATION)
        return {int(row.ProcessId) for row in rows
                if row.CommandLine and "/AUTOMATION" in row.CommandLine.upper()}
    except Exception:
        return None


def automation_pids() -> set:
    """PIDs of the headless PowerPoint instances running right now."""
    if os.name != "nt":
        return set()
    found = _automation_pids_wmi()
    if found is not None:
        return found
    try:
        done = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             _PS_AUTOMATION_PIDS],
            capture_output=True, text=True, timeout=30)
    except Exception:
        return set()
    return {int(line) for line in done.stdout.split() if line.strip().isdigit()}


def _terminate(pid: int, grace_ms: int) -> None:
    """Give the process `grace_ms` to exit on its own, then end it."""
    import ctypes

    PROCESS_TERMINATE, SYNCHRONIZE, WAIT_OBJECT_0 = 0x0001, 0x00100000, 0
    handle = ctypes.windll.kernel32.OpenProcess(
        PROCESS_TERMINATE | SYNCHRONIZE, False, pid)
    if not handle:
        return  # already gone, which is the whole point
    try:
        if ctypes.windll.kernel32.WaitForSingleObject(
                handle, grace_ms) != WAIT_OBJECT_0:
            ctypes.windll.kernel32.TerminateProcess(handle, 1)
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def sweep_automation(grace_ms: int = 2000) -> set:
    """End every headless PowerPoint on the host BEFORE a run starts driving
    one, and return the pids ended.

    This is the hole the teardown snapshot cannot cover, and it is the reason a
    single interrupted run made the machine refuse to render for the rest of the
    day (real host, 24/08/2026: one /AUTOMATION instance left from 10:05, and
    every design page after it answered "PowerPoint would not start").

    Two measured facts make it necessary:

    DispatchEx DOES NOT START A SECOND POWERPOINT. Two DispatchEx calls in one
    process return one pid (checked 24/08/2026), because PowerPoint is
    single-instance for automation. So a run that begins with a leftover alive
    does not get a clean instance - it ATTACHES to the wedged one, and inherits
    whatever state wedged it.

    AND A LEFTOVER IS THEREFORE INVISIBLE TO force_quit. Its snapshot is taken
    before the dispatch, so a pre-existing pid is in `started` and can never
    appear in the difference; teardown asks Quit, a wedged instance ignores it,
    and the leftover survives to break the next run too. Self-perpetuating,
    which is why it has to be cleared at the START of a run rather than the end.

    Killing rather than asking is deliberate: an instance that answers Quit is
    not the one this exists for. What makes it safe is the /AUTOMATION filter -
    a person opening a deck never produces one - and the render lock, which is
    held by the only caller that can be driving PowerPoint from this process.

    RETURNS WHAT IS STILL ALIVE, not what it ended, because that is exactly the
    snapshot force_quit wants and a render should not pay for a second process
    listing to ask the same question. Normally empty. Each listing is a WMI or
    PowerShell round trip that has to read every process's command line, and
    four of them per render was 2s of a 7.7s wait spent asking Windows the same
    thing twice over (measured 24/08/2026); the clean case now costs one.
    """
    if os.name != "nt":
        return set()
    leftovers = automation_pids()
    if not leftovers:
        return leftovers          # the normal case: nothing to kill, nothing to re-ask
    for pid in leftovers:
        _terminate(pid, grace_ms)
    return automation_pids()      # only the ones that refused to go


def force_quit(app, started: set | None = None, grace_ms: int = 5000) -> None:
    """Quit PowerPoint, and make sure it actually went.

    Quit() cannot close an instance that is WEDGED - one sitting on a repair
    prompt, or halfway through a failed Presentations.Open. The call returns, the
    exception is swallowed, and a windowless POWERPNT.EXE survives. Every later
    automation attempt on the host then fails against it with -2147417856 System
    call failed and leaks another one, so a single bad render makes the machine
    unusable until somebody finds them in Task Manager (three in one afternoon,
    23/08/2026, each from a failed render on the review page).

    `started` is automation_pids() taken BEFORE the instance was created; the
    kill is limited to headless instances that appeared since. Without it this
    only asks Quit nicely, which is the old behaviour and the reason for the
    leak - so callers driving PowerPoint should always pass it.

    The snapshot covers instances THIS run created and nothing else, by design.
    An instance that was already alive is in `started` and survives here however
    wedged it is, which is why every caller sweeps first (sweep_automation): the
    two halves together are what keep the host clean.

    Quit is asked first and given `grace_ms` to land, because a clean exit
    flushes PowerPoint's own state; the kill is only for when it does not."""
    try:
        app.Quit()
    except Exception:
        pass
    if started is None or os.name != "nt":
        return
    for pid in automation_pids() - started:
        _terminate(pid, grace_ms)


def com_failure_advice(exc) -> str:
    """One sentence a designer can act on, with the raw error kept for a bug
    report. Unknown HRESULTs fall through to the error itself rather than to a
    guess: a wrong instruction wastes more time than no instruction."""
    args = getattr(exc, "args", ())
    advice = _COM_ADVICE.get(args[0] if args else None)
    if not advice:
        return f"PowerPoint automation failed: {exc}"
    return f"{advice} (The error itself: {exc}.)"


# ---------------------------------------------------------------- package IO

def _read_parts(deck_bytes: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(deck_bytes)) as z:
        return {n: z.read(n) for n in z.namelist()}


def _write_parts(parts: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return out.getvalue()


def _rels_name(part: str) -> str:
    d, _, base = part.rpartition("/")
    return f"{d}/_rels/{base}.rels"


def _rel_targets(parts: dict, part: str, rel_type: str) -> list[str]:
    """Absolute part names this part references via rel_type."""
    rels = parts.get(_rels_name(part))
    if rels is None:
        return []
    base_dir = part.rpartition("/")[0]
    out = []
    for rel in etree.fromstring(rels):
        if rel.get("Type") == rel_type:
            out.append(_resolve(base_dir, rel.get("Target")))
    return out


def _resolve(base_dir: str, target: str) -> str:
    """Resolve a relationship target ('../slideLayouts/x.xml') to a part name."""
    if target.startswith("/"):
        return target.lstrip("/")
    segs = base_dir.split("/")
    for piece in target.split("/"):
        if piece == "..":
            segs.pop()
        elif piece not in (".", ""):
            segs.append(piece)
    return "/".join(segs)


# ------------------------------------------------------- structural identity

def _canonical(xml_bytes: bytes, strip_layout_ids: bool = False) -> bytes:
    """Canonicalized XML with relationship ids normalized, so two clone parts
    hash equal even when their rIds differ."""
    root = etree.fromstring(xml_bytes)
    if strip_layout_ids:
        for lst in root.iter(f"{{{P}}}sldLayoutIdLst"):
            lst.getparent().remove(lst)
    for el in root.iter():
        for attr in list(el.attrib):
            if attr.startswith(f"{{{R}}}"):
                el.set(attr, "#rid")
    return etree.tostring(root, method="c14n")


def _layout_signature(parts: dict, layout: str) -> str:
    return hashlib.sha256(_canonical(parts[layout])).hexdigest()


def _master_signature(parts: dict, master: str) -> str:
    h = hashlib.sha256(_canonical(parts[master], strip_layout_ids=True))
    # sort by canonical CONTENT, not part name: clones number their layout
    # parts differently, and part-name order must not change the hash
    for layout_c14n in sorted(_canonical(parts[la])
                              for la in _rel_targets(parts, master, REL_SLIDE_LAYOUT)):
        h.update(layout_c14n)
    for theme in _rel_targets(parts, master, REL_THEME):
        h.update(_canonical(parts[theme]))
    return h.hexdigest()


# ------------------------------------------------------------------ analysis

@dataclass
class ForeignSlide:
    slide_index: int          # zero-based, presentation order
    slide_part: str
    master_part: str
    layout_part: str
    layout_name: str
    twin_layout_part: str | None = None   # structural twin on the dominant master
    twin_layout_name: str | None = None
    name_match_layout: str | None = None  # same-NAME layout on the dominant master


@dataclass
class UnifyAnalysis:
    masters: dict[str, list[int]] = field(default_factory=dict)  # part -> slide idxs
    dominant: str | None = None
    clone_masters: set = field(default_factory=set)  # masters structurally == dominant
    foreign: list[ForeignSlide] = field(default_factory=list)

    @property
    def multiple_masters(self) -> bool:
        return len(self.masters) > 1


def _slide_parts_in_order(parts: dict) -> list[str]:
    """Slide part names in presentation order (sldIdLst -> rels)."""
    pres = etree.fromstring(parts["ppt/presentation.xml"])
    rels = etree.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    by_id = {rel.get("Id"): _resolve("ppt", rel.get("Target")) for rel in rels}
    out = []
    lst = pres.find(f"{{{P}}}sldIdLst")
    if lst is None:
        return out
    for sld in lst:
        part = by_id.get(sld.get(f"{{{R}}}id"))
        if part:
            out.append(part)
    return out


def _layout_display_name(parts: dict, layout: str) -> str:
    try:
        root = etree.fromstring(parts[layout])
        return root.find(f".//{{{P}}}cSld").get("name") or layout
    except Exception:
        return layout


def analyze(deck_bytes: bytes) -> UnifyAnalysis:
    """Map every slide to its master; identify the dominant master, clone
    masters, and per-foreign-slide fix routes (twin repoint vs name match)."""
    parts = _read_parts(deck_bytes)
    a = UnifyAnalysis()

    slide_layout: dict[str, str] = {}
    layout_master: dict[str, str] = {}
    for s_idx, slide in enumerate(_slide_parts_in_order(parts)):
        layouts = _rel_targets(parts, slide, REL_SLIDE_LAYOUT)
        if not layouts:
            continue  # broken chain; master_slide.no_usable_master covers it
        layout = layouts[0]
        if layout not in layout_master:
            masters = _rel_targets(parts, layout, REL_SLIDE_MASTER)
            if not masters:
                continue
            layout_master[layout] = masters[0]
        slide_layout[slide] = layout
        a.masters.setdefault(layout_master[layout], []).append(s_idx)

    if len(a.masters) < 2:
        a.dominant = next(iter(a.masters), None)
        return a

    # dominant = most slides; ties break by sldMasterIdLst order (part order
    # here is a stable proxy: masters were inserted in slide order, so sort
    # explicitly by (-count, part name) for determinism)
    a.dominant = sorted(a.masters.items(), key=lambda kv: (-len(kv[1]), kv[0]))[0][0]

    dom_layouts = _rel_targets(parts, a.dominant, REL_SLIDE_LAYOUT)
    dom_by_sig = {_layout_signature(parts, la): la for la in dom_layouts}
    dom_by_name: dict[str, str] = {}
    for la in dom_layouts:
        dom_by_name.setdefault(_layout_display_name(parts, la), la)
    dom_sig = _master_signature(parts, a.dominant)

    for master, slide_idxs in a.masters.items():
        if master == a.dominant:
            continue
        if _master_signature(parts, master) == dom_sig:
            a.clone_masters.add(master)

    slides_in_order = _slide_parts_in_order(parts)
    for master, slide_idxs in a.masters.items():
        if master == a.dominant:
            continue
        for s_idx in slide_idxs:
            slide = slides_in_order[s_idx]
            layout = slide_layout[slide]
            name = _layout_display_name(parts, layout)
            fs = ForeignSlide(slide_index=s_idx, slide_part=slide,
                              master_part=master, layout_part=layout,
                              layout_name=name)
            # A twin repoint is only a visual no-op when the ENTIRE master
            # chain is identical: a byte-identical layout under a different
            # master/theme still inherits different fonts and colors.
            twin = (dom_by_sig.get(_layout_signature(parts, layout))
                    if master in a.clone_masters else None)
            if twin is not None:
                fs.twin_layout_part = twin
                fs.twin_layout_name = _layout_display_name(parts, twin)
            elif name in dom_by_name:
                fs.name_match_layout = name
            a.foreign.append(fs)
    a.foreign.sort(key=lambda f: f.slide_index)
    return a


# --------------------------------------------------------------- dedup fixer

def dedup(deck_bytes: bytes, repoints: dict[int, str]) -> bytes:
    """Re-target the given slides (zero-based index -> dominant-master layout
    part name) and delete any master left with no slides. The slides' own XML
    is never touched; only relationship targets and package plumbing change."""
    parts = _read_parts(deck_bytes)
    slides = _slide_parts_in_order(parts)

    for s_idx, target_layout in repoints.items():
        slide = slides[s_idx]
        rels_name = _rels_name(slide)
        root = etree.fromstring(parts[rels_name])
        for rel in root:
            if rel.get("Type") == REL_SLIDE_LAYOUT:
                # target relative to ppt/slides/ -> ../slideLayouts/xxx.xml
                rel.set("Target", "../" + target_layout[len("ppt/"):])
        parts[rels_name] = etree.tostring(root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)

    # masters that now own zero slides get removed with their layouts + theme
    survivors = analyze(_write_parts(parts))
    referenced = set(survivors.masters)
    pres = etree.fromstring(parts["ppt/presentation.xml"])
    pres_rels = etree.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    all_masters = [_resolve("ppt", r.get("Target")) for r in pres_rels
                   if r.get("Type") == REL_SLIDE_MASTER]

    doomed: set[str] = set()
    for master in all_masters:
        if master in referenced:
            continue
        layouts = _rel_targets(parts, master, REL_SLIDE_LAYOUT)
        themes = _rel_targets(parts, master, REL_THEME)
        # keep a theme that another (surviving) master shares
        shared = set()
        for other in all_masters:
            if other != master and other in referenced:
                shared.update(_rel_targets(parts, other, REL_THEME))
        doomed.add(master)
        doomed.add(_rels_name(master))
        for la in layouts:
            doomed.add(la)
            doomed.add(_rels_name(la))
        for th in themes:
            if th not in shared:
                doomed.add(th)

        rel_id = next(r.get("Id") for r in pres_rels
                      if r.get("Type") == REL_SLIDE_MASTER
                      and _resolve("ppt", r.get("Target")) == master)
        lst = pres.find(f"{{{P}}}sldMasterIdLst")
        for entry in list(lst):
            if entry.get(f"{{{R}}}id") == rel_id:
                lst.remove(entry)
        for r in list(pres_rels):
            if r.get("Id") == rel_id:
                pres_rels.remove(r)

    if doomed:
        parts["ppt/presentation.xml"] = etree.tostring(
            pres, xml_declaration=True, encoding="UTF-8", standalone=True)
        parts["ppt/_rels/presentation.xml.rels"] = etree.tostring(
            pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ctypes = etree.fromstring(parts["[Content_Types].xml"])
        for o in list(ctypes):
            if o.get("PartName", "").lstrip("/") in doomed:
                ctypes.remove(o)
        parts["[Content_Types].xml"] = etree.tostring(
            ctypes, xml_declaration=True, encoding="UTF-8", standalone=True)
        for name in doomed:
            parts.pop(name, None)

    return _write_parts(parts)


# ----------------------------------------------------------------- COM fixer

def com_unify(deck_bytes: bytes, assignments: dict[int, str]) -> tuple[bytes | None, dict[int, str]]:
    """Re-apply layouts via desktop PowerPoint: for each zero-based slide
    index, assign the named layout from the deck's dominant design (dominant =
    most slides, matching analyze()). Returns (new_bytes, per-slide errors);
    new_bytes is None when PowerPoint itself is unavailable.

    PowerPoint runs its own placeholder matching here: content is preserved
    and unmatched placeholders are orphaned in place, but the visual result
    can differ, which is why these fixes require explicit designer approval.
    """
    errors: dict[int, str] = {}
    if not com_available():
        return None, {i: "needs desktop PowerPoint (run on the LAN box)"
                      for i in assignments}

    import gc

    import pythoncom
    import win32com.client

    from .render import _RENDER_LOCK  # PowerPoint is single-instance

    with _RENDER_LOCK:
        pythoncom.CoInitialize()
        fd, path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        app = None
        try:
            with open(path, "wb") as f:
                f.write(deck_bytes)
            # leftovers first: DispatchEx attaches to a running /AUTOMATION
            # instance rather than starting a clean one. What survives the sweep
            # is the teardown's snapshot, so that is one listing, not two.
            started = sweep_automation()
            try:
                app = win32com.client.DispatchEx("PowerPoint.Application")
                app.DisplayAlerts = 1  # ppAlertsNone
            except Exception as exc:
                return None, {i: f"PowerPoint automation unavailable: {exc}"
                              for i in assignments}
            pres = app.Presentations.Open(path, 0, 0, 0)
            try:
                # dominant design = the one with most slides (analyze()'s rule)
                counts: dict[int, int] = {}
                for i in range(1, pres.Slides.Count + 1):
                    idx = pres.Slides(i).Design.Index
                    counts[idx] = counts.get(idx, 0) + 1
                dom = pres.Designs(max(counts, key=lambda k: counts[k]))
                layouts = {dom.SlideMaster.CustomLayouts(i).Name:
                           dom.SlideMaster.CustomLayouts(i)
                           for i in range(1, dom.SlideMaster.CustomLayouts.Count + 1)}
                for s_idx, layout_name in sorted(assignments.items()):
                    if s_idx >= pres.Slides.Count:
                        errors[s_idx] = "slide index out of range"
                        continue
                    target = layouts.get(layout_name)
                    if target is None:
                        errors[s_idx] = (f"dominant master has no layout named "
                                         f"'{layout_name}'")
                        continue
                    try:
                        pres.Slides(s_idx + 1).CustomLayout = target
                    except Exception as exc:
                        errors[s_idx] = f"layout apply failed: {exc}"
                # drop designs left with no slides
                used = {pres.Slides(i).Design.Index
                        for i in range(1, pres.Slides.Count + 1)}
                for i in range(pres.Designs.Count, 0, -1):
                    if i not in used:
                        try:
                            pres.Designs(i).Delete()
                        except Exception:
                            pass  # keep it; harmless leftover
                pres.Save()
            finally:
                pres.Close()
            with open(path, "rb") as f:
                return f.read(), errors
        finally:
            if app is not None:
                # The Presentations.Count guard is why this cannot just call
                # Quit: another deck open in the SAME instance would be closed
                # under someone's feet. When this instance is ours alone, it goes
                # for certain (force_quit) rather than as far as Quit manages.
                try:
                    idle = app.Presentations.Count == 0
                except Exception:
                    idle = True  # unreachable: wedged, and force_quit's business
                if idle:
                    force_quit(app, started)
            app = None
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            try:
                os.unlink(path)
            except OSError:
                pass
