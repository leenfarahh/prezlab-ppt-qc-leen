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
    """Desktop PowerPoint reachable for COM automation on this host."""
    if os.name != "nt":
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application").Close()
        return True
    except OSError:
        return False


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
                try:
                    if app.Presentations.Count == 0:
                        app.Quit()
                except Exception:
                    pass
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
