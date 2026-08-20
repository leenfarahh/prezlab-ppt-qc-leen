"""U4 (part 2): Arabic / RTL detection.

Detects Arabic content per run/paragraph so direction-sensitive operations and
font substitution can be guarded (PRD Section 8). A missed Arabic run is the
dangerous failure mode, so detection is deliberately broad: Unicode block
ranges OR the rtl attribute on a:pPr both count as hits.
"""

from dataclasses import dataclass, field

from .ns import find, qn

# PRD 8.2: Arabic Unicode blocks including presentation forms.
ARABIC_RANGES = (
    (0x0600, 0x06FF),  # Arabic
    (0x0750, 0x077F),  # Arabic Supplement
    (0x08A0, 0x08FF),  # Arabic Extended-A
    (0xFB50, 0xFDFF),  # Arabic Presentation Forms-A
    (0xFE70, 0xFEFF),  # Arabic Presentation Forms-B
)


def contains_arabic(text: str) -> bool:
    return any(lo <= ord(ch) <= hi for ch in text for lo, hi in ARABIC_RANGES)


def paragraph_is_rtl(paragraph) -> bool:
    """True when a:pPr carries rtl='1' (or 'true')."""
    pPr = find(paragraph._p, "a:pPr")
    return pPr is not None and pPr.get("rtl") in ("1", "true")


@dataclass
class ArabicHit:
    slide_index: int  # zero-based, per Appendix A.2
    shape_id: str
    paragraph_index: int
    run_index: int | None  # None => paragraph-level rtl hit with no Arabic text
    reason: str  # "unicode" | "rtl_attr" | "unicode+rtl_attr"
    sample: str = field(default="", repr=False)


def _iter_text_shapes(shapes):
    """Yield shapes with text frames, descending into groups (child shapes of a
    group carry their own text; group traversal is required or Arabic inside a
    group is silently missed)."""
    for shape in shapes:
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from _iter_text_shapes(shape.shapes)
        elif shape.has_text_frame:
            yield shape


def scan_presentation(prs) -> list[ArabicHit]:
    hits: list[ArabicHit] = []
    for s_idx, slide in enumerate(prs.slides):
        for shape in _iter_text_shapes(slide.shapes):
            shape_id = str(shape.shape_id)
            for p_idx, para in enumerate(shape.text_frame.paragraphs):
                rtl = paragraph_is_rtl(para)
                para_had_unicode_hit = False
                for r_idx, run in enumerate(para.runs):
                    if contains_arabic(run.text):
                        para_had_unicode_hit = True
                        hits.append(
                            ArabicHit(
                                slide_index=s_idx,
                                shape_id=shape_id,
                                paragraph_index=p_idx,
                                run_index=r_idx,
                                reason="unicode+rtl_attr" if rtl else "unicode",
                                sample=run.text[:40],
                            )
                        )
                if rtl and not para_had_unicode_hit:
                    hits.append(
                        ArabicHit(
                            slide_index=s_idx,
                            shape_id=shape_id,
                            paragraph_index=p_idx,
                            run_index=None,
                            reason="rtl_attr",
                        )
                    )
    return hits


def cs_typeface(run) -> str | None:
    """Complex-script typeface on the run (a:rPr/a:cs/@typeface).

    python-pptx Font.name reads/writes only a:latin, so Arabic font auditing
    must go through this raw-XML path (PRD 8.4)."""
    rPr = find(run._r, "a:rPr")
    cs = find(rPr, "a:cs")
    return cs.get("typeface") if cs is not None else None


def set_cs_typeface(run, typeface: str) -> None:
    """Write the complex-script typeface. Insertion order matters for schema
    validity: a:cs must follow a:latin/a:ea within a:rPr."""
    rPr = run._r.get_or_add_rPr()
    cs = find(rPr, "a:cs")
    if cs is None:
        cs = rPr.makeelement(qn("a:cs"), {})
        anchor = find(rPr, "a:ea") or find(rPr, "a:latin")
        if anchor is not None:
            anchor.addnext(cs)
        else:
            rPr.append(cs)
    cs.set("typeface", typeface)
