"""U4 (part 1b): effective color resolution.

python-pptx does not resolve schemeClr references to RGB (accessing .rgb on a
scheme color raises), and does not apply lumMod/lumOff/tint/shade transforms.
A palette checker that compares only literal srgbClr values false-flags decks
that are perfectly on-palette via theme colors (PRD Section 6.5, verified).

This module resolves: schemeClr -> clrMap -> theme1.xml clrScheme RGB ->
transform math, and provides CIEDE2000 for perceptual nearest-match.

Transform math note: lum ops are computed in HSL and tint/shade per-channel in
sRGB byte space. These are the standard approximations; spike success
criterion U4 requires hand-checking a sample against PowerPoint's rendered
values before v1 relies on them.
"""

import colorsys
import math

from lxml import etree

from .ns import find

_SCHEME_SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink")


# --- theme + clrMap --------------------------------------------------------


def _theme_element(master):
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            part = rel.target_part
            el = getattr(part, "_element", None)
            return el if el is not None else etree.fromstring(part.blob)
    return None


def color_scheme(master) -> dict[str, tuple[int, int, int]]:
    """{'accent1': (r,g,b), ...} from theme1.xml a:clrScheme."""
    theme = _theme_element(master)
    scheme = find(theme, ".//a:clrScheme")
    out = {}
    for slot in _SCHEME_SLOTS:
        el = find(scheme, f"a:{slot}")
        if el is None:
            continue
        srgb = find(el, "a:srgbClr")
        sys = find(el, "a:sysClr")
        hexval = srgb.get("val") if srgb is not None else (
            sys.get("lastClr") if sys is not None else None)
        if hexval:
            out[slot] = tuple(int(hexval[i:i + 2], 16) for i in (0, 2, 4))
    return out


def clr_map(master) -> dict[str, str]:
    """p:clrMap on the master maps bg1/tx1/bg2/tx2 to theme slots."""
    cm = find(master.element, "p:clrMap")
    if cm is None:
        return {}
    return dict(cm.attrib)


# --- transforms ------------------------------------------------------------


def _apply_transforms(rgb: tuple[int, int, int], color_el) -> tuple[int, int, int]:
    """Apply a:tint / a:shade / a:lumMod / a:lumOff children of a color element.
    Values are in thousandths of a percent (e.g. val='75000' = 75%)."""

    def pct(name):
        el = find(color_el, f"a:{name}")
        return int(el.get("val")) / 100_000.0 if el is not None else None

    r, g, b = rgb
    tint = pct("tint")
    if tint is not None:
        r, g, b = (round(c * tint + 255 * (1 - tint)) for c in (r, g, b))
    shade = pct("shade")
    if shade is not None:
        r, g, b = (round(c * shade) for c in (r, g, b))

    lum_mod, lum_off = pct("lumMod"), pct("lumOff")
    if lum_mod is not None or lum_off is not None:
        h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        l = l * (lum_mod if lum_mod is not None else 1.0) + (lum_off or 0.0)
        l = min(1.0, max(0.0, l))
        r, g, b = (round(c * 255) for c in colorsys.hls_to_rgb(h, l, s))
    return (int(r), int(g), int(b))


def resolve_color_element(color_el, master) -> tuple[int, int, int] | None:
    """Resolve ONE a:srgbClr / a:schemeClr element to final RGB, transforms
    applied. Split out from resolve_solid_fill because some colors are not
    wrapped in a:solidFill at all: p:bgRef carries its color element directly."""
    if color_el is None:
        return None
    tag = etree.QName(color_el).localname
    if tag == "srgbClr":
        hexval = color_el.get("val")
        if not hexval:
            return None
        rgb = tuple(int(hexval[i:i + 2], 16) for i in (0, 2, 4))
        return _apply_transforms(rgb, color_el)
    if tag == "schemeClr":
        name = color_el.get("val")
        mapping = clr_map(master)
        slot = mapping.get(name, name)  # bg1->lt1 etc.; accents map to themselves
        base = color_scheme(master).get(slot)
        if base is None:
            return None
        return _apply_transforms(base, color_el)
    return None


def resolve_solid_fill(fill_parent_el, master) -> tuple[int, int, int] | None:
    """Resolve the a:solidFill under fill_parent_el (spPr or rPr) to final RGB.
    Returns None when there is no solid fill (gradient/pattern/picture/none:
    no single color exists, flag-only per PRD 4.1)."""
    solid = find(fill_parent_el, "a:solidFill")
    if solid is None:
        return None
    for child in ("a:srgbClr", "a:schemeClr"):
        el = find(solid, child)
        if el is not None:
            return resolve_color_element(el, master)
    return None


# --- perceptual distance (CIEDE2000) ---------------------------------------


def _srgb_to_lab(rgb: tuple[int, int, int]) -> tuple[float, float, float]:
    def lin(c):
        c /= 255.0
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (lin(c) for c in rgb)
    # sRGB D65
    x = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / 0.95047
    y = 0.2126729 * r + 0.7151522 * g + 0.0721750 * b
    z = (0.0193339 * r + 0.1191920 * g + 0.9503041 * b) / 1.08883

    def f(t):
        return t ** (1 / 3) if t > 0.008856 else 7.787 * t + 16 / 116

    fx, fy, fz = f(x), f(y), f(z)
    return (116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz))


def ciede2000(rgb1: tuple[int, int, int], rgb2: tuple[int, int, int]) -> float:
    L1, a1, b1 = _srgb_to_lab(rgb1)
    L2, a2, b2 = _srgb_to_lab(rgb2)
    C1, C2 = math.hypot(a1, b1), math.hypot(a2, b2)
    Cm = (C1 + C2) / 2
    G = 0.5 * (1 - math.sqrt(Cm ** 7 / (Cm ** 7 + 25 ** 7)))
    a1p, a2p = a1 * (1 + G), a2 * (1 + G)
    C1p, C2p = math.hypot(a1p, b1), math.hypot(a2p, b2)

    def hp(a, b):
        if a == 0 and b == 0:
            return 0.0
        h = math.degrees(math.atan2(b, a))
        return h + 360 if h < 0 else h

    h1p, h2p = hp(a1p, b1), hp(a2p, b2)
    dLp = L2 - L1
    dCp = C2p - C1p
    if C1p * C2p == 0:
        dhp = 0.0
    elif abs(h2p - h1p) <= 180:
        dhp = h2p - h1p
    elif h2p - h1p > 180:
        dhp = h2p - h1p - 360
    else:
        dhp = h2p - h1p + 360
    dHp = 2 * math.sqrt(C1p * C2p) * math.sin(math.radians(dhp) / 2)
    Lpm = (L1 + L2) / 2
    Cpm = (C1p + C2p) / 2
    if C1p * C2p == 0:
        hpm = h1p + h2p
    elif abs(h1p - h2p) <= 180:
        hpm = (h1p + h2p) / 2
    elif h1p + h2p < 360:
        hpm = (h1p + h2p + 360) / 2
    else:
        hpm = (h1p + h2p - 360) / 2
    T = (1 - 0.17 * math.cos(math.radians(hpm - 30))
         + 0.24 * math.cos(math.radians(2 * hpm))
         + 0.32 * math.cos(math.radians(3 * hpm + 6))
         - 0.20 * math.cos(math.radians(4 * hpm - 63)))
    d_theta = 30 * math.exp(-(((hpm - 275) / 25) ** 2))
    RC = 2 * math.sqrt(Cpm ** 7 / (Cpm ** 7 + 25 ** 7))
    SL = 1 + (0.015 * (Lpm - 50) ** 2) / math.sqrt(20 + (Lpm - 50) ** 2)
    SC = 1 + 0.045 * Cpm
    SH = 1 + 0.015 * Cpm * T
    RT = -math.sin(math.radians(2 * d_theta)) * RC
    return math.sqrt(
        (dLp / SL) ** 2 + (dCp / SC) ** 2 + (dHp / SH) ** 2
        + RT * (dCp / SC) * (dHp / SH)
    )


def nearest_palette_match(rgb, palette: dict[str, tuple[int, int, int]]):
    """(name, deltaE) of the perceptually closest palette entry."""
    best = min(palette.items(), key=lambda kv: ciede2000(rgb, kv[1]))
    return best[0], ciede2000(rgb, best[1])
