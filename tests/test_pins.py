"""Numbered pins: badge N in the findings list = box N on the slide preview
(and in the visual audit PDF). The numbering rule lives once in
qc.render.pin_numbers and is consumed by the report rows, the preview
rects, and the PDF, so these tests pin the contract between them."""

import re

from fastapi.testclient import TestClient

from qc.render import audit_rects, pin_numbers
from qc.web import app

client = TestClient(app)


def _rec(record_id, slide, shape, module="font", action="flagged"):
    return {"record_id": record_id, "slide_index": slide, "shape_id": shape,
            "module": module, "action": action, "issue_type": f"{module}.x",
            "severity": "warning", "arabic_flag": False}


def test_pin_numbers_shared_shape_and_exclusions():
    records = [
        _rec("a", 0, "10"),                       # slide 1, shape 10 -> pin 1
        _rec("b", 0, "11"),                       # slide 1, shape 11 -> pin 2
        _rec("c", 0, "10"),                       # same shape -> same pin 1
        _rec("d", 1, "10"),                       # numbering restarts per slide
        _rec("e", 0, "-"),                        # slide-level: no pin
        _rec("f", 0, "12", module="preflight"),   # preflight: no pin
        _rec("g", 0, "13", action="changed"),     # applied fix: no pin
    ]
    pins = pin_numbers(records)
    assert pins == {"a": 1, "b": 2, "c": 1, "d": 1}


def test_rect_pins_agree_with_list_pins(fixtures_dir):
    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = client.post("/audit",
                    files={"deck": ("mixed_layouts.pptx", deck,
                                    "application/octet-stream")},
                    data={"profile": "prezlab_en"})
    assert r.status_code == 200
    job_id = r.text.split("/manifest/", 1)[1].split('"', 1)[0]
    records = client.get(f"/manifest/{job_id}").json()["records"]

    pins = pin_numbers(records)
    rects = audit_rects(deck, records)
    assert pins and rects

    for slide_idx, slide_rects in rects.items():
        for rect in slide_rects:
            assert rect["pin"] >= 1
            for rid in rect["record_ids"]:
                assert pins[rid] == rect["pin"], \
                    f"slide {slide_idx}: list badge and box number diverge"

    # report rows carry the same numbers as data-pin attributes
    for m in re.finditer(r'data-record="([0-9a-f]+)"[^>]*data-pin="(\d+)"',
                         r.text):
        assert pins[m.group(1)] == int(m.group(2))
    assert 'class="pinno' in r.text


def test_whole_slide_findings_marked(fixtures_dir):
    deck = (fixtures_dir / "mixed_layouts.pptx").read_bytes()
    r = client.post("/audit",
                    files={"deck": ("mixed_layouts.pptx", deck,
                                    "application/octet-stream")},
                    data={"profile": "prezlab_en", "modules": ["master_slide"]})
    assert r.status_code == 200
    # layout outliers are slide-level (shape_id "-"): whole-slide marker,
    # never a numbered pin
    assert 'pinno whole' in r.text
