"""Shared XML namespace helpers for walking PresentationML/DrawingML trees."""

from pptx.oxml.ns import qn

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSMAP = {"a": A, "p": P}


def find(el, path):
    return el.find(path, NSMAP) if el is not None else None


def findall(el, path):
    return el.findall(path, NSMAP) if el is not None else []


def attr(el, name, default=None):
    return el.get(name, default) if el is not None else default


__all__ = ["qn", "A", "P", "NSMAP", "find", "findall", "attr"]
