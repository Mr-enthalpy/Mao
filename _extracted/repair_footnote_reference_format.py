#!/usr/bin/env python3
"""Normalize body footnote references in annotated DOCX files.

Fixes two mechanical problems from the original insertion scripts:

1. The text run after a split footnote reference sometimes has an empty w:rPr,
   so it can render smaller than the surrounding body text.
2. Body footnote references used bare Word-generated numbers. This normalizes
   the body marker to a superscripted bracket form: [n].

The true footnote relationship is preserved via w:footnoteReference.
"""

from __future__ import annotations

import copy
import os
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def run_text(run: etree._Element) -> str:
    return "".join(run.xpath(".//w:t/text()", namespaces=NS))


def contains_footnote_ref(run: etree._Element) -> bool:
    return bool(run.xpath(".//w:footnoteReference", namespaces=NS))


def is_text_run(run: etree._Element) -> bool:
    return bool(run.xpath(".//w:t", namespaces=NS)) and not contains_footnote_ref(run)


def get_rpr(run: etree._Element) -> etree._Element | None:
    return run.find(w("rPr"))


def set_rpr(run: etree._Element, rpr: etree._Element | None) -> bool:
    existing = get_rpr(run)
    if rpr is None:
        return False
    new_rpr = copy.deepcopy(rpr)
    if existing is not None:
        if etree.tostring(existing) == etree.tostring(new_rpr):
            return False
        run.remove(existing)
    run.insert(0, new_rpr)
    return True


def ensure_child(parent: etree._Element, tag: str) -> etree._Element:
    child = parent.find(w(tag))
    if child is None:
        child = etree.SubElement(parent, w(tag))
    return child


def normalize_ref_rpr(run: etree._Element) -> bool:
    changed = False
    rpr = run.find(w("rPr"))
    if rpr is None:
        rpr = etree.Element(w("rPr"))
        run.insert(0, rpr)
        changed = True

    rstyle = rpr.find(w("rStyle"))
    if rstyle is None:
        rstyle = etree.Element(w("rStyle"))
        rpr.insert(0, rstyle)
        changed = True
    if rstyle.get(w("val")) != "FootnoteReference":
        rstyle.set(w("val"), "FootnoteReference")
        changed = True

    va = rpr.find(w("vertAlign"))
    if va is None:
        va = etree.SubElement(rpr, w("vertAlign"))
        changed = True
    if va.get(w("val")) != "superscript":
        va.set(w("val"), "superscript")
        changed = True

    return changed


def normalize_ref_brackets(run: etree._Element) -> bool:
    """Make a footnote-reference run display as [<generated-number>]."""
    existing_children = [c for c in run if c.tag != w("rPr")]
    if (
        len(existing_children) == 3
        and existing_children[0].tag == w("t")
        and existing_children[0].text == "["
        and existing_children[1].tag == w("footnoteReference")
        and existing_children[2].tag == w("t")
        and existing_children[2].text == "]"
    ):
        return False

    refs = run.xpath("./w:footnoteReference", namespaces=NS)
    if len(refs) != 1:
        refs = run.xpath(".//w:footnoteReference", namespaces=NS)
    if len(refs) != 1:
        return False
    ref = refs[0]
    ref_id = ref.get(w("id"))

    rpr = get_rpr(run)
    # Remove non-rPr children and rebuild as [ footnoteReference ].
    for child in list(run):
        if child is not rpr:
            run.remove(child)

    left = etree.SubElement(run, w("t"))
    left.text = "["
    new_ref = etree.SubElement(run, w("footnoteReference"))
    if ref_id is not None:
        new_ref.set(w("id"), ref_id)
    right = etree.SubElement(run, w("t"))
    right.text = "]"
    return True


def nearest_previous_body_rpr(run: etree._Element) -> etree._Element | None:
    prev = run.getprevious()
    while prev is not None:
        if etree.QName(prev).localname == "r" and is_text_run(prev) and run_text(prev):
            rpr = get_rpr(prev)
            if rpr is not None and len(rpr):
                return rpr
        prev = prev.getprevious()
    return None


def patch_docx(path: Path) -> bool:
    with zipfile.ZipFile(path, "r") as zin:
        names = set(zin.namelist())
        if "word/document.xml" not in names:
            return False
        doc_root = etree.fromstring(zin.read("word/document.xml"))
        changed = False

        for ref_run in doc_root.xpath(".//w:footnoteReference/ancestor::w:r[1]", namespaces=NS):
            if normalize_ref_rpr(ref_run):
                changed = True
            if normalize_ref_brackets(ref_run):
                changed = True

            source_rpr = nearest_previous_body_rpr(ref_run)
            next_run = ref_run.getnext()
            if (
                next_run is not None
                and etree.QName(next_run).localname == "r"
                and is_text_run(next_run)
                and run_text(next_run)
            ):
                next_rpr = get_rpr(next_run)
                if next_rpr is None or not len(next_rpr):
                    if set_rpr(next_run, source_rpr):
                        changed = True

        if not changed:
            return False

        overrides = {"word/document.xml": xml_bytes(doc_root)}
        fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(path.parent))
        os.close(fd)
        Path(tmp_name).unlink(missing_ok=True)
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in overrides:
                    zout.writestr(info, overrides[info.filename])
                else:
                    zout.writestr(info, zin.read(info.filename))
        shutil.move(tmp_name, path)
        return True


def verify(path: Path) -> tuple[int, int, int, int]:
    with zipfile.ZipFile(path, "r") as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
    total = styled = bracketed = bad_next = 0
    for ref_run in root.xpath(".//w:footnoteReference/ancestor::w:r[1]", namespaces=NS):
        total += 1
        rpr = get_rpr(ref_run)
        rstyle = rpr.find(w("rStyle")) if rpr is not None else None
        va = rpr.find(w("vertAlign")) if rpr is not None else None
        if (
            rstyle is not None
            and rstyle.get(w("val")) == "FootnoteReference"
            and va is not None
            and va.get(w("val")) == "superscript"
        ):
            styled += 1
        children = [c for c in ref_run if c.tag != w("rPr")]
        if (
            len(children) == 3
            and children[0].tag == w("t")
            and children[0].text == "["
            and children[1].tag == w("footnoteReference")
            and children[2].tag == w("t")
            and children[2].text == "]"
        ):
            bracketed += 1
        next_run = ref_run.getnext()
        if (
            next_run is not None
            and etree.QName(next_run).localname == "r"
            and is_text_run(next_run)
            and run_text(next_run)
        ):
            next_rpr = get_rpr(next_run)
            if next_rpr is None or not len(next_rpr):
                bad_next += 1
    return total, styled, bracketed, bad_next


def main() -> None:
    touched: list[Path] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        if patch_docx(path):
            touched.append(path.relative_to(ROOT))
    print(f"patched_docs={len(touched)}")
    for rel in touched:
        print(rel)

    totals = [0, 0, 0, 0]
    bad: list[str] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        result = verify(path)
        totals = [a + b for a, b in zip(totals, result)]
        total, styled, bracketed, bad_next = result
        if total != styled or total != bracketed or bad_next:
            bad.append(
                f"{path.relative_to(ROOT)} total={total} styled={styled} "
                f"bracketed={bracketed} bad_next={bad_next}"
            )
    print(f"footnote_refs={totals[0]}")
    print(f"styled_superscript_refs={totals[1]}")
    print(f"bracketed_refs={totals[2]}")
    print(f"empty_next_text_rpr={totals[3]}")
    if bad:
        raise SystemExit("Bad footnote refs:\n" + "\n".join(bad))


if __name__ == "__main__":
    main()
