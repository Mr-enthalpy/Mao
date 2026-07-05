#!/usr/bin/env python3
"""Use static visible [n] markers plus hidden true footnote references.

Some Word/WPS builds render automatic footnote numbers differently, especially
after literal brackets are placed around w:footnoteReference. This script makes
the visible marker deterministic:

- Body text: visible superscript static "[n]" + hidden w:footnoteReference.
- Footnote text area: visible static "[n]" at the beginning of each note.

The true footnote references are preserved so Word keeps note relationships.
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


def make_rpr(style: str | None = None, superscript: bool = False, hidden: bool = False) -> etree._Element:
    rpr = etree.Element(w("rPr"))
    if style:
        rstyle = etree.SubElement(rpr, w("rStyle"))
        rstyle.set(w("val"), style)
    if superscript:
        va = etree.SubElement(rpr, w("vertAlign"))
        va.set(w("val"), "superscript")
    if hidden:
        etree.SubElement(rpr, w("vanish"))
        etree.SubElement(rpr, w("specVanish"))
    return rpr


def make_text_run(text: str, *, superscript: bool = False) -> etree._Element:
    run = etree.Element(w("r"))
    run.append(make_rpr("FootnoteReference", superscript=superscript))
    t = etree.SubElement(run, w("t"))
    t.text = text
    return run


def make_hidden_ref_run(ref_id: str | None) -> etree._Element:
    run = etree.Element(w("r"))
    run.append(make_rpr("FootnoteReference", superscript=True, hidden=True))
    ref = etree.SubElement(run, w("footnoteReference"))
    if ref_id is not None:
        ref.set(w("id"), ref_id)
    return run


def make_footnote_label_run(text: str) -> etree._Element:
    run = etree.Element(w("r"))
    run.append(make_rpr("FootnoteReference"))
    t = etree.SubElement(run, w("t"))
    t.text = text
    return run


def strip_leading_extra_closing_bracket(p: etree._Element, label: str) -> bool:
    """Remove a residual literal ']' immediately after a footnote label."""
    changed = False
    runs = p.xpath("./w:r", namespaces=NS)
    label_index = None
    for i, run in enumerate(runs[:3]):
        if visible_text(run) == label:
            label_index = i
            break
    if label_index is None or label_index + 1 >= len(runs):
        return False

    next_run = runs[label_index + 1]
    texts = next_run.xpath("./w:t", namespaces=NS)
    if not texts:
        return False

    first_text = texts[0].text or ""
    if first_text.startswith("]"):
        texts[0].text = first_text[1:]
        changed = True

    if not visible_text(next_run):
        p.remove(next_run)
    return changed


def is_hidden_ref_run(run: etree._Element) -> bool:
    return bool(run.xpath("./w:footnoteReference", namespaces=NS)) and bool(
        run.xpath("./w:rPr/w:vanish | ./w:rPr/w:specVanish", namespaces=NS)
    )


def visible_text(run: etree._Element) -> str:
    return "".join(run.xpath("./w:t/text()", namespaces=NS))


def normalize_body(root: etree._Element) -> tuple[bool, dict[str, int]]:
    changed = False
    id_to_num: dict[str, int] = {}
    refs = list(root.xpath(".//w:footnoteReference", namespaces=NS))
    ordinal = 1
    for ref in refs:
        ref_run = ref.getparent()
        if ref_run is None or etree.QName(ref_run).localname != "r":
            continue
        ref_id = ref.get(w("id"))
        if ref_id is None:
            continue
        if ref_id not in id_to_num:
            id_to_num[ref_id] = ordinal
            ordinal += 1
        label = f"[{id_to_num[ref_id]}]"
        parent = ref_run.getparent()
        if parent is None:
            continue

        prev_run = ref_run.getprevious()
        next_run = ref_run.getnext()
        already_static = (
            prev_run is not None
            and etree.QName(prev_run).localname == "r"
            and visible_text(prev_run) == label
            and is_hidden_ref_run(ref_run)
        )
        if already_static:
            continue

        insert_at = parent.index(ref_run)

        # Remove literal bracket text from the old combined run by replacing it
        # with a hidden reference-only run.
        parent.remove(ref_run)
        parent.insert(insert_at, make_hidden_ref_run(ref_id))

        # If an earlier repair left literal brackets in adjacent runs, remove
        # those only when they are immediately next to this reference position.
        if prev_run is not None and etree.QName(prev_run).localname == "r" and visible_text(prev_run) in ("[", label):
            parent.remove(prev_run)
            insert_at -= 1
        if next_run is not None and etree.QName(next_run).localname == "r" and visible_text(next_run) == "]":
            parent.remove(next_run)

        parent.insert(insert_at, make_text_run(label, superscript=True))
        changed = True

    return changed, id_to_num


def normalize_footnotes(root: etree._Element, id_to_num: dict[str, int]) -> bool:
    changed = False
    for footnote in root.xpath(".//w:footnote", namespaces=NS):
        fid = footnote.get(w("id"))
        if fid not in id_to_num:
            continue
        label = f"[{id_to_num[fid]}]"
        paragraphs = footnote.xpath("./w:p", namespaces=NS)
        if not paragraphs:
            continue
        p = paragraphs[0]
        existing_runs = p.xpath("./w:r", namespaces=NS)
        if (
            existing_runs
            and visible_text(existing_runs[0]) == label
            and not p.xpath("./w:r[.//w:footnoteRef]", namespaces=NS)
        ):
            if strip_leading_extra_closing_bracket(p, label):
                changed = True
            continue

        # Remove existing automatic footnoteRef runs and any immediately
        # adjacent literal bracket label runs previously added by repair passes.
        for run in list(p.xpath("./w:r[.//w:footnoteRef]", namespaces=NS)):
            p.remove(run)
            changed = True
        for run in list(p.xpath("./w:r[w:t]", namespaces=NS))[:3]:
            if visible_text(run).strip() in ("[", "]", label):
                p.remove(run)
                changed = True

        insert_at = 1 if len(p) and p[0].tag == w("pPr") else 0
        p.insert(insert_at, make_footnote_label_run(label))
        strip_leading_extra_closing_bracket(p, label)
        changed = True
    return changed


def patch_docx(path: Path) -> bool:
    with zipfile.ZipFile(path, "r") as zin:
        names = set(zin.namelist())
        if "word/document.xml" not in names or "word/footnotes.xml" not in names:
            return False
        doc_root = etree.fromstring(zin.read("word/document.xml"))
        fn_root = etree.fromstring(zin.read("word/footnotes.xml"))

        changed_body, id_to_num = normalize_body(doc_root)
        changed_notes = normalize_footnotes(fn_root, id_to_num)
        if not (changed_body or changed_notes):
            return False

        overrides = {
            "word/document.xml": xml_bytes(doc_root),
            "word/footnotes.xml": xml_bytes(fn_root),
        }
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


def verify(path: Path) -> tuple[int, int, int, int, int]:
    with zipfile.ZipFile(path, "r") as zin:
        doc_root = etree.fromstring(zin.read("word/document.xml"))
        fn_root = etree.fromstring(zin.read("word/footnotes.xml"))
    body_refs = list(doc_root.xpath(".//w:footnoteReference", namespaces=NS))
    visible_body = hidden_refs = nested_body = 0
    id_to_num: dict[str, int] = {}
    ordinal = 1
    for ref in body_refs:
        run = ref.getparent()
        fid = ref.get(w("id"))
        if fid is None or run is None:
            continue
        if fid not in id_to_num:
            id_to_num[fid] = ordinal
            ordinal += 1
        label = f"[{id_to_num[fid]}]"
        if is_hidden_ref_run(run):
            hidden_refs += 1
        prev_run = run.getprevious()
        if prev_run is not None and visible_text(prev_run) == label:
            visible_body += 1
        if prev_run is not None and "[[" in visible_text(prev_run):
            nested_body += 1

    visible_notes = 0
    for footnote in fn_root.xpath(".//w:footnote", namespaces=NS):
        fid = footnote.get(w("id"))
        if fid not in id_to_num:
            continue
        label = f"[{id_to_num[fid]}]"
        p = footnote.find(w("p"))
        if p is not None:
            texts = [visible_text(r) for r in p.xpath("./w:r", namespaces=NS)]
            if label in texts[:2]:
                visible_notes += 1
    return len(body_refs), hidden_refs, visible_body, visible_notes, nested_body


def main() -> None:
    touched: list[Path] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        if patch_docx(path):
            touched.append(path.relative_to(ROOT))
    print(f"patched_docs={len(touched)}")
    for rel in touched:
        print(rel)

    totals = [0, 0, 0, 0, 0]
    bad: list[str] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        result = verify(path)
        totals = [a + b for a, b in zip(totals, result)]
        body_refs, hidden_refs, visible_body, visible_notes, nested_body = result
        if not (
            body_refs == hidden_refs == visible_body == visible_notes
            and nested_body == 0
        ):
            bad.append(
                f"{path.relative_to(ROOT)} refs={body_refs} hidden={hidden_refs} "
                f"body_labels={visible_body} note_labels={visible_notes} nested={nested_body}"
            )
    print(f"body_refs={totals[0]}")
    print(f"hidden_true_refs={totals[1]}")
    print(f"visible_body_labels={totals[2]}")
    print(f"visible_note_labels={totals[3]}")
    print(f"nested_body_labels={totals[4]}")
    if bad:
        raise SystemExit("Bad visible footnote markers:\n" + "\n".join(bad))


if __name__ == "__main__":
    main()
