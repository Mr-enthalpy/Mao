#!/usr/bin/env python3
"""Repair Word comment anchors inserted by add_review_comments_from_todo.py.

The first pass added valid comments, but the visible comment-reference run did
not explicitly use Word's CommentReference character style. In Word this can
inherit surrounding paragraph formatting and disturb apparent body/comment font
sizes. This repair makes the reference run explicit and also moves any range
start that accidentally precedes paragraph properties after w:pPr.
"""

from __future__ import annotations

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


def ensure_comment_reference_style(run: etree._Element) -> bool:
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
    if rstyle.get(w("val")) != "CommentReference":
        rstyle.set(w("val"), "CommentReference")
        changed = True
    return changed


def ensure_comment_text_style(comments_root: etree._Element) -> bool:
    changed = False
    for p in comments_root.xpath(".//w:comment/w:p", namespaces=NS):
        ppr = p.find(w("pPr"))
        if ppr is None:
            ppr = etree.Element(w("pPr"))
            p.insert(0, ppr)
            changed = True
        pstyle = ppr.find(w("pStyle"))
        if pstyle is None:
            pstyle = etree.Element(w("pStyle"))
            ppr.insert(0, pstyle)
            changed = True
        if pstyle.get(w("val")) != "CommentText":
            pstyle.set(w("val"), "CommentText")
            changed = True
    return changed


def move_range_start_after_ppr(p: etree._Element) -> bool:
    if len(p) < 2:
        return False
    if p[0].tag != w("commentRangeStart") or p[1].tag != w("pPr"):
        return False
    start = p[0]
    p.remove(start)
    p.insert(1, start)
    return True


def patch_docx(path: Path) -> bool:
    with zipfile.ZipFile(path, "r") as zin:
        names = set(zin.namelist())
        if "word/document.xml" not in names or "word/comments.xml" not in names:
            return False
        doc_root = etree.fromstring(zin.read("word/document.xml"))
        comments_root = etree.fromstring(zin.read("word/comments.xml"))

        changed = False
        for p in doc_root.xpath(".//w:p[.//w:commentReference]", namespaces=NS):
            if move_range_start_after_ppr(p):
                changed = True
        for run in doc_root.xpath(".//w:commentReference/ancestor::w:r[1]", namespaces=NS):
            if ensure_comment_reference_style(run):
                changed = True
        if ensure_comment_text_style(comments_root):
            changed = True

        if not changed:
            return False

        overrides = {
            "word/document.xml": xml_bytes(doc_root),
            "word/comments.xml": xml_bytes(comments_root),
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


def verify(path: Path) -> tuple[int, int]:
    with zipfile.ZipFile(path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            return (0, 0)
        doc_root = etree.fromstring(zf.read("word/document.xml"))
    refs = doc_root.xpath(".//w:commentReference/ancestor::w:r[1]", namespaces=NS)
    styled = 0
    for run in refs:
        rstyle = run.find(f"{w('rPr')}/{w('rStyle')}")
        if rstyle is not None and rstyle.get(w("val")) == "CommentReference":
            styled += 1
    return (len(refs), styled)


def main() -> None:
    touched: list[Path] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        if patch_docx(path):
            touched.append(path.relative_to(ROOT))
    print(f"patched_docs={len(touched)}")
    for rel in touched:
        print(rel)

    total_refs = 0
    total_styled = 0
    bad: list[str] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        refs, styled = verify(path)
        total_refs += refs
        total_styled += styled
        if refs != styled:
            bad.append(f"{path.relative_to(ROOT)} refs={refs} styled={styled}")
    print(f"comment_reference_runs={total_refs}")
    print(f"styled_comment_reference_runs={total_styled}")
    if bad:
        raise SystemExit("Unstyled comment references:\n" + "\n".join(bad))


if __name__ == "__main__":
    main()
