#!/usr/bin/env python3
"""Add true Word comments for existing "本篇待人工确认事项" paragraphs.

The initial annotation batches stored uncertain items as visible end-of-document
paragraphs. This script keeps those paragraphs, but also wires each item as a
real Word comment anchored to that item paragraph, so the DOCX contains
word/comments.xml and standard comment anchors.
"""

from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
AUTHOR = "LLM初注待核"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
NS = {"w": W_NS, "pr": PKG_REL_NS}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def px(tag: str) -> str:
    return f"{{{PKG_REL_NS}}}{tag}"


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def paragraph_text(p: etree._Element) -> str:
    return "".join(p.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)).strip()


def ensure_comments_root(existing: bytes | None) -> etree._Element:
    if existing:
        return etree.fromstring(existing)
    return etree.Element(w("comments"), nsmap={"w": W_NS})


def next_comment_id(comments_root: etree._Element) -> int:
    ids: list[int] = []
    for c in comments_root.xpath("//w:comment", namespaces=NS):
        try:
            ids.append(int(c.get(w("id"))))
        except Exception:
            pass
    return max(ids) + 1 if ids else 0


def has_our_comments(comments_root: etree._Element) -> bool:
    for c in comments_root.xpath("//w:comment", namespaces=NS):
        if c.get(w("author")) == AUTHOR:
            return True
    return False


def ensure_content_type(ct_root: etree._Element) -> None:
    existing = ct_root.xpath(
        "//*[local-name()='Override' and @PartName='/word/comments.xml']"
    )
    if existing:
        return
    override = etree.Element("Override")
    override.set("PartName", "/word/comments.xml")
    override.set("ContentType", COMMENTS_CONTENT_TYPE)
    ct_root.append(override)


def ensure_document_rels(rels_root: etree._Element) -> None:
    for rel in rels_root.xpath("//pr:Relationship", namespaces=NS):
        if rel.get("Type") == COMMENTS_REL_TYPE and rel.get("Target") == "comments.xml":
            return
    ids: list[int] = []
    for rel in rels_root.xpath("//pr:Relationship", namespaces=NS):
        m = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if m:
            ids.append(int(m.group(1)))
    rel = etree.SubElement(rels_root, px("Relationship"))
    rel.set("Id", f"rId{max(ids) + 1 if ids else 1}")
    rel.set("Type", COMMENTS_REL_TYPE)
    rel.set("Target", "comments.xml")


def add_comment_anchor(p: etree._Element, comment_id: int) -> None:
    start = etree.Element(w("commentRangeStart"))
    start.set(w("id"), str(comment_id))
    p.insert(0, start)

    end = etree.Element(w("commentRangeEnd"))
    end.set(w("id"), str(comment_id))
    p.append(end)

    run = etree.SubElement(p, w("r"))
    ref = etree.SubElement(run, w("commentReference"))
    ref.set(w("id"), str(comment_id))


def append_comment(comments_root: etree._Element, comment_id: int, text: str) -> None:
    comment = etree.SubElement(comments_root, w("comment"))
    comment.set(w("id"), str(comment_id))
    comment.set(w("author"), AUTHOR)
    comment.set(w("date"), dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z")
    p = etree.SubElement(comment, w("p"))
    r_el = etree.SubElement(p, w("r"))
    t = etree.SubElement(r_el, w("t"))
    t.text = text


def collect_todo_paragraphs(doc_root: etree._Element) -> list[tuple[etree._Element, str]]:
    paragraphs = doc_root.xpath(".//w:p", namespaces=NS)
    results: list[tuple[etree._Element, str]] = []
    in_todo = False
    item_re = re.compile(r"^\s*(?:\d+[.．、]\s*)?(【[^】]+】.+|.+需核.+|.+核查.+|.+确认.+)")

    for p in paragraphs:
        text = paragraph_text(p)
        if "本篇待人工确认事项" in text:
            in_todo = True
            continue
        if not in_todo:
            continue
        if not text:
            continue
        if item_re.match(text):
            cleaned = re.sub(r"^\s*\d+[.．、]\s*", "", text).strip()
            results.append((p, cleaned))
    return results


def patch_docx(path: Path) -> int:
    with zipfile.ZipFile(path, "r") as zin:
        names = set(zin.namelist())
        if "word/document.xml" not in names:
            return 0
        doc_root = etree.fromstring(zin.read("word/document.xml"))
        comments_root = ensure_comments_root(
            zin.read("word/comments.xml") if "word/comments.xml" in names else None
        )
        if has_our_comments(comments_root):
            return 0

        targets = collect_todo_paragraphs(doc_root)
        if not targets:
            return 0

        ct_root = etree.fromstring(zin.read("[Content_Types].xml"))
        ensure_content_type(ct_root)

        rels_name = "word/_rels/document.xml.rels"
        rels_root = etree.fromstring(zin.read(rels_name))
        ensure_document_rels(rels_root)

        cid = next_comment_id(comments_root)
        for p, item_text in targets:
            add_comment_anchor(p, cid)
            append_comment(comments_root, cid, item_text)
            cid += 1

        overrides = {
            "word/document.xml": xml_bytes(doc_root),
            "word/comments.xml": xml_bytes(comments_root),
            "[Content_Types].xml": xml_bytes(ct_root),
            rels_name: xml_bytes(rels_root),
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
            if "word/comments.xml" not in names:
                zout.writestr("word/comments.xml", overrides["word/comments.xml"])
        shutil.move(tmp_name, path)
        return len(targets)


def verify_docx(path: Path) -> tuple[int, int, int]:
    with zipfile.ZipFile(path, "r") as zf:
        if "word/comments.xml" not in zf.namelist():
            return (0, 0, 0)
        comments_root = etree.fromstring(zf.read("word/comments.xml"))
        doc_root = etree.fromstring(zf.read("word/document.xml"))
    comment_ids = [
        c.get(w("id"))
        for c in comments_root.xpath("//w:comment", namespaces=NS)
        if c.get(w("author")) == AUTHOR
    ]
    starts = {
        e.get(w("id"))
        for e in doc_root.xpath(".//w:commentRangeStart", namespaces=NS)
    }
    refs = {
        e.get(w("id"))
        for e in doc_root.xpath(".//w:commentReference", namespaces=NS)
    }
    anchored = sum(1 for cid in comment_ids if cid in starts and cid in refs)
    return (len(comment_ids), anchored, len(refs))


def main() -> None:
    total_docs = 0
    total_comments = 0
    touched: list[tuple[Path, int]] = []
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        added = patch_docx(path)
        if added:
            total_docs += 1
            total_comments += added
            touched.append((path.relative_to(ROOT), added))

    print(f"patched_docs={total_docs}")
    print(f"added_comments={total_comments}")
    for rel, count in touched:
        print(f"{rel}\t{count}")

    bad: list[str] = []
    verified_comments = 0
    for path in sorted(ROOT.glob("1952.*/*__LLM初注.docx")):
        comments, anchored, _refs = verify_docx(path)
        verified_comments += comments
        if comments and comments != anchored:
            bad.append(f"{path.relative_to(ROOT)} comments={comments} anchored={anchored}")
    print(f"verified_comments_by_author={verified_comments}")
    if bad:
        raise SystemExit("Unanchored comments:\n" + "\n".join(bad))


if __name__ == "__main__":
    main()
