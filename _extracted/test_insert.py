#!/usr/bin/env python3
"""Test footnote insertion into DOCX."""
from docx import Document
import os, shutil
from lxml import etree

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
jan_dir = r'E:\1952年大传\1952年大传\1952.1'

files = sorted([f for f in os.listdir(jan_dir) if f.endswith('.docx')])
src = os.path.join(jan_dir, files[0])
dst = os.path.join(jan_dir, '__test3.docx')
shutil.copy2(src, dst)

doc = Document(dst)

# Get footnotes part
fn_part = None
for rel in doc.part.rels.values():
    if 'footnote' in rel.reltype.lower():
        fn_part = rel.target_part
        break

fn_root = etree.fromstring(fn_part.blob)
max_id = max(int(fn.get('{%s}id' % WNS, '0')) for fn in fn_root.findall('{%s}footnote' % WNS))
new_id = max_id + 1
print('New footnote ID:', new_id)

# Find text and insert footnote reference
found = False
for para in doc.paragraphs:
    for run in para.runs:
        idx = run.text.find('梁思成')
        if idx >= 0:
            before = run.text[:idx + 3]
            after = run.text[idx + 3:]
            run.text = before

            ref_xml = (
                '<w:r xmlns:w="%s">'
                '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                '<w:footnoteReference w:id="%d"/>'
                '</w:r>'
            ) % (WNS, new_id)
            ref_elem = etree.fromstring(ref_xml)
            run._element.addnext(ref_elem)

            if after:
                after_xml = (
                    '<w:r xmlns:w="%s">'
                    '<w:rPr></w:rPr>'
                    '<w:t xml:space="preserve">%s</w:t>'
                    '</w:r>'
                ) % (WNS, after)
                after_elem = etree.fromstring(after_xml)
                ref_elem.addnext(after_elem)

            found = True
            break
    if found:
        break

if not found:
    print('ERROR: Text not found')
    exit(1)

# Add footnote content
fn_text = '梁思成（1901—1972），广东新会人，中国近代建筑学家、建筑教育家，曾任清华大学建筑系主任。'
fn_xml = (
    '<w:footnote xmlns:w="%s" w:id="%d">'
    '<w:p>'
    '<w:pPr>'
    '<w:rPr>'
    '<w:rFonts w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'
    '<w:kern w:val="2"/>'
    '<w:sz w:val="21"/><w:szCs w:val="21"/>'
    '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
    '</w:rPr>'
    '</w:pPr>'
    '<w:r>'
    '<w:rPr>'
    '<w:rFonts w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'
    '<w:kern w:val="2"/>'
    '<w:sz w:val="21"/><w:szCs w:val="21"/>'
    '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
    '</w:rPr>'
    '<w:footnoteRef/>'
    '</w:r>'
    '<w:r>'
    '<w:rPr>'
    '<w:rFonts w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'
    '<w:kern w:val="2"/>'
    '<w:sz w:val="21"/><w:szCs w:val="21"/>'
    '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
    '</w:rPr>'
    '<w:t xml:space="preserve">%s</w:t>'
    '</w:r>'
    '</w:p>'
    '</w:footnote>'
) % (WNS, new_id, fn_text)

fn_new_elem = etree.fromstring(fn_xml)
fn_root.append(fn_new_elem)
fn_part._blob = etree.tostring(fn_root, xml_declaration=True, encoding='UTF-8', standalone=True)

doc.save(dst)
print('Saved to', dst)
print('Done!')
