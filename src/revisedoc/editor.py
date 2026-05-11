"""
Core .docx editing operations with change tracking, comments, and revision restoration.
"""

import copy
import json
import os
import shutil
import zipfile
from datetime import datetime, timezone

from lxml import etree
from docx import Document


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OPC_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _ns(tag):
    return f"{{{W}}}{tag}"


def _find_body(doc):
    return doc.element.body


def _run_text(r_elem):
    t = r_elem.find(_ns("t"))
    if t is not None and t.text:
        return t.text
    return ""


def _set_run_text(r_elem, text):
    t = r_elem.find(_ns("t"))
    if t is None:
        t = etree.SubElement(r_elem, _ns("t"))
    t.text = text


def _next_revision_id(body_elem):
    ids = [0]
    for attr_name in (_ns("id"),):
        for elem in body_elem.iter():
            val = elem.get(attr_name)
            if val is not None:
                try:
                    ids.append(int(val))
                except ValueError:
                    pass
    return str(max(ids) + 1)


# ─── Core operations ───────────────────────────────────────────────────────────


def _replace_in_paragraph(p, old_text, new_text, author, start_rev_id):
    rev_id = start_rev_id

    runs = list(p.iterchildren(_ns("r")))
    para_text = ""
    run_boundaries = []
    for idx, r in enumerate(runs):
        t = r.find(_ns("t"))
        if t is not None and t.text:
            start = len(para_text)
            para_text += t.text
            run_boundaries.append((idx, start, len(para_text)))

    if not para_text:
        return 0, rev_id

    segments = []
    pos = 0
    match_count = 0
    while True:
        m = para_text.find(old_text, pos)
        if m == -1:
            if pos < len(para_text):
                segments.append((para_text[pos:], False, pos))
            break
        if m > pos:
            segments.append((para_text[pos:m], False, pos))
        segments.append((para_text[m:m + len(old_text)], True, m))
        match_count += 1
        pos = m + len(old_text)

    if match_count == 0:
        return 0, rev_id

    for r in runs:
        p.remove(r)

    def _rpr_at(text_pos):
        for idx, rs, re in run_boundaries:
            if rs <= text_pos < re:
                rpr = runs[idx].find(_ns("rPr"))
                if rpr is not None:
                    return copy.deepcopy(rpr)
                return None
        return None

    for text, is_match, text_pos in segments:
        rpr = _rpr_at(text_pos)
        if is_match:
            del_elem = etree.Element(_ns("del"))
            del_elem.set(_ns("id"), rev_id)
            del_elem.set(_ns("author"), author)
            del_elem.set(_ns("date"), datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
            r_del = etree.SubElement(del_elem, _ns("r"))
            dt = etree.SubElement(r_del, _ns("delText"))
            dt.text = text
            if rpr is not None:
                r_del.insert(0, copy.deepcopy(rpr))

            ins = etree.Element(_ns("ins"))
            ins.set(_ns("id"), rev_id)
            ins.set(_ns("author"), author)
            ins.set(_ns("date"), datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
            r_ins = etree.SubElement(ins, _ns("r"))
            _set_run_text(r_ins, new_text)
            if rpr is not None:
                r_ins.insert(0, copy.deepcopy(rpr))

            p.append(del_elem)
            p.append(ins)
            rev_id = str(int(rev_id) + 1)
        else:
            r = etree.SubElement(p, _ns("r"))
            _set_run_text(r, text)
            if rpr is not None:
                r.insert(0, copy.deepcopy(rpr))

    return match_count, rev_id


def replace_text(doc, old_text, new_text, author="Editor"):
    body = _find_body(doc)

    if not old_text:
        raise ValueError("old_text must not be empty")

    found = False
    for p in body.iterchildren(_ns("p")):
        if old_text in "".join(_run_text(r) for r in p.iterchildren(_ns("r"))):
            found = True
            break
    if not found:
        raise ValueError(f"Text {old_text!r} not found in document")

    next_id = _next_revision_id(body)
    for p in list(body.iterchildren(_ns("p"))):
        count, next_id = _replace_in_paragraph(p, old_text, new_text, author, next_id)

    return doc


def _find_target_paragraph(body, target_text):
    for p in body.iterchildren(_ns("p")):
        para_text = ""
        run_data = []
        for r in p.iterchildren(_ns("r")):
            t = r.find(_ns("t"))
            if t is not None and t.text:
                start = len(para_text)
                para_text += t.text
                run_data.append((r, start, len(para_text)))

        pos = para_text.find(target_text)
        if pos != -1:
            end_pos = pos + len(target_text)
            start_idx = None
            end_idx = None
            for idx, (r, rs, re) in enumerate(run_data):
                if rs <= pos < re or (rs < end_pos and rs >= pos):
                    if start_idx is None:
                        start_idx = idx
                    end_idx = idx
            return p, start_idx, end_idx, run_data

    raise ValueError(f"Text {target_text!r} not found in document")


def _inject_comments_part(docx_path, comments_xml_str):
    tmp = docx_path + ".tmp"
    found = False
    OPC_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == "[Content_Types].xml":
                    ct_root = etree.fromstring(data)
                    parts = [ov.get("PartName") for ov in ct_root.findall(f"{{{CT_NS}}}Override")]
                    if "/word/comments.xml" not in parts:
                        ov = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
                        ov.set("PartName", "/word/comments.xml")
                        ov.set("ContentType",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
                        data = etree.tostring(ct_root, xml_declaration=True,
                                              encoding="UTF-8", standalone=True)
                    zout.writestr(item, data)

                elif item.filename == "word/_rels/document.xml.rels":
                    rels_root = etree.fromstring(data)
                    rels = rels_root.findall(f"{{{OPC_REL_NS}}}Relationship")
                    existing = [r.get("Target") for r in rels]
                    if "comments.xml" not in existing:
                        rel = etree.SubElement(rels_root, f"{{{OPC_REL_NS}}}Relationship")
                        rel.set("Id", "rComments1")
                        rel.set("Type",
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
                        rel.set("Target", "comments.xml")
                        data = etree.tostring(rels_root, xml_declaration=True,
                                              encoding="UTF-8", standalone=True)
                    else:
                        found = True
                    zout.writestr(item, data)

                elif item.filename == "word/comments.xml":
                    found = True
                    existing_root = etree.fromstring(data)
                    new_root = etree.fromstring(comments_xml_str)
                    for c in new_root.iter(_ns("comment")):
                        existing_root.append(c)
                    data = etree.tostring(existing_root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                    zout.writestr(item, data)

                else:
                    zout.writestr(item, data)

            if not found:
                zout.writestr("word/comments.xml", comments_xml_str)

    shutil.move(tmp, docx_path)


def add_comment(doc, target_text, comment_text, author="Editor"):
    body = _find_body(doc)
    p, start_idx, end_idx, run_data = _find_target_paragraph(body, target_text)

    existing_ids = [0]
    for elem in body.iter():
        for attr in (_ns("id"),):
            val = elem.get(attr)
            if val is not None:
                try:
                    cid = int(val)
                    tag = elem.tag
                    if tag in (_ns("commentRangeStart"), _ns("commentRangeEnd"),
                               _ns("commentReference")):
                        existing_ids.append(cid)
                except ValueError:
                    pass

    comment_id = str(max(existing_ids) + 1)

    cs = etree.Element(_ns("commentRangeStart"))
    cs.set(_ns("id"), comment_id)

    ce = etree.Element(_ns("commentRangeEnd"))
    ce.set(_ns("id"), comment_id)

    cr_r = etree.Element(_ns("r"))
    cr = etree.SubElement(cr_r, _ns("commentReference"))
    cr.set(_ns("id"), comment_id)

    all_children = list(p)
    first_pos = None
    last_pos = None
    run_count = 0
    for ci, child in enumerate(all_children):
        if child.tag == _ns("r"):
            if run_count == start_idx:
                first_pos = ci
            if run_count == end_idx:
                last_pos = ci
            run_count += 1

    if first_pos is not None:
        p.insert(first_pos, cs)
        if last_pos is not None:
            if last_pos >= first_pos:
                last_pos += 1
            p.insert(last_pos + 1, ce)
            p.insert(last_pos + 2, cr_r)
        else:
            p.append(ce)
            p.append(cr_r)
    else:
        p.insert(0, cs)
        p.append(ce)
        p.append(cr_r)

    comment_elem = etree.Element(_ns("comment"))
    comment_elem.set(_ns("id"), comment_id)
    comment_elem.set(_ns("author"), author)
    comment_elem.set(_ns("date"),
                     datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
    cp = etree.SubElement(comment_elem, _ns("p"))
    cr2 = etree.SubElement(cp, _ns("r"))
    _set_run_text(cr2, comment_text)

    if not hasattr(doc.part, "_pending_comments"):
        doc.part._pending_comments = []
    doc.part._pending_comments.append(comment_elem)

    return doc


def finalize_comments(output_path, pending_comments=None):
    if pending_comments is None:
        pending_path = output_path + ".comments.json"
        if os.path.exists(pending_path):
            with open(pending_path) as f:
                pending_data = json.load(f)
            pending_comments = [etree.fromstring(c) for c in pending_data]
            os.remove(pending_path)

    if not pending_comments:
        return

    comments_root = etree.Element(_ns("comments"))
    for c in pending_comments:
        if isinstance(c, (str, bytes)):
            c = etree.fromstring(c)
        comments_root.append(c)

    comments_xml = etree.tostring(comments_root, xml_declaration=True,
                                  encoding="UTF-8", standalone=True)
    _inject_comments_part(output_path, comments_xml)


def list_revisions(doc):
    body = _find_body(doc)
    revisions = []
    rev_counter = 0

    for p in body.iterchildren(_ns("p")):
        full_para_text = _get_effective_text(p)

        for ins in p.iterchildren(_ns("ins")):
            rev_counter += 1
            ins_text = ""
            for r in ins.iterchildren(_ns("r")):
                ins_text += _run_text(r)
            revisions.append({
                "id": ins.get(_ns("id"), f"ins_{rev_counter}"),
                "type": "insertion",
                "author": ins.get(_ns("author"), ""),
                "date": ins.get(_ns("date"), ""),
                "text": ins_text,
                "paragraph_text": full_para_text,
            })

        for del_elem in p.iterchildren(_ns("del")):
            rev_counter += 1
            del_text = ""
            for r in del_elem.iterchildren(_ns("r")):
                dt = r.find(_ns("delText"))
                if dt is not None and dt.text:
                    del_text += dt.text
            revisions.append({
                "id": del_elem.get(_ns("id"), f"del_{rev_counter}"),
                "type": "deletion",
                "author": del_elem.get(_ns("author"), ""),
                "date": del_elem.get(_ns("date"), ""),
                "text": del_text,
                "paragraph_text": full_para_text,
            })

    return revisions


def restore_revision(doc, revision_id, restore_type="deletion"):
    body = _find_body(doc)
    found = False

    for p in body.iterchildren(_ns("p")):
        if restore_type == "deletion":
            for del_elem in list(p.iterchildren(_ns("del"))):
                if del_elem.get(_ns("id")) == revision_id:
                    found = True
                    idx = list(p).index(del_elem)
                    for r in list(del_elem.iterchildren(_ns("r"))):
                        for dt in r.iterchildren(_ns("delText")):
                            dt.tag = _ns("t")
                        p.insert(idx, r)
                        idx += 1
                    p.remove(del_elem)

        for ins in list(p.iterchildren(_ns("ins"))):
            if ins.get(_ns("id")) == revision_id:
                found = True
                p.remove(ins)

        if restore_type == "insertion":
            for del_elem in list(p.iterchildren(_ns("del"))):
                if del_elem.get(_ns("id")) == revision_id:
                    found = True
                    p.remove(del_elem)

    if not found:
        raise ValueError(f"Revision with id {revision_id!r} not found")

    return doc


def _get_effective_text(p_elem):
    parts = []
    for child in p_elem.iterchildren():
        if child.tag == _ns("r"):
            parts.append(_run_text(child))
        elif child.tag == _ns("ins"):
            for r in child.iterchildren(_ns("r")):
                parts.append(_run_text(r))
    return "".join(parts)


def get_full_text(doc):
    body = _find_body(doc)
    paras = []
    for p in body.iterchildren(_ns("p")):
        paras.append(_get_effective_text(p))
    return "\n".join(paras)
