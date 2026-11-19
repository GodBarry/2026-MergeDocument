"""Merge Word review comments independently of changes to the paragraph text."""

from collections import defaultdict
from copy import deepcopy
from difflib import SequenceMatcher
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
W16CID = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ID = W + "id"
MARKERS = {W + name for name in ("commentRangeStart", "commentRangeEnd", "commentReference")}
TEXT = {W + "t", W + "delText", W + "instrText"}
SINGLE_CHAR = {W + name for name in ("tab", "br", "cr", "drawing", "object", "pict", "sym")}


def _characters(node):
    if node.tag in TEXT:
        return node.text or ""
    return "\ufffc" if node.tag in SINGLE_CHAR else ""


def _read_positions(body):
    paragraphs = list(body.iter(W + "p"))
    texts = []
    anchors = defaultdict(dict)
    for index, paragraph in enumerate(paragraphs):
        fragments = []
        offset = 0
        for node in paragraph.iter():
            if node.tag in MARKERS:
                anchors[node.get(ID)][node.tag] = (index, offset)
            fragment = _characters(node)
            if fragment:
                fragments.append(fragment)
                offset += len(fragment)
        texts.append("".join(fragments))
    return paragraphs, texts, anchors


def _map_offset(old, new, offset, left_boundary):
    if old == new:
        return offset
    # A range starts before its first selected character and ends immediately
    # after its last. Edits outside either edge must stay outside the range.
    character = offset if left_boundary else offset - 1
    if character < 0:
        return 0
    for op, i, j, a, b in SequenceMatcher(None, old, new, autojunk=False).get_opcodes():
        if i <= character < j:
            if op == "equal":
                return a + offset - i
            return a + round((offset - i) * (b - a) / (j - i))
    return len(new)


def _insert_marker(paragraph, offset, marker):
    """Place a marker at a Word run boundary, splitting only when the range requires it."""
    if marker.tag == W + "commentReference":
        reference_run = etree.Element(W + "r")
        reference_run.append(marker)
        marker = reference_run
    cursor = 0
    for run in paragraph.iter(W + "r"):
        for child in list(run):
            data = _characters(child)
            if not data:
                continue
            if offset < cursor + len(data):
                parent = run.getparent()
                cut = offset - cursor
                preceding = [item for item in list(run)[:run.index(child)] if item.tag != W + "rPr"]
                if preceding or cut:
                    left = etree.Element(W + "r", attrib=dict(run.attrib))
                    properties = run.find(W + "rPr")
                    if properties is not None:
                        left.append(deepcopy(properties))
                    for item in preceding:
                        left.append(item)
                    if cut:
                        text = deepcopy(child)
                        text.text, child.text = data[:cut], data[cut:]
                        text.set(SPACE, "preserve")
                        child.set(SPACE, "preserve")
                        left.append(text)
                    parent.insert(parent.index(run), left)
                parent.insert(parent.index(run), marker)
                return
            cursor += len(data)
    paragraph.append(marker)


def _part_write(path, root):
    etree.ElementTree(root).write(str(path), encoding="UTF-8", xml_declaration=True, standalone=True)


def _identity(comment):
    # Word can renumber comments and split runs on save.
    paragraphs = tuple("".join(t.text or "" for t in p.iter(W + "t"))
                       for p in comment.iter(W + "p"))
    return comment.get(W + "author"), comment.get(W + "date"), paragraphs


def _anchor_key(tags, texts, para_ids):
    return tuple(
        (tag, para_ids[p] or p, texts[p], offset)
        for tag, (p, offset) in sorted(tags.items())
    )


class WordComments:
    """Allocate one ID space for original, imported and generated comments."""

    def __init__(self, original, copies):
        self.next_id = 0
        self.by_copy = []
        self.sources = []
        self.originals = {}
        self.imports = []
        self.thread_entries = []
        self.used_para_ids = set()
        self.new_para_id = 1
        original_ids = defaultdict(list)
        by_durable, by_para, base_durables, original_anchors = {}, {}, {}, {}
        versions = defaultdict(dict)
        for index, path in enumerate([original, *(path for path, _ in copies)], start=-1):
            mapping = {}
            if index >= 0:
                self.by_copy.append(mapping)
            with ZipFile(path) as package:
                if "word/comments.xml" not in package.namelist():
                    continue
                comments = etree.fromstring(package.read("word/comments.xml"))
                source_body = etree.fromstring(package.read("word/document.xml")).find(W + "body")
                source_paras, texts, anchors = _read_positions(source_body)
                para_ids = [p.get(W14 + "paraId") for p in source_paras]
                self.sources.append((mapping, texts, para_ids, anchors))
                thread_ids = {}
                imported_ids = set()
                reused = set()
                durables = {}
                if "word/commentsIds.xml" in package.namelist():
                    stable_ids = etree.fromstring(package.read("word/commentsIds.xml"))
                    durables = {entry.get(W16CID + "paraId"): entry.get(W16CID + "durableId")
                                for entry in stable_ids.findall(W16CID + "commentId")}
                for comment in comments.findall(W + "comment"):
                    local = comment.get(ID)
                    identity = _identity(comment)
                    comment_paras = list(comment.iter(W + "p"))
                    last_para = comment_paras[-1].get(W14 + "paraId") if comment_paras else None
                    durable = durables.get(last_para)
                    anchor = _anchor_key(anchors.get(local, {}), texts, para_ids)
                    if index < 0:
                        self.originals[local] = comment
                        original_ids[identity].append(local)
                        original_anchors[local] = anchor
                        if last_para:
                            by_para[last_para] = local
                        if durable:
                            by_durable[durable] = local
                            base_durables[local] = durable
                        self.next_id = max(self.next_id, int(local) + 1)
                        mapping[local] = local
                        self.used_para_ids.update(p.get(W14 + "paraId") for p in comment.iter(W + "p"))
                        continue
                    base_id = by_durable.get(durable) or by_para.get(last_para)
                    if durable and base_id in base_durables and base_durables[base_id] != durable:
                        base_id = None
                    if base_id is None:
                        # Classic Word comments lack persistent IDs. A matching opinion
                        # at another location is a different comment, even with the same local ID.
                        candidates = [cid for cid in original_ids[identity] if cid not in reused
                                      and not (last_para and any(p.get(W14 + "paraId")
                                               for p in self.originals[cid].iter(W + "p")))
                                      and anchor and original_anchors[cid] == anchor]
                        base_id = local if local in candidates else None
                        if base_id is None and len(candidates) == 1:
                            base_id = candidates[0]
                    inherited = base_id is not None
                    replace_base = False
                    if inherited and _identity(self.originals[base_id]) != identity:
                        previous = versions[base_id]
                        if identity in previous:
                            mapping[local] = previous[identity]
                            continue
                        replace_base = not previous
                        inherited = replace_base
                    cid = base_id if inherited else str(self.next_id)
                    if base_id is not None and _identity(self.originals[base_id]) != identity:
                        versions[base_id][identity] = cid
                    mapping[local] = cid
                    if inherited:
                        reused.add(base_id)
                    else:
                        self.next_id += 1
                    if not inherited or replace_base:
                        comment.set(ID, cid)
                        self.imports.append((path, comment))
                    for p_index, paragraph in enumerate(comment_paras):
                        old = paragraph.get(W14 + "paraId")
                        if old is None:
                            continue
                        if inherited:
                            base_paras = list(self.originals[base_id].iter(W + "p"))
                            # Word identifies a thread by the final paragraph in its comment.
                            if p_index == len(comment_paras) - 1:
                                thread_ids[old] = base_paras[-1].get(W14 + "paraId", old)
                            elif p_index < len(base_paras) - 1:
                                thread_ids[old] = base_paras[p_index].get(W14 + "paraId", old)
                            if replace_base and old in thread_ids:
                                paragraph.set(W14 + "paraId", thread_ids[old])
                            continue
                        new = old
                        while new in self.used_para_ids:
                            new = f"{self.new_para_id:08X}"
                            self.new_para_id += 1
                        self.used_para_ids.add(new)
                        paragraph.set(W14 + "paraId", new)
                        thread_ids[old] = new
                        imported_ids.add(old)
                if imported_ids and "word/commentsExtended.xml" in package.namelist():
                    extended = etree.fromstring(package.read("word/commentsExtended.xml"))
                    for entry in extended.findall(W15 + "commentEx"):
                        old = entry.get(W15 + "paraId")
                        if old in imported_ids:
                            entry.set(W15 + "paraId", thread_ids[old])
                            parent = entry.get(W15 + "paraIdParent")
                            if parent is not None and parent in thread_ids:
                                entry.set(W15 + "paraIdParent", thread_ids[parent])
                            self.thread_entries.append(entry)

    def remap(self, element, copy_index):
        if not self.sources:
            return
        mapping = self.by_copy[copy_index]
        for node in element.iter(*MARKERS):
            local = node.get(ID)
            if local not in mapping:
                raise ValueError(f"副本批注引用缺少正文，编号 {local}")
            node.set(ID, mapping[local])

    def restore_ranges(self, body):
        """Keep copied anchors intact and transfer comment-only edits to the merged text."""
        if not self.sources:
            return
        placed = set()
        for node in list(body.iter(*MARKERS)):
            key = node.get(ID), node.tag
            if key in placed:
                node.getparent().remove(node)
            else:
                placed.add(key)
        paragraphs, target_texts, _ = _read_positions(body)
        by_para_id = defaultdict(list)
        for index, p in enumerate(paragraphs):
            by_para_id[p.get(W14 + "paraId")].append(index)
        for mapping, texts, para_ids, anchors in self.sources:
            missing = [(mapping[cid], tag, position) for cid, tags in anchors.items() if cid in mapping
                       for tag, position in tags.items() if (mapping[cid], tag) not in placed]
            if not missing:
                continue
            if not paragraphs:
                raise ValueError("合稿缺少可定位人工批注的正文段落")
            aligned = {}
            for _, i, j, a, b in SequenceMatcher(None, texts, target_texts, autojunk=False).get_opcodes():
                for source in range(i, j):
                    aligned[source] = min(a + min(source - i, max(0, b - a - 1)), len(paragraphs) - 1)
            for cid, tag, (index, offset) in missing:
                target = aligned[index]
                candidates = by_para_id.get(para_ids[index], []) if para_ids[index] else []
                if candidates:
                    target = min(candidates, key=lambda n: (target_texts[n] != texts[index], abs(n - target)))
                new_offset = _map_offset(texts[index], target_texts[target], offset, tag == W + "commentRangeStart")
                _insert_marker(paragraphs[target], new_offset, etree.Element(tag, {ID: cid}))
                placed.add((cid, tag))

    def write(self, extract_dir, clone_resources):
        if not self.imports:
            return
        directory = Path(extract_dir) / "word"
        path = directory / "comments.xml"
        root = etree.parse(str(path)).getroot() if path.exists() else etree.Element(W + "comments", nsmap={"w": W[1:-1]})
        existing = {c.get(ID): c for c in root.findall(W + "comment")}
        for source, comment in self.imports:
            clone_resources(comment, source, extract_dir, owner="comments.xml")
            old = existing.get(comment.get(ID))
            if old is None:
                root.append(comment)
            else:
                root.replace(old, comment)
        _part_write(path, root)
        if not self.thread_entries:
            return
        extended_path = directory / "commentsExtended.xml"
        if extended_path.exists():
            extended = etree.parse(str(extended_path)).getroot()
        else:
            extended = etree.Element(W15 + "commentsEx", nsmap={"w15": W15[1:-1]})
            rels_path = directory / "_rels/document.xml.rels"
            rels = etree.parse(str(rels_path)).getroot()
            used_ids = {rel.get("Id") for rel in rels}
            rid = "rIdCommentsExtended"
            while rid in used_ids:
                rid += "_"
            etree.SubElement(rels, REL + "Relationship", Id=rid,
                             Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                             Target="commentsExtended.xml")
            _part_write(rels_path, rels)
            types_path = Path(extract_dir) / "[Content_Types].xml"
            types = etree.parse(str(types_path)).getroot()
            etree.SubElement(types, CT + "Override", PartName="/word/commentsExtended.xml",
                             ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml")
            _part_write(types_path, types)
        extended.extend(self.thread_entries)
        _part_write(extended_path, extended)
