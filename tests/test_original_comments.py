"""Exercise review comments through the same merge command used by the desktop UI."""

from collections import Counter
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree

from test_merge_regressions import CONTENT_TYPES, OFFICE_REL, PACKAGE_REL, ROOT, WORD_NS, write_docx


W = f"{{{WORD_NS}}}"
REFERENCE = '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{}"/></w:r>'


def reviewed_paragraph(text, cid):
    run = f"<w:r><w:t>{text}</w:t></w:r>"
    if cid is not None:
        run = (f'<w:commentRangeStart w:id="{cid}"/>' + run + f'<w:commentRangeEnd w:id="{cid}"/>'
               + REFERENCE.format(cid))
    return "<w:p>" + run + "</w:p>"


def review_doc(path, blocks, comments):
    write_docx(path, blocks)
    if not comments:
        return
    with ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    types = etree.fromstring(parts["[Content_Types].xml"])
    etree.SubElement(types, f"{{{CONTENT_TYPES}}}Override", PartName="/word/comments.xml",
                     ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    parts["[Content_Types].xml"] = etree.tostring(types)
    relations = etree.Element(f"{{{PACKAGE_REL}}}Relationships")
    etree.SubElement(relations, f"{{{PACKAGE_REL}}}Relationship", Id="rId1", Type=OFFICE_REL + "/comments",
                     Target="comments.xml")
    parts["word/_rels/document.xml.rels"] = etree.tostring(relations)
    root = etree.Element(W + "comments", nsmap={"w": WORD_NS})
    for cid, author, text in comments:
        c = etree.SubElement(root, W + "comment", {W + "id": cid, W + "author": author})
        paragraph = etree.SubElement(c, W + "p")
        run = etree.SubElement(paragraph, W + "r")
        etree.SubElement(run, W + "t").text = text
    parts["word/comments.xml"] = etree.tostring(root)
    with ZipFile(path, "w") as output:
        for name, data in parts.items():
            output.writestr(name, data)


class OriginalCommentsTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.base = self.folder / "base.docx"
        self.copies = self.folder / "copies"
        self.copies.mkdir()

    def merged(self):
        output = self.folder / "merged.docx"
        cmd = [sys.executable, "-X", "utf8", "-m", "mddm.run_merge", str(self.base), str(self.copies), "-o", str(output)]
        result = subprocess.run(cmd, cwd=ROOT, capture_output=True, encoding="utf-8", timeout=30)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with ZipFile(output) as package:
            body = etree.fromstring(package.read("word/document.xml"))
            comments = etree.fromstring(package.read("word/comments.xml"))
        by_id = {c.get(W + "id"): (c.get(W + "author"), "".join(c.itertext())) for c in comments}
        self.assertEqual(len(by_id), len(comments))
        for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            ids = Counter(n.get(W + "id") for n in body.iter(W + tag))
            self.assertEqual(ids, Counter(by_id.keys()))
        return body, by_id

    def test_two_reviewers_on_unchanged_paragraph_keep_separate_overlapping_ranges(self):
        review_doc(self.base, [reviewed_paragraph("Please review July budget.", None)], [])
        for author, left, selected, right in (("Finance", "Please review ", "July", " budget."),
                                               ("Director", "Please ", "review July budget", ".")):
            block = reviewed_paragraph(selected, "0").replace(
                "<w:p>", f"<w:p><w:r><w:t>{left}</w:t></w:r>").replace(
                "</w:p>", f"<w:r><w:t>{right}</w:t></w:r></w:p>")
            review_doc(self.copies / (author + ".docx"), [block], [("0", author, selected)])
        body, comments = self.merged()
        self.assertEqual(len(comments), 2)
        self.assertEqual(len(list(body.iter(W + "p"))), 1)
        active = set()
        selected = {key: "" for key in comments}
        for n in body.iter():
            if n.tag == W + "commentRangeStart":
                active.add(n.get(W + "id"))
            elif n.tag == W + "commentRangeEnd":
                active.remove(n.get(W + "id"))
            elif n.tag == W + "t":
                for cid in active:
                    selected[cid] += n.text
        self.assertEqual(selected, {cid: text for cid, (_, text) in comments.items()})

    def test_original_review_is_kept_once_alongside_new_and_generated_comments(self):
        inherited = ("0", "QA", "Shared original opinion")
        first = reviewed_paragraph("First paragraph.", "0")
        review_doc(self.base, [first, reviewed_paragraph("Second paragraph.", None)], [inherited])
        for name in ("Finance", "Office"):
            changed = name == "Finance"
            inherited_id = "0" if changed else "7"
            review_doc(self.copies / (name + ".docx"), [
                reviewed_paragraph("First paragraph.", inherited_id),
                reviewed_paragraph("Second paragraph changed." if changed else "Second paragraph.", "1")
            ], [(inherited_id, *inherited[1:]), ("1", name, name + " opinion")])
        body, comments = self.merged()
        notes = list(comments.values())
        self.assertEqual(notes.count(inherited[1:]), 1)
        self.assertIn(("Finance", "Finance opinion"), notes)
        self.assertIn(("Office", "Office opinion"), notes)
        self.assertTrue(any("修改" in note for _, note in notes))
        self.assertIn("Second paragraph changed.", "".join(body.itertext()))
