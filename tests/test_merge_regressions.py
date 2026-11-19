import hashlib
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c63f8cfc0f01f00050001ff89993d1d0000000049454e44ae426082"
)


def paragraph(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def table(rows):
    joined_rows = "".join(
        "<w:tr>" + "".join(f"<w:tc>{paragraph(text)}</w:tc>" for text in row) + "</w:tr>"
        for row in rows
    )
    return f'<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>{joined_rows}</w:tbl>'


def write_docx(path, blocks, embed_png=False, declare_png=True):
    default_png = '<Default Extension="png" ContentType="image/png"/>' if embed_png and declare_png else ""
    with ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", f"""
            <Types xmlns="{CONTENT_TYPES}">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Override PartName="/word/document.xml"
                ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              {default_png}
            </Types>""")
        package.writestr("_rels/.rels", f"""
            <Relationships xmlns="{PACKAGE_REL}">
              <Relationship Id="rId1" Type="{OFFICE_REL}/officeDocument" Target="word/document.xml"/>
            </Relationships>""")
        package.writestr("word/document.xml", f"""
            <w:document xmlns:w="{WORD_NS}" xmlns:r="{OFFICE_REL}" xmlns:a="{DRAWING}"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
              <w:body>{''.join(blocks)}<w:sectPr/></w:body>
            </w:document>""")
        if embed_png:
            package.writestr("word/_rels/document.xml.rels", f"""
                <Relationships xmlns="{PACKAGE_REL}">
                  <Relationship Id="rId1" Type="{OFFICE_REL}/image" Target="media/image1.png"/>
                </Relationships>""")
            package.writestr("word/media/image1.png", PNG)


class MergeRegressions(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        self.copies = self.folder / "copies"
        self.copies.mkdir()
        self.base = self.folder / "base.docx"
        self.output = self.folder / "merged.docx"
        self.report = self.folder / "merged_conflicts.txt"
        self.first = paragraph("Alpha department completed 10 projects.")
        self.second = paragraph("Bravo department updated 20 records.")
        write_docx(self.base, [self.first, self.second])

    def run_merge(self, output=None):
        return subprocess.run(
            [sys.executable, "-X", "utf8", str(ROOT / "run_merge.py"), str(self.base),
             str(self.copies), "-o", str(output or self.output)],
            capture_output=True, encoding="utf-8", timeout=30,
        )

    def assert_success(self, result):
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def read_document(self):
        with ZipFile(self.output) as package:
            return etree.fromstring(package.read("word/document.xml"))

    def test_bad_copy_aborts_without_replacing_existing_output(self):
        self.output.write_bytes(b"previous output")
        bad_copy = self.copies / "broken.docx"
        bad_copy.write_bytes(b"not a Word file")
        result = self.run_merge()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(bad_copy.name, result.stdout)
        self.assertEqual(self.output.read_bytes(), b"previous output")
        self.assertFalse(self.report.exists())

    def test_output_cannot_overwrite_copy(self):
        copy = self.copies / "archive.docx"
        write_docx(copy, [self.first.replace("10", "11"), self.second])
        original = copy.read_bytes()
        result = self.run_merge(output=copy)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(copy.read_bytes(), original)

    def test_conflict_report_distinguishes_same_and_different_paragraphs(self):
        a = self.copies / "a.docx"
        b = self.copies / "b.docx"
        write_docx(a, [self.first.replace("10", "11"), self.second])
        write_docx(b, [self.first, self.second.replace("20", "21")])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 0", self.report.read_text(encoding="utf-8"))

        write_docx(b, [self.first.replace("10", "12"), self.second])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 1", self.report.read_text(encoding="utf-8"))
        result_text = "".join(self.read_document().itertext())
        self.assertIn("11 projects", result_text)
        self.assertIn("12 projects", result_text)

    def test_table_row_addition_and_removal(self):
        header = ["Team", "Count"]
        existing = ["Alpha", "10"]
        for rows in ([header, existing, ["Bravo", "20"]], [header], [header, existing, ["", ""]]):
            with self.subTest(rows=rows):
                write_docx(self.base, [self.first, table([header, existing])])
                write_docx(self.copies / "a.docx", [self.first, table(rows)])
                self.assert_success(self.run_merge())
                out = self.read_document()
                actual = [
                    ["".join(cell.itertext()).strip() for cell in row.findall(f"{{{WORD_NS}}}tc")]
                    for row in out.findall(f".//{{{WORD_NS}}}tr")
                ]
                self.assertEqual(actual, rows)

    def test_table_conflicts_keep_both_submissions(self):
        header = ["Team", "Count"]
        write_docx(self.base, [self.first, table([header, ["Alpha", "10"]])])
        write_docx(self.copies / "a.docx", [self.first, table([header, ["Alpha", "15"]])])
        write_docx(self.copies / "b.docx", [self.first, table([header, ["Alpha", "10"], ["Bravo", "20"]])])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 1", self.report.read_text(encoding="utf-8"))
        output_tables = ["".join(table_node.itertext()) for table_node in self.read_document().iter(f"{{{WORD_NS}}}tbl")]
        self.assertTrue(any("Alpha15" in text for text in output_tables))
        self.assertTrue(any("Bravo20" in text for text in output_tables))

    def test_delete_and_edit_keep_original_marked_and_edited_version(self):
        write_docx(self.copies / "a.docx", [self.second])
        write_docx(self.copies / "b.docx", [self.first.replace("10", "11"), self.second])
        self.assert_success(self.run_merge())
        self.assertIn("冲突数量: 1", self.report.read_text(encoding="utf-8"))
        output = self.read_document()
        self.assertTrue(output.findall(f".//{{{WORD_NS}}}strike"))
        text = "".join(output.itertext())
        self.assertIn("10 projects", text)
        self.assertIn("11 projects", text)

    def test_first_png_has_content_type_and_missing_type_aborts(self):
        image = (
            '<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData>'
            '<a:blip r:embed="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
        )
        copy = self.copies / "a.docx"
        write_docx(copy, [self.first, self.second, image], embed_png=True)
        self.assert_success(self.run_merge())
        with ZipFile(self.output) as package:
            types = etree.fromstring(package.read("[Content_Types].xml"))
            matches = [entry.get("PartName") for entry in types if entry.get("ContentType") == "image/png"]
            self.assertEqual(len(matches), 1)
            self.assertEqual(package.read(matches[0].lstrip("/")), PNG)

        old_output = hashlib.sha256(self.output.read_bytes()).hexdigest()
        write_docx(copy, [self.first, self.second, image], embed_png=True, declare_png=False)
        result = self.run_merge()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("资源缺少内容类型", result.stdout)
        self.assertEqual(hashlib.sha256(self.output.read_bytes()).hexdigest(), old_output)


if __name__ == "__main__":
    unittest.main()
